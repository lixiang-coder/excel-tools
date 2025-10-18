package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.*;

/**
 * 流式应用映射处理器 - 解决百万数据OOM问题
 */
public class StreamingMappingApplier {
    private static final Logger logger = LogManager.getLogger(StreamingMappingApplier.class);
    private final MemoryConfigManager memoryConfig;

    public StreamingMappingApplier() {
        this.memoryConfig = MemoryConfigManager.getInstance();
    }

    /**
     * 流式应用映射关系
     */
    public void applyMapping(File inputFile, String lookupField, String fillField,
                           Map<String, String> mappingRelations, File outputFile,
                           ProgressCallback callback) throws Exception {

        logger.info("开始流式应用映射: {}", inputFile.getName());
        logger.info("映射关系数量: {}", mappingRelations.size());

        // 创建输出工作簿
        int windowSize = memoryConfig.getMemoryWindowSize();
        logger.info("使用内存窗口大小: {}", windowSize);

        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(windowSize)) {
            if (memoryConfig.isEnableCompression()) {
                outputWorkbook.setCompressTempFiles(true);
            }

            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

            // 使用SAX流式读取并应用映射
            processInputFileWithMapping(inputFile, outputSheet, lookupField, fillField,
                                       mappingRelations, callback);

            // 保存输出文件
            logger.info("保存输出文件: {}", outputFile.getName());
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }

        logger.info("流式映射应用完成");
    }

    /**
     * 使用SAX流式处理输入文件并应用映射
     */
    private void processInputFileWithMapping(File inputFile, Sheet outputSheet,
                                            String lookupField, String fillField,
                                            Map<String, String> mappingRelations,
                                            ProgressCallback callback) throws Exception {

        try (FileInputStream fis = new FileInputStream(inputFile);
             OPCPackage pkg = OPCPackage.open(fis)) {

            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            // 获取第一个工作表
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
            if (iter.hasNext()) {
                InputStream sheetStream = iter.next();

                // 创建处理器
                MappingSheetHandler handler = new MappingSheetHandler(
                    outputSheet, lookupField, fillField, mappingRelations, callback, memoryConfig
                );

                // 解析工作表
                processSheet(styles, sst, handler, sheetStream);

                sheetStream.close();

                logger.info("处理完成，更新了 {} 条记录", handler.getUpdatedCount());
            }
        }
    }

    /**
     * 解析单个工作表
     */
    private void processSheet(StylesTable styles, SharedStringsTable sst,
                            MappingSheetHandler handler, InputStream sheetInputStream) throws Exception {

        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);

        try {
            XMLReader sheetParser = XMLReaderFactory.createXMLReader();
            ContentHandler contentHandler = new XSSFSheetXMLHandler(
                styles, sst, handler, formatter, false
            );
            sheetParser.setContentHandler(contentHandler);
            sheetParser.parse(sheetSource);
        } catch (Exception e) {
            logger.error("解析工作表失败", e);
            throw e;
        }
    }

    /**
     * SAX事件处理器 - 应用映射
     */
    private static class MappingSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final Sheet outputSheet;
        private final String lookupField;
        private final String fillField;
        private final Map<String, String> mappingRelations;
        private final ProgressCallback callback;
        private final MemoryConfigManager memoryConfig;

        private Map<String, Integer> headerIndexMap;
        private Integer lookupIndex;
        private Integer fillIndex;
        private List<String> headerNames;

        private int currentRowNum = 0;
        private int outputRowNum = 0;
        private Map<Integer, String> currentRowData;
        private int updatedCount = 0;
        private int lastGCCheck = 0;

        public MappingSheetHandler(Sheet outputSheet, String lookupField, String fillField,
                                  Map<String, String> mappingRelations,
                                  ProgressCallback callback, MemoryConfigManager memoryConfig) {
            this.outputSheet = outputSheet;
            this.lookupField = lookupField;
            this.fillField = fillField;
            this.mappingRelations = mappingRelations;
            this.callback = callback;
            this.memoryConfig = memoryConfig;
            this.currentRowData = new HashMap<>();
            this.headerNames = new ArrayList<>();
        }

        @Override
        public void startRow(int rowNum) {
            currentRowNum = rowNum;
            currentRowData.clear();
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum == 0) {
                // 处理表头行
                processHeaderRow();
            } else {
                // 处理数据行
                processDataRow();

                // 更新进度
                if (callback != null && outputRowNum % 100 == 0) {
                    callback.updateProgress(outputRowNum, 0);  // 总数未知，传0
                }

                // 定期内存检查和GC
                if (outputRowNum - lastGCCheck >= 1000) {
                    if (memoryConfig.checkMemoryWarning()) {
                        memoryConfig.forceGarbageCollection();
                    }
                    lastGCCheck = outputRowNum;

                    // 清理SXSSF缓存
                    try {
                        if (outputSheet instanceof org.apache.poi.xssf.streaming.SXSSFSheet) {
                            ((org.apache.poi.xssf.streaming.SXSSFSheet) outputSheet).flushRows();
                        }
                    } catch (IOException e) {
                        // 忽略刷新错误
                    }
                }
            }
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                        org.apache.poi.xssf.usermodel.XSSFComment comment) {
            // 提取列索引
            int colIndex = getColumnIndex(cellReference);
            if (formattedValue != null) {
                currentRowData.put(colIndex, formattedValue.trim());
            }
        }

        /**
         * 处理表头行
         */
        private void processHeaderRow() {
            headerIndexMap = new HashMap<>();

            // 构建表头索引映射
            for (Map.Entry<Integer, String> entry : currentRowData.entrySet()) {
                String headerName = entry.getValue();
                headerIndexMap.put(headerName, entry.getKey());
                headerNames.add(headerName);
            }

            // 获取查找字段和填充字段的索引
            lookupIndex = headerIndexMap.get(lookupField);
            fillIndex = headerIndexMap.get(fillField);

            if (lookupIndex == null) {
                throw new RuntimeException("找不到查找字段: " + lookupField);
            }
            if (fillIndex == null) {
                throw new RuntimeException("找不到填充字段: " + fillField);
            }

            // 写入表头到输出
            Row headerRow = outputSheet.createRow(outputRowNum++);
            for (Map.Entry<Integer, String> entry : currentRowData.entrySet()) {
                Cell cell = headerRow.createCell(entry.getKey());
                cell.setCellValue(entry.getValue());
            }
        }

        /**
         * 处理数据行
         */
        private void processDataRow() {
            Row outputRow = outputSheet.createRow(outputRowNum++);

            // 复制所有单元格数据
            for (Map.Entry<Integer, String> entry : currentRowData.entrySet()) {
                Cell cell = outputRow.createCell(entry.getKey());
                cell.setCellValue(entry.getValue());
            }

            // 获取查找字段的值
            String lookupValue = currentRowData.get(lookupIndex);
            if (lookupValue != null && !lookupValue.isEmpty()) {
                // 查找映射值
                String mappedValue = null;

                // 优先用key查找
                if (mappingRelations.containsKey(lookupValue)) {
                    mappedValue = mappingRelations.get(lookupValue);
                } else {
                    // 否则遍历value查找key
                    for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                        if (entry.getValue().equals(lookupValue)) {
                            mappedValue = entry.getKey();
                            break;
                        }
                    }
                }

                // 填充映射值
                Cell fillCell = outputRow.getCell(fillIndex);
                if (fillCell == null) {
                    fillCell = outputRow.createCell(fillIndex);
                }

                if (mappedValue != null) {
                    fillCell.setCellValue(mappedValue);
                    updatedCount++;
                } else {
                    // 如果找不到映射，保持原值
                    fillCell.setCellValue(lookupValue);
                }
            }
        }

        /**
         * 从单元格引用中提取列索引
         */
        private int getColumnIndex(String cellReference) {
            int col = 0;
            for (int i = 0; i < cellReference.length(); i++) {
                char c = cellReference.charAt(i);
                if (Character.isDigit(c)) {
                    break;
                }
                col = col * 26 + (c - 'A' + 1);
            }
            return col - 1;
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            // 不处理页眉页脚
        }

        public int getUpdatedCount() {
            return updatedCount;
        }
    }

    /**
     * 进度回调接口
     */
    public interface ProgressCallback {
        void updateProgress(int current, int total);
    }
}
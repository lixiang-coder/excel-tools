package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.*;

/**
 * 流式Excel处理器 - 解决大文件OOM问题
 * 使用SAX模式读取，SXSSF模式写入
 */
public class StreamingExcelProcessor {
    private static final Logger logger = LogManager.getLogger(StreamingExcelProcessor.class);
    private final MemoryConfigManager memoryConfig;

    public StreamingExcelProcessor() {
        this.memoryConfig = MemoryConfigManager.getInstance();
    }

    /**
     * 流式处理Excel文件
     */
    public void processExcel(File inputFile, File templateFile, File outputFile,
                           Map<String, String> mappings, ProgressCallback callback) throws Exception {

        logger.info("开始流式处理Excel文件: {}", inputFile.getName());

        // 1. 读取模板文件的表头（模板文件通常很小，可以直接读取）
        List<String> templateHeaders = readTemplateHeaders(templateFile);
        logger.info("模板表头数量: {}", templateHeaders.size());

        // 2. 转换映射关系
        Map<String, List<String>> targetToSourceMap = new HashMap<>();
        for (Map.Entry<String, String> entry : mappings.entrySet()) {
            targetToSourceMap.computeIfAbsent(entry.getValue(), k -> new ArrayList<>()).add(entry.getKey());
        }

        // 3. 创建输出工作簿
        int windowSize = memoryConfig.getMemoryWindowSize();
        logger.info("使用内存窗口大小: {}", windowSize);

        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(windowSize)) {
            if (memoryConfig.isEnableCompression()) {
                outputWorkbook.setCompressTempFiles(true);
            }

            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

            // 4. 写入表头
            Row headerRow = outputSheet.createRow(0);
            for (int i = 0; i < templateHeaders.size(); i++) {
                headerRow.createCell(i).setCellValue(templateHeaders.get(i));
            }

            // 5. 使用SAX流式读取源文件
            processInputFileStreaming(inputFile, outputSheet, templateHeaders,
                                    targetToSourceMap, callback);

            // 6. 保存输出文件
            logger.info("保存输出文件: {}", outputFile.getName());
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }

        logger.info("流式处理完成");
    }

    /**
     * 读取模板表头
     */
    private List<String> readTemplateHeaders(File templateFile) throws IOException {
        List<String> headers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(templateFile);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        headers.add(cell.getStringCellValue());
                    }
                }
            }
        }
        return headers;
    }

    /**
     * 使用SAX流式处理输入文件
     */
    private void processInputFileStreaming(File inputFile, Sheet outputSheet,
                                         List<String> templateHeaders,
                                         Map<String, List<String>> targetToSourceMap,
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
                StreamingSheetHandler handler = new StreamingSheetHandler(
                    outputSheet, templateHeaders, targetToSourceMap, callback, memoryConfig
                );

                // 解析工作表
                processSheet(styles, sst, handler, sheetStream);

                sheetStream.close();
            }
        }
    }

    /**
     * 解析单个工作表
     */
    private void processSheet(StylesTable styles, SharedStringsTable sst,
                            StreamingSheetHandler handler, InputStream sheetInputStream) throws Exception {

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
     * SAX事件处理器
     */
    private static class StreamingSheetHandler implements SheetContentsHandler {
        private final Sheet outputSheet;
        private final List<String> templateHeaders;
        private final Map<String, List<String>> targetToSourceMap;
        private final ProgressCallback callback;
        private final MemoryConfigManager memoryConfig;

        private Map<String, Integer> sourceHeaderIndexMap;
        private int currentRowNum = 0;
        private int outputRowNum = 1;
        private Map<Integer, String> currentRowData;
        private int totalRows = 0;
        private int processedRows = 0;
        private int lastGCCheck = 0;

        public StreamingSheetHandler(Sheet outputSheet, List<String> templateHeaders,
                                    Map<String, List<String>> targetToSourceMap,
                                    ProgressCallback callback, MemoryConfigManager memoryConfig) {
            this.outputSheet = outputSheet;
            this.templateHeaders = templateHeaders;
            this.targetToSourceMap = targetToSourceMap;
            this.callback = callback;
            this.memoryConfig = memoryConfig;
            this.currentRowData = new HashMap<>();
        }

        @Override
        public void startRow(int rowNum) {
            currentRowNum = rowNum;
            currentRowData.clear();
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum == 0) {
                // 处理表头行，建立索引映射
                sourceHeaderIndexMap = new HashMap<>();
                for (Map.Entry<Integer, String> entry : currentRowData.entrySet()) {
                    sourceHeaderIndexMap.put(entry.getValue(), entry.getKey());
                }
            } else {
                // 处理数据行
                processDataRow();
                processedRows++;

                // 更新进度
                if (callback != null && processedRows % 100 == 0) {
                    callback.updateProgress(processedRows, totalRows > 0 ? totalRows : processedRows);
                }

                // 定期内存检查和GC
                if (processedRows - lastGCCheck >= 1000) {
                    if (memoryConfig.checkMemoryWarning()) {
                        memoryConfig.forceGarbageCollection();
                    }
                    lastGCCheck = processedRows;

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
         * 处理数据行
         */
        private void processDataRow() {
            Row outputRow = outputSheet.createRow(outputRowNum++);

            // 遍历每个目标表头
            for (int i = 0; i < templateHeaders.size(); i++) {
                String targetHeader = templateHeaders.get(i);
                List<String> sourceHeaders = targetToSourceMap.get(targetHeader);

                if (sourceHeaders != null && !sourceHeaders.isEmpty()) {
                    if (sourceHeaders.size() == 1) {
                        // 单个映射
                        String sourceHeader = sourceHeaders.get(0);
                        Integer sourceIndex = sourceHeaderIndexMap.get(sourceHeader);
                        if (sourceIndex != null) {
                            String value = currentRowData.get(sourceIndex);
                            if (value != null && !value.isEmpty()) {
                                outputRow.createCell(i).setCellValue(value);
                            }
                        }
                    } else {
                        // 多个映射，需要合并
                        StringBuilder mergedValue = new StringBuilder();
                        for (String sourceHeader : sourceHeaders) {
                            Integer sourceIndex = sourceHeaderIndexMap.get(sourceHeader);
                            if (sourceIndex != null) {
                                String value = currentRowData.get(sourceIndex);
                                if (value != null && !value.isEmpty()) {
                                    if (mergedValue.length() > 0) {
                                        mergedValue.append("|");
                                    }
                                    mergedValue.append(sourceHeader).append("：").append(value);
                                }
                            }
                        }
                        if (mergedValue.length() > 0) {
                            outputRow.createCell(i).setCellValue(mergedValue.toString());
                        }
                    }
                }
            }
        }

        /**
         * 从单元格引用中提取列索引 (如 "A1" -> 0, "B1" -> 1)
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
    }

    /**
     * 进度回调接口
     */
    public interface ProgressCallback {
        void updateProgress(int current, int total);
    }
}
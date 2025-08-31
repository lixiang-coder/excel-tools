package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 分块处理映射关系，解决大量映射关系导致的内存问题
 */
public class ChunkedMappingProcessor {
    private static final Logger logger = LogManager.getLogger(ChunkedMappingProcessor.class);
    
    private final MemoryConfigManager memoryConfig;
    private final Path tempDir;
    
    // 使用临时文件存储映射关系
    private final Map<String, Path> mappingChunkFiles = new ConcurrentHashMap<>();
    private int chunkCounter = 0;
    
    public ChunkedMappingProcessor() throws IOException {
        this.memoryConfig = MemoryConfigManager.getInstance();
        this.tempDir = Files.createTempDirectory("excel_mapping_");
        
        // 程序退出时清理临时文件
        Runtime.getRuntime().addShutdownHook(new Thread(this::cleanup));
    }
    
    /**
     * 分块保存映射关系到临时文件
     */
    public void saveMappingInChunks(File sourceFile, String sourceField, String targetField,
                                   ProgressCallback callback) throws IOException {
        
        int chunkSize = memoryConfig.getBatchSize();
        logger.info("开始分块处理映射关系，块大小: {}", chunkSize);
        
        try (FileInputStream fis = new FileInputStream(sourceFile);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            
            // 获取表头索引
            Map<String, Integer> headerMap = getHeaderIndexMap(sheet);
            Integer sourceIndex = headerMap.get(sourceField);
            Integer targetIndex = headerMap.get(targetField);
            
            if (sourceIndex == null || targetIndex == null) {
                throw new IOException("找不到指定的字段: " + sourceField + " 或 " + targetField);
            }
            
            int totalRows = sheet.getLastRowNum();
            int processedRows = 0;
            Map<String, String> currentChunk = new HashMap<>();
            
            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                
                Cell sourceCell = row.getCell(sourceIndex);
                Cell targetCell = row.getCell(targetIndex);
                
                if (sourceCell != null && targetCell != null) {
                    String sourceValue = getCellValueAsString(sourceCell);
                    String targetValue = getCellValueAsString(targetCell);
                    
                    if (!sourceValue.isEmpty() && !targetValue.isEmpty()) {
                        currentChunk.put(sourceValue, targetValue);
                    }
                }
                
                processedRows++;
                
                // 当块达到指定大小或处理完所有行时，保存到文件
                if (currentChunk.size() >= chunkSize || i == totalRows) {
                    if (!currentChunk.isEmpty()) {
                        saveChunkToFile(currentChunk);
                        currentChunk.clear();
                        
                        // 强制垃圾回收
                        if (memoryConfig.checkMemoryWarning()) {
                            memoryConfig.forceGarbageCollection();
                        }
                    }
                }
                
                // 更新进度
                if (callback != null && processedRows % 1000 == 0) {
                    callback.updateProgress(processedRows, totalRows);
                }
            }
            
            if (callback != null) {
                callback.updateProgress(totalRows, totalRows);
            }
            
            logger.info("映射关系处理完成，共处理{}行，生成{}个文件块", processedRows, mappingChunkFiles.size());
        }
    }
    
    /**
     * 应用映射关系到目标文件
     */
    public void applyMappingFromChunks(File targetFile, String lookupField, String fillField,
                                     File outputFile, ProgressCallback callback) throws IOException {
        
        logger.info("开始应用分块映射关系");
        
        try (FileInputStream fis = new FileInputStream(targetFile);
             Workbook inputWorkbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {
            
            Sheet inputSheet = inputWorkbook.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");
            
            // 获取表头索引
            Map<String, Integer> headerMap = getHeaderIndexMap(inputSheet);
            Integer lookupIndex = headerMap.get(lookupField);
            Integer fillIndex = headerMap.get(fillField);
            
            if (lookupIndex == null || fillIndex == null) {
                throw new IOException("找不到指定的字段: " + lookupField + " 或 " + fillField);
            }
            
            // 复制表头
            copyRow(inputSheet.getRow(0), outputSheet.createRow(0));
            
            int totalRows = inputSheet.getLastRowNum();
            int processedRows = 0;
            int updatedCount = 0;
            
            for (int i = 1; i <= totalRows; i++) {
                Row inputRow = inputSheet.getRow(i);
                if (inputRow == null) continue;
                
                Row outputRow = outputSheet.createRow(i);
                copyRow(inputRow, outputRow);
                
                // 查找映射值
                Cell lookupCell = inputRow.getCell(lookupIndex);
                if (lookupCell != null) {
                    String lookupValue = getCellValueAsString(lookupCell);
                    String mappedValue = findMappingValue(lookupValue);
                    
                    if (mappedValue != null) {
                        Cell fillCell = outputRow.createCell(fillIndex);
                        fillCell.setCellValue(mappedValue);
                        updatedCount++;
                    }
                }
                
                processedRows++;
                
                // 更新进度和内存检查
                if (processedRows % 1000 == 0) {
                    if (callback != null) {
                        callback.updateProgress(processedRows, totalRows);
                    }
                    if (memoryConfig.checkMemoryWarning()) {
                        memoryConfig.forceGarbageCollection();
                    }
                }
            }
            
            // 保存输出文件
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
            
            if (callback != null) {
                callback.updateProgress(totalRows, totalRows);
            }
            
            logger.info("映射应用完成，共处理{}行，更新{}条记录", processedRows, updatedCount);
        }
    }
    
    /**
     * 从所有块文件中查找映射值
     */
    private String findMappingValue(String key) throws IOException {
        for (Path chunkFile : mappingChunkFiles.values()) {
            try (BufferedReader reader = Files.newBufferedReader(chunkFile)) {
                String line;
                while ((line = reader.readLine()) != null) {
                    String[] parts = line.split("\t", 2);
                    if (parts.length == 2 && parts[0].equals(key)) {
                        return parts[1];
                    }
                }
            }
        }
        return null;
    }
    
    /**
     * 保存当前块到临时文件
     */
    private void saveChunkToFile(Map<String, String> chunk) throws IOException {
        String chunkFileName = "mapping_chunk_" + (++chunkCounter) + ".txt";
        Path chunkFile = tempDir.resolve(chunkFileName);
        
        try (BufferedWriter writer = Files.newBufferedWriter(chunkFile)) {
            for (Map.Entry<String, String> entry : chunk.entrySet()) {
                writer.write(entry.getKey() + "\t" + entry.getValue());
                writer.newLine();
            }
        }
        
        mappingChunkFiles.put(chunkFileName, chunkFile);
        logger.debug("保存映射块到文件: {}, 大小: {}", chunkFileName, chunk.size());
    }
    
    /**
     * 获取表头索引映射
     */
    private Map<String, Integer> getHeaderIndexMap(Sheet sheet) {
        Map<String, Integer> headerMap = new HashMap<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    headerMap.put(cell.getStringCellValue().trim(), i);
                }
            }
        }
        return headerMap;
    }
    
    /**
     * 复制行数据
     */
    private void copyRow(Row sourceRow, Row targetRow) {
        if (sourceRow == null) return;
        
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(i);
                copyCellValue(sourceCell, targetCell);
            }
        }
    }
    
    /**
     * 复制单元格值
     */
    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }
    
    /**
     * 获取单元格值为字符串
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double value = cell.getNumericCellValue();
                return value == Math.floor(value) ? 
                    String.format("%.0f", value) : String.valueOf(value);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
    
    /**
     * 清理临时文件
     */
    public void cleanup() {
        try {
            if (Files.exists(tempDir)) {
                Files.walk(tempDir)
                     .sorted(Comparator.reverseOrder())
                     .map(Path::toFile)
                     .forEach(File::delete);
                logger.info("清理临时文件完成");
            }
        } catch (IOException e) {
            logger.error("清理临时文件失败", e);
        }
    }
    
    /**
     * 进度回调接口
     */
    public interface ProgressCallback {
        void updateProgress(int current, int total);
    }
}
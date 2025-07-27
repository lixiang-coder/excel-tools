package com.tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * 将多个Excel文件合并到一个新的Excel文件中
 */
public class ExcelMerger {

    private static final Logger log = Logger.getLogger(ExcelMerger.class.getName());

    public static void main(String[] args) {
        // 直接硬编码路径
        String folderPath = "D:/edgedownload/temp/数据迁移";

        File folder = new File(folderPath);
        if (!folder.exists() || !folder.isDirectory()) {
            log.log(Level.SEVERE, "无效的文件夹路径！");
            return;
        }

        // 输出文件
        String outputFilePath = folderPath + File.separator + "combined_output.xlsx";
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet combinedSheet = outputWorkbook.createSheet("CombinedData");

            DataFormatter formatter = new DataFormatter(); // 使用 DataFormatter 格式化数字
            int globalRowCount = 0; // 记录全局行号

            for (File file : Objects.requireNonNull(folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx")))) {
                processFile(combinedSheet, globalRowCount, file, formatter);

                // 更新全局行号
                try (Workbook inputWorkbook = WorkbookFactory.create(file)) {
                    Sheet sheet = inputWorkbook.getSheetAt(0);
                    globalRowCount += sheet.getLastRowNum() + 1; // 计算当前文件的总行数并更新全局行号
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
            }

            log.log(Level.INFO, "合并完成！输出文件路径：" + outputFilePath);

        } catch (IOException e) {
            log.log(Level.SEVERE, "写出文件时发生错误", e.getMessage());
            throw new RuntimeException("处理Excel文件失败: " + e.getMessage());
        }
    }

    private static void processFile(Sheet combinedSheet, int startRow, File file, DataFormatter formatter) throws IOException {
        try (Workbook inputWorkbook = WorkbookFactory.create(file)) {
            Sheet inputSheet = inputWorkbook.getSheetAt(0); // 只处理第一个工作表

            for (int i = 0; i <= inputSheet.getLastRowNum(); i++) { // 遍历每一行
                Row row = inputSheet.getRow(i);
                Cell cell = row.getCell(0); // 第一列
                String value = "";
                if (cell != null) {
                    value = formatter.formatCellValue(cell); // 使用 DataFormatter 格式化单元格值
                }

                // 将数据写入到新的工作表中
                Row outputRow = combinedSheet.createRow(startRow++);
                outputRow.createCell(0).setCellValue(value);
            }
        }
    }
}
package com.xzy.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Excel工具类：公共方法提取
 */
public class ExcelUtils {
    private static final Logger logger = LogManager.getLogger(ExcelUtils.class);
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final DateTimeFormatter DATE_TIME_MINUTE_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm");
    private static final DateTimeFormatter DATE_TIME_SECOND_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");

    /**
     * 将单元格内容转为字符串
     */
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatExcelDate(cell.getLocalDateTimeCellValue());
                }
                return numericText(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return formatExcelDate(cell.getLocalDateTimeCellValue());
                            }
                            return numericText(cell.getNumericCellValue());
                        case BOOLEAN:
                            return String.valueOf(cell.getBooleanCellValue());
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    logger.warn("获取公式计算结果时出错: " + e.getMessage());
                    return "";
                }
            default:
                return "";
        }
    }

    static String numericText(double value) {
        return value == Math.floor(value) ? String.format(Locale.ROOT, "%.0f", value) : String.valueOf(value);
    }

    /**
     * 将Excel日期序列值转换为稳定、无时区信息的字符串。
     * 纯日期不补零点时间；存在时间时按实际精度保留到分钟或秒。
     */
    public static String formatExcelDate(double excelDate) {
        return formatExcelDate(excelDate, false);
    }

    public static String formatExcelDate(double excelDate, boolean use1904Windowing) {
        if (!DateUtil.isValidExcelDate(excelDate)) {
            return String.valueOf(excelDate);
        }
        // roundSeconds=true，避免Excel浮点误差将整分钟解析为上一秒的末尾。
        return formatExcelDate(DateUtil.getLocalDateTime(excelDate, use1904Windowing, true));
    }

    private static String formatExcelDate(LocalDateTime dateTime) {
        if (dateTime.getHour() == 0 && dateTime.getMinute() == 0 && dateTime.getSecond() == 0) {
            return DATE_FORMATTER.format(dateTime);
        }
        if (dateTime.getSecond() == 0) {
            return DATE_TIME_MINUTE_FORMATTER.format(dateTime);
        }
        return DATE_TIME_SECOND_FORMATTER.format(dateTime);
    }

    /** 保留日期值及显示格式，同一输出工作簿按格式串复用样式。 */
    public static void setDateCellValue(Cell target, LocalDateTime value, String format,
                                       Map<String, CellStyle> dateStyles) {
        target.setCellValue(value);
        CellStyle style = dateStyles.computeIfAbsent(format, key -> {
            CellStyle created = target.getSheet().getWorkbook().createCellStyle();
            created.setDataFormat(target.getSheet().getWorkbook().createDataFormat().getFormat(key));
            return created;
        });
        target.setCellStyle(style);
    }

    /**
     * 复制单元格值（保留类型）
     */
    public static void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            case ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
        }
    }

    /** 界面按名称选择字段，不把空列当作可选项。 */
    public static List<String> readNamedHeaders(File file) throws IOException {
        List<String> names = new ArrayList<>(readHeadersOnly(file));
        names.removeIf(name -> name.trim().isEmpty());
        return names;
    }

    /** 返回按物理列号排列的表头，空列使用空字符串占位。 */
    public static List<String> readHeadersOnly(File file) throws IOException {
        if (file.getName().toLowerCase().endsWith(".xlsx")) {
            return XlsxSheetReader.readHeaders(file);
        }
        return readXLSHeaders(file);
    }

    /**
     * 读取XLS文件的表头
     */
    private static List<String> readXLSHeaders(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            return readHeaderRow(sheet.getRow(0));
        }
    }

    public static List<String> readHeaderRow(Row row) {
        List<String> headers = new ArrayList<>();
        if (row != null) {
            for (int i = 0; i < row.getLastCellNum(); i++) {
                headers.add(getCellValueAsString(row.getCell(i)));
            }
        }
        return headers;
    }

    /**
     * 去除所有不可见字符（空格、制表符、换行、BOM）
     */
    public static String cleanString(String s) {
        if (s == null) return null;
        return s.replaceAll("\\s+", "").replace("﻿", "");
    }

}

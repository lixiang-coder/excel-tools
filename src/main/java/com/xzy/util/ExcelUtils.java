package com.xzy.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeMap;

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

    public static String getCheckedCellText(Cell cell) throws FormulaCache.FormulaCacheException {
        if (cell != null && cell.getCellType() == CellType.FORMULA) {
            FormulaCache.read(cell);
        }
        return getCellValueAsString(cell);
    }

    public static String getCheckedCellText(File file, Cell cell) throws IOException {
        try {
            return getCheckedCellText(cell);
        } catch (FormulaCache.FormulaCacheException e) {
            throw e.inFile(file);
        }
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
        return readNamedHeaders(file, Comparator.naturalOrder());
    }

    public static List<String> readNamedHeaders(File file, Comparator<String> comparison) throws IOException {
        List<String> headers = readHeadersOnly(file);
        Map<String, Integer> indices = new TreeMap<>(comparison);
        List<String> names = new ArrayList<>();
        for (int column = 0; column < headers.size(); column++) {
            String name = headers.get(column);
            if (name.trim().isEmpty() || comparison.compare(name, "") == 0) continue;
            putUniqueHeader(indices, name, column, file.getAbsolutePath());
            names.add(name);
        }
        return names;
    }

    public static Map<String, Integer> indexNamedHeaders(List<String> headers, String location) throws IOException {
        Map<String, Integer> indices = new LinkedHashMap<>();
        for (int column = 0; column < headers.size(); column++) {
            putUniqueHeader(indices, headers.get(column), column, location);
        }
        return indices;
    }

    /** 调用方按实际匹配规则规范化字段名；保存从零开始的物理列号。 */
    public static void putUniqueHeader(Map<String, Integer> indices, String key, int column, String location)
            throws IOException {
        if (key == null || key.trim().isEmpty()) return;
        Integer previous = indices.putIfAbsent(key, column);
        if (previous != null) {
            throw new IOException(location + "：重复表头“" + key + "”位于 "
                    + CellReference.convertNumToColString(previous) + " 列和 "
                    + CellReference.convertNumToColString(column) + " 列，请修改表头后重试。");
        }
    }

    public static void markForRecalculation(Workbook workbook) {
        if (workbook instanceof XSSFWorkbook) {
            // 保留用户设置的计算模式，让依赖更新值的公式在打开工作簿时刷新。
            CTWorkbook xml = ((XSSFWorkbook) workbook).getCTWorkbook();
            (xml.isSetCalcPr() ? xml.getCalcPr() : xml.addNewCalcPr()).setFullCalcOnLoad(true);
        } else {
            workbook.setForceFormulaRecalculation(true);
        }
    }

    public static void writeFilledWorkbook(Workbook workbook, File destination, OutputStream stream) throws IOException {
        boolean xlsOutput = destination.getName().toLowerCase(Locale.ROOT).endsWith(".xls");
        if (xlsOutput && !(workbook instanceof HSSFWorkbook)) {
            throw new IOException("请将此工作簿保存为 .xlsx，以保留原格式。");
        }
        if (workbook instanceof HSSFWorkbook && !xlsOutput) {
            LegacyWorkbookConverter.writeXlsx((HSSFWorkbook) workbook, stream);
        } else {
            workbook.write(stream);
        }
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

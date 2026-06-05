package com.xzy.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel工具类：公共方法提取
 */
public class ExcelUtils {
    private static final Logger logger = LogManager.getLogger(ExcelUtils.class);

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
                    return cell.getDateCellValue().toString();
                }
                double value = cell.getNumericCellValue();
                if (value == Math.floor(value)) {
                    return String.format("%.0f", value);
                }
                return String.valueOf(value);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return cell.getDateCellValue().toString();
                            }
                            double formulaValue = cell.getNumericCellValue();
                            if (formulaValue == Math.floor(formulaValue)) {
                                return String.format("%.0f", formulaValue);
                            }
                            return String.valueOf(formulaValue);
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

    /**
     * 使用SAX模式只读取Excel文件的表头行
     */
    public static List<String> readHeadersOnly(File file) throws IOException {
        if (file.getName().toLowerCase().endsWith(".xlsx")) {
            return readXLSXHeaders(file);
        } else {
            return readXLSHeaders(file);
        }
    }

    /**
     * 使用SAX模式读取XLSX文件的表头
     */
    private static List<String> readXLSXHeaders(File file) throws IOException {
        List<String> headers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             OPCPackage pkg = OPCPackage.open(fis)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();

            if (it.hasNext()) {
                InputStream sheet = it.next();
                try {
                    XMLReader parser = javax.xml.parsers.SAXParserFactory.newInstance()
                            .newSAXParser().getXMLReader();
                    try {
                        parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                        parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
                        parser.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
                    } catch (SAXNotRecognizedException | SAXNotSupportedException e) {
                        throw new RuntimeException("不安全的XML解析器，无法设置XXE防护", e);
                    }

                    HeaderHandler handler = new HeaderHandler(sst);
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheet));
                    headers = handler.getHeaders();
                } catch (SAXException e) {
                    throw new IOException("解析Excel文件时出错: " + e.getMessage(), e);
                } finally {
                    sheet.close();
                }
            }
        } catch (OpenXML4JException | javax.xml.parsers.ParserConfigurationException e) {
            throw new IOException("打开Excel文件时出错: " + e.getMessage(), e);
        }
        return headers;
    }

    /**
     * 读取XLS文件的表头
     */
    private static List<String> readXLSHeaders(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            List<String> headers = new ArrayList<>();
            if (headerRow != null) {
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        headers.add(cell.getStringCellValue().trim());
                    }
                }
            }
            return headers;
        }
    }

    /**
     * 去除所有不可见字符（空格、制表符、换行、BOM）
     */
    public static String cleanString(String s) {
        if (s == null) return null;
        return s.replaceAll("\\s+", "").replace("﻿", "");
    }

    /**
     * SAX处理器，用于解析XLSX文件的表头，支持共享字符串
     */
    private static class HeaderHandler extends DefaultHandler {
        private final List<String> headers = new ArrayList<>();
        private final SharedStringsTable sst;
        private boolean isHeaderRow = true;
        private String cellType = null;
        private StringBuilder currentValue = new StringBuilder();
        private boolean isCellValue = false;

        public HeaderHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("c".equals(qName)) {
                cellType = attributes.getValue("t");
                isCellValue = true;
                currentValue.setLength(0);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCellValue) {
                currentValue.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (!isHeaderRow) return;

            if ("c".equals(qName)) {
                String value = currentValue.toString().trim();
                if ("s".equals(cellType)) {
                    try {
                        int idx = Integer.parseInt(value);
                        value = sst.getItemAt(idx).getString();
                    } catch (Exception e) {
                        // 解析共享字符串失败，使用原始值
                    }
                }
                headers.add(value);
                isCellValue = false;
            } else if ("row".equals(qName)) {
                isHeaderRow = false;
            }
        }

        public List<String> getHeaders() {
            return headers;
        }
    }
}

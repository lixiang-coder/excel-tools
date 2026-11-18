package com.xzy.util;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.XMLConstants;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/** 首个工作表的流式读取；表头和数据行使用相同的单元格位置及缓存值规则。 */
public final class XlsxSheetReader {
    private XlsxSheetReader() {
    }

    /** 保留空列占位，读完首个实际存在的行即停止。 */
    public static List<String> readHeaders(File file) throws IOException {
        List<String> headers = new ArrayList<>();
        read(file, true, -1, -1, (number, row, key, preservedCell) -> headers.addAll(row));
        return headers;
    }

    /** 跳过表头后逐行回调，不把整个工作表的数据保存在内存中。 */
    public static void readDataRows(File file, Consumer<List<String>> rowConsumer) throws IOException {
        read(file, false, -1, -1, (number, row, key, preservedCell) -> rowConsumer.accept(row));
    }

    /** 应用映射时同时提供实际行号、与关系表一致的查找键和填充列原值。 */
    public static void readMappingRows(File file, int lookupColumn, int fillColumn,
                                       MappingRowConsumer rowConsumer) throws IOException {
        read(file, false, lookupColumn, fillColumn, rowConsumer);
    }

    @FunctionalInterface
    public interface MappingRowConsumer {
        void accept(int rowNumber, List<String> row, String lookupKey, PreservedCell preservedCell);
    }

    /** 原单元格或公式缓存的标量，用于需要跳过修改的目标格。 */
    public static final class PreservedCell {
        private final CellType type;
        private final String value;
        private final String dateFormat;
        private final boolean date1904;

        private PreservedCell(CellType type, String value, String dateFormat, boolean date1904) {
            this.type = type;
            this.value = value;
            this.dateFormat = dateFormat;
            this.date1904 = date1904;
        }

        public void writeTo(Cell target, Map<String, CellStyle> dateStyles) {
            target.setBlank();
            switch (type) {
                case STRING:
                    target.setCellValue(value);
                    break;
                case BOOLEAN:
                    target.setCellValue("1".equals(value) || "true".equalsIgnoreCase(value));
                    break;
                case ERROR:
                    target.setCellErrorValue(FormulaError.forString(value).getCode());
                    break;
                case NUMERIC:
                    double number = Double.parseDouble(value);
                    if (dateFormat == null) {
                        target.setCellValue(number);
                    } else {
                        ExcelUtils.setDateCellValue(target, DateUtil.getLocalDateTime(number, date1904),
                                dateFormat, dateStyles);
                    }
                    break;
                default:
                    break;
            }
        }
    }

    private static void read(File file, boolean headersOnly, int lookupColumn, int preservedColumn,
                             MappingRowConsumer rowConsumer) throws IOException {
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings strings = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            boolean date1904 = uses1904Dates(reader);
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            if (!sheets.hasNext()) {
                throw new IOException("源文件中没有找到工作表: " + file.getAbsolutePath());
            }
            try (InputStream sheet = sheets.next()) {
                SheetHandler handler = new SheetHandler(file, sheets.getSheetName(), strings, styles,
                        date1904, headersOnly, lookupColumn, preservedColumn, rowConsumer);
                XMLReader parser = secureParser();
                parser.setContentHandler(handler);
                try {
                    parser.parse(new InputSource(sheet));
                } catch (HeaderComplete expected) {
                    // 表头加载不解析后续数据，后续公式没有缓存不影响选择字段。
                }
            }
        } catch (OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IOException("读取Excel文件失败: " + e.getMessage(), e);
        }
    }

    private static boolean uses1904Dates(XSSFReader reader)
            throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        final boolean[] date1904 = {false};
        XMLReader parser = secureParser();
        parser.setContentHandler(new DefaultHandler() {
            @Override
            public void startElement(String uri, String localName, String qName, Attributes attributes) {
                if ("workbookPr".equals(localName)) {
                    String value = attributes.getValue("date1904");
                    date1904[0] = "1".equals(value) || "true".equalsIgnoreCase(value);
                }
            }
        });
        try (InputStream workbook = reader.getWorkbookData()) {
            parser.parse(new InputSource(workbook));
        }
        return date1904[0];
    }

    private static XMLReader secureParser() throws ParserConfigurationException, SAXException {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        factory.setNamespaceAware(true);
        factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
        XMLReader parser = factory.newSAXParser().getXMLReader();
        parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
        parser.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        parser.setProperty(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        parser.setProperty(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        return parser;
    }

    private static final class HeaderComplete extends SAXException {
        private static final long serialVersionUID = 1L;
    }

    private static final class SheetHandler extends DefaultHandler {
        private final File file;
        private final String sheetName;
        private final SharedStrings strings;
        private final StylesTable styles;
        private final byte[] dateStyles;
        private final boolean date1904;
        private final boolean headersOnly;
        private final int lookupColumn;
        private final int preservedColumn;
        private final MappingRowConsumer rowConsumer;
        private final StringBuilder value = new StringBuilder();
        private List<String> row = new ArrayList<>();
        private boolean firstRow = true;
        private int rowNumber;
        private int column;
        private int nextColumn;
        private String reference;
        private String type;
        private int styleIndex;
        private boolean inCell;
        private boolean inValue;
        private boolean inInlineString;
        private boolean inInlineText;
        private int phoneticDepth;
        private boolean formula;
        private boolean valuePresent;
        private PreservedCell preservedCell;
        private String lookupKey;

        private SheetHandler(File file, String sheetName, SharedStrings strings, StylesTable styles,
                             boolean date1904, boolean headersOnly, int lookupColumn, int preservedColumn,
                             MappingRowConsumer rowConsumer) {
            this.file = file;
            this.sheetName = sheetName;
            this.strings = strings;
            this.styles = styles;
            this.dateStyles = new byte[styles == null ? 0 : styles.getNumCellStyles()];
            this.date1904 = date1904;
            this.headersOnly = headersOnly;
            this.lookupColumn = lookupColumn;
            this.preservedColumn = preservedColumn;
            this.rowConsumer = rowConsumer;
        }

        @Override
        public void startElement(String uri, String name, String qName, Attributes attributes) throws SAXException {
            if ("row".equals(name)) {
                row = new ArrayList<>();
                preservedCell = null;
                lookupKey = null;
                nextColumn = 0;
                String index = attributes.getValue("r");
                rowNumber = index == null ? rowNumber + 1 : integer(index, "行号");
            } else if ("c".equals(name)) {
                inCell = true;
                reference = attributes.getValue("r");
                column = reference == null ? nextColumn : columnIndex(reference);
                if (column >= 16384) {
                    throw cellError("列号超出 XLSX 范围");
                }
                if (reference == null) {
                    reference = CellReference.convertNumToColString(column) + rowNumber;
                }
                nextColumn = column + 1;
                type = attributes.getValue("t");
                String style = attributes.getValue("s");
                styleIndex = style == null ? 0 : integer(style, "样式索引");
                formula = false;
                valuePresent = false;
                inValue = false;
                inInlineString = false;
                inInlineText = false;
                phoneticDepth = 0;
                value.setLength(0);
            } else if (inCell) {
                if ("f".equals(name)) {
                    formula = true;
                } else if ("v".equals(name)) {
                    valuePresent = true;
                    inValue = true;
                } else if ("is".equals(name)) {
                    inInlineString = true;
                } else if ("rPh".equals(name)) {
                    phoneticDepth++;
                } else if ("t".equals(name) && inInlineString && phoneticDepth == 0) {
                    inInlineText = true;
                }
            }
        }

        @Override
        public void characters(char[] characters, int start, int length) {
            if (inValue || inInlineText) {
                value.append(characters, start, length);
            }
        }

        @Override
        public void endElement(String uri, String name, String qName) throws SAXException {
            if ("c".equals(name)) {
                while (row.size() <= column) {
                    row.add("");
                }
                String decoded = cellValue();
                row.set(column, decoded);
                if (column == lookupColumn && !firstRow) {
                    lookupKey = lookupKey(decoded);
                }
                if (column == preservedColumn) {
                    preservedCell = preservedCellValue(decoded);
                }
                inCell = false;
            } else if ("row".equals(name)) {
                if (firstRow) {
                    firstRow = false;
                    if (headersOnly) {
                        rowConsumer.accept(rowNumber, row, lookupKey, preservedCell);
                        throw new HeaderComplete();
                    }
                } else {
                    rowConsumer.accept(rowNumber, row, lookupKey, preservedCell);
                }
            } else if ("v".equals(name)) {
                inValue = false;
            } else if ("is".equals(name)) {
                inInlineString = false;
            } else if ("t".equals(name)) {
                inInlineText = false;
            } else if ("rPh".equals(name)) {
                phoneticDepth--;
            }
        }

        // 与 WorkbookFactory 读取关系表时的值表示一致，同时保留原行文本供正常导出。
        private String lookupKey(String decoded) throws SAXException {
            if ("b".equals(type)) {
                return String.valueOf("1".equals(decoded) || "true".equalsIgnoreCase(decoded));
            }
            if (!formula && !decoded.isEmpty() && (type == null || "n".equals(type)) && !isDateStyle()) {
                try {
                    return ExcelUtils.numericText(Double.parseDouble(decoded));
                } catch (NumberFormatException e) {
                    throw cellError("无效的查找列数值: " + decoded);
                }
            }
            return decoded.trim();
        }

        private PreservedCell preservedCellValue(String decoded) {
            String exact = value.toString();
            if ("s".equals(type)) {
                return new PreservedCell(CellType.STRING, decoded, null, date1904);
            }
            if ("inlineStr".equals(type) || "str".equals(type) || "d".equals(type)) {
                return new PreservedCell(CellType.STRING, exact, null, date1904);
            }
            String raw = exact.trim();
            if (raw.isEmpty()) {
                return new PreservedCell(CellType.BLANK, "", null, date1904);
            }
            if ("b".equals(type)) {
                return new PreservedCell(CellType.BOOLEAN, raw, null, date1904);
            }
            if ("e".equals(type)) {
                return new PreservedCell(CellType.ERROR, raw, null, date1904);
            }
            String format = isDateStyle() ? styles.getStyleAt(styleIndex).getDataFormatString() : null;
            return new PreservedCell(CellType.NUMERIC, raw, format, date1904);
        }

        private String cellValue() throws SAXException {
            String raw = value.toString().trim();
            if (formula && (!valuePresent || (raw.isEmpty() && !"str".equals(type)))) {
                throw cellError("公式缺少已保存的计算结果，请使用 Excel 或 WPS 重新计算并保存文件后再试");
            }
            if ("s".equals(type)) {
                if (raw.isEmpty()) {
                    return "";
                }
                try {
                    if (strings == null) {
                        throw cellError("缺少共享字符串表");
                    }
                    return strings.getItemAt(Integer.parseInt(raw)).getString();
                } catch (NumberFormatException | IndexOutOfBoundsException e) {
                    throw cellError("无效的共享字符串索引: " + raw);
                }
            }
            if ("inlineStr".equals(type) || "str".equals(type) || "e".equals(type) || "d".equals(type)) {
                return raw;
            }
            if (raw.isEmpty()) {
                return "";
            }
            if ("b".equals(type)) {
                if (!formula) return raw;
                if ("1".equals(raw) || "true".equalsIgnoreCase(raw)) return "true";
                if ("0".equals(raw) || "false".equalsIgnoreCase(raw)) return "false";
                throw cellError("无效的布尔值: " + raw);
            }
            if (isDateStyle()) {
                try {
                    return ExcelUtils.formatExcelDate(Double.parseDouble(raw), date1904);
                } catch (NumberFormatException e) {
                    throw cellError("无效的日期值: " + raw);
                }
            }
            if (formula) {
                try {
                    double number = Double.parseDouble(raw);
                    return ExcelUtils.numericText(number);
                } catch (NumberFormatException e) {
                    throw cellError("无效的公式数值缓存: " + raw);
                }
            }
            // 普通数字保持原 XLSX 文本表现，避免改变已有主键的匹配规则。
            return raw;
        }

        private boolean isDateStyle() {
            if (styleIndex < 0 || styleIndex >= dateStyles.length) return false;
            if (dateStyles[styleIndex] == 0) {
                CellStyle style = styles.getStyleAt(styleIndex);
                dateStyles[styleIndex] = (byte) (DateUtil.isADateFormat(style.getDataFormat(),
                        style.getDataFormatString()) ? 2 : 1);
            }
            return dateStyles[styleIndex] == 2;
        }

        private int columnIndex(String cellReference) throws SAXException {
            int index = 0;
            int end = 0;
            while (end < cellReference.length()) {
                char character = Character.toUpperCase(cellReference.charAt(end));
                if (character < 'A' || character > 'Z') break;
                index = index * 26 + character - 'A' + 1;
                if (index > 16384) throw cellError("列号超出 XLSX 范围: " + cellReference);
                end++;
            }
            if (index == 0 || end == cellReference.length() || cellReference.charAt(end) == '0') {
                throw cellError("无效的单元格位置: " + cellReference);
            }
            for (int position = end; position < cellReference.length(); position++) {
                char character = cellReference.charAt(position);
                if (character < '0' || character > '9') {
                    throw cellError("无效的单元格位置: " + cellReference);
                }
            }
            return index - 1;
        }

        private int integer(String text, String label) throws SAXException {
            try {
                return Integer.parseInt(text);
            } catch (NumberFormatException e) {
                throw cellError("无效的" + label + ": " + text);
            }
        }

        private SAXException cellError(String message) {
            return new SAXException("源文件 " + file.getAbsolutePath() + "，工作表 " + sheetName
                    + "，单元格 " + reference + "：" + message);
        }
    }
}

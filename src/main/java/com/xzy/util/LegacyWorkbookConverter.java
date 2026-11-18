package com.xzy.util;

import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.CalcModeRecord;
import org.apache.poi.hssf.record.CalcCountRecord;
import org.apache.poi.hssf.record.ColumnInfoRecord;
import org.apache.poi.hssf.record.DeltaRecord;
import org.apache.poi.hssf.record.IterationRecord;
import org.apache.poi.hssf.record.NameRecord;
import org.apache.poi.hssf.record.PrecisionRecord;
import org.apache.poi.hssf.record.ProtectRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordBase;
import org.apache.poi.hssf.record.SCLRecord;
import org.apache.poi.hssf.record.SupBookRecord;
import org.apache.poi.hssf.record.WindowProtectRecord;
import org.apache.poi.hssf.record.aggregates.ColumnInfoRecordsAggregate;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCalcMode;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STVisibility;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 普通 BIFF 工作簿的兼容导出；不重新计算公式，也不修改传入的工作簿。
 * 无法保真的对象明确拒绝，调用方应提供同格式 .xls 导出，并用 SafeExport 发布文件。
 */
public final class LegacyWorkbookConverter {
    private LegacyWorkbookConverter() {
    }

    public static void writeXlsx(HSSFWorkbook source, OutputStream stream) throws IOException {
        try {
            checkSupported(source);
            try (XSSFWorkbook target = new XSSFWorkbook()) {
                target.getCTWorkbook().getWorkbookPr()
                        .setDate1904(source.getWorkbook().isUsing1904DateWindowing());
                for (int i = 0; i < source.getNumberOfSheets(); i++) {
                    target.createSheet(source.getSheetName(i));
                    target.setSheetVisibility(i, source.getSheetVisibility(i));
                }
                copyNames(source, target);
                Styles styles = new Styles(source, target);
                for (int i = 0; i < source.getNumberOfSheets(); i++) {
                    copySheet(source.getSheetAt(i), target.getSheetAt(i), styles);
                }
                target.setActiveSheet(source.getActiveSheetIndex());
                target.setFirstVisibleTab(source.getFirstVisibleTab());
                target.getCTWorkbook().getBookViews().getWorkbookViewArray(0)
                        .setVisibility(source.isHidden() ? STVisibility.HIDDEN : STVisibility.VISIBLE);
                target.setForceFormulaRecalculation(true);
                PrecisionRecord precision = (PrecisionRecord) source.getWorkbook()
                        .findFirstRecordBySid(PrecisionRecord.sid);
                if (precision != null) {
                    target.getCTWorkbook().getCalcPr().setFullPrecision(precision.getFullPrecision());
                }
                // Excel 的工作簿计算参数在 BIFF 工作表中存储；首张原表为准，新增报告表使用默认参数。
                if (source.getNumberOfSheets() > 0) {
                    HSSFSheet first = source.getSheetAt(0);
                    IterationRecord iteration = (IterationRecord) first.getSheet()
                            .findFirstRecordBySid(IterationRecord.sid);
                    CalcCountRecord count = (CalcCountRecord) first.getSheet()
                            .findFirstRecordBySid(CalcCountRecord.sid);
                    DeltaRecord delta = (DeltaRecord) first.getSheet().findFirstRecordBySid(DeltaRecord.sid);
                    if (iteration != null) target.getCTWorkbook().getCalcPr().setIterate(iteration.getIteration());
                    if (count != null) target.getCTWorkbook().getCalcPr().setIterateCount(count.getIterations() & 0xffff);
                    if (delta != null) target.getCTWorkbook().getCalcPr().setIterateDelta(delta.getMaxChange());
                }
                // BIFF 的重算 getter 与其 setter 含义不一致；保留手动计算模式但标记缓存需要更新。
                for (int i = 0; i < source.getNumberOfSheets(); i++) {
                    CalcModeRecord mode = (CalcModeRecord) source.getSheetAt(i).getSheet()
                            .findFirstRecordBySid(CalcModeRecord.sid);
                    if (mode != null && mode.getCalcMode() == CalcModeRecord.MANUAL) {
                        target.getCTWorkbook().getCalcPr().setCalcMode(STCalcMode.MANUAL);
                        break;
                    }
                    if (mode != null && mode.getCalcMode() == CalcModeRecord.AUTOMATIC_EXCEPT_TABLES) {
                        target.getCTWorkbook().getCalcPr().setCalcMode(STCalcMode.AUTO_NO_TABLE);
                    }
                }
                // 全部转换成功后才写入；流的生命周期和安全发布由调用方负责。
                target.write(stream);
            }
        } catch (RuntimeException e) {
            throw new IOException("无法安全转换为 .xlsx，请另存为同格式 .xls。原因：" + e.getMessage(), e);
        }
    }

    private static IOException unsupported(String description) {
        return new IOException(description + "，无法安全转换为 .xlsx，请另存为同格式 .xls 以保留原内容。");
    }

    private static void checkSupported(HSSFWorkbook source) throws IOException {
        DirectoryNode directory = source.getDirectory();
        if (directory != null) {
            for (Entry entry : directory) {
                String name = entry.getName();
                if (!"Workbook".equals(name) && !"Book".equals(name)
                        && !"\005SummaryInformation".equals(name)
                        && !"\005DocumentSummaryInformation".equals(name)) {
                    throw unsupported("工作簿包含宏、嵌入对象或其他附加内容（" + name + "）");
                }
            }
        }
        if (source.isWriteProtected() || !source.getAllPictures().isEmpty()
                || !source.getAllEmbeddedObjects().isEmpty()) {
            throw unsupported("工作簿包含写保护、图片或嵌入对象");
        }
        for (Record record : source.getWorkbook().getRecords()) {
            if (record instanceof SupBookRecord && !((SupBookRecord) record).isInternalReferences()) {
                throw unsupported("工作簿包含外部工作簿引用或加载项函数");
            }
            if ((record instanceof ProtectRecord && ((ProtectRecord) record).getProtect())
                    || (record instanceof WindowProtectRecord && ((WindowProtectRecord) record).getProtect())) {
                throw unsupported("工作簿包含结构或窗口保护");
            }
        }
        for (int i = 0; i < source.getWorkbook().getNumNames(); i++) {
            NameRecord record = source.getWorkbook().getNameRecord(i);
            if (record.isFunctionName() || record.isHiddenName()
                    || (record.isBuiltInName() && record.getBuiltInName() != NameRecord.BUILTIN_PRINT_AREA
                    && record.getBuiltInName() != NameRecord.BUILTIN_PRINT_TITLE)) {
                throw unsupported("工作簿包含特殊名称或筛选（" + record.getNameText() + "）");
            }
        }
        for (int i = 0; i < source.getNumberOfSheets(); i++) {
            HSSFSheet sheet = source.getSheetAt(i);
            String location = "工作表“" + sheet.getSheetName() + "”";
            BOFRecord bof = (BOFRecord) sheet.getSheet().findFirstRecordBySid(BOFRecord.sid);
            if (bof != null && bof.getType() != BOFRecord.TYPE_WORKSHEET) {
                throw unsupported(location + "不是普通工作表");
            }
            if (sheet.getDrawingPatriarch() != null || HSSFChart.getSheetCharts(sheet).length > 0) {
                throw unsupported(location + "包含图形、图表或批注");
            }
            if (sheet.getProtect() || sheet.getObjectProtect() || sheet.getScenarioProtect()
                    || !sheet.getDataValidations().isEmpty()
                    || sheet.getSheetConditionalFormatting().getNumConditionalFormattings() != 0) {
                throw unsupported(location + "包含保护、数据验证或条件格式");
            }
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.isPartOfArrayFormulaGroup()) {
                        throw unsupported(location + "的 " + cell.getAddress() + " 包含数组公式");
                    }
                }
            }
        }
    }

    private static void copyNames(HSSFWorkbook source, XSSFWorkbook target) {
        Map<String, NameRecord> records = new HashMap<>();
        for (int i = 0; i < source.getWorkbook().getNumNames(); i++) {
            NameRecord record = source.getWorkbook().getNameRecord(i);
            records.put(record.getSheetNumber() + ":" + record.getNameText(), record);
        }
        Map<HSSFName, XSSFName> names = new LinkedHashMap<>();
        for (HSSFName original : source.getAllNames()) {
            NameRecord record = records.get((original.getSheetIndex() + 1) + ":" + original.getNameName());
            // BIFF 打印标题的全列范围到 IV；OOXML 必须按新的行列上限重新生成。
            if (record != null && record.isBuiltInName()) continue;
            XSSFName copy = target.createName();
            copy.setSheetIndex(original.getSheetIndex());
            copy.setNameName(original.getNameName());
            copy.setComment(original.getComment());
            names.put(original, copy);
        }
        // 先注册全部名称，以便名称或公式引用后面定义的名称。
        for (Map.Entry<HSSFName, XSSFName> name : names.entrySet()) {
            String formula = name.getKey().getRefersToFormula();
            if (formula != null) {
                name.getValue().setRefersToFormula(formula);
            }
        }
        for (int i = 0; i < source.getNumberOfSheets(); i++) {
            HSSFSheet sheet = source.getSheetAt(i);
            XSSFSheet copy = target.getSheetAt(i);
            copy.setRepeatingRows(sheet.getRepeatingRows());
            copy.setRepeatingColumns(sheet.getRepeatingColumns());
            String printArea = source.getPrintArea(i);
            if (printArea != null) {
                XSSFName area = target.createName();
                area.setSheetIndex(i);
                area.setNameName(XSSFName.BUILTIN_PRINT_AREA);
                area.setRefersToFormula(printArea);
            }
        }
    }

    private static void copySheet(HSSFSheet source, XSSFSheet target, Styles styles) throws IOException {
        target.setDefaultColumnWidth(source.getDefaultColumnWidth());
        target.setDefaultRowHeight(source.getDefaultRowHeight());
        ColumnInfoRecordsAggregate columns = null;
        for (RecordBase record : source.getSheet().getRecords()) {
            if (record instanceof ColumnInfoRecordsAggregate) {
                columns = (ColumnInfoRecordsAggregate) record;
                break;
            }
        }
        for (int c = 0; c < 256; c++) {
            target.setColumnWidth(c, source.getColumnWidth(c));
            target.setColumnHidden(c, source.isColumnHidden(c));
            CellStyle columnStyle = source.getColumnStyle(c);
            if (columnStyle != null) {
                target.setDefaultColumnStyle(c, styles.cellStyle(columnStyle));
            }
            ColumnInfoRecord column = columns == null ? null : columns.findColumnInfo(c);
            if (column != null) {
                CTCol copied = target.getColumnHelper().getColumn(c, true);
                copied.setOutlineLevel((short) column.getOutlineLevel());
                copied.setCollapsed(column.getCollapsed());
                target.getCTWorksheet().getSheetFormatPr().setOutlineLevelCol((short) Math.max(
                        target.getCTWorksheet().getSheetFormatPr().getOutlineLevelCol(), column.getOutlineLevel()));
            }
        }
        for (Row row : source) {
            XSSFRow copy = target.createRow(row.getRowNum());
            copy.setHeight(row.getHeight());
            copy.setZeroHeight(row.getZeroHeight());
            copy.getCTRow().setOutlineLevel((short) row.getOutlineLevel());
            copy.getCTRow().setCollapsed(source.getSheet().getRow(row.getRowNum()).getColapsed());
            target.getCTWorksheet().getSheetFormatPr().setOutlineLevelRow((short) Math.max(
                    target.getCTWorksheet().getSheetFormatPr().getOutlineLevelRow(), row.getOutlineLevel()));
            if (row.isFormatted()) {
                copy.setRowStyle(styles.cellStyle(row.getRowStyle()));
            }
            for (Cell cell : row) {
                XSSFCell copied = copy.createCell(cell.getColumnIndex());
                copyCell((HSSFCell) cell, copied, styles);
                copied.setCellStyle(styles.cellStyle(cell.getCellStyle()));
            }
        }
        for (HSSFHyperlink original : source.getHyperlinkList()) {
            XSSFRow row = target.getRow(original.getFirstRow());
            if (row == null) row = target.createRow(original.getFirstRow());
            XSSFCell cell = row.getCell(original.getFirstColumn());
            if (cell == null) cell = row.createCell(original.getFirstColumn());
            XSSFHyperlink copy = new XSSFHyperlink(original);
            cell.setHyperlink(copy);
            // setHyperlink 会先收窄到锚点单元格；恢复原范围，包括没有创建单元格的空白位置。
            copy.setLastRow(original.getLastRow());
            copy.setLastColumn(original.getLastColumn());
        }
        for (CellRangeAddress range : source.getMergedRegions()) {
            target.addMergedRegionUnsafe(range.copy());
        }
        copySheetSettings(source, target);
    }

    private static void copyCell(HSSFCell source, XSSFCell target, Styles styles) throws IOException {
        if (source.getCellType() == CellType.FORMULA) {
            CellValue cache = FormulaCache.read(source);
            target.setCellFormula(source.getCellFormula());
            // setCellValue 在公式格上写缓存；不能使用 copyCell，否则缓存会变成缺失的默认 0。
            switch (cache.getCellType()) {
                case NUMERIC: target.setCellValue(cache.getNumberValue()); break;
                case STRING: target.setCellValue(cache.getStringValue()); break;
                case BOOLEAN: target.setCellValue(cache.getBooleanValue()); break;
                case ERROR: target.setCellErrorValue(cache.getErrorValue()); break;
                default: throw unsupported("公式缓存类型无法转换：" + source.getAddress());
            }
            return;
        }
        switch (source.getCellType()) {
            case NUMERIC: target.setCellValue(source.getNumericCellValue()); break;
            case BOOLEAN: target.setCellValue(source.getBooleanCellValue()); break;
            case ERROR: target.setCellErrorValue(source.getErrorCellValue()); break;
            case STRING:
                HSSFRichTextString original = source.getRichStringCellValue();
                XSSFRichTextString text = new XSSFRichTextString(original.getString());
                for (int i = 0; i < original.numFormattingRuns(); i++) {
                    int start = original.getIndexOfFormattingRun(i);
                    int end = i + 1 < original.numFormattingRuns()
                            ? original.getIndexOfFormattingRun(i + 1) : original.length();
                    if (start < end) {
                        text.applyFont(start, end, styles.font(original.getFontOfFormattingRun(i)));
                    }
                }
                target.setCellValue(text);
                break;
            case BLANK: target.setBlank(); break;
            default: throw unsupported("单元格类型无法转换：" + source.getAddress());
        }
    }

    private static void copySheetSettings(HSSFSheet source, XSSFSheet target) {
        target.setSelected(source.isSelected());
        target.setRightToLeft(source.isRightToLeft());
        target.setDisplayGridlines(source.isDisplayGridlines());
        target.setDisplayFormulas(source.isDisplayFormulas());
        target.setDisplayRowColHeadings(source.isDisplayRowColHeadings());
        target.setDisplayZeros(source.isDisplayZeros());
        target.setDisplayGuts(source.getDisplayGuts());
        target.setRowSumsBelow(source.getRowSumsBelow());
        target.setRowSumsRight(source.getRowSumsRight());
        target.setForceFormulaRecalculation(source.getForceFormulaRecalculation());
        target.showInPane(source.getTopRow() & 0xffff, source.getLeftCol() & 0xffff);
        target.setActiveCell(source.getActiveCell());
        SCLRecord zoom = (SCLRecord) source.getSheet().findFirstRecordBySid(SCLRecord.sid);
        if (zoom != null && zoom.getDenominator() != 0) {
            target.setZoom(100 * zoom.getNumerator() / zoom.getDenominator());
        }
        PaneInformation pane = source.getPaneInformation();
        if (pane != null) {
            if (pane.isFreezePane()) {
                target.createFreezePane(pane.getVerticalSplitPosition(), pane.getHorizontalSplitPosition(),
                        pane.getVerticalSplitLeftColumn(), pane.getHorizontalSplitTopRow());
            } else {
                target.createSplitPane(pane.getVerticalSplitPosition(), pane.getHorizontalSplitPosition(),
                        pane.getVerticalSplitLeftColumn(), pane.getHorizontalSplitTopRow(), pane.getActivePaneType());
            }
        }
        target.setPrintGridlines(source.isPrintGridlines());
        target.setPrintRowAndColumnHeadings(source.isPrintRowAndColumnHeadings());
        target.setHorizontallyCenter(source.getHorizontallyCenter());
        target.setVerticallyCenter(source.getVerticallyCenter());
        target.setAutobreaks(source.getAutobreaks());
        target.setFitToPage(source.getFitToPage());
        for (short margin = 0; margin <= 5; margin++) {
            target.setMargin(margin, source.getMargin(margin));
        }
        target.getHeader().setLeft(source.getHeader().getLeft());
        target.getHeader().setCenter(source.getHeader().getCenter());
        target.getHeader().setRight(source.getHeader().getRight());
        target.getFooter().setLeft(source.getFooter().getLeft());
        target.getFooter().setCenter(source.getFooter().getCenter());
        target.getFooter().setRight(source.getFooter().getRight());
        for (int row : source.getRowBreaks()) target.setRowBreak(row);
        for (int column : source.getColumnBreaks()) target.setColumnBreak(column);
        PrintSetup print = source.getPrintSetup();
        PrintSetup copy = target.getPrintSetup();
        copy.setPaperSize(print.getPaperSize());
        if (print.getScale() >= 10 && print.getScale() <= 400) copy.setScale(print.getScale());
        copy.setPageStart(print.getPageStart());
        copy.setFitWidth(print.getFitWidth());
        copy.setFitHeight(print.getFitHeight());
        copy.setLandscape(print.getLandscape());
        copy.setLeftToRight(print.getLeftToRight());
        copy.setNoColor(print.getNoColor());
        copy.setDraft(print.getDraft());
        copy.setNotes(print.getNotes());
        copy.setNoOrientation(print.getNoOrientation());
        copy.setUsePage(print.getUsePage());
        copy.setHResolution(print.getHResolution());
        copy.setVResolution(print.getVResolution());
        copy.setHeaderMargin(print.getHeaderMargin());
        copy.setFooterMargin(print.getFooterMargin());
        copy.setCopies(print.getCopies());
    }

    private static final class Styles {
        private final HSSFWorkbook source;
        private final XSSFWorkbook target;
        private final Map<Short, XSSFCellStyle> styles = new HashMap<>();
        private final Map<Integer, XSSFFont> fonts = new HashMap<>();

        private Styles(HSSFWorkbook source, XSSFWorkbook target) {
            this.source = source;
            this.target = target;
        }

        private XSSFCellStyle cellStyle(CellStyle original) {
            XSSFCellStyle copy = styles.get(original.getIndex());
            if (copy != null) return copy;
            copy = target.createCellStyle();
            styles.put(original.getIndex(), copy);
            copy.setFont(font(original.getFontIndex()));
            copy.setDataFormat(target.createDataFormat().getFormat(original.getDataFormatString()));
            copy.setAlignment(original.getAlignment());
            copy.setVerticalAlignment(original.getVerticalAlignment());
            copy.setWrapText(original.getWrapText());
            copy.setRotation(original.getRotation());
            copy.setIndention(original.getIndention());
            copy.setShrinkToFit(original.getShrinkToFit());
            copy.setLocked(original.getLocked());
            copy.setHidden(original.getHidden());
            copy.setQuotePrefixed(original.getQuotePrefixed());
            copy.setReadingOrder(ReadingOrder.forLong(((HSSFCellStyle) original).getReadingOrder()));
            copy.setBorderLeft(original.getBorderLeft());
            copy.setBorderRight(original.getBorderRight());
            copy.setBorderTop(original.getBorderTop());
            copy.setBorderBottom(original.getBorderBottom());
            copy.setLeftBorderColor(color(original.getLeftBorderColor()));
            copy.setRightBorderColor(color(original.getRightBorderColor()));
            copy.setTopBorderColor(color(original.getTopBorderColor()));
            copy.setBottomBorderColor(color(original.getBottomBorderColor()));
            copy.setFillForegroundColor(color(original.getFillForegroundColor()));
            copy.setFillBackgroundColor(color(original.getFillBackgroundColor()));
            copy.setFillPattern(original.getFillPattern());
            return copy;
        }

        private XSSFFont font(int index) {
            XSSFFont copy = fonts.get(index);
            if (copy != null) return copy;
            HSSFFont original = source.getFontAt(index);
            copy = target.createFont();
            fonts.put(index, copy);
            copy.setFontName(original.getFontName());
            copy.setFontHeight(original.getFontHeight());
            copy.setBold(original.getBold());
            copy.setItalic(original.getItalic());
            copy.setStrikeout(original.getStrikeout());
            copy.setUnderline(original.getUnderline());
            copy.setTypeOffset(original.getTypeOffset());
            copy.setCharSet(original.getCharSet());
            XSSFColor color = color(original.getColor());
            if (color.isIndexed()) copy.setColor(color.getIndexed());
            else if (color.isAuto()) copy.setColor(IndexedColors.AUTOMATIC.getIndex());
            else copy.setColor(color);
            return copy;
        }

        private XSSFColor color(short index) {
            HSSFColor original = source.getCustomPalette().getColor(index);
            HSSFColor standard = HSSFColor.getIndexHash().get((int) index);
            if (original != null && (standard == null
                    || !Arrays.equals(original.getTriplet(), standard.getTriplet()))) {
                short[] rgb = original.getTriplet();
                return new XSSFColor(new byte[]{(byte) rgb[0], (byte) rgb[1], (byte) rgb[2]});
            }
            XSSFColor color = new XSSFColor();
            if (index == Font.COLOR_NORMAL) color.setAuto(true);
            else color.setIndexed(index);
            return color;
        }
    }
}

package com.xzy;

import com.xzy.model.MappingRelations;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static org.junit.Assert.*;

@RunWith(Parameterized.class)
public class AppliedMappingWorkbookTest {
    private static final String MAIN_SHEET = "订单 主表";

    @Parameterized.Parameters(name = "{0} -> {1}")
    public static Collection<Object[]> formats() {
        return Arrays.asList(new Object[][]{{"xlsx", "xlsx"}, {"xls", "xls"}, {"xls", "xlsx"}});
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final String inputFormat;
    private final String outputFormat;

    public AppliedMappingWorkbookTest(String inputFormat, String outputFormat) {
        this.inputFormat = inputFormat;
        this.outputFormat = outputFormat;
    }

    @Test
    public void mappingPreservesWorkbookAndAppendsConflictReportWithoutOverwritingExistingSheet() throws Exception {
        File source = targetWorkbook();
        File relation = relationWorkbook(true);
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        byte[] relationsBefore = Files.readAllBytes(relation.toPath());
        File output = new File(temporary.getRoot(), "result." + outputFormat);

        int[] stats = export(source, output, relation);

        assertEquals(1, stats[1]);
        assertEquals(1, stats[2]);
        assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
        assertArrayEquals(relationsBefore, Files.readAllBytes(relation.toPath()));
        try (FileInputStream originalInput = new FileInputStream(source);
             Workbook original = WorkbookFactory.create(originalInput);
             FileInputStream resultInput = new FileInputStream(output);
             Workbook result = WorkbookFactory.create(resultInput)) {
            assertOutputFormat(result);
            assertEquals(original.getNumberOfSheets() + 1, result.getNumberOfSheets());
            assertPreservedWorkbook(original, result, new HashSet<>(Arrays.asList(1, 8)));
            Sheet main = result.getSheetAt(0);
            assertEquals("mapped value", main.getRow(1).getCell(1).getStringCellValue());
            assertEquals(CellType.NUMERIC, main.getRow(4).getCell(1).getCellType());
            assertEquals(42.25, main.getRow(4).getCell(1).getNumericCellValue(), 0.0);
            assertEquals("not-found", main.getRow(8).getCell(1).getStringCellValue());

            Sheet report = result.getSheetAt(result.getNumberOfSheets() - 1);
            assertNotEquals("映射冲突", report.getSheetName());
            assertNotSame(result.getSheet("映射冲突"), report);
            assertEquals("目标行号", report.getRow(0).getCell(0).getStringCellValue());
            assertEquals("关系行号", report.getRow(0).getCell(5).getStringCellValue());
            assertEquals(2, report.getLastRowNum());
            Set<String> candidates = new HashSet<>();
            for (int row = 1; row <= report.getLastRowNum(); row++) {
                assertEquals(5.0, report.getRow(row).getCell(0).getNumericCellValue(), 0.0);
                assertEquals("conflict", report.getRow(row).getCell(1).getStringCellValue());
                candidates.add(report.getRow(row).getCell(3).getStringCellValue());
            }
            assertEquals(new HashSet<>(Arrays.asList("candidate A", "candidate B")), candidates);
        }
    }

    @Test
    public void blankLookupKeysPreserveFillCellsAndZeroAndFalseStillMatch() throws Exception {
        File source = new File(temporary.getRoot(), "blank-keys." + inputFormat);
        try (Workbook workbook = "xls".equals(inputFormat) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(MAIN_SHEET);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Key");
            header.createCell(1).setCellValue("Fill");
            CellStyle style = customStyle(workbook);
            for (int index = 1; index <= 6; index++) {
                Cell fill = sheet.createRow(index).createCell(1);
                fill.setCellFormula("2+3");
                fill.setCellStyle(style);
            }
            // 第一行查找格缺失，其他行包含不同形式的空值。
            sheet.getRow(2).createCell(0).setBlank();
            sheet.getRow(3).createCell(0).setCellValue("");
            sheet.getRow(4).createCell(0).setCellValue(" \t ");
            sheet.getRow(5).createCell(0).setCellValue("\u00a0\u3000");
            sheet.getRow(6).createCell(0).setCellFormula("IF(1=1,\"\",\"x\")");
            sheet.createRow(7).createCell(0).setCellValue("   ");
            sheet.createRow(8).createCell(0).setCellValue(0);
            sheet.createRow(9).createCell(0).setCellValue(false);
            sheet.createRow(10).createCell(0).setCellValue("  unmatched  ");
            Row verified = sheet.createRow(11);
            verified.createCell(0, CellType.BLANK).setCellStyle(style);
            verified.createCell(1).setCellValue("已核对");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            write(workbook, source);
        }
        MappingRelations mappings = new MappingRelations();
        mappings.add("0", "zero", source, 1);
        mappings.add("false", "boolean", source, 2);
        File output = new File(temporary.getRoot(), "blank-key-result." + outputFormat);

        int[] stats = ExcelTools.exportAppliedMapping(source, output, "Key", "Fill", mappings,
                (percent, message) -> { }, message -> { });

        assertArrayEquals(new int[] {11, 2, 0}, stats);
        try (Workbook original = WorkbookFactory.create(source, null, true);
             Workbook result = WorkbookFactory.create(output, null, true)) {
            Sheet sheet = result.getSheetAt(0);
            for (int index = 1; index <= 6; index++) {
                Cell before = original.getSheetAt(0).getRow(index).getCell(1);
                Cell after = sheet.getRow(index).getCell(1);
                assertCellEquals(before, after);
                assertStyleEquals(original, before, result, after);
                assertEquals(5, after.getNumericCellValue(), 0);
            }
            assertNull(sheet.getRow(7).getCell(1));
            assertEquals("zero", sheet.getRow(8).getCell(1).getStringCellValue());
            assertEquals("boolean", sheet.getRow(9).getCell(1).getStringCellValue());
            assertEquals("xlsx".equals(inputFormat) ? "  unmatched  " : "unmatched",
                    sheet.getRow(10).getCell(1).getStringCellValue());
            Row verified = sheet.getRow(11);
            assertEquals(CellType.BLANK, verified.getCell(0).getCellType());
            assertStyleEquals(original, original.getSheetAt(0).getRow(11).getCell(0), result, verified.getCell(0));
            assertEquals("已核对", verified.getCell(1).getStringCellValue());
        }
    }

    @Test
    public void conflictFreeMappingKeepsOriginalSheetCountAndUntouchedWorkbookContent() throws Exception {
        File source = targetWorkbook();
        File relation = relationWorkbook(false);
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        File output = new File(temporary.getRoot(), "without-conflicts." + outputFormat);

        int[] stats = export(source, output, relation);

        assertEquals(1, stats[1]);
        assertEquals(0, stats[2]);
        assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
        try (FileInputStream originalInput = new FileInputStream(source);
             Workbook original = WorkbookFactory.create(originalInput);
             FileInputStream resultInput = new FileInputStream(output);
             Workbook result = WorkbookFactory.create(resultInput)) {
            assertOutputFormat(result);
            assertEquals(original.getNumberOfSheets(), result.getNumberOfSheets());
            assertPreservedWorkbook(original, result, new HashSet<>(Arrays.asList(1, 4, 8)));
            assertEquals("mapped value", result.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
            assertEquals("conflict", result.getSheetAt(0).getRow(4).getCell(1).getStringCellValue());
            assertEquals("not-found", result.getSheetAt(0).getRow(8).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void keepsCommentsInNativeFormatAndProtectsOutputWhenLegacyConversionCannotKeepThem() throws Exception {
        File input = temporary.newFile("notes." + inputFormat);
        try (Workbook book = "xls".equals(inputFormat) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet first = book.createSheet("订单");
            Row header = first.createRow(0);
            header.createCell(0).setCellValue("Key");
            header.createCell(1).setCellValue("Fill");
            first.createRow(1).createCell(0).setCellValue("unique");
            Sheet notes = book.createSheet("说明");
            CreationHelper helper = book.getCreationHelper();
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol2(3);
            anchor.setRow2(3);
            Comment comment = notes.createDrawingPatriarch().createCellComment(anchor);
            comment.setString(helper.createRichTextString("此处保留审核意见"));
            notes.createRow(0).createCell(0).setCellComment(comment);
            write(book, input);
        }
        File relation = relationWorkbook(false);
        byte[] before = Files.readAllBytes(input.toPath());
        File output = temporary.newFile("with-notes." + outputFormat);
        if (inputFormat.equals(outputFormat)) {
            export(input, output, relation);
            try (Workbook saved = WorkbookFactory.create(output, null, true)) {
                assertEquals("mapped value", saved.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
                assertEquals("此处保留审核意见", saved.getSheet("说明").getRow(0).getCell(0)
                        .getCellComment().getString().getString());
            }
        } else {
            byte[] previous = {12, 34, 56};
            Files.write(output.toPath(), previous);
            IOException failure = assertThrows(IOException.class, () -> export(input, output, relation));
            assertTrue(failure.getMessage(), failure.getMessage().contains(".xls"));
            assertArrayEquals(previous, Files.readAllBytes(output.toPath()));
        }
        assertArrayEquals(before, Files.readAllBytes(input.toPath()));
    }

    private static int[] export(File source, File output, File relation) throws Exception {
        MappingRelations mapping = new MappingRelations();
        ExcelTools.readMappingRelations(relation, "From", "To", mapping);
        return ExcelTools.exportAppliedMapping(source, output, "Key", "Fill", mapping,
                (percent, message) -> { }, message -> { });
    }

    private void assertOutputFormat(Workbook result) {
        if ("xls".equals(outputFormat)) assertTrue(result instanceof HSSFWorkbook);
        else assertTrue(result instanceof XSSFWorkbook);
    }

    private static void assertPreservedWorkbook(Workbook original, Workbook result, Set<Integer> changedFillRows) {
        for (int index = 0; index < original.getNumberOfSheets(); index++) {
            Sheet before = original.getSheetAt(index);
            Sheet after = result.getSheetAt(index);
            String context = "sheet " + before.getSheetName();
            assertEquals(context, before.getSheetName(), after.getSheetName());
            assertEquals(context, original.getSheetVisibility(index), result.getSheetVisibility(index));
            assertEquals(context, mergedRegions(before), mergedRegions(after));
            assertEquals(context, before.getLastRowNum(), after.getLastRowNum());
            int columnCount = 0;
            for (Row beforeRow : before) {
                Row afterRow = after.getRow(beforeRow.getRowNum());
                assertNotNull(context + ", row " + beforeRow.getRowNum(), afterRow);
                assertEquals(context, beforeRow.getHeight(), afterRow.getHeight());
                assertEquals(context, beforeRow.getZeroHeight(), afterRow.getZeroHeight());
                columnCount = Math.max(columnCount, beforeRow.getLastCellNum());
                for (Cell beforeCell : beforeRow) {
                    Cell afterCell = afterRow.getCell(beforeCell.getColumnIndex());
                    assertNotNull(context + ", cell " + beforeCell.getAddress(), afterCell);
                    if (index == 0 && beforeCell.getColumnIndex() == 1
                            && changedFillRows.contains(beforeRow.getRowNum())) continue;
                    assertCellEquals(beforeCell, afterCell);
                }
            }
            for (int column = 0; column < columnCount; column++) {
                assertEquals(context + ", width " + column, before.getColumnWidth(column), after.getColumnWidth(column));
                assertEquals(context + ", hidden column " + column, before.isColumnHidden(column), after.isColumnHidden(column));
            }
        }
        assertStyleEquals(original, original.getSheetAt(0).getRow(1).getCell(1),
                result, result.getSheetAt(0).getRow(1).getCell(1));
        assertStyleEquals(original, original.getSheetAt(0).getRow(1).getCell(2),
                result, result.getSheetAt(0).getRow(1).getCell(2));
        assertStyleEquals(original, original.getSheetAt(0).getRow(4).getCell(1),
                result, result.getSheetAt(0).getRow(4).getCell(1));
        assertStyleEquals(original, original.getSheetAt(1).getRow(1).getCell(0),
                result, result.getSheetAt(1).getRow(1).getCell(0));
    }

    private static void assertCellEquals(Cell before, Cell after) {
        String context = before.getSheet().getSheetName() + "!" + before.getAddress();
        assertEquals(context, before.getCellType(), after.getCellType());
        switch (before.getCellType()) {
            case STRING:
                assertEquals(context, before.getStringCellValue(), after.getStringCellValue());
                break;
            case NUMERIC:
                assertEquals(context, before.getNumericCellValue(), after.getNumericCellValue(), 0.0);
                break;
            case BOOLEAN:
                assertEquals(context, before.getBooleanCellValue(), after.getBooleanCellValue());
                break;
            case ERROR:
                assertEquals(context, before.getErrorCellValue(), after.getErrorCellValue());
                break;
            case FORMULA:
                assertEquals(context, before.getCellFormula(), after.getCellFormula());
                break;
            default:
                break;
        }
    }

    private static void assertStyleEquals(Workbook original, Cell beforeCell, Workbook result, Cell afterCell) {
        CellStyle before = beforeCell.getCellStyle();
        CellStyle after = afterCell.getCellStyle();
        assertEquals(before.getDataFormatString(), after.getDataFormatString());
        assertEquals(before.getFillPattern(), after.getFillPattern());
        assertEquals(before.getFillForegroundColor(), after.getFillForegroundColor());
        assertEquals(before.getAlignment(), after.getAlignment());
        assertEquals(before.getVerticalAlignment(), after.getVerticalAlignment());
        assertEquals(before.getWrapText(), after.getWrapText());
        assertEquals(before.getBorderBottom(), after.getBorderBottom());
        assertEquals(before.getBottomBorderColor(), after.getBottomBorderColor());
        Font beforeFont = original.getFontAt(before.getFontIndex());
        Font afterFont = result.getFontAt(after.getFontIndex());
        assertEquals(beforeFont.getFontName(), afterFont.getFontName());
        assertEquals(beforeFont.getFontHeight(), afterFont.getFontHeight());
        assertEquals(beforeFont.getBold(), afterFont.getBold());
        assertEquals(beforeFont.getColor(), afterFont.getColor());
    }

    private static List<String> mergedRegions(Sheet sheet) {
        List<String> regions = new ArrayList<>();
        for (CellRangeAddress region : sheet.getMergedRegions()) regions.add(region.formatAsString());
        return regions;
    }

    private File targetWorkbook() throws IOException {
        File file = new File(temporary.getRoot(), "target." + inputFormat);
        try (Workbook workbook = "xls".equals(inputFormat) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            CellStyle textStyle = customStyle(workbook);
            CellStyle numericStyle = workbook.createCellStyle();
            numericStyle.cloneStyleFrom(textStyle);
            numericStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            Sheet main = workbook.createSheet(MAIN_SHEET);
            Row header = main.createRow(0);
            String[] columns = {"Key", "Fill", "Amount", "Computed", "Notes", "Merge left", "Merge right"};
            for (int column = 0; column < columns.length; column++) {
                header.createCell(column).setCellValue(columns[column]);
                main.setColumnWidth(column, (15 + column) * 256 + 64);
            }
            header.setHeightInPoints(30);
            main.setColumnHidden(4, true);
            dataRow(main, 1, "unique", "old unique", 12.5, "C2*2");
            dataRow(main, 4, "conflict", "old conflict", 20.0, "C5+1");
            dataRow(main, 8, "not-found", "old fallback", 99.0, "C9/3");
            main.getRow(1).getCell(1).setCellStyle(textStyle);
            main.getRow(1).getCell(2).setCellStyle(numericStyle);
            main.getRow(4).getCell(1).setCellValue(42.25);
            main.getRow(4).getCell(1).setCellStyle(numericStyle);
            main.getRow(1).createCell(5).setCellValue("merged content");
            main.getRow(1).createCell(6).setBlank();
            main.addMergedRegion(new CellRangeAddress(1, 1, 5, 6));
            Row hiddenRow = main.createRow(6);
            hiddenRow.setZeroHeight(true);
            hiddenRow.setHeightInPoints(24);
            hiddenRow.createCell(4).setCellValue("hidden row content");

            Sheet summary = workbook.createSheet("汇总 数据");
            summary.createRow(0).createCell(0).setCellValue("other sheet must remain");
            Row summaryRow = summary.createRow(1);
            summaryRow.createCell(0).setCellValue(123.5);
            summaryRow.getCell(0).setCellStyle(numericStyle);
            summaryRow.createCell(1).setCellFormula("'" + MAIN_SHEET + "'!C2+'" + MAIN_SHEET + "'!C5");
            summaryRow.createCell(2).setCellValue(true);
            summary.setColumnWidth(0, 31 * 256);
            summary.setColumnWidth(1, 29 * 256);
            summary.createRow(3).createCell(0).setCellValue("merged summary");
            summary.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));

            Sheet existingReport = workbook.createSheet("映射冲突");
            existingReport.createRow(0).createCell(0).setCellValue("用户原有报告，不得覆盖");
            existingReport.createRow(1).createCell(0).setCellFormula("1+1");
            existingReport.getRow(1).createCell(1).setCellValue("existing-report-data");
            Sheet hidden = workbook.createSheet("内部 数据");
            hidden.createRow(0).createCell(0).setCellValue("hidden-data");
            hidden.createRow(1).createCell(0).setCellErrorValue(FormulaError.NA.getCode());
            workbook.setSheetVisibility(3, SheetVisibility.HIDDEN);
            Sheet veryHidden = workbook.createSheet("内部 规则");
            veryHidden.createRow(0).createCell(0).setCellFormula("10+20");
            workbook.setSheetVisibility(4, SheetVisibility.VERY_HIDDEN);
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            write(workbook, file);
        }
        return file;
    }

    private static CellStyle customStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setDataFormat(workbook.createDataFormat().getFormat("@"));
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.RED.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    private static void dataRow(Sheet sheet, int index, String key, String fill, double amount, String formula) {
        Row row = sheet.createRow(index);
        row.setHeightInPoints(22.5f);
        row.createCell(0).setCellValue(key);
        row.createCell(1).setCellValue(fill);
        row.createCell(2).setCellValue(amount);
        row.createCell(3).setCellFormula(formula);
        row.createCell(4).setCellValue("  unchanged text  ");
    }

    private File relationWorkbook(boolean withConflict) throws IOException {
        File file = new File(temporary.getRoot(), "relations.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Relations");
            sheet.createRow(0).createCell(0).setCellValue("From");
            sheet.getRow(0).createCell(1).setCellValue("To");
            sheet.createRow(1).createCell(0).setCellValue("unique");
            sheet.getRow(1).createCell(1).setCellValue("mapped value");
            if (withConflict) {
                sheet.createRow(2).createCell(0).setCellValue("conflict");
                sheet.getRow(2).createCell(1).setCellValue("candidate A");
                sheet.createRow(3).createCell(0).setCellValue("conflict");
                sheet.getRow(3).createCell(1).setCellValue("candidate B");
            }
            write(workbook, file);
        }
        return file;
    }

    private static void write(Workbook workbook, File file) throws IOException {
        try (FileOutputStream output = new FileOutputStream(file)) {
            workbook.write(output);
        }
    }
}

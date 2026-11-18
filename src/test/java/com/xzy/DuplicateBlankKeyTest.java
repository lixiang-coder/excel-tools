package com.xzy;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;

import static org.junit.Assert.*;

@RunWith(Parameterized.class)
public class DuplicateBlankKeyTest {
    private static final String MARKER_LIKE_KEY = "__BLANK_KEY_1__";
    private enum Empty { BLANK, FORMULA }

    @Parameterized.Parameters(name = "input={0}")
    public static Collection<Object[]> formats() {
        return Arrays.asList(new Object[][] {{"xls"}, {"xlsx"}});
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final String extension;

    public DuplicateBlankKeyTest(String extension) {
        this.extension = extension;
    }

    @Test
    public void reportCountsMeaningfulRowsAndKeepsBlankKeysInEncounterOrder() throws Exception {
        File input = mixedInput();
        File output = new File(temporary.getRoot(), "reported.xlsx");
        byte[] before = Files.readAllBytes(input.toPath());

        DuplicateColumnMerger.MergeReport report = new DuplicateColumnMerger().mergeWithReport(
                input.getPath(), output.getPath(), "id", Collections.singletonList("value"), null);

        assertEquals(13, report.getInputRows());
        assertEquals(10, report.getOutputRows());
        assertEquals(5, report.getBlankKeyRows());
        assertRows(output, expectedMixed(false));
        assertArrayEquals(before, Files.readAllBytes(input.toPath()));
    }

    @Test
    public void allBlankKeysRemainSeparateEvenWhenTheirOtherValuesAreIdentical() throws Exception {
        File input = temporary.newFile("blank-only." + extension);
        try (Workbook workbook = newWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            row(sheet, 0, "id", "value", "note");
            row(sheet, 1, null, "same", "same");
            row(sheet, 2, Empty.BLANK, "same", "same");
            row(sheet, 3, " \t ", "same", "same");
            row(sheet, 4, Empty.FORMULA, "same", "same");
            addEmptyRows(sheet, 5);
            save(workbook, input);
        }
        File output = new File(temporary.getRoot(), "blank-only-result.xlsx");

        DuplicateColumnMerger.MergeReport report = new DuplicateColumnMerger().mergeWithReport(
                input.getPath(), output.getPath(), "id", Arrays.asList("value", "note"), null);

        assertEquals(4, report.getInputRows());
        assertEquals(4, report.getOutputRows());
        assertEquals(4, report.getBlankKeyRows());
        assertRows(output, new String[][] {
                {"", "same", "same"}, {"", "same", "same"},
                {"", "same", "same"}, {"", "same", "same"}
        });
    }

    @Test
    public void bothLegacyEntryPointsPreserveBlankRowsAndTheirExistingMergeRules() throws Exception {
        File input = mixedInput();
        for (boolean mergeAllColumns : new boolean[] {false, true}) {
            File output = new File(temporary.getRoot(), "legacy-" + mergeAllColumns + ".xlsx");
            DuplicateColumnMerger merger = new DuplicateColumnMerger();
            if (mergeAllColumns) {
                merger.mergeByKeyColumn(input.getPath(), output.getPath(), "id");
            } else {
                merger.mergeByCustomerIdWithProgress(input.getPath(), output.getPath(), "id",
                        Collections.singletonList("value"), null);
            }
            assertRows(output, expectedMixed(mergeAllColumns));
        }
    }

    @Test
    public void blankKeyRowsKeepValuesOutsideHeadersAndExcelErrors() throws Exception {
        File input = temporary.newFile("sparse." + extension);
        try (Workbook workbook = newWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            row(sheet, 0, "id", "value");
            row(sheet, 1, "\u3000", null, "  备注  ");
            sheet.createRow(2).createCell(1).setCellErrorValue(FormulaError.DIV0.getCode());
            sheet.createRow(3).createCell(0).setCellErrorValue(FormulaError.NA.getCode());
            save(workbook, input);
        }
        File output = new File(temporary.getRoot(), "sparse-result.xlsx");
        DuplicateColumnMerger.MergeReport report = new DuplicateColumnMerger().mergeWithReport(
                input.getPath(), output.getPath(), "id", Collections.singletonList("value"), null);

        assertEquals(3, report.getInputRows());
        assertEquals(3, report.getOutputRows());
        assertEquals(2, report.getBlankKeyRows());
        assertRows(output, new String[][] {{"", "", "  备注  "}, {"", "#DIV/0!"}, {"#N/A", ""}});
    }

    private File mixedInput() throws Exception {
        File input = temporary.newFile("mixed." + extension);
        try (Workbook workbook = newWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            row(sheet, 0, "id", "value", "note");
            row(sheet, 1, null, "甲", "first blank");
            row(sheet, 2, "001", "a", null);
            row(sheet, 3, " \t ", "乙", "second blank");
            row(sheet, 4, "001", "b", "first nonempty");
            row(sheet, 5, MARKER_LIKE_KEY, "s1", "marker first");
            row(sheet, 6, Empty.FORMULA, "丙", "formula blank");
            row(sheet, 7, "001", "a", "later");
            row(sheet, 8, MARKER_LIKE_KEY, "s2", "marker second");
            row(sheet, 9, "1", "different", "distinct");
            row(sheet, 10, null, 0, null);
            row(sheet, 11, null, false, null);
            row(sheet, 12, 0, "zero-key", "valid");
            row(sheet, 13, false, "false-key", "valid");
            addEmptyRows(sheet, 14);
            save(workbook, input);
        }
        return input;
    }

    private String[][] expectedMixed(boolean mergeAllColumns) {
        return new String[][] {
                {"", "甲", "first blank"},
                {"001", "a|b", mergeAllColumns ? "first nonempty|later" : "first nonempty"},
                {"", "乙", "second blank"},
                {MARKER_LIKE_KEY, "s1|s2", mergeAllColumns ? "marker first|marker second" : "marker first"},
                {"", "丙", "formula blank"},
                {"1", "different", "distinct"},
                {"", "0", ""},
                {"", "false", ""},
                {"0", "zero-key", "valid"},
                {"false", "false-key", "valid"}
        };
    }

    private void addEmptyRows(Sheet sheet, int start) {
        sheet.createRow(start);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setWrapText(true);
        sheet.createRow(start + 1).createCell(1, CellType.BLANK).setCellStyle(style);
        row(sheet, start + 2, Empty.FORMULA, " \t ", Empty.BLANK);
    }

    private void row(Sheet sheet, int index, Object... values) {
        Row row = sheet.createRow(index);
        for (int column = 0; column < values.length; column++) {
            Object value = values[column];
            if (value == null) continue;
            Cell cell = row.createCell(column);
            if (value == Empty.FORMULA) cell.setCellFormula("\"\"");
            else if (value == Empty.BLANK) cell.setBlank();
            else if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
            else if (value instanceof Boolean) cell.setCellValue((Boolean) value);
            else cell.setCellValue((String) value);
        }
    }

    private Workbook newWorkbook() {
        return "xls".equals(extension) ? new HSSFWorkbook() : new XSSFWorkbook();
    }

    private void save(Workbook workbook, File file) throws Exception {
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
        try (FileOutputStream output = new FileOutputStream(file)) {
            workbook.write(output);
        }
    }

    private void assertRows(File file, String[][] expected) throws Exception {
        try (FileInputStream input = new FileInputStream(file); Workbook workbook = WorkbookFactory.create(input)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(expected.length + 1, sheet.getPhysicalNumberOfRows());
            assertEquals(expected.length, sheet.getLastRowNum());
            for (int row = 0; row < expected.length; row++) {
                assertNotNull("output row " + (row + 1), sheet.getRow(row + 1));
                for (int column = 0; column < expected[row].length; column++) {
                    Cell cell = sheet.getRow(row + 1).getCell(column);
                    String actual = cell == null || cell.getCellType() == CellType.BLANK
                            ? "" : cell.getStringCellValue();
                    assertEquals("row " + (row + 1) + ", column " + column, expected[row][column], actual);
                }
            }
        }
    }
}

package com.xzy;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.Assert.*;

@RunWith(Parameterized.class)
public class TableMergerFormulaTest {
    @Parameterized.Parameters(name = "{0}, multipleFiles={1}, xlsx={2}")
    public static Collection<Object[]> cases() {
        List<Object[]> cases = new ArrayList<>();
        for (TableMerger.MergeStrategy strategy : new TableMerger.MergeStrategy[]{
                TableMerger.MergeStrategy.HEADER_BASED, TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER}) {
            for (boolean multipleFiles : new boolean[]{false, true}) {
                for (boolean xlsx : new boolean[]{false, true}) {
                    cases.add(new Object[]{strategy, multipleFiles, xlsx});
                }
            }
        }
        return cases;
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final TableMerger.MergeStrategy strategy;
    private final boolean multipleFiles;
    private final boolean xlsx;

    public TableMergerFormulaTest(TableMerger.MergeStrategy strategy, boolean multipleFiles, boolean xlsx) {
        this.strategy = strategy;
        this.multipleFiles = multipleFiles;
        this.xlsx = xlsx;
    }

    @Test
    public void mergesSavedResultsWithTheirTypesAfterMovingRowsAndColumns() throws Exception {
        List<File> inputs = createInputs((sheet, index) -> {
            boolean reorder = hasHeaders() && index == 1;
            header(sheet, reorder ? new String[]{"result", "base", "text", "boolean", "error", "zero", "empty"}
                    : new String[]{"base", "result", "text", "boolean", "error", "zero", "empty"});
            if (hasHeaders()) sheet.getRow(0).getCell(4).setCellFormula("1/0");
            Row row = sheet.createRow(firstDataRow());
            int baseColumn = reorder ? 1 : 0;
            row.createCell(baseColumn).setCellValue(index == 0 ? 2 : 7);
            String reference = (reorder ? "B" : "A") + (firstDataRow() + 1);
            row.createCell(reorder ? 0 : 1).setCellFormula(reference + "*2");
            row.createCell(2).setCellFormula("\"text \"&" + reference);
            row.createCell(3).setCellFormula(reference + ">0");
            row.createCell(4).setCellFormula("1/0");
            row.createCell(5).setCellFormula(reference + "-" + reference);
            row.createCell(6).setCellFormula("IF(" + reference + ">0,\"\",\"unused\")");
        }, true);
        Map<File, byte[]> originals = snapshot(inputs);
        File output = output(inputs);

        merge(inputs, output);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Sheet sheet = dataSheet(workbook);
            if (hasHeaders()) {
                assertEquals("#DIV/0!", sheet.getRow(0).getCell(4).getStringCellValue());
            }
            int columnOffset = hasHeaders() ? 0 : 1;
            Row second = sheet.getRow(firstDataRow() + 1);
            assertNumeric(second.getCell(columnOffset), 7);
            assertNumeric(second.getCell(columnOffset + 1), 14);
            assertEquals(CellType.STRING, second.getCell(columnOffset + 2).getCellType());
            assertEquals("text 7", second.getCell(columnOffset + 2).getStringCellValue());
            assertEquals(CellType.BOOLEAN, second.getCell(columnOffset + 3).getCellType());
            assertTrue(second.getCell(columnOffset + 3).getBooleanCellValue());
            assertError(second.getCell(columnOffset + 4));
            assertNumeric(second.getCell(columnOffset + 5), 0);
            assertEquals(CellType.STRING, second.getCell(columnOffset + 6).getCellType());
            assertEquals("", second.getCell(columnOffset + 6).getStringCellValue());
            for (Row row : sheet) {
                for (Cell cell : row) {
                    assertNotEquals(CellType.FORMULA, cell.getCellType());
                }
            }
        }
        assertUnchanged(originals);
    }

    @Test
    public void usesSavedCacheWithoutRecalculatingChangedInputs() throws Exception {
        List<File> inputs = createInputs((sheet, index) -> {
            header(sheet, "base", "result");
            Row row = sheet.createRow(firstDataRow());
            row.createCell(0).setCellValue(7);
            Cell formula = row.createCell(1);
            formula.setCellFormula("A" + (firstDataRow() + 1) + "*2");
            sheet.getWorkbook().getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(formula);
            // 修改输入后保留已计算的 14，验证合并读取的是保存的结果，而非自行重算 18。
            row.getCell(0).setCellValue(9);
        }, false);
        File output = output(inputs);

        merge(inputs, output);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Row second = dataSheet(workbook).getRow(firstDataRow() + 1);
            int offset = hasHeaders() ? 0 : 1;
            assertNumeric(second.getCell(offset), 9);
            assertNumeric(second.getCell(offset + 1), 14);
        }
    }

    @Test
    public void retainsFormulaOnlyRowsIncludingErrorEmptyZeroAndFalseResults() throws Exception {
        List<File> inputs = createInputs((sheet, index) -> {
            header(sheet, "value");
            String[] formulas = {"1/0", "\"\"", "0", "FALSE()"};
            for (int i = 0; i < formulas.length; i++) {
                sheet.createRow(firstDataRow() + i).createCell(0).setCellFormula(formulas[i]);
            }
            // 有重算标志不代表已保存的 0 或空字符串缓存缺失。
            sheet.getWorkbook().setForceFormulaRecalculation(true);
        }, true);
        File output = output(inputs);

        merge(inputs, output);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Sheet sheet = dataSheet(workbook);
            assertEquals(firstDataRow() + 8, sheet.getPhysicalNumberOfRows());
            int column = hasHeaders() ? 0 : 1;
            for (int i = 0; i < 2; i++) {
                int row = firstDataRow() + i * 4;
                assertError(sheet.getRow(row).getCell(column));
                Cell empty = sheet.getRow(row + 1).getCell(column);
                assertEquals(CellType.STRING, empty.getCellType());
                assertEquals("", empty.getStringCellValue());
                assertNumeric(sheet.getRow(row + 2).getCell(column), 0);
                Cell bool = sheet.getRow(row + 3).getCell(column);
                assertEquals(CellType.BOOLEAN, bool.getCellType());
                assertFalse(bool.getBooleanCellValue());
            }
        }
    }

    @Test
    public void missingCacheAbortsBatchButStoredBiffZeroRemainsValid() throws Exception {
        if (!xlsx) {
            assertStoredBiffZeroIsAccepted(false);
            return;
        }
        for (String missingKind : new String[]{"n", "n-empty", "str", "b", "e"}) {
            assertMissingCacheProtectsOutput(missingKind, false);
        }
    }

    @Test
    public void sampleAndHeaderCachesRejectMissingXmlValueButAcceptStoredBiffZero() throws Exception {
        if (!xlsx) {
            assertStoredBiffZeroIsAccepted(true);
            return;
        }
        assertMissingCacheProtectsOutput("n", true);
    }

    private void assertStoredBiffZeroIsAccepted(boolean inFirstRow) throws Exception {
        List<File> inputs = createInputs((sheet, index) -> {
            header(sheet, "value");
            int formulaRow = inFirstRow ? 0 : firstDataRow();
            Row row = sheet.getRow(formulaRow);
            if (row == null) row = sheet.createRow(formulaRow);
            Cell formula = row.createCell(0);
            formula.setCellFormula("7*2");
            // BIFF 的固定缓存字段保存为 0，不能因公式或重算标志推断它缺失。
            formula.setCellValue(0);
            sheet.getWorkbook().setForceFormulaRecalculation(inFirstRow);
            if (hasHeaders() && inFirstRow) sheet.createRow(1).createCell(0).setCellValue(7);
        }, false);
        File output = output(inputs);

        merge(inputs, output);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Sheet sheet = dataSheet(workbook);
            int column = hasHeaders() ? 0 : 1;
            assertNumeric(sheet.getRow(firstDataRow()).getCell(column), hasHeaders() && inFirstRow ? 7 : 0);
            if (hasHeaders() && inFirstRow) {
                assertEquals("0", sheet.getRow(0).getCell(0).getStringCellValue());
            }
        }
    }

    private void assertMissingCacheProtectsOutput(String missingKind, boolean inFirstRow) throws Exception {
        int missingRow = inFirstRow ? 0 : firstDataRow() + 1;
        List<File> inputs = createInputs((sheet, index) -> {
            header(sheet, "value");
            sheet.createRow(firstDataRow()).createCell(0).setCellValue(7);
            if (index == 1) {
                Row row = sheet.getRow(missingRow);
                if (row == null) row = sheet.createRow(missingRow);
                XSSFCell formula = (XSSFCell) row.createCell(0);
                formula.setCellFormula("7*2");
                formula.getCTCell().setT("str".equals(missingKind) ? STCellType.STR
                        : "b".equals(missingKind) ? STCellType.B
                        : "e".equals(missingKind) ? STCellType.E : STCellType.N);
                if ("n-empty".equals(missingKind)) {
                    formula.getCTCell().setV("");
                } else if (formula.getCTCell().isSetV()) {
                    formula.getCTCell().unsetV();
                }
            }
        }, false);
        Map<File, byte[]> originals = snapshot(inputs);
        File existing = output(inputs);
        byte[] previous = "previous complete output".getBytes(StandardCharsets.UTF_8);
        Files.write(existing.toPath(), previous);
        File fresh = new File(existing.getParentFile(), "new.xlsx");
        Set<String> filesBefore = new HashSet<>(Arrays.asList(existing.getParentFile().list()));

        for (File output : new File[]{existing, fresh}) {
            IOException failure = assertThrows(IOException.class, () -> merge(inputs, output));
            assertTrue(failure.getMessage(), failure.getMessage().contains(inputs.get(inputs.size() - 1).getAbsolutePath()));
            assertTrue(failure.getMessage(), failure.getMessage().contains("Second"));
            assertTrue(failure.getMessage(), failure.getMessage().contains("A" + (missingRow + 1)));
            assertTrue(failure.getMessage(), failure.getMessage().contains("重新计算"));
            assertTrue(failure.getMessage(), failure.getMessage().contains("保存"));
        }

        assertArrayEquals(previous, Files.readAllBytes(existing.toPath()));
        assertFalse(fresh.exists());
        assertEquals(filesBefore, new HashSet<>(Arrays.asList(existing.getParentFile().list())));
        assertUnchanged(originals);
    }

    private List<File> createInputs(SheetPopulator populator, boolean evaluate) throws Exception {
        File directory = temporary.newFolder();
        List<File> inputs = new ArrayList<>();
        int fileCount = multipleFiles ? 2 : 1;
        for (int fileIndex = 0; fileIndex < fileCount; fileIndex++) {
            File file = new File(directory, "input" + fileIndex + (xlsx ? ".xlsx" : ".xls"));
            try (Workbook workbook = xlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
                int sheetCount = multipleFiles ? 1 : 2;
                for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                    int index = multipleFiles ? fileIndex : sheetIndex;
                    populator.populate(workbook.createSheet(index == 0 ? "First" : "Second"), index);
                }
                if (evaluate) workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
                try (FileOutputStream stream = new FileOutputStream(file)) {
                    workbook.write(stream);
                }
            }
            inputs.add(file);
        }
        return inputs;
    }

    private void header(Sheet sheet, String... names) {
        if (hasHeaders()) {
            Row row = sheet.createRow(0);
            for (int i = 0; i < names.length; i++) row.createCell(i).setCellValue(names[i]);
        }
    }

    private void merge(List<File> inputs, File output) throws IOException {
        TableMerger merger = new TableMerger();
        if (multipleFiles) merger.mergeFiles(inputs, output.getPath(), strategy, null);
        else merger.mergeSheetsInWorkbook(inputs.get(0).getPath(), output.getPath(), strategy, null);
    }

    private boolean hasHeaders() {
        return strategy == TableMerger.MergeStrategy.HEADER_BASED;
    }

    private int firstDataRow() {
        return hasHeaders() ? 1 : 0;
    }

    private Sheet dataSheet(Workbook workbook) {
        return workbook.getSheet(hasHeaders() ? "合并数据" : "总表");
    }

    private File output(List<File> inputs) {
        return new File(inputs.get(0).getParentFile(), "output.xlsx");
    }

    private static void assertNumeric(Cell cell, double expected) {
        assertEquals(CellType.NUMERIC, cell.getCellType());
        assertEquals(expected, cell.getNumericCellValue(), 0);
    }

    private static void assertError(Cell cell) {
        assertEquals(CellType.ERROR, cell.getCellType());
        assertEquals(FormulaError.DIV0.getCode(), cell.getErrorCellValue());
    }

    private static Map<File, byte[]> snapshot(List<File> files) throws IOException {
        Map<File, byte[]> result = new HashMap<>();
        for (File file : files) result.put(file, Files.readAllBytes(file.toPath()));
        return result;
    }

    private static void assertUnchanged(Map<File, byte[]> originals) throws IOException {
        for (Map.Entry<File, byte[]> entry : originals.entrySet()) {
            assertArrayEquals(entry.getValue(), Files.readAllBytes(entry.getKey().toPath()));
        }
    }

    private interface SheetPopulator {
        void populate(Sheet sheet, int index);
    }
}

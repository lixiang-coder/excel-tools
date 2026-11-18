package com.xzy;

import com.xzy.model.MappingRelations;
import com.xzy.util.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCalcMode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import static org.junit.Assert.*;

public class ExcelToolsColumnFormulaTest {
    private static final String FORMULA_SHEET = "计算数据";
    private static final byte[] OLD_OUTPUT = "existing output must survive".getBytes(StandardCharsets.UTF_8);
    private static final BiConsumer<Integer, String> NO_PROGRESS = (percent, message) -> { };
    private static final Consumer<String> NO_LOG = message -> { };

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void xlsxHeaderReaderPreservesOmittedMiddleCell() throws Exception {
        assertHeaderPositions("xlsx", false);
    }

    @Test
    public void xlsxHeaderReaderPreservesExplicitBlankMiddleCell() throws Exception {
        assertHeaderPositions("xlsx", true);
    }

    @Test
    public void xlsHeaderReaderPreservesOmittedMiddleCell() throws Exception {
        assertHeaderPositions("xls", false);
    }

    @Test
    public void xlsHeaderReaderPreservesExplicitBlankMiddleCell() throws Exception {
        assertHeaderPositions("xls", true);
    }

    @Test
    public void xlsxProcessorPreservesOmittedMiddleSourceAndTemplateColumns() throws Exception {
        assertProcessorColumnPositions("xlsx", false);
    }

    @Test
    public void xlsxProcessorPreservesExplicitBlankMiddleSourceAndTemplateColumns() throws Exception {
        assertProcessorColumnPositions("xlsx", true);
    }

    @Test
    public void xlsProcessorPreservesOmittedMiddleSourceAndTemplateColumns() throws Exception {
        assertProcessorColumnPositions("xls", false);
    }

    @Test
    public void xlsProcessorPreservesExplicitBlankMiddleSourceAndTemplateColumns() throws Exception {
        assertProcessorColumnPositions("xls", true);
    }

    @Test
    public void xlsxAppliedMappingPreservesOmittedMiddleColumn() throws Exception {
        assertAppliedColumnPositions("xlsx", false);
    }

    @Test
    public void xlsxAppliedMappingPreservesExplicitBlankMiddleColumn() throws Exception {
        assertAppliedColumnPositions("xlsx", true);
    }

    @Test
    public void xlsAppliedMappingPreservesOmittedMiddleColumn() throws Exception {
        assertAppliedColumnPositions("xls", false);
    }

    @Test
    public void xlsAppliedMappingPreservesExplicitBlankMiddleColumn() throws Exception {
        assertAppliedColumnPositions("xls", true);
    }

    @Test
    public void processorUsesOnlyCachedFormulaResultWithoutFormulaText() throws Exception {
        File source = formulaWorkbook("cached-three", "1+2", "3", STCellType.N);
        File template = simpleTemplate();
        File output = new File(temporary.getRoot(), "processed.xlsx");
        byte[] before = Files.readAllBytes(source.toPath());

        assertEquals(1, ExcelTools.exportProcessedExcel(source, template, output,
                Collections.singletonMap("Key", "Result"), NO_PROGRESS, NO_LOG));

        assertEquals("3", cellText(output, 1, 0));
        assertArrayEquals(before, Files.readAllBytes(source.toPath()));
    }

    @Test
    public void appliedMappingLooksUpCachedFormulaResult() throws Exception {
        assertFormulaLookup("cached-three", "1+2", "3", "three");
    }

    @Test
    public void appliedMappingAcceptsCachedZeroAsLookupValue() throws Exception {
        assertFormulaLookup("cached-zero", "1-1", "0", "zero");
    }

    @Test
    public void processorAcceptsCachedZero() throws Exception {
        assertFormulaOutput("cached-zero", "1-1", "0", STCellType.N, "0");
    }

    @Test
    public void processorAcceptsCachedEmptyString() throws Exception {
        assertFormulaOutput("cached-empty", "IF(1=1,\"\",\"unused\")", "", STCellType.STR, "");
    }

    @Test
    public void processorAcceptsCachedText() throws Exception {
        assertFormulaOutput("cached-text", "\"answer\"", "answer", STCellType.STR, "answer");
    }

    @Test
    public void processorMissingCacheReportsLocationAndRecalculationWithoutPublishing() throws Exception {
        assertMissingCacheProtected(false, null);
    }

    @Test
    public void appliedMappingMissingCacheReportsLocationAndRecalculationWithoutPublishing() throws Exception {
        assertMissingCacheProtected(true, null);
    }

    @Test
    public void processorEmptyNumericCacheReportsLocationAndRecalculationWithoutPublishing() throws Exception {
        assertMissingCacheProtected(false, "");
    }

    @Test
    public void appliedMappingEmptyNumericCacheReportsLocationAndRecalculationWithoutPublishing() throws Exception {
        assertMissingCacheProtected(true, "");
    }

    @Test
    public void appliedMappingKeepsUnusedFormulasAndRequestsRecalculationInTheOriginalMode() throws Exception {
        File source = temporary.newFile("manual.xlsx");
        try (XSSFWorkbook book = new XSSFWorkbook()) {
            Sheet sheet = book.createSheet("原表");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Key");
            header.createCell(1).setCellValue("Value");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("A");
            data.createCell(2).setCellFormula("LEN(B2)");
            book.getCTWorkbook().addNewCalcPr().setCalcMode(STCalcMode.MANUAL);
            write(book, source);
        }
        File output = new File(temporary.getRoot(), "mapped.xlsx");
        ExcelTools.exportAppliedMapping(source, output, "Key", "Value", new MappingRelations(), NO_PROGRESS, NO_LOG);
        try (Workbook result = WorkbookFactory.create(output, null, true)) {
            XSSFWorkbook book = (XSSFWorkbook) result;
            assertEquals("LEN(B2)", book.getSheetAt(0).getRow(1).getCell(2).getCellFormula());
            assertEquals(STCalcMode.MANUAL, book.getCTWorkbook().getCalcPr().getCalcMode());
            assertTrue(book.getCTWorkbook().getCalcPr().getFullCalcOnLoad());
        }
    }

    private void assertHeaderPositions(String extension, boolean explicitBlank) throws Exception {
        File input = gapWorkbook("headers", extension, explicitBlank, false);
        assertEquals(Arrays.asList("Key", "", "Value"), ExcelUtils.readHeadersOnly(input));
    }

    private void assertProcessorColumnPositions(String extension, boolean explicitBlank) throws Exception {
        File source = gapWorkbook("source", extension, explicitBlank, false);
        File template = gapWorkbook("template", extension, explicitBlank, true);
        File output = new File(temporary.getRoot(), "processed.xlsx");
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        byte[] templateBefore = Files.readAllBytes(template.toPath());

        assertEquals(1, ExcelTools.exportProcessedExcel(source, template, output,
                Collections.singletonMap("Value", "Result"), NO_PROGRESS, NO_LOG));

        assertEquals("Result", cellText(output, 0, 2));
        assertEquals("", cellText(output, 0, 1));
        assertEquals("from-physical-C", cellText(output, 1, 2));
        assertEquals("", cellText(output, 1, 1));
        assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
        assertArrayEquals(templateBefore, Files.readAllBytes(template.toPath()));
    }

    private void assertAppliedColumnPositions(String extension, boolean explicitBlank) throws Exception {
        File source = gapWorkbook("target", extension, explicitBlank, false);
        File output = new File(temporary.getRoot(), "applied.xlsx");
        byte[] before = Files.readAllBytes(source.toPath());

        MappingRelations relations = new MappingRelations();
        relations.add("lookup", "matched", source, 2);
        int[] stats = ExcelTools.exportAppliedMapping(source, output, "Key", "Value",
                relations, NO_PROGRESS, NO_LOG);

        assertArrayEquals(new int[] {1, 1, 0}, stats);
        assertEquals("Value", cellText(output, 0, 2));
        assertEquals("", cellText(output, 0, 1));
        assertEquals("matched", cellText(output, 1, 2));
        assertEquals("keep-physical-B", cellText(output, 1, 1));
        assertArrayEquals(before, Files.readAllBytes(source.toPath()));
    }

    private void assertFormulaLookup(String name, String formula, String cache, String mapped) throws Exception {
        File source = formulaWorkbook(name, formula, cache, STCellType.N);
        File output = new File(temporary.getRoot(), "applied.xlsx");

        MappingRelations relations = new MappingRelations();
        relations.add(cache, mapped, source, 2);
        int[] stats = ExcelTools.exportAppliedMapping(source, output, "Key", "Value",
                relations, NO_PROGRESS, NO_LOG);

        assertArrayEquals(new int[] {1, 1, 0}, stats);
        try (Workbook original = WorkbookFactory.create(source, null, true);
             Workbook result = WorkbookFactory.create(output, null, true)) {
            Cell originalLookup = original.getSheetAt(0).getRow(1).getCell(0);
            Cell outputLookup = result.getSheetAt(0).getRow(1).getCell(0);
            assertEquals(CellType.FORMULA, outputLookup.getCellType());
            assertEquals(originalLookup.getCellFormula(), outputLookup.getCellFormula());
            assertEquals(originalLookup.getCachedFormulaResultType(), outputLookup.getCachedFormulaResultType());
            assertEquals(originalLookup.getNumericCellValue(), outputLookup.getNumericCellValue(), 0);
            assertEquals(mapped, result.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        }
    }

    private void assertFormulaOutput(String name, String formula, String cache,
                                     STCellType.Enum cacheType, String expected) throws Exception {
        File source = formulaWorkbook(name, formula, cache, cacheType);
        File output = new File(temporary.getRoot(), "processed.xlsx");

        assertEquals(1, ExcelTools.exportProcessedExcel(source, simpleTemplate(), output,
                Collections.singletonMap("Key", "Result"), NO_PROGRESS, NO_LOG));
        assertEquals(expected, cellText(output, 1, 0));
    }

    private void assertMissingCacheProtected(boolean appliedMapping, String cache) throws Exception {
        File source = formulaWorkbook("missing-cache", "1+2", cache, STCellType.N);
        File template = simpleTemplate();
        File existing = temporary.newFile("existing.xlsx");
        Files.write(existing.toPath(), OLD_OUTPUT);
        File absent = new File(temporary.getRoot(), "absent.xlsx");
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        Set<String> before = directoryEntries();

        for (File output : Arrays.asList(existing, absent)) {
            Exception failure = assertThrows(Exception.class, () -> {
                if (appliedMapping) {
                    ExcelTools.exportAppliedMapping(source, output, "Key", "Value",
                            new MappingRelations(), NO_PROGRESS, NO_LOG);
                } else {
                    ExcelTools.exportProcessedExcel(source, template, output,
                            Collections.singletonMap("Key", "Result"), NO_PROGRESS, NO_LOG);
                }
            });
            String messages = exceptionMessages(failure);
            assertTrue(messages, messages.contains(FORMULA_SHEET));
            assertTrue(messages, messages.contains("A2"));
            assertTrue(messages, messages.contains("重算") || messages.contains("重新计算"));
            assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(existing.toPath()));
            assertFalse(absent.exists());
            assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
            assertEquals(before, directoryEntries());
        }
    }

    private File gapWorkbook(String name, String extension, boolean explicitBlank, boolean template) throws IOException {
        File file = new File(temporary.getRoot(), name + "." + extension);
        try (Workbook workbook = "xls".equals(extension) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(template ? "Unused" : "Key");
            if (explicitBlank) {
                header.createCell(1, CellType.BLANK);
            }
            header.createCell(2).setCellValue(template ? "Result" : "Value");
            if (!template) {
                Row data = sheet.createRow(1);
                data.createCell(0).setCellValue("lookup");
                data.createCell(1).setCellValue("keep-physical-B");
                data.createCell(2).setCellValue("from-physical-C");
            }
            write(workbook, file);
        }
        return file;
    }

    private File formulaWorkbook(String name, String formula, String cache, STCellType.Enum cacheType) throws IOException {
        File file = new File(temporary.getRoot(), name + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(FORMULA_SHEET);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Key");
            header.createCell(1).setCellValue("Value");
            Row data = sheet.createRow(1);
            XSSFCell formulaCell = (XSSFCell) data.createCell(0);
            formulaCell.setCellFormula(formula);
            formulaCell.getCTCell().setT(cacheType);
            if (cache == null) {
                if (formulaCell.getCTCell().isSetV()) {
                    formulaCell.getCTCell().unsetV();
                }
            } else {
                // Set the saved result explicitly; POI must not evaluate this fixture on write.
                formulaCell.getCTCell().setV(cache);
            }
            data.createCell(1).setCellValue("old-value");
            write(workbook, file);
        }
        return file;
    }

    private File simpleTemplate() throws IOException {
        File file = new File(temporary.getRoot(), "template.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Template").createRow(0).createCell(0).setCellValue("Result");
            write(workbook, file);
        }
        return file;
    }

    private void write(Workbook workbook, File file) throws IOException {
        try (FileOutputStream output = new FileOutputStream(file)) {
            workbook.write(output);
        }
    }

    private String cellText(File file, int rowIndex, int column) throws IOException {
        try (FileInputStream input = new FileInputStream(file); Workbook workbook = WorkbookFactory.create(input)) {
            Row row = workbook.getSheetAt(0).getRow(rowIndex);
            if (row == null) {
                return "";
            }
            Cell cell = row.getCell(column);
            return cell == null || cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
        }
    }

    private String exceptionMessages(Throwable failure) {
        StringBuilder messages = new StringBuilder();
        for (Throwable current = failure; current != null; current = current.getCause()) {
            messages.append(current.getMessage()).append('\n');
        }
        return messages.toString();
    }

    private Set<String> directoryEntries() {
        String[] names = temporary.getRoot().list();
        assertNotNull(names);
        return new HashSet<>(Arrays.asList(names));
    }
}

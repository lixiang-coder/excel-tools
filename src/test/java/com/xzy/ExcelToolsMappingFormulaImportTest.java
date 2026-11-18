package com.xzy;

import com.xzy.model.MappingRelations;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Collections;

import static org.junit.Assert.*;

public class ExcelToolsMappingFormulaImportTest {
    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void missingCacheInEitherSelectedColumnRejectsTheWholeImport() throws Exception {
        for (int column : new int[]{0, 1}) {
            for (String kind : new String[]{"n", "n-empty", "str", "b", "e"}) {
                File input = new File(temporary.newFolder(), "relations.xlsx");
                try (Workbook workbook = new XSSFWorkbook()) {
                    Sheet sheet = header(workbook);
                    pair(sheet, 1, "existing", "value");
                    Row row = pair(sheet, 6, "new", "new-value");
                    XSSFCell cell = (XSSFCell) row.getCell(column);
                    cell.setCellFormula("7*2");
                    cell.getCTCell().setT("str".equals(kind) ? STCellType.STR
                            : "b".equals(kind) ? STCellType.B
                            : "e".equals(kind) ? STCellType.E : STCellType.N);
                    if ("n-empty".equals(kind)) {
                        cell.getCTCell().setV("");
                    } else if (cell.getCTCell().isSetV()) {
                        cell.getCTCell().unsetV();
                    }
                    save(workbook, input);
                }
                byte[] originalBytes = Files.readAllBytes(input.toPath());
                File existingFile = new File(input.getParentFile(), "previous.xlsx");
                MappingRelations relations = new MappingRelations();
                relations.add("existing", "value", existingFile, 3);

                IOException error = assertThrows(IOException.class,
                        () -> ExcelTools.readMappingRelations(input, "From", "To", relations));

                assertTrue(error.getMessage(), error.getMessage().contains(input.getAbsolutePath()));
                assertTrue(error.getMessage(), error.getMessage().contains("关系页"));
                assertTrue(error.getMessage(), error.getMessage().contains(column == 0 ? "A7" : "B7"));
                assertTrue(error.getMessage(), error.getMessage().contains("重新计算"));
                assertTrue(error.getMessage(), error.getMessage().contains("保存"));
                assertEquals(1, relations.size());
                assertEquals("value", relations.resolve("existing").getValue());
                assertEquals(1, relations.resolve("existing").getCandidates().get("value").size());
                assertEquals(Collections.singleton(existingFile), relations.getSourceFiles());
                assertNull(relations.resolve("new").getValue());
                assertArrayEquals(originalBytes, Files.readAllBytes(input.toPath()));
            }
        }
    }

    @Test
    public void missingCacheInAnUnselectedColumnDoesNotBlockImport() throws Exception {
        File input = temporary.newFile("unselected.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = header(workbook);
            pair(sheet, 1, "A", "Alpha").createCell(2).setCellFormula("7*2");
            pair(sheet, 6, "B", "Beta").createCell(2).setCellFormula("1+1");
            save(workbook, input);
        }
        MappingRelations relations = new MappingRelations();

        assertEquals(2, ExcelTools.readMappingRelations(input, "From", "To", relations));

        assertEquals("Alpha", relations.resolve("A").getValue());
        assertEquals("Beta", relations.resolve("B").getValue());
        assertEquals(7, relations.resolve("B").getCandidates().get("Beta").get(0).getRowNumber());
    }

    @Test
    public void importsSavedValuesIncludingZeroAndIgnoresValidEmptyText() throws Exception {
        for (boolean xlsx : new boolean[]{false, true}) {
            File input = temporary.newFile(xlsx ? "valid.xlsx" : "valid.xls");
            try (Workbook workbook = xlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
                Sheet sheet = header(workbook);
                pair(sheet, 1, "number", "").getCell(1).setCellFormula("7*2");
                pair(sheet, 2, "zero", "").getCell(1).setCellFormula("0");
                pair(sheet, 3, "boolean", "").getCell(1).setCellFormula("TRUE()");
                pair(sheet, 4, "text", "").getCell(1).setCellFormula("\"saved text\"");
                pair(sheet, 5, "empty", "").getCell(1).setCellFormula("\"\"");
                Cell sourceFormula = pair(sheet, 6, "", "source-result").getCell(0);
                sourceFormula.setCellFormula("0");
                workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
                workbook.setForceFormulaRecalculation(true);
                save(workbook, input);
            }
            MappingRelations relations = new MappingRelations();

            assertEquals(5, ExcelTools.readMappingRelations(input, "From", "To", relations));

            assertEquals("14", relations.resolve("number").getValue());
            assertEquals("0", relations.resolve("zero").getValue());
            assertEquals("true", relations.resolve("boolean").getValue());
            assertEquals("saved text", relations.resolve("text").getValue());
            assertEquals("source-result", relations.resolve("0").getValue());
            assertNull(relations.resolve("empty").getValue());
        }
    }

    @Test
    public void importsTheSavedResultWithoutRecalculatingChangedReferences() throws Exception {
        for (boolean xlsx : new boolean[]{false, true}) {
            File input = temporary.newFile(xlsx ? "cached.xlsx" : "cached.xls");
            try (Workbook workbook = xlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
                Row row = pair(header(workbook), 1, "A", "");
                Cell base = row.createCell(2);
                base.setCellValue(7);
                row.getCell(1).setCellFormula("C2*2");
                workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(row.getCell(1));
                base.setCellValue(9);
                save(workbook, input);
            }
            MappingRelations relations = new MappingRelations();

            ExcelTools.readMappingRelations(input, "From", "To", relations);

            assertEquals("14", relations.resolve("A").getValue());
        }
    }

    private static Sheet header(Workbook workbook) {
        Sheet sheet = workbook.createSheet("关系页");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("From");
        row.createCell(1).setCellValue("To");
        row.createCell(2).setCellValue("Unused");
        return sheet;
    }

    private static Row pair(Sheet sheet, int rowNumber, String source, String target) {
        Row row = sheet.createRow(rowNumber);
        row.createCell(0).setCellValue(source);
        row.createCell(1).setCellValue(target);
        return row;
    }

    private static void save(Workbook workbook, File file) throws IOException {
        try (FileOutputStream output = new FileOutputStream(file)) {
            workbook.write(output);
        }
    }
}

package com.xzy;

import com.xzy.model.MappingRelations;
import com.xzy.util.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.*;

public class MappingConflictCellValueTest {
    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void preservesXlsValuesAndFormulaResults() throws Exception {
        assertOriginalValues("xls", new HSSFWorkbook());
    }

    @Test
    public void preservesXlsxValuesAndFormulaResultsWith1904Dates() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.getCTWorkbook().getWorkbookPr().setDate1904(true);
        assertOriginalValues("xlsx", workbook);
    }

    private void assertOriginalValues(String extension, Workbook inputWorkbook) throws Exception {
        File input = temporary.newFile("target." + extension);
        LocalDateTime timestamp = LocalDateTime.of(2025, 5, 6, 7, 30);
        try (Workbook workbook = inputWorkbook) {
            Sheet sheet = newTarget(workbook);
            fillCell(sheet, 1).setCellValue("  前后有空格  ");
            fillCell(sheet, 2).setCellValue(19.75);
            fillCell(sheet, 3).setCellValue(false);
            fillCell(sheet, 4).setCellErrorValue(FormulaError.NA.getCode());
            Cell date = fillCell(sheet, 5);
            date.setCellValue(timestamp);
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm"));
            date.setCellStyle(dateStyle);
            fillCell(sheet, 6).setCellFormula("4*3");
            fillCell(sheet, 7).setCellFormula("1=1");
            fillCell(sheet, 8).setCellFormula("1/0");
            fillCell(sheet, 9).setCellFormula("\"  缓存文本  \"");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            try (FileOutputStream stream = new FileOutputStream(input)) {
                workbook.write(stream);
            }
        }

        try (Workbook result = export(input, 9)) {
            Sheet sheet = result.getSheetAt(0);
            assertEquals("  前后有空格  ", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals(19.75, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.0);
            assertFalse(sheet.getRow(3).getCell(1).getBooleanCellValue());
            assertEquals(FormulaError.NA.getCode(), sheet.getRow(4).getCell(1).getErrorCellValue());
            Cell date = sheet.getRow(5).getCell(1);
            assertTrue(DateUtil.isCellDateFormatted(date));
            assertEquals(timestamp, date.getLocalDateTimeCellValue());
            assertEquals(CellType.NUMERIC, sheet.getRow(6).getCell(1).getCellType());
            assertEquals(12, sheet.getRow(6).getCell(1).getNumericCellValue(), 0.0);
            assertTrue(sheet.getRow(7).getCell(1).getBooleanCellValue());
            assertEquals(FormulaError.DIV0.getCode(), sheet.getRow(8).getCell(1).getErrorCellValue());
            assertEquals("  缓存文本  ", sheet.getRow(9).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void preservesSpacesInXlsxInlineString() throws Exception {
        File input = temporary.newFile("inline.xlsx");
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            fillCell(newTarget(workbook), 1).setCellValue("  首尾空格  ");
            try (FileOutputStream stream = new FileOutputStream(input)) {
                workbook.write(stream);
            } finally {
                workbook.dispose();
            }
        }
        try (Workbook result = export(input, 1)) {
            assertEquals("  首尾空格  ", result.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void numericTextAndDateHeadersKeepDistinctNamesWithoutChangingFallbackData() throws Exception {
        File input = temporary.newFile("numeric-headers.xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.getCTWorkbook().getWorkbookPr().setDate1904(true);
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(2025);
            header.createCell(1).setCellValue("2025.0");
            Cell date = header.createCell(2);
            date.setCellValue(1 + 59.6 / 86400);
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
            date.setCellStyle(dateStyle);
            for (int index = 1; index <= 2; index++) {
                Row row = sheet.createRow(index);
                row.createCell(0).setCellValue(6 + index);
                row.createCell(2).setCellValue(6 + index);
            }
            try (FileOutputStream stream = new FileOutputStream(input)) { workbook.write(stream); }
        }
        List<String> names = ExcelUtils.readNamedHeaders(input);
        assertEquals(Arrays.asList("2025", "2025.0", "1904/01/02 00:01"), names);
        MappingRelations relations = new MappingRelations();
        relations.add("7", "七", input, 2);
        File output = new File(temporary.getRoot(), "filled.xlsx");
        for (int key : new int[]{0, 2}) {
            ExcelTools.exportAppliedMapping(input, output, names.get(key), names.get(1), relations,
                    (percent, message) -> { }, message -> { });
            try (Workbook result = WorkbookFactory.create(output, null, true)) {
                Sheet sheet = result.getSheetAt(0);
                assertEquals("七", sheet.getRow(1).getCell(1).getStringCellValue());
                assertEquals("8.0", sheet.getRow(2).getCell(1).getStringCellValue());
                assertEquals(7, sheet.getRow(1).getCell(0).getNumericCellValue(), 0);
            }
        }
    }

    @Test
    public void usesSameLookupKeysForXlsxNativeNumbersBooleansAndTextAsRelationImport() throws Exception {
        File relationFile = temporary.newFile("typed-relations.xls");
        File targetFile = temporary.newFile("typed-target.xlsx");
        try (Workbook relations = new HSSFWorkbook(); Workbook target = new XSSFWorkbook()) {
            Sheet relationSheet = newTarget(relations);
            Sheet targetSheet = newTarget(target);
            for (int group = 0; group < 3; group++) {
                Row targetRow = targetSheet.createRow(group + 1);
                setLookup(targetRow.createCell(0), group);
                targetRow.createCell(1).setCellValue("保留");
                for (int candidate = 0; candidate < 2; candidate++) {
                    Row relation = relationSheet.createRow(1 + group * 2 + candidate);
                    setLookup(relation.createCell(0), group);
                    relation.createCell(1).setCellValue("候选" + group + candidate);
                }
            }
            try (FileOutputStream stream = new FileOutputStream(relationFile)) { relations.write(stream); }
            try (FileOutputStream stream = new FileOutputStream(targetFile)) { target.write(stream); }
        }
        MappingRelations relations = new MappingRelations();
        ExcelTools.readMappingRelations(relationFile, "Key", "Value", relations);
        File output = new File(temporary.getRoot(), "typed-output.xlsx");
        int[] stats = ExcelTools.exportAppliedMapping(targetFile, output, "Key", "Value", relations,
                (percent, message) -> { }, message -> { });
        assertArrayEquals(new int[]{3, 0, 3}, stats);
        try (Workbook result = WorkbookFactory.create(output, null, true)) {
            Sheet conflicts = result.getSheet("映射冲突");
            for (int group = 0; group < 3; group++) {
                assertEquals("保留", result.getSheetAt(0).getRow(group + 1).getCell(1).getStringCellValue());
                assertEquals("候选" + group + "0", conflicts.getRow(1 + group * 2).getCell(3).getStringCellValue());
            }
        }
    }

    private void setLookup(Cell cell, int group) {
        switch (group) {
            case 0: cell.setCellValue(1.0); break;
            case 1: cell.setCellValue(true); break;
            default: cell.setCellValue("  编号  ");
        }
    }

    @Test
    public void uniqueMatchAndFallbackReplaceXlsFillFormulaWithText() throws Exception {
        File input = temporary.newFile("formulas.xls");
        try (Workbook book = new HSSFWorkbook()) {
            Sheet sheet = newTarget(book);
            fillCell(sheet, 1).setCellFormula("1+1");
            fillCell(sheet, 2).setCellFormula("1+2");
            sheet.getRow(2).getCell(0).setCellValue("unknown");
            book.getCreationHelper().createFormulaEvaluator().evaluateAll();
            try (FileOutputStream stream = new FileOutputStream(input)) { book.write(stream); }
        }
        MappingRelations relations = new MappingRelations();
        relations.add("A", "唯一", input, 2);
        File output = new File(temporary.getRoot(), "filled.xlsx");
        ExcelTools.exportAppliedMapping(input, output, "Key", "Value", relations,
                (percent, message) -> { }, message -> { });
        try (Workbook book = WorkbookFactory.create(output, null, true)) {
            Cell matched = book.getSheetAt(0).getRow(1).getCell(1);
            Cell fallback = book.getSheetAt(0).getRow(2).getCell(1);
            assertEquals(CellType.STRING, matched.getCellType());
            assertEquals("唯一", matched.getStringCellValue());
            assertEquals(CellType.STRING, fallback.getCellType());
            assertEquals("unknown", fallback.getStringCellValue());
        }
    }

    @Test
    public void replacesExistingInlineTextForBothMatchedAndUnmatchedRows() throws Exception {
        File input = temporary.newFile("inline-fill.xlsx");
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            try {
                Sheet sheet = newTarget(workbook);
                fillCell(sheet, 1).setCellValue("旧值一");
                fillCell(sheet, 2).setCellValue("旧值二");
                sheet.getRow(2).getCell(0).setCellValue("missing");
                try (FileOutputStream stream = new FileOutputStream(input)) { workbook.write(stream); }
            } finally {
                workbook.dispose();
            }
        }
        MappingRelations relations = new MappingRelations();
        relations.add("A", "甲", input, 2);
        File output = new File(temporary.getRoot(), "filled.xlsx");
        int[] stats = ExcelTools.exportAppliedMapping(input, output, "Key", "Value", relations,
                (percent, message) -> { }, message -> { });
        assertArrayEquals(new int[]{2, 1, 0}, stats);
        try (Workbook saved = WorkbookFactory.create(output, null, true)) {
            Sheet sheet = saved.getSheetAt(0);
            assertEquals("甲", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("missing", sheet.getRow(2).getCell(1).getStringCellValue());
        }
    }

    private Sheet newTarget(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Key");
        header.createCell(1).setCellValue("Value");
        return sheet;
    }

    private Cell fillCell(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        row.createCell(0).setCellValue("A");
        return row.createCell(1);
    }

    private Workbook export(File input, int conflictCount) throws Exception {
        MappingRelations relations = new MappingRelations();
        File relationFile = new File(temporary.getRoot(), "relation.xlsx");
        relations.add("A", "甲", relationFile, 2);
        relations.add("A", "乙", relationFile, 3);
        File output = new File(temporary.getRoot(), "result.xlsx");
        int[] stats = ExcelTools.exportAppliedMapping(input, output, "Key", "Value",
                relations, (percent, message) -> { }, message -> { });
        assertArrayEquals(new int[]{conflictCount, 0, conflictCount}, stats);
        return WorkbookFactory.create(output, null, true);
    }
}

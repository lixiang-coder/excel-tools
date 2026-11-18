package com.xzy.util;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.record.DateWindow1904Record;
import org.apache.poi.hssf.record.CalcModeRecord;
import org.apache.poi.hssf.record.CalcCountRecord;
import org.apache.poi.hssf.record.DeltaRecord;
import org.apache.poi.hssf.record.IterationRecord;
import org.apache.poi.hssf.record.PrecisionRecord;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCalcMode;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import static org.junit.Assert.*;

public class LegacyWorkbookConverterTest {
    @Rule public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void preservesSheetsCoordinatesDimensionsStylesAndNames() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            HSSFSheet first = source.createSheet("主数据");
            source.createSheet("其他表").createRow(0).createCell(0).setCellValue("保留");
            source.createSheet("隐藏表");
            source.setSheetVisibility(1, SheetVisibility.HIDDEN);
            source.setSheetVisibility(2, SheetVisibility.VERY_HIDDEN);
            first.setDefaultColumnWidth(12);
            first.setDefaultRowHeightInPoints(19);
            first.setColumnWidth(8, 6000);
            first.setColumnHidden(9, true);
            first.createFreezePane(2, 3);
            first.addMergedRegion(new CellRangeAddress(4, 5, 1, 2));
            first.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
            source.setPrintArea(0, 0, 10, 0, 15);
            first.getHeader().setCenter("标题");
            first.getPrintSetup().setLandscape(true);

            source.getCustomPalette().setColorAtIndex((short) 40, (byte) 12, (byte) 34, (byte) 56);
            HSSFFont font = source.createFont();
            font.setFontName("宋体");
            font.setFontHeightInPoints((short) 13);
            font.setBold(true);
            font.setColor((short) 40);
            HSSFCellStyle style = source.createCellStyle();
            style.setFont(font);
            style.setFillForegroundColor((short) 40);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor((short) 40);
            style.setDataFormat(source.createDataFormat().getFormat("0.000"));
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.RIGHT);
            first.setDefaultColumnStyle(8, style);
            HSSFRow row = first.createRow(7);
            row.setHeightInPoints(31);
            row.setZeroHeight(true);
            row.setRowStyle(style);
            HSSFCell cell = row.createCell(8);
            cell.setCellValue(42.5);
            cell.setCellStyle(style);
            row.createCell(10).setCellStyle(style);
            first.groupRow(7, 8);
            first.setRowGroupCollapsed(7, true);
            first.groupColumn(8, 9);
            first.setColumnGroupCollapsed(8, true);

            Name global = source.createName();
            global.setNameName("数据范围");
            global.setRefersToFormula("'主数据'!$I$8");
            Name local = source.createName();
            local.setSheetIndex(1);
            local.setNameName("局部范围");
            local.setRefersToFormula("'其他表'!$A$1");

            // 先保存再读入真实 BIFF，避免只检验内存中新建对象的行为。
            try (HSSFWorkbook loaded = saved(source); XSSFWorkbook converted = converted(loaded)) {
                assertEquals(3, converted.getNumberOfSheets());
                assertEquals("主数据", converted.getSheetName(0));
                assertEquals("其他表", converted.getSheetName(1));
                assertEquals("保留", converted.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
                assertEquals(SheetVisibility.HIDDEN, converted.getSheetVisibility(1));
                assertEquals(SheetVisibility.VERY_HIDDEN, converted.getSheetVisibility(2));
                XSSFSheet sheet = converted.getSheetAt(0);
                assertNull(sheet.getRow(0));
                assertEquals(6000, sheet.getColumnWidth(8));
                assertTrue(sheet.isColumnHidden(9));
                assertEquals(31, sheet.getRow(7).getHeightInPoints(), 0);
                assertTrue(sheet.getRow(7).getZeroHeight());
                assertEquals(1, sheet.getRow(7).getOutlineLevel());
                assertEquals(1, sheet.getColumnOutlineLevel(8));
                assertEquals("B5:C6", sheet.getMergedRegion(0).formatAsString());
                assertEquals(3, sheet.getPaneInformation().getHorizontalSplitPosition());
                assertEquals("1:2", sheet.getRepeatingRows().formatAsString());
                assertTrue(converted.getPrintArea(0).contains("$A$1:$K$16"));
                assertEquals("标题", sheet.getHeader().getCenter());
                assertTrue(sheet.getPrintSetup().getLandscape());
                XSSFCell copied = sheet.getRow(7).getCell(8);
                assertEquals(42.5, copied.getNumericCellValue(), 0);
                XSSFCellStyle copiedStyle = copied.getCellStyle();
                assertEquals("0.000", copiedStyle.getDataFormatString());
                assertEquals(HorizontalAlignment.RIGHT, copiedStyle.getAlignment());
                assertTrue(copiedStyle.getWrapText());
                assertArrayEquals(new byte[]{12, 34, 56}, copiedStyle.getFillForegroundXSSFColor().getRGB());
                assertArrayEquals(new byte[]{12, 34, 56}, copiedStyle.getBottomBorderXSSFColor().getRGB());
                assertArrayEquals(new byte[]{12, 34, 56}, copiedStyle.getFont().getXSSFColor().getRGB());
                assertTrue(copiedStyle.getFont().getBold());
                assertEquals("宋体", copiedStyle.getFont().getFontName());
                assertEquals(CellType.BLANK, sheet.getRow(7).getCell(10).getCellType());
                assertEquals(copiedStyle.getIndex(), sheet.getColumnStyle(8).getIndex());
                assertEquals("主数据!$I$8", converted.getName("数据范围").getRefersToFormula());
                assertEquals(1, converted.getName("局部范围").getSheetIndex());
            }
        }
    }

    @Test
    public void preservesFormulaTextAndAllCachedResultTypesWithoutRecalculation() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            HSSFSheet sheet = source.createSheet("数据");
            source.createSheet("参数").createRow(0).createCell(0).setCellValue(7);
            HSSFRow row = sheet.createRow(8);
            String[] formulas = {"'参数'!A1*2", "\"缓存文字\"", "1=1", "1/0", "1-1", "\"\""};
            for (int i = 0; i < formulas.length; i++) row.createCell(i).setCellFormula(formulas[i]);
            source.getCreationHelper().createFormulaEvaluator().evaluateAll();
            // 保留已保存缓存 14，即使输入后来改为 9，也不能擅自重新计算成 18。
            source.getSheetAt(1).getRow(0).getCell(0).setCellValue(9);
            try (HSSFWorkbook loaded = saved(source); XSSFWorkbook converted = converted(loaded)) {
                XSSFRow copied = converted.getSheetAt(0).getRow(8);
                for (int i = 0; i < formulas.length; i++) {
                    assertEquals(CellType.FORMULA, copied.getCell(i).getCellType());
                    assertEquals(loaded.getSheetAt(0).getRow(8).getCell(i).getCellFormula(),
                            copied.getCell(i).getCellFormula());
                }
                assertEquals(14, copied.getCell(0).getNumericCellValue(), 0);
                assertEquals("缓存文字", copied.getCell(1).getStringCellValue());
                assertTrue(copied.getCell(2).getBooleanCellValue());
                assertEquals(FormulaError.DIV0.getCode(), copied.getCell(3).getErrorCellValue());
                assertEquals(0, copied.getCell(4).getNumericCellValue(), 0);
                assertEquals("", copied.getCell(5).getStringCellValue());
            }
        }
    }

    @Test
    public void preserves1904DatesRichTextAndHyperlinks() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            ((DateWindow1904Record) source.getWorkbook().findFirstRecordBySid(DateWindow1904Record.sid))
                    .setWindowing((short) 1);
            HSSFRow row = source.createSheet("内容").createRow(0);
            HSSFCell date = row.createCell(0);
            date.setCellValue(45000.5);
            HSSFCellStyle dateStyle = source.createCellStyle();
            dateStyle.setDataFormat(source.createDataFormat().getFormat("yyyy-mm-dd hh:mm"));
            date.setCellStyle(dateStyle);
            HSSFFont bold = source.createFont();
            bold.setBold(true);
            HSSFRichTextString text = new HSSFRichTextString("前后文字");
            text.applyFont(0, 2, bold);
            row.createCell(1).setCellValue(text);
            HSSFHyperlink link = source.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress("https://example.org/data");
            row.getCell(1).setHyperlink(link);
            link.setLastRow(3);
            link.setLastColumn(2);
            try (HSSFWorkbook loaded = saved(source); XSSFWorkbook converted = converted(loaded)) {
                assertTrue(converted.isDate1904());
                XSSFRow copied = converted.getSheetAt(0).getRow(0);
                assertEquals(45000.5, copied.getCell(0).getNumericCellValue(), 0);
                assertEquals(loaded.getSheetAt(0).getRow(0).getCell(0).getLocalDateTimeCellValue(),
                        copied.getCell(0).getLocalDateTimeCellValue());
                assertEquals("前后文字", copied.getCell(1).getStringCellValue());
                assertTrue(copied.getCell(1).getRichStringCellValue().getFontAtIndex(0).getBold());
                assertEquals("https://example.org/data", copied.getCell(1).getHyperlink().getAddress());
                assertEquals(3, copied.getCell(1).getHyperlink().getLastRow());
                assertEquals(2, copied.getCell(1).getHyperlink().getLastColumn());
            }
        }
    }

    @Test
    public void preservesStandardColorsAndManualCalculationMode() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            HSSFSheet sheet = source.createSheet("手动计算");
            HSSFCell cell = sheet.createRow(0).createCell(0);
            cell.setCellFormula("7*2");
            source.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
            HSSFCellStyle style = source.createCellStyle();
            HSSFFont font = source.createFont();
            font.setColor(IndexedColors.DARK_BLUE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBottomBorderColor(IndexedColors.RED.getIndex());
            cell.setCellStyle(style);
            ((CalcModeRecord) sheet.getSheet().findFirstRecordBySid(CalcModeRecord.sid))
                    .setCalcMode(CalcModeRecord.MANUAL);
            ((PrecisionRecord) source.getWorkbook().findFirstRecordBySid(PrecisionRecord.sid))
                    .setFullPrecision(false);
            ((IterationRecord) sheet.getSheet().findFirstRecordBySid(IterationRecord.sid)).setIteration(true);
            ((CalcCountRecord) sheet.getSheet().findFirstRecordBySid(CalcCountRecord.sid)).setIterations((short) 250);
            int deltaIndex = sheet.getSheet().findFirstRecordLocBySid(DeltaRecord.sid);
            sheet.getSheet().getRecords().set(deltaIndex, new DeltaRecord(0.00001));
            source.setForceFormulaRecalculation(true);
            try (XSSFWorkbook converted = converted(source)) {
                assertTrue(converted.getCTWorkbook().getCalcPr().getFullCalcOnLoad());
                assertEquals(STCalcMode.MANUAL, converted.getCTWorkbook().getCalcPr().getCalcMode());
                assertFalse(converted.getCTWorkbook().getCalcPr().getFullPrecision());
                assertTrue(converted.getCTWorkbook().getCalcPr().getIterate());
                assertEquals(250, converted.getCTWorkbook().getCalcPr().getIterateCount());
                assertEquals(0.00001, converted.getCTWorkbook().getCalcPr().getIterateDelta(), 0);
                XSSFCell copy = converted.getSheetAt(0).getRow(0).getCell(0);
                assertEquals(14, copy.getNumericCellValue(), 0);
                assertEquals(IndexedColors.LIGHT_YELLOW.getIndex(), copy.getCellStyle().getFillForegroundColor());
                assertEquals(IndexedColors.RED.getIndex(), copy.getCellStyle().getBottomBorderColor());
                assertEquals(IndexedColors.DARK_BLUE.getIndex(), copy.getCellStyle().getFont().getColor());
            }
        }
    }

    @Test
    public void copiesNewNamesAfterInMemoryPrintAreaChanges() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            HSSFSheet sheet = source.createSheet("数据");
            sheet.setRepeatingColumns(new CellRangeAddress(-1, -1, 0, 1));
            source.setPrintArea(0, 0, 10, 0, 15);
            Name named = source.createName();
            named.setNameName("数值");
            named.setRefersToFormula("数据!$C$1");
            sheet.createRow(0).createCell(2).setCellValue(7);
            sheet.getRow(0).createCell(3).setCellFormula("数值*2");
            source.getCreationHelper().createFormulaEvaluator().evaluateAll();
            try (XSSFWorkbook converted = converted(source)) {
                assertNotNull(converted.getName("数值"));
                assertEquals("A:B", converted.getSheetAt(0).getRepeatingColumns().formatAsString());
                assertTrue(converted.getPrintArea(0).contains("$A$1:$K$16"));
                assertEquals(14, converted.getSheetAt(0).getRow(0).getCell(3).getNumericCellValue(), 0);
            }
        }
    }

    @Test
    public void rejectsDrawingsWithoutPublishingExistingOrNewOutput() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            HSSFSheet sheet = source.createSheet("含图形");
            sheet.createDrawingPatriarch().createSimpleShape(new HSSFClientAnchor())
                    .setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            File input = temporary.newFile("source.xls");
            try (FileOutputStream out = new FileOutputStream(input)) { source.write(out); }
            byte[] original = Files.readAllBytes(input.toPath());
            for (boolean exists : new boolean[]{true, false}) {
                File output = new File(temporary.getRoot(), exists ? "existing.xlsx" : "new.xlsx");
                byte[] previous = "已有输出".getBytes(StandardCharsets.UTF_8);
                if (exists) Files.write(output.toPath(), previous);
                try (SafeExport export = SafeExport.begin(output, input)) {
                    try (FileOutputStream stream = new FileOutputStream(export.temporaryFile())) {
                        LegacyWorkbookConverter.writeXlsx(source, stream);
                    }
                    export.commit();
                    fail("应拒绝无法保真的图形转换");
                } catch (IOException e) {
                    assertTrue(e.getMessage().contains("同格式 .xls"));
                    assertTrue(e.getMessage().contains("含图形"));
                }
                if (exists) assertArrayEquals(previous, Files.readAllBytes(output.toPath()));
                else assertFalse(output.exists());
                assertArrayEquals(original, Files.readAllBytes(input.toPath()));
            }
        }
    }

    @Test
    public void rejectsMacrosBeforeWritingAnyBytes() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook()) {
            source.createSheet("数据");
            ByteArrayOutputStream bytes = new ByteArrayOutputStream();
            source.write(bytes);
            try (POIFSFileSystem serialized = new POIFSFileSystem(new ByteArrayInputStream(bytes.toByteArray()))) {
                serialized.getRoot().createDirectory("_VBA_PROJECT_CUR");
                try (HSSFWorkbook macroWorkbook = new HSSFWorkbook(serialized)) {
                    assertRejected(macroWorkbook, "宏");
                }
            }
        }
    }

    @Test
    public void rejectsValidationConditionalFormattingArraysAndProtection() throws Exception {
        for (int feature = 0; feature < 4; feature++) {
            try (HSSFWorkbook source = new HSSFWorkbook()) {
                HSSFSheet sheet = source.createSheet("特殊表");
                sheet.createRow(0).createCell(0).setCellValue(1);
                if (feature == 0) {
                    DataValidationHelper helper = sheet.getDataValidationHelper();
                    sheet.addValidationData(helper.createValidation(helper.createExplicitListConstraint(
                            new String[]{"甲", "乙"}), new CellRangeAddressList(0, 4, 0, 0)));
                } else if (feature == 1) {
                    SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();
                    formatting.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(0, 2, 0, 0)},
                            formatting.createConditionalFormattingRule("A1>0"));
                } else if (feature == 2) {
                    sheet.setArrayFormula("A1*2", new CellRangeAddress(1, 2, 1, 1));
                } else {
                    sheet.protectSheet("password");
                }
                assertRejected(source, "特殊表");
            }
        }
    }

    @Test
    public void rejectsExternalWorkbookReferences() throws Exception {
        try (HSSFWorkbook source = new HSSFWorkbook(); HSSFWorkbook external = new HSSFWorkbook()) {
            external.createSheet("外部").createRow(0).createCell(0).setCellValue(7);
            source.linkExternalWorkbook("external.xls", external);
            source.createSheet("数据").createRow(0).createCell(0).setCellFormula("'[external.xls]外部'!A1");
            assertRejected(source, "外部");
        }
    }

    private static void assertRejected(HSSFWorkbook source, String message) throws Exception {
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        try {
            LegacyWorkbookConverter.writeXlsx(source, stream);
            fail("应拒绝无法保真的转换");
        } catch (IOException e) {
            assertTrue(e.getMessage(), e.getMessage().contains(message));
            assertTrue(e.getMessage(), e.getMessage().contains("同格式 .xls"));
            assertEquals(0, stream.size());
        }
    }

    private static HSSFWorkbook saved(HSSFWorkbook source) throws IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        source.write(bytes);
        return new HSSFWorkbook(new ByteArrayInputStream(bytes.toByteArray()));
    }

    private static XSSFWorkbook converted(HSSFWorkbook source) throws IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        LegacyWorkbookConverter.writeXlsx(source, bytes);
        // 使用 OOXML 专用读取器，不能将扩展名为 .xlsx 的 BIFF 文件误判为成功。
        return new XSSFWorkbook(new ByteArrayInputStream(bytes.toByteArray()));
    }
}

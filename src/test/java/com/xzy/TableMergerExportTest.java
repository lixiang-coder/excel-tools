package com.xzy;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

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
public class TableMergerExportTest {
    private static final byte[] OLD_OUTPUT = "已有结果不能因导出失败丢失".getBytes(StandardCharsets.UTF_8);

    @Parameterized.Parameters(name = "{0}, multipleFiles={1}, headerData={2}")
    public static Collection<Object[]> cases() {
        List<Object[]> cases = new ArrayList<>();
        for (boolean multipleFiles : new boolean[] {false, true}) {
            cases.add(new Object[] {TableMerger.MergeStrategy.HEADER_BASED, multipleFiles, true});
            cases.add(new Object[] {TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER, multipleFiles, false});
            cases.add(new Object[] {TableMerger.MergeStrategy.AUTO, multipleFiles, true});
            cases.add(new Object[] {TableMerger.MergeStrategy.AUTO, multipleFiles, false});
        }
        return cases;
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final TableMerger.MergeStrategy strategy;
    private final boolean multipleFiles;
    private final boolean headerData;
    private List<File> inputs;
    private Map<File, byte[]> originalInputs;

    public TableMergerExportTest(TableMerger.MergeStrategy strategy, boolean multipleFiles, boolean headerData) {
        this.strategy = strategy;
        this.multipleFiles = multipleFiles;
        this.headerData = headerData;
    }

    @Before
    public void createInputs() throws IOException {
        inputs = new ArrayList<>();
        inputs.add(writeWorkbook("first.xlsx", multipleFiles ? new int[] {101} : new int[] {101, 102}));
        if (multipleFiles) {
            inputs.add(writeWorkbook("second.xlsx", new int[] {102}));
        }
        originalInputs = new HashMap<>();
        for (File input : inputs) {
            originalInputs.put(input, Files.readAllBytes(input.toPath()));
        }
    }

    @Test
    public void rejectsEveryInputAsOutputIncludingNonFirstInputAndPathAlias() throws Exception {
        Set<String> before = directoryEntries();
        for (File input : inputs) {
            File alias = new File(input.getParentFile(), "." + File.separator + input.getName());
            assertThrows(IOException.class, () -> merge(alias, null));
            assertInputsUnchanged();
            assertEquals(before, directoryEntries());
        }
    }

    @Test
    public void preservesExistingOutputWhenProcessingFailsMidway() throws Exception {
        File output = temporary.newFile("existing.xlsx");
        Files.write(output.toPath(), OLD_OUTPUT);
        Set<String> before = directoryEntries();

        assertMidwayFailure(output);

        assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
        assertInputsUnchanged();
        assertEquals(before, directoryEntries());
    }

    @Test
    public void leavesNoNewOutputWhenProcessingFailsMidway() throws Exception {
        File output = new File(temporary.getRoot(), "new.xlsx");
        Set<String> before = directoryEntries();

        assertMidwayFailure(output);

        assertFalse(output.exists());
        assertInputsUnchanged();
        assertEquals(before, directoryEntries());
    }

    @Test
    public void publishesCompleteDataAndReportForNewAndExistingOutput() throws Exception {
        File output = new File(temporary.getRoot(), "result.xlsx");
        merge(output, null);
        assertCompleteResult(output);

        Files.write(output.toPath(), OLD_OUTPUT);
        Set<String> before = directoryEntries();
        merge(output, null);

        assertCompleteResult(output);
        assertInputsUnchanged();
        assertEquals(before, directoryEntries());
    }

    @Test
    public void keepsFailureReportWhileMergingReadableInputs() throws Exception {
        if (multipleFiles) {
            inputs.add(1, new File(temporary.getRoot(), "missing.xlsx"));
        } else {
            File input = writeWorkbook("with-empty-sheet.xlsx", new int[] {101, 102}, true);
            inputs = Arrays.asList(input);
            originalInputs.put(input, Files.readAllBytes(input.toPath()));
        }
        File output = new File(temporary.getRoot(), "partial-success.xlsx");

        merge(output, null);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Sheet data = workbook.getSheet(headerData ? "合并数据" : "总表");
            assertEquals(headerData ? 3 : 2, data.getPhysicalNumberOfRows());
            Sheet report = workbook.getSheet(headerData ? "合并报告" : "报告");
            if (headerData) {
                assertEquals("成功", report.getRow(1).getCell(2).getStringCellValue());
                assertEquals("失败", report.getRow(2).getCell(2).getStringCellValue());
                assertEquals("成功", report.getRow(3).getCell(2).getStringCellValue());
                assertEquals("成功 2 / 失败 1", report.getRow(4).getCell(2).getStringCellValue());
            } else {
                assertEquals("成功", report.getRow(4).getCell(3).getStringCellValue());
                assertTrue(report.getRow(5).getCell(3).getStringCellValue().startsWith("失败"));
                assertEquals("成功", report.getRow(6).getCell(3).getStringCellValue());
                assertTrue(report.getRow(1).getCell(1).getStringCellValue().contains("成功合并 2 个工作表，失败 1 个"));
            }
        }
        assertInputsUnchanged();
    }

    @Test
    public void allInputsFailedDoesNotReplaceOrCreateOutput() throws Exception {
        File empty = writeEmptyWorkbook("empty.xlsx");
        inputs = Arrays.asList(empty);
        originalInputs.put(empty, Files.readAllBytes(empty.toPath()));
        File existing = temporary.newFile("existing.xlsx");
        Files.write(existing.toPath(), OLD_OUTPUT);
        File fresh = new File(temporary.getRoot(), "new.xlsx");
        Set<String> before = directoryEntries();

        IOException existingFailure = assertThrows(IOException.class, () -> merge(existing, null));
        IOException newFailure = assertThrows(IOException.class, () -> merge(fresh, null));

        String expectedReason = strategy == TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER
                ? "没有数据" : "未找到有效表头";
        assertTrue(existingFailure.getMessage().contains(expectedReason));
        assertTrue(newFailure.getMessage().contains(expectedReason));
        assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(existing.toPath()));
        assertFalse(fresh.exists());
        assertEquals(before, directoryEntries());
        assertInputsUnchanged();
    }

    @Test
    public void validHeaderWithoutDataIsStillASuccessfulInput() throws Exception {
        File input = temporary.newFile("header-only.xlsx");
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream stream = new FileOutputStream(input)) {
            Row header = workbook.createSheet("Sheet1").createRow(0);
            header.createCell(0).setCellValue("编号");
            header.createCell(1).setCellValue("姓名");
            workbook.write(stream);
        }
        inputs = Arrays.asList(input);
        originalInputs.put(input, Files.readAllBytes(input.toPath()));
        File output = new File(temporary.getRoot(), "header-only-result.xlsx");

        merge(output, null);

        try (Workbook workbook = WorkbookFactory.create(output)) {
            if (strategy == TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER) {
                assertEquals(1, workbook.getSheet("总表").getPhysicalNumberOfRows());
                assertEquals("成功", workbook.getSheet("报告").getRow(4).getCell(3).getStringCellValue());
            } else {
                assertEquals(1, workbook.getSheet("合并数据").getPhysicalNumberOfRows());
                Sheet report = workbook.getSheet("合并报告");
                assertEquals("成功", report.getRow(1).getCell(2).getStringCellValue());
                assertEquals(0, report.getRow(1).getCell(4).getNumericCellValue(), 0);
            }
        }
        assertInputsUnchanged();
    }

    private void assertMidwayFailure(File output) {
        int[] notifications = {0};
        Exception failure = assertThrows(Exception.class, () -> merge(output, (current, total) -> {
            notifications[0]++;
            throw new IllegalStateException("injected processing failure");
        }));
        assertEquals(1, notifications[0]);
        assertTrue(failure.getMessage().contains("injected processing failure"));
    }

    private void merge(File output, TableMerger.ProgressListener listener) throws IOException {
        TableMerger merger = new TableMerger();
        if (multipleFiles) {
            merger.mergeFiles(inputs, output.getPath(), strategy, listener);
        } else {
            merger.mergeSheetsInWorkbook(inputs.get(0).getPath(), output.getPath(), strategy, listener);
        }
    }

    private File writeWorkbook(String name, int[] ids) throws IOException {
        return writeWorkbook(name, ids, false);
    }

    private File writeWorkbook(String name, int[] ids, boolean includeEmptySheet) throws IOException {
        File file = temporary.newFile(name);
        try (Workbook workbook = new XSSFWorkbook()) {
            for (int i = 0; i < ids.length; i++) {
                Sheet sheet = workbook.createSheet("Sheet" + (i + 1));
                if (headerData) {
                    Row header = sheet.createRow(0);
                    header.createCell(0).setCellValue("编号");
                    header.createCell(1).setCellValue("姓名");
                }
                Row row = sheet.createRow(headerData ? 1 : 0);
                row.createCell(0).setCellValue(ids[i]);
                row.createCell(1).setCellValue(ids[i] == 101 ? "甲" : "乙");
                if (includeEmptySheet && i == 0) {
                    workbook.createSheet("Empty");
                }
            }
            try (FileOutputStream stream = new FileOutputStream(file)) {
                workbook.write(stream);
            }
        }
        return file;
    }

    private File writeEmptyWorkbook(String name) throws IOException {
        File file = temporary.newFile(name);
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream stream = new FileOutputStream(file)) {
            workbook.createSheet("Empty");
            workbook.write(stream);
        }
        return file;
    }

    private void assertCompleteResult(File output) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(output)) {
            assertEquals(2, workbook.getNumberOfSheets());
            Sheet data = workbook.getSheet(headerData ? "合并数据" : "总表");
            assertNotNull(data);
            assertEquals(headerData ? 3 : 2, data.getPhysicalNumberOfRows());
            int firstDataRow = headerData ? 1 : 0;
            int firstDataColumn = headerData ? 0 : 1;
            assertEquals(101, data.getRow(firstDataRow).getCell(firstDataColumn).getNumericCellValue(), 0);
            assertEquals("甲", data.getRow(firstDataRow).getCell(firstDataColumn + 1).getStringCellValue());
            assertEquals(102, data.getRow(firstDataRow + 1).getCell(firstDataColumn).getNumericCellValue(), 0);
            assertEquals("乙", data.getRow(firstDataRow + 1).getCell(firstDataColumn + 1).getStringCellValue());
            Sheet report = workbook.getSheet(headerData ? "合并报告" : "报告");
            assertNotNull(report);
            if (headerData) {
                assertEquals("编号", data.getRow(0).getCell(0).getStringCellValue());
                assertEquals("成功 2 / 失败 0", report.getRow(3).getCell(2).getStringCellValue());
                assertEquals(2, report.getRow(3).getCell(4).getNumericCellValue(), 0);
            } else {
                assertEquals("[1]Sheet1", data.getRow(0).getCell(0).getStringCellValue());
                assertEquals(multipleFiles ? "[2]Sheet1" : "[1]Sheet2", data.getRow(1).getCell(0).getStringCellValue());
                assertTrue(report.getRow(1).getCell(1).getStringCellValue().contains("成功合并 2 个工作表，失败 0 个，共 2 行数据"));
            }
        }
    }

    private void assertInputsUnchanged() throws IOException {
        for (Map.Entry<File, byte[]> original : originalInputs.entrySet()) {
            assertArrayEquals(original.getValue(), Files.readAllBytes(original.getKey().toPath()));
        }
    }

    private Set<String> directoryEntries() {
        return new HashSet<>(Arrays.asList(temporary.getRoot().list()));
    }
}

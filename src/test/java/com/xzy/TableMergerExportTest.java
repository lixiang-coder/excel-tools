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

        TableMerger.MergeSummary summary = merge(output, null);
        assertEquals(2, summary.getSuccessCount());
        assertEquals(0, summary.getPartialCount());
        assertEquals(1, summary.getFailureCount());

        try (Workbook workbook = WorkbookFactory.create(output)) {
            Sheet data = workbook.getSheet(headerData ? "合并数据" : "总表");
            assertEquals(headerData ? 3 : 2, data.getPhysicalNumberOfRows());
            Sheet report = workbook.getSheet(headerData ? "合并报告" : "报告");
            if (headerData) {
                assertEquals("成功", report.getRow(1).getCell(2).getStringCellValue());
                assertEquals("失败", report.getRow(2).getCell(2).getStringCellValue());
                assertEquals("成功", report.getRow(3).getCell(2).getStringCellValue());
                assertEquals(summary.formatSummary(), report.getRow(4).getCell(2).getStringCellValue());
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
    public void findsHeadersBelowBlankRowsAndReturnsSeparatePartialAndFailureCounts() throws Exception {
        if (multipleFiles) inputs.add(temporary.newFile("empty.xlsx"));
        setTableData(new Object[][] {null, {"  "}, {"编号", "姓名"}, {101, "甲"}},
                new Object[][] {{}, null, {"编号"}, {102}}, new Object[][] {{}});
        File output = new File(temporary.getRoot(), "leading-blank-rows.xlsx");

        TableMerger.MergeSummary summary = merge(output, null);

        assertEquals(1, summary.getFailureCount());
        try (Workbook workbook = WorkbookFactory.create(output)) {
            if (strategy == TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER) {
                assertEquals(2, summary.getSuccessCount());
                assertEquals(0, summary.getPartialCount());
                assertEquals(4, workbook.getSheet("总表").getPhysicalNumberOfRows());
            } else {
                assertEquals(1, summary.getSuccessCount());
                assertEquals(1, summary.getPartialCount());
                Sheet data = workbook.getSheet("合并数据");
                assertEquals(3, data.getPhysicalNumberOfRows());
                assertEquals("编号", data.getRow(0).getCell(0).getStringCellValue());
                assertEquals(101, data.getRow(1).getCell(0).getNumericCellValue(), 0);
                assertEquals(102, data.getRow(2).getCell(0).getNumericCellValue(), 0);
                Sheet report = workbook.getSheet("合并报告");
                assertEquals("部分成功", report.getRow(2).getCell(2).getStringCellValue());
                assertEquals("失败", report.getRow(3).getCell(2).getStringCellValue());
                assertEquals("成功 1 / 部分成功 1 / 失败 1", summary.formatSummary());
                assertEquals(summary.formatSummary(), report.getRow(4).getCell(2).getStringCellValue());
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
    public void singleTextRowNeedsAnExplicitHeaderChoice() throws Exception {
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

        if (strategy == TableMerger.MergeStrategy.AUTO) {
            assertThrows(TableMerger.AmbiguousHeadersException.class, () -> merge(output, null));
            assertFalse(output.exists());
            assertInputsUnchanged();
            return;
        }
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

    @Test
    public void textOnlyListsRequireAChoiceAndDoNotCreateFalseSuccessfulRows() throws Exception {
        setTableData(new String[][] {{"张三", "北京"}, {"李四", "上海"}},
                new String[][] {{"王五", "深圳"}, {"赵六", "成都"}});
        File output = temporary.newFile("text-lists.xlsx");
        Files.write(output.toPath(), OLD_OUTPUT);

        if (strategy == TableMerger.MergeStrategy.AUTO) {
            IOException failure = assertThrows(TableMerger.AmbiguousHeadersException.class, () -> merge(output, null));
            assertTrue(failure.getMessage().contains("无表头合并"));
            assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
        } else {
            merge(output, null);
            try (Workbook workbook = WorkbookFactory.create(output)) {
                if (strategy == TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER) {
                    Sheet data = workbook.getSheet("总表");
                    assertEquals(4, data.getPhysicalNumberOfRows());
                    String[] names = {"张三", "李四", "王五", "赵六"};
                    for (int i = 0; i < names.length; i++) {
                        assertEquals(names[i], data.getRow(i).getCell(1).getStringCellValue());
                    }
                } else {
                    assertEquals(2, workbook.getSheet("合并数据").getPhysicalNumberOfRows());
                    Sheet report = workbook.getSheet("合并报告");
                    assertEquals("失败", report.getRow(2).getCell(2).getStringCellValue());
                    assertEquals(0, report.getRow(2).getCell(4).getNumericCellValue(), 0);
                    assertTrue(report.getRow(2).getCell(5).getStringCellValue().contains("没有共同列"));
                    assertEquals("成功 1 / 部分成功 0 / 失败 1", report.getRow(3).getCell(2).getStringCellValue());
                }
            }
        }
        assertInputsUnchanged();
    }

    @Test
    public void identicalFirstRecordsAreInsufficientEvidenceOfHeaders() throws Exception {
        File output = temporary.newFile("identical-first-rows.xlsx");
        for (Object firstValue : new Object[] {"张三", 7}) {
            setTableData(new Object[][] {{firstValue, "北京"}, {"李四", "上海"}},
                    new Object[][] {{firstValue, "北京"}, {"赵六", "成都"}});
            Files.write(output.toPath(), OLD_OUTPUT);
            if (strategy == TableMerger.MergeStrategy.AUTO && firstValue instanceof String) {
                assertThrows(TableMerger.AmbiguousHeadersException.class, () -> merge(output, null));
                assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
            } else {
                merge(output, null);
                try (Workbook workbook = WorkbookFactory.create(output)) {
                    if (strategy == TableMerger.MergeStrategy.HEADER_BASED) {
                        assertEquals(3, workbook.getSheet("合并数据").getPhysicalNumberOfRows());
                    } else {
                        Sheet data = workbook.getSheet("总表");
                        assertEquals(4, data.getPhysicalNumberOfRows());
                        assertEquals("北京", data.getRow(0).getCell(2).getStringCellValue());
                        assertEquals("北京", data.getRow(2).getCell(2).getStringCellValue());
                    }
                }
            }
            assertInputsUnchanged();
        }
    }

    @Test
    public void duplicateNormalizedHeadersInALaterTableProtectTheWholeExport() throws Exception {
        setTableData(new Object[][] {{"ID", "Amount"}, {1, "10"}},
                new Object[][] {{"ID", "Amount", "A mount"}, {2, "20", "30"}});
        File output = temporary.newFile("duplicate-headers.xlsx");
        Files.write(output.toPath(), OLD_OUTPUT);

        if (strategy == TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER) {
            merge(output, null);
            try (Workbook workbook = WorkbookFactory.create(output)) {
                Row row = workbook.getSheet("总表").getRow(3);
                assertEquals("20", row.getCell(2).getStringCellValue());
                assertEquals("30", row.getCell(3).getStringCellValue());
            }
        } else {
            IOException failure = assertThrows(IOException.class, () -> merge(output, null));
            String message = failure.getMessage();
            assertTrue(message, message.contains("重复表头") && message.contains("amount"));
            assertTrue(message, message.contains("B 列") && message.contains("C 列"));
            assertTrue(message, message.contains("Table2"));
            assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
        }
        assertInputsUnchanged();
    }

    private void setTableData(Object[][]... tables) throws IOException {
        for (int fileIndex = 0; fileIndex < inputs.size(); fileIndex++) {
            File input = inputs.get(fileIndex);
            try (Workbook workbook = new XSSFWorkbook(); FileOutputStream out = new FileOutputStream(input)) {
                int sheetCount = multipleFiles ? 1 : tables.length;
                for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                    int index = multipleFiles ? fileIndex : sheetIndex;
                    Sheet sheet = workbook.createSheet("Table" + (index + 1));
                    for (int r = 0; r < tables[index].length; r++) {
                        Object[] values = tables[index][r];
                        if (values == null) continue;
                        Row row = sheet.createRow(r);
                        for (int c = 0; c < values.length; c++) {
                            if (values[c] instanceof Number) {
                                row.createCell(c).setCellValue(((Number) values[c]).doubleValue());
                            } else {
                                row.createCell(c).setCellValue((String) values[c]);
                            }
                        }
                    }
                }
                workbook.write(out);
            }
            originalInputs.put(input, Files.readAllBytes(input.toPath()));
        }
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

    private TableMerger.MergeSummary merge(File output, TableMerger.ProgressListener listener) throws IOException {
        TableMerger merger = new TableMerger();
        if (multipleFiles) {
            return merger.mergeFilesWithSummary(inputs, output.getPath(), strategy, listener);
        }
        return merger.mergeSheetsInWorkbookWithSummary(inputs.get(0).getPath(), output.getPath(), strategy, listener);
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
                assertEquals("成功 2 / 部分成功 0 / 失败 0", report.getRow(3).getCell(2).getStringCellValue());
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

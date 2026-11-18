package com.xzy;

import com.xzy.model.MappingRelations;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
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
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

@RunWith(Parameterized.class)
public class ExcelToolsMappingConflictTest {
    private static final String[] RELATION_HEADERS = {"From", "To"};
    private static final String[] TARGET_HEADERS = {"Key", "Fill", "Marker"};
    private static final List<String> REPORT_HEADERS = Arrays.asList(
            "目标行号", "查找值", "匹配方向", "候选值", "关系文件", "关系行号");
    private static final DataFormatter FORMATTER = new DataFormatter();

    @Parameterized.Parameters(name = "target={0}")
    public static Collection<Object[]> formats() {
        return Arrays.asList(new Object[][]{{"xls"}, {"xlsx"}});
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final String extension;

    public ExcelToolsMappingConflictTest(String extension) {
        this.extension = extension;
    }

    @Test
    public void forwardConflictWithinOneFileKeepsOriginalAndReportsActualRowNumbers() throws Exception {
        File relation = workbook("relations", RELATION_HEADERS,
                row(2, "A", "Alpha"), row(5, "A", "Amber"), row(8, "Safe", "Result"));
        File target = workbook("target", TARGET_HEADERS,
                row(4, "A", "keep-forward", "conflict"),
                row(7, "Safe", "old", "matched"), row(11, "missing", "old", "unmatched"));
        Map<File, byte[]> originals = originalBytes(target, relation);
        File output = output();

        int[] stats = apply(target, output, load(relation));

        assertArrayEquals(new int[]{3, 1, 1}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("keep-forward", text(findRow(result.getSheetAt(0), 0, "A"), 1));
            assertEquals("conflict", text(findRow(result.getSheetAt(0), 0, "A"), 2));
            assertEquals("Result", text(findRow(result.getSheetAt(0), 0, "Safe"), 1));
            assertEquals("missing", text(findRow(result.getSheetAt(0), 0, "missing"), 1));
            assertConflictReport(result, Arrays.asList(
                    conflict(4, "A", "正向", "Alpha", relation, 2),
                    conflict(4, "A", "正向", "Amber", relation, 5)));
        }
        assertUnchanged(originals);
    }

    @Test
    public void reverseConflictAcrossXlsAndXlsxFilesRetainsEveryCandidateAndOrigin() throws Exception {
        File first = workbook("first-relations", "xls", RELATION_HEADERS, row(3, "OldA", "Common"));
        File second = workbook("second-relations", "xlsx", RELATION_HEADERS, row(7, "OldB", "Common"));
        File target = workbook("target", TARGET_HEADERS,
                row(2, "OldA", "old", "forward"), row(9, "Common", "keep-reverse", "reverse"));
        File output = output();

        int[] stats = apply(target, output, load(first, second));

        assertArrayEquals(new int[]{2, 1, 1}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("Common", text(findRow(result.getSheetAt(0), 0, "OldA"), 1));
            assertEquals("keep-reverse", text(findRow(result.getSheetAt(0), 0, "Common"), 1));
            assertConflictReport(result, Arrays.asList(
                    conflict(9, "Common", "反向", "OldA", first, 3),
                    conflict(9, "Common", "反向", "OldB", second, 7)));
        }
    }

    @Test
    public void forwardConflictsAlsoAccumulateAcrossFilesAndReverseConflictsWithinAFile() throws Exception {
        File first = workbook("first-relations", RELATION_HEADERS,
                row(2, "A", "One"), row(4, "ReverseA", "Shared"), row(7, "ReverseB", "Shared"));
        File second = workbook("second-relations", RELATION_HEADERS, row(5, "A", "Two"));
        File target = workbook("target", TARGET_HEADERS,
                row(3, "A", "keep-forward", "forward"), row(8, "Shared", "keep-reverse", "reverse"));
        File output = output();

        int[] stats = apply(target, output, load(first, second));

        assertArrayEquals(new int[]{2, 0, 2}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("keep-forward", text(findRow(result.getSheetAt(0), 0, "A"), 1));
            assertEquals("keep-reverse", text(findRow(result.getSheetAt(0), 0, "Shared"), 1));
            assertConflictReport(result, Arrays.asList(
                    conflict(3, "A", "正向", "One", first, 2),
                    conflict(3, "A", "正向", "Two", second, 5),
                    conflict(8, "Shared", "反向", "ReverseA", first, 4),
                    conflict(8, "Shared", "反向", "ReverseB", first, 7)));
        }
    }

    @Test
    public void repeatedIdenticalRelationsRemainUniqueAcrossRowsFilesAndImports() throws Exception {
        File first = workbook("first-relations", RELATION_HEADERS,
                row(2, "A", "Alpha"), row(6, "A", "Alpha"));
        File second = workbook("second-relations", RELATION_HEADERS, row(3, "A", "Alpha"));
        MappingRelations relations = load(first, first, second);
        File target = workbook("target", TARGET_HEADERS,
                row(2, "A", "old", "forward"), row(5, "Alpha", "old", "reverse"),
                row(8, "unmatched", "old", "fallback"));
        File output = output();

        int[] stats = apply(target, output, relations);

        assertArrayEquals(new int[]{3, 2, 0}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("Alpha", text(findRow(result.getSheetAt(0), 0, "A"), 1));
            assertEquals("A", text(findRow(result.getSheetAt(0), 0, "Alpha"), 1));
            assertEquals("unmatched", text(findRow(result.getSheetAt(0), 0, "unmatched"), 1));
            assertNoConflictReport(result);
        }
    }

    @Test
    public void uniqueForwardMatchTakesPriorityOverConflictingReverseMatches() throws Exception {
        File relation = workbook("relations", RELATION_HEADERS,
                row(2, "Pivot", "Chosen"), row(3, "BackA", "Pivot"), row(4, "BackB", "Pivot"));
        File target = workbook("target", TARGET_HEADERS, row(6, "Pivot", "old", "priority"));
        File output = output();

        int[] stats = apply(target, output, load(relation));

        assertArrayEquals(new int[]{1, 1, 0}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("Chosen", text(findRow(result.getSheetAt(0), 0, "Pivot"), 1));
            assertNoConflictReport(result);
        }
    }

    @Test
    public void forwardConflictCannotFallBackToAUniqueReverseMatch() throws Exception {
        File relation = workbook("relations", RELATION_HEADERS,
                row(2, "Pivot", "ChoiceA"), row(3, "Pivot", "ChoiceB"), row(4, "Back", "Pivot"));
        File target = workbook("target", TARGET_HEADERS, row(5, "Pivot", "keep-original", "priority"));
        File output = output();

        int[] stats = apply(target, output, load(relation));

        assertArrayEquals(new int[]{1, 0, 1}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("keep-original", text(findRow(result.getSheetAt(0), 0, "Pivot"), 1));
            assertConflictReport(result, Arrays.asList(
                    conflict(5, "Pivot", "正向", "ChoiceA", relation, 2),
                    conflict(5, "Pivot", "正向", "ChoiceB", relation, 3)));
        }
    }

    @Test
    public void conflictsPreserveBlankFillCellsAndCountAffectedTargetRows() throws Exception {
        File relation = workbook("relations", RELATION_HEADERS,
                row(2, "A", "Alpha"), row(3, "A", "Amber"));
        File target = workbook("target", TARGET_HEADERS,
                row(2, "A", null, "missing-cell"), row(6, "A", "", "empty-string"),
                row(9, "A", "keep-original", "existing-value"));
        File output = output();

        int[] stats = apply(target, output, load(relation));

        assertArrayEquals(new int[]{3, 0, 3}, stats);
        try (FileInputStream input = new FileInputStream(output); Workbook result = WorkbookFactory.create(input)) {
            assertEquals("", text(findRow(result.getSheetAt(0), 2, "missing-cell"), 1));
            assertEquals("", text(findRow(result.getSheetAt(0), 2, "empty-string"), 1));
            assertEquals("keep-original", text(findRow(result.getSheetAt(0), 2, "existing-value"), 1));
            List<List<String>> expected = new ArrayList<>();
            for (int targetRow : new int[]{2, 6, 9}) {
                expected.add(conflict(targetRow, "A", "正向", "Alpha", relation, 2));
                expected.add(conflict(targetRow, "A", "正向", "Amber", relation, 3));
            }
            assertConflictReport(result, expected);
        }
    }

    @Test
    public void newMappingModelProtectsTargetAndEveryRelationSourceFromOverwrite() throws Exception {
        File first = workbook("first-relations", RELATION_HEADERS, row(2, "A", "Alpha"));
        File second = workbook("second-relations", RELATION_HEADERS, row(4, "A", "Amber"));
        File target = workbook("target", TARGET_HEADERS, row(2, "A", "keep", "target"));
        MappingRelations relations = load(first, second);
        Map<File, byte[]> originals = originalBytes(target, first, second);

        for (File protectedFile : originals.keySet()) {
            File alias = new File(protectedFile.getParentFile(), "." + File.separator + protectedFile.getName());
            for (File destination : Arrays.asList(protectedFile, alias)) {
                assertThrows(IOException.class, () -> apply(target, destination, relations));
                assertUnchanged(originals);
            }
        }
    }

    private MappingRelations load(File... files) throws IOException {
        MappingRelations relations = new MappingRelations();
        for (File file : files) {
            ExcelTools.readMappingRelations(file, "From", "To", relations);
        }
        return relations;
    }

    private static int[] apply(File target, File output, MappingRelations relations) throws Exception {
        return ExcelTools.exportAppliedMapping(target, output, "Key", "Fill", relations,
                (percent, message) -> { }, message -> { });
    }

    private File output() {
        return new File(temporary.getRoot(), "output.xlsx");
    }

    private File workbook(String name, String[] headers, FixtureRow... rows) throws IOException {
        return workbook(name, extension, headers, rows);
    }

    private File workbook(String name, String format, String[] headers, FixtureRow... rows) throws IOException {
        File file = new File(temporary.getRoot(), name + "." + format);
        try (Workbook workbook = "xls".equals(format) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            writeRow(sheet.createRow(0), headers);
            for (FixtureRow fixture : rows) {
                writeRow(sheet.createRow(fixture.excelRow - 1), fixture.values);
            }
            try (FileOutputStream output = new FileOutputStream(file)) {
                workbook.write(output);
            }
        }
        return file;
    }

    private static void writeRow(Row row, String[] values) {
        for (int column = 0; column < values.length; column++) {
            if (values[column] != null) row.createCell(column).setCellValue(values[column]);
        }
    }

    private static FixtureRow row(int excelRow, String... values) {
        return new FixtureRow(excelRow, values);
    }

    private static Row findRow(Sheet sheet, int column, String value) {
        for (int row = 1; row <= sheet.getLastRowNum(); row++) {
            if (value.equals(text(sheet.getRow(row), column))) return sheet.getRow(row);
        }
        fail("Missing output row with column " + column + " = " + value);
        return null;
    }

    private static String text(Row row, int column) {
        return row == null ? "" : FORMATTER.formatCellValue(row.getCell(column));
    }

    private static List<String> conflict(int targetRow, String lookup, String direction, String candidate,
                                         File relation, int relationRow) {
        return Arrays.asList(String.valueOf(targetRow), lookup, direction, candidate,
                relation.getAbsolutePath(), String.valueOf(relationRow));
    }

    private static void assertConflictReport(Workbook result, List<List<String>> expected) {
        assertEquals(2, result.getNumberOfSheets());
        Sheet report = result.getSheet("映射冲突");
        assertNotNull("A conflict report must be present", report);
        List<String> headers = new ArrayList<>();
        for (int column = 0; column < REPORT_HEADERS.size(); column++) headers.add(text(report.getRow(0), column));
        assertEquals(REPORT_HEADERS, headers);
        List<List<String>> actual = new ArrayList<>();
        for (int row = 1; row <= report.getLastRowNum(); row++) {
            List<String> values = new ArrayList<>();
            for (int column = 0; column < REPORT_HEADERS.size(); column++) values.add(text(report.getRow(row), column));
            actual.add(values);
        }
        assertEquals("Report must neither omit nor repeat candidate origins", expected.size(), actual.size());
        assertEquals(new HashSet<>(expected), new HashSet<>(actual));
    }

    private static void assertNoConflictReport(Workbook result) {
        assertEquals(1, result.getNumberOfSheets());
        assertNull(result.getSheet("映射冲突"));
    }

    private static Map<File, byte[]> originalBytes(File... files) throws IOException {
        Map<File, byte[]> originals = new LinkedHashMap<>();
        for (File file : files) originals.put(file, Files.readAllBytes(file.toPath()));
        return originals;
    }

    private static void assertUnchanged(Map<File, byte[]> originals) throws IOException {
        for (Map.Entry<File, byte[]> entry : originals.entrySet()) {
            assertArrayEquals(entry.getKey().getName(), entry.getValue(), Files.readAllBytes(entry.getKey().toPath()));
        }
    }

    private static final class FixtureRow {
        private final int excelRow;
        private final String[] values;

        private FixtureRow(int excelRow, String[] values) {
            this.excelRow = excelRow;
            this.values = values;
        }
    }
}

package com.xzy;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import static org.junit.Assert.*;

public class CrossAndDuplicateExportTest {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void crossSheetRejectsOutputMatchingEitherInput() throws Exception {
        Fixture fixture = fixture();
        for (File output : new File[]{fixture.source, fixture.target}) {
            assertFailedExportPreservesFiles(fixture, output,
                    () -> fill(fixture, output, "id"));
        }
    }

    @Test
    public void crossSheetFailurePreservesExistingOutput() throws Exception {
        Fixture fixture = fixture();
        File output = existingOutput(fixture);
        assertFailedExportPreservesFiles(fixture, output,
                () -> fill(fixture, output, "missing"));
    }

    @Test
    public void crossSheetFailureDoesNotLeaveNewOutput() throws Exception {
        Fixture fixture = fixture();
        File output = new File(fixture.directory, "output.xlsx");
        assertFailedExportPreservesFiles(fixture, output,
                () -> fill(fixture, output, "missing"));
    }

    @Test
    public void crossSheetExportsExpectedDataAndPreservesInputs() throws Exception {
        for (boolean replaceExisting : new boolean[]{false, true}) {
            Fixture fixture = fixture();
            File output = replaceExisting ? existingOutput(fixture)
                    : new File(fixture.directory, "output.xlsx");
            byte[] sourceBefore = Files.readAllBytes(fixture.source.toPath());
            byte[] targetBefore = Files.readAllBytes(fixture.target.toPath());

            fill(fixture, output, "id");

            try (Workbook workbook = WorkbookFactory.create(output)) {
                Sheet sheet = workbook.getSheetAt(0);
                assertEquals(3, sheet.getLastRowNum());
                assertEquals("existing|first|second|first", text(sheet, 1, 1));
                assertEquals("untouched", text(sheet, 2, 1));
                assertEquals("first|second|first", text(sheet, 3, 1));
                assertEquals("target-note", text(sheet, 1, 2));
            }
            assertArrayEquals(sourceBefore, Files.readAllBytes(fixture.source.toPath()));
            assertArrayEquals(targetBefore, Files.readAllBytes(fixture.target.toPath()));
            assertEquals(new HashSet<>(Arrays.asList("source.xlsx", "target.xlsx", "output.xlsx")),
                    fileNames(fixture.directory));
        }
    }

    @Test
    public void bothDuplicateEntryPointsRejectInputAsOutput() throws Exception {
        for (boolean legacy : new boolean[]{false, true}) {
            Fixture fixture = fixture();
            assertFailedExportPreservesFiles(fixture, fixture.source,
                    () -> merge(fixture, fixture.source, "id", legacy));
        }
    }

    @Test
    public void bothDuplicateEntryPointsPreserveExistingOutputOnFailure() throws Exception {
        for (boolean legacy : new boolean[]{false, true}) {
            Fixture fixture = fixture();
            File output = existingOutput(fixture);
            assertFailedExportPreservesFiles(fixture, output,
                    () -> merge(fixture, output, "missing", legacy));
        }
    }

    @Test
    public void bothDuplicateEntryPointsDoNotLeaveNewOutputOnFailure() throws Exception {
        for (boolean legacy : new boolean[]{false, true}) {
            Fixture fixture = fixture();
            File output = new File(fixture.directory, "output.xlsx");
            assertFailedExportPreservesFiles(fixture, output,
                    () -> merge(fixture, output, "missing", legacy));
        }
    }

    @Test
    public void bothDuplicateEntryPointsExportExpectedDataAndPreserveInputs() throws Exception {
        for (boolean legacy : new boolean[]{false, true}) {
            for (boolean replaceExisting : new boolean[]{false, true}) {
                Fixture fixture = fixture();
                File output = replaceExisting ? existingOutput(fixture)
                        : new File(fixture.directory, "output.xlsx");
                byte[] sourceBefore = Files.readAllBytes(fixture.source.toPath());

                merge(fixture, output, "id", legacy);

                try (Workbook workbook = WorkbookFactory.create(output)) {
                    Sheet sheet = workbook.getSheetAt(0);
                    assertEquals(2, sheet.getLastRowNum());
                    assertEquals("A", text(sheet, 1, 0));
                    assertEquals("first|second", text(sheet, 1, 1));
                    assertEquals(legacy ? "kept|later" : "kept", text(sheet, 1, 2));
                    assertEquals("B", text(sheet, 2, 0));
                    assertEquals("third", text(sheet, 2, 1));
                }
                assertArrayEquals(sourceBefore, Files.readAllBytes(fixture.source.toPath()));
                assertEquals(new HashSet<>(Arrays.asList("source.xlsx", "target.xlsx", "output.xlsx")),
                        fileNames(fixture.directory));
            }
        }
    }

    @Test
    public void crossSheetCellLimitFailurePreservesExistingOutput() throws Exception {
        Fixture fixture = fixture();
        writeOverlongAggregate(fixture.source);
        File output = existingOutput(fixture);
        byte[] outputBefore = Files.readAllBytes(output.toPath());
        Set<String> filesBefore = fileNames(fixture.directory);

        Exception failure = assertThrows(Exception.class, () -> fill(fixture, output, "id"));

        assertTrue(failure.getMessage(), failure.getMessage().contains("32767"));
        assertArrayEquals(outputBefore, Files.readAllBytes(output.toPath()));
        assertEquals(filesBefore, fileNames(fixture.directory));
    }

    @Test
    public void bothDuplicateEntryPointsPreserveExistingOutputOnCellLimitFailure() throws Exception {
        for (boolean legacy : new boolean[]{false, true}) {
            Fixture fixture = fixture();
            writeOverlongAggregate(fixture.source);
            File output = existingOutput(fixture);
            byte[] outputBefore = Files.readAllBytes(output.toPath());
            Set<String> filesBefore = fileNames(fixture.directory);

            Exception failure = assertThrows(Exception.class,
                    () -> merge(fixture, output, "id", legacy));

            assertTrue(failure.getMessage(), failure.getMessage().contains("32767"));
            assertArrayEquals(outputBefore, Files.readAllBytes(output.toPath()));
            assertEquals(filesBefore, fileNames(fixture.directory));
        }
    }

    private void assertFailedExportPreservesFiles(Fixture fixture, File output,
                                                   ExportOperation operation) throws Exception {
        byte[] sourceBefore = Files.readAllBytes(fixture.source.toPath());
        byte[] targetBefore = Files.readAllBytes(fixture.target.toPath());
        byte[] outputBefore = output.exists() ? Files.readAllBytes(output.toPath()) : null;
        Set<String> filesBefore = fileNames(fixture.directory);

        assertThrows(IOException.class, operation::run);

        assertArrayEquals(sourceBefore, Files.readAllBytes(fixture.source.toPath()));
        assertArrayEquals(targetBefore, Files.readAllBytes(fixture.target.toPath()));
        if (outputBefore == null) {
            assertFalse("失败时不应留下输出文件", output.exists());
        } else {
            assertArrayEquals(outputBefore, Files.readAllBytes(output.toPath()));
        }
        assertEquals("失败后应清理导出临时文件", filesBefore, fileNames(fixture.directory));
    }

    private void fill(Fixture fixture, File output, String sourceKey) throws IOException {
        new CrossSheetFiller().fill(fixture.source.getPath(), fixture.target.getPath(),
                output.getPath(), sourceKey, "value", "id", "value", null);
    }

    private void merge(Fixture fixture, File output, String key, boolean legacy) throws IOException {
        DuplicateColumnMerger merger = new DuplicateColumnMerger();
        if (legacy) {
            merger.mergeByKeyColumn(fixture.source.getPath(), output.getPath(), key);
        } else {
            merger.mergeByCustomerIdWithProgress(fixture.source.getPath(), output.getPath(), key,
                    Arrays.asList("value"), null);
        }
    }

    private Fixture fixture() throws IOException {
        File directory = temporaryFolder.newFolder();
        File source = new File(directory, "source.xlsx");
        File target = new File(directory, "target.xlsx");
        writeWorkbook(source, new String[][]{
                {"id", "value", "note"}, {"A", "first", "kept"},
                {"A", "second", "later"}, {"A", "first", "kept"}, {"B", "third", "b-note"}
        });
        writeWorkbook(target, new String[][]{
                {"id", "value", "note"}, {"A", "existing", "target-note"},
                {"C", "untouched", "unmatched"}, {"A", "", "repeated"}
        });
        return new Fixture(directory, source, target);
    }

    private File existingOutput(Fixture fixture) throws IOException {
        File output = new File(fixture.directory, "output.xlsx");
        writeWorkbook(output, new String[][]{{"previous"}, {"keep on failure"}});
        return output;
    }

    private void writeOverlongAggregate(File source) throws IOException {
        char[] chars = new char[20000];
        Arrays.fill(chars, 'x');
        String value = new String(chars);
        writeWorkbook(source, new String[][]{
                {"id", "value"}, {"A", value}, {"A", "y" + value}
        });
    }

    private void writeWorkbook(File file, String[][] rows) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream out = new FileOutputStream(file)) {
            Sheet sheet = workbook.createSheet("Data");
            for (int r = 0; r < rows.length; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < rows[r].length; c++) {
                    row.createCell(c).setCellValue(rows[r][c]);
                }
            }
            workbook.write(out);
        }
    }

    private static String text(Sheet sheet, int row, int column) {
        return sheet.getRow(row).getCell(column).getStringCellValue();
    }

    private static Set<String> fileNames(File directory) {
        String[] names = directory.list();
        assertNotNull(names);
        return new HashSet<>(Arrays.asList(names));
    }

    private interface ExportOperation {
        void run() throws Exception;
    }

    private static class Fixture {
        private final File directory;
        private final File source;
        private final File target;

        private Fixture(File directory, File source, File target) {
            this.directory = directory;
            this.source = source;
            this.target = target;
        }
    }
}

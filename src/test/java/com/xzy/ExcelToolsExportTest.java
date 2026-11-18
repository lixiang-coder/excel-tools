package com.xzy;

import com.xzy.model.MappingRelations;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import static org.junit.Assert.*;

@RunWith(Parameterized.class)
public class ExcelToolsExportTest {
    private static final byte[] OLD_OUTPUT = "existing output must survive".getBytes(StandardCharsets.UTF_8);
    private static final BiConsumer<Integer, String> NO_PROGRESS = (percent, message) -> { };
    private static final Consumer<String> NO_LOG = message -> { };

    @Parameterized.Parameters(name = "input={0}")
    public static Collection<Object[]> formats() {
        return Arrays.asList(new Object[][]{{"xls"}, {"xlsx"}});
    }

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    private final String extension;

    public ExcelToolsExportTest(String extension) {
        this.extension = extension;
    }

    @Test
    public void processorPublishesReadableOutputAndPreservesBothInputs() throws Exception {
        File source = workbook("source", new String[][]{{"Name"}, {"Alice"}});
        File template = workbook("template", new String[][]{{"Result"}});
        File output = existingOutput("processed.xlsx");
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        byte[] templateBefore = Files.readAllBytes(template.toPath());
        Set<String> before = directoryEntries();

        int processed = ExcelTools.exportProcessedExcel(source, template, output,
                Collections.singletonMap("Name", "Result"), stillUnpublished(output), NO_LOG);

        assertEquals(1, processed);
        assertEquals("Result", cell(output, 0, 0));
        assertEquals("Alice", cell(output, 1, 0));
        assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
        assertArrayEquals(templateBefore, Files.readAllBytes(template.toPath()));
        assertEquals(before, directoryEntries());
    }

    @Test
    public void processorRejectsSourceTemplateAndTheirNormalizedAliases() throws Exception {
        File source = workbook("source", new String[][]{{"Name"}, {"Alice"}});
        File template = workbook("template", new String[][]{{"Result"}});
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        byte[] templateBefore = Files.readAllBytes(template.toPath());
        Set<String> before = directoryEntries();

        for (File output : Arrays.asList(source, alias(source), template, alias(template))) {
            assertThrows(IOException.class, () -> ExcelTools.exportProcessedExcel(source, template, output,
                    Collections.singletonMap("Name", "Result"), NO_PROGRESS, NO_LOG));
            assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
            assertArrayEquals(templateBefore, Files.readAllBytes(template.toPath()));
            assertEquals(before, directoryEntries());
        }
    }

    @Test
    public void processorFailurePreservesExistingOutputAndDoesNotPublishNewOutput() throws Exception {
        String longValue = repeated('x', 20000);
        File source = workbook("source", new String[][]{{"Left", "Right"}, {longValue, longValue}});
        File template = workbook("template", new String[][]{{"Result"}});
        Map<String, String> columns = new HashMap<>();
        columns.put("Left", "Result");
        columns.put("Right", "Result");
        byte[] sourceBefore = Files.readAllBytes(source.toPath());
        byte[] templateBefore = Files.readAllBytes(template.toPath());
        File existing = existingOutput("existing.xlsx");
        File absent = new File(temporary.getRoot(), "absent.xlsx");
        Set<String> before = directoryEntries();

        for (File output : Arrays.asList(existing, absent)) {
            // Each source value is valid, but their merged output exceeds Excel's cell text limit.
            assertThrows(IllegalArgumentException.class, () -> ExcelTools.exportProcessedExcel(
                    source, template, output, columns, NO_PROGRESS, NO_LOG));
            assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(existing.toPath()));
            assertFalse(absent.exists());
            assertArrayEquals(sourceBefore, Files.readAllBytes(source.toPath()));
            assertArrayEquals(templateBefore, Files.readAllBytes(template.toPath()));
            assertEquals(before, directoryEntries());
        }
    }

    @Test
    public void appliedMappingPublishesReadableOutputAndPreservesTarget() throws Exception {
        File target = workbook("target", new String[][]{{"Key", "Value"}, {"A", "old"}, {"missing", "old"}});
        File output = existingOutput("mapped.xlsx");
        byte[] targetBefore = Files.readAllBytes(target.toPath());
        Set<String> before = directoryEntries();

        MappingRelations relations = new MappingRelations();
        relations.add("A", "Alpha", target, 2);
        int[] stats = ExcelTools.exportAppliedMapping(target, output, "Key", "Value",
                relations, stillUnpublished(output), NO_LOG);

        assertArrayEquals(new int[]{2, 1, 0}, stats);
        assertEquals("Alpha", cell(output, 1, 1));
        assertEquals("missing", cell(output, 2, 1));
        assertArrayEquals(targetBefore, Files.readAllBytes(target.toPath()));
        assertEquals(before, directoryEntries());
    }

    @Test
    public void appliedMappingProtectsTargetAndEveryContributingMappingFile() throws Exception {
        File target = workbook("target", new String[][]{{"Key", "Value"}, {"A", "old"}});
        File firstMapping = workbook("first-mapping", new String[][]{{"From", "To"}, {"A", "Alpha"}});
        File secondMapping = workbook("second-mapping", new String[][]{{"From", "To"}, {"B", "Beta"}});
        File emptyMapping = workbook("empty-mapping", new String[][]{{"From", "To"}, {"", ""}});
        MappingRelations relations = new MappingRelations();

        assertEquals(1, ExcelTools.readMappingRelations(firstMapping, "From", "To", relations));
        assertEquals(1, ExcelTools.readMappingRelations(secondMapping, "From", "To", relations));
        assertEquals(0, ExcelTools.readMappingRelations(emptyMapping, "From", "To", relations));
        assertEquals(new HashSet<>(Arrays.asList(firstMapping, secondMapping)), relations.getSourceFiles());
        assertEquals("Alpha", relations.resolve("A").getValue());
        assertEquals("Beta", relations.resolve("B").getValue());
        Set<String> before = directoryEntries();

        for (File protectedFile : Arrays.asList(target, firstMapping, secondMapping)) {
            byte[] original = Files.readAllBytes(protectedFile.toPath());
            for (File output : Arrays.asList(protectedFile, alias(protectedFile))) {
                assertThrows(IOException.class, () -> ExcelTools.exportAppliedMapping(target, output, "Key", "Value",
                        relations, NO_PROGRESS, NO_LOG));
                assertArrayEquals(original, Files.readAllBytes(protectedFile.toPath()));
                assertEquals(before, directoryEntries());
            }
        }
    }

    @Test
    public void appliedMappingFailurePreservesExistingOutputAndDoesNotPublishNewOutput() throws Exception {
        File target = workbook("target", new String[][]{{"Key", "Value"}, {"A", "old"}});
        byte[] targetBefore = Files.readAllBytes(target.toPath());
        File existing = existingOutput("existing.xlsx");
        File absent = new File(temporary.getRoot(), "absent.xlsx");
        MappingRelations relations = new MappingRelations();
        relations.add("A", repeated('x', 32768), target, 2);
        Set<String> before = directoryEntries();

        for (File output : Arrays.asList(existing, absent)) {
            assertThrows(IllegalArgumentException.class, () -> ExcelTools.exportAppliedMapping(
                    target, output, "Key", "Value", relations, NO_PROGRESS, NO_LOG));
            assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(existing.toPath()));
            assertFalse(absent.exists());
            assertArrayEquals(targetBefore, Files.readAllBytes(target.toPath()));
            assertEquals(before, directoryEntries());
        }
    }

    @Test
    public void corruptInputsDoNotReplaceExistingOutputs() throws Exception {
        File source = temporary.newFile("corrupt." + extension);
        Files.write(source.toPath(), "not an Excel workbook".getBytes(StandardCharsets.UTF_8));
        File template = workbook("template", new String[][]{{"Result"}});
        File output = existingOutput("existing.xlsx");
        Set<String> before = directoryEntries();

        assertThrows(Exception.class, () -> ExcelTools.exportProcessedExcel(source, template, output,
                Collections.singletonMap("Name", "Result"), NO_PROGRESS, NO_LOG));
        assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
        assertEquals(before, directoryEntries());

        assertThrows(Exception.class, () -> ExcelTools.exportAppliedMapping(source, output, "Key", "Value",
                new MappingRelations(), NO_PROGRESS, NO_LOG));
        assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
        assertEquals(before, directoryEntries());
    }

    private BiConsumer<Integer, String> stillUnpublished(File output) {
        return (percent, message) -> {
            assertTrue("Progress must not announce completion before commit", percent < 100);
            try {
                assertArrayEquals(OLD_OUTPUT, Files.readAllBytes(output.toPath()));
            } catch (IOException e) {
                throw new AssertionError(e);
            }
        };
    }

    private File workbook(String name, String[][] values) throws IOException {
        File file = new File(temporary.getRoot(), name + "." + extension);
        try (Workbook workbook = "xls".equals(extension) ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            for (int rowIndex = 0; rowIndex < values.length; rowIndex++) {
                Row row = sheet.createRow(rowIndex);
                for (int column = 0; column < values[rowIndex].length; column++) {
                    row.createCell(column).setCellValue(values[rowIndex][column]);
                }
            }
            try (FileOutputStream output = new FileOutputStream(file)) {
                workbook.write(output);
            }
        }
        return file;
    }

    private File existingOutput(String name) throws IOException {
        File output = temporary.newFile(name);
        Files.write(output.toPath(), OLD_OUTPUT);
        return output;
    }

    private File alias(File file) {
        return new File(file.getParentFile(), "." + File.separator + file.getName());
    }

    private Set<String> directoryEntries() {
        String[] names = temporary.getRoot().list();
        assertNotNull(names);
        return new HashSet<>(Arrays.asList(names));
    }

    private String cell(File file, int row, int column) throws IOException {
        try (FileInputStream input = new FileInputStream(file); Workbook workbook = WorkbookFactory.create(input)) {
            return workbook.getSheetAt(0).getRow(row).getCell(column).getStringCellValue();
        }
    }

    private static String repeated(char value, int count) {
        char[] characters = new char[count];
        Arrays.fill(characters, value);
        return new String(characters);
    }
}

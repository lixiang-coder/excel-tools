package com.xzy;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;

import static org.junit.Assert.*;

public class TableMergerWriteFailureTest {
    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void unreadableInputIsReportedButWriteFailurePreservesThePreviousCompleteOutput() throws Exception {
        File good = workbook("good.xlsx", 1);
        File damaged = temporary.newFile("damaged.xlsx");
        Files.write(damaged.toPath(), new byte[] {1, 2, 3});
        File wide = workbook("wide.xlsx", SpreadsheetVersion.EXCEL2007.getLastColumnIndex());
        File output = new File(temporary.getRoot(), "output.xlsx");
        TableMerger merger = new TableMerger();
        TableMerger.MergeStrategy strategy = TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER;

        TableMerger.MergeSummary summary = merger.mergeFilesWithSummary(Arrays.asList(good, damaged),
                output.getPath(), strategy, null);
        assertEquals(1, summary.getSuccessCount());
        assertEquals(1, summary.getFailureCount());
        byte[] previous = Files.readAllBytes(output.toPath());

        IOException failure = assertThrows(IOException.class, () -> merger.mergeFilesWithSummary(
                Arrays.asList(good, wide), output.getPath(), strategy, null));

        assertTrue(failure.getMessage(), failure.getMessage().contains(wide.getAbsolutePath()));
        assertArrayEquals(previous, Files.readAllBytes(output.toPath()));
        assertEquals(4, temporary.getRoot().list().length);
    }

    private File workbook(String name, int lastColumn) throws IOException {
        File file = temporary.newFile(name);
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream stream = new FileOutputStream(file)) {
            Row row = workbook.createSheet("数据").createRow(0);
            row.createCell(0).setCellValue("首列");
            row.createCell(lastColumn).setCellValue("末列");
            workbook.write(stream);
        }
        return file;
    }
}

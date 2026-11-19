package com.xzy.panel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import javax.swing.JComboBox;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import static org.junit.Assert.*;

public class CrossSheetPanelTest {
    @Rule public TemporaryFolder temp = new TemporaryFolder();

    @Test
    public void targetMustBeChosenAgainAfterReloadAndAStaleLoadCannotRestoreOldColumns() throws Exception {
        File first = workbook("first.xls", "编号", "备注");
        File second = workbook("second.xls", "工号", "姓名");
        File missing = new File(temp.getRoot(), "missing.xls");
        JTextField path = new JTextField();
        JTextArea preview = new JTextArea();
        JComboBox<String> key = new JComboBox<>();
        JComboBox<String> value = new JComboBox<>();
        CrossSheetPanel.WorkbookFields fields = new CrossSheetPanel.WorkbookFields(path, preview, key, value, true);

        CountDownLatch firstDone = new CountDownLatch(1);
        SwingUtilities.invokeAndWait(() -> fields.load(first, message -> {}, firstDone::countDown));
        assertTrue(firstDone.await(15, TimeUnit.SECONDS));

        CountDownLatch bothDone = new CountDownLatch(2);
        SwingUtilities.invokeAndWait(() -> {
            assertEquals("编号", key.getSelectedItem());
            assertNull(value.getSelectedItem());
            value.setSelectedItem("备注");
            fields.load(first, message -> {}, bothDone::countDown);
            assertNull(value.getSelectedItem());
            fields.load(second, message -> {}, bothDone::countDown);
            assertEquals(second.getAbsolutePath(), path.getText());
            assertEquals(0, key.getItemCount());
            assertEquals(0, value.getItemCount());
        });
        assertTrue(bothDone.await(15, TimeUnit.SECONDS));

        CountDownLatch failed = new CountDownLatch(1);
        SwingUtilities.invokeAndWait(() -> {
            assertEquals("工号", key.getSelectedItem());
            assertNull(value.getSelectedItem());
            assertEquals("姓名", value.getItemAt(1));
            value.setSelectedItem("姓名");
            fields.load(missing, message -> {}, failed::countDown);
        });
        assertTrue(failed.await(15, TimeUnit.SECONDS));
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(missing.getAbsolutePath(), path.getText());
            assertEquals(0, key.getItemCount());
            assertEquals(0, value.getItemCount());
            assertTrue(preview.getText().contains("读取目标表表头失败"));
        });
    }

    @Test
    public void onlyAnExplicitChoiceCanAppendToTheMatchKey() {
        assertFalse(CrossSheetPanel.allowTargetColumn("编号", null,
                message -> assertTrue(message.contains("填入")), message -> { throw new AssertionError(message); }));
        assertFalse(CrossSheetPanel.allowTargetColumn("编号", "编号", message -> fail(message), message -> {
            assertTrue(message.contains("编号"));
            assertTrue(message.contains("主键"));
            return false;
        }));
        assertTrue(CrossSheetPanel.allowTargetColumn("编号", "编号", message -> fail(message), message -> true));
        assertTrue(CrossSheetPanel.allowTargetColumn("编号", "备注", message -> fail(message),
                message -> { throw new AssertionError(message); }));
    }

    private File workbook(String filename, String key, String value) throws IOException {
        File file = temp.newFile(filename);
        try (HSSFWorkbook workbook = new HSSFWorkbook(); FileOutputStream stream = new FileOutputStream(file)) {
            Row row = workbook.createSheet().createRow(0);
            row.createCell(0).setCellValue(key);
            row.createCell(1).setCellValue(value);
            workbook.write(stream);
        }
        return file;
    }
}

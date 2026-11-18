package com.xzy.ui;

import org.junit.Test;

import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;
import java.awt.FontMetrics;
import java.awt.Insets;

import static org.junit.Assert.assertTrue;

public class WorkbenchLayoutTest {
    @Test
    public void toolbarFitsFullChineseButtonLabelsAtLargeFontSize() throws Exception {
        SwingUtilities.invokeAndWait(() -> {
            JButton select = Workbench.primaryButton("选择文件");
            JButton reload = new JButton("重新导入文件");
            JButton[] buttons = {select, reload};
            JPanel bar = Workbench.actions(buttons);
            for (JButton button : buttons) {
                button.setFont(button.getFont().deriveFont(28f));
            }
            bar.setSize(bar.getPreferredSize());
            bar.doLayout();

            for (JButton button : buttons) {
                Insets insets = button.getInsets();
                FontMetrics metrics = button.getFontMetrics(button.getFont());
                int textWidth = button.getWidth() - insets.left - insets.right;
                assertTrue(button.getText(), textWidth >= metrics.stringWidth(button.getText()));
            }
            assertTrue(select.getX() + select.getWidth() < reload.getX());
            assertTrue(reload.getX() + reload.getWidth() <= bar.getWidth());
        });
    }
}

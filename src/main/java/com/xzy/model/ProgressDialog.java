package com.xzy.model;

import javax.swing.*;
import java.awt.*;

/**
 * 进度对话框包装类
 */
public class ProgressDialog {
    public final JDialog dialog;
    public final JLabel label;
    public final JProgressBar progressBar;

    public ProgressDialog(JDialog dialog, JLabel label, JProgressBar progressBar) {
        this.dialog = dialog;
        this.label = label;
        this.progressBar = progressBar;
    }

    public static ProgressDialog create(JFrame parent, String title) {
        JDialog dialog = new JDialog(parent, title, true);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.setResizable(false);
        dialog.setUndecorated(false);

        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JLabel label = new JLabel("正在处理...");
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);

        panel.add(label, BorderLayout.NORTH);
        panel.add(progressBar, BorderLayout.CENTER);

        dialog.add(panel);
        dialog.pack();
        dialog.setLocationRelativeTo(parent);

        return new ProgressDialog(dialog, label, progressBar);
    }
}

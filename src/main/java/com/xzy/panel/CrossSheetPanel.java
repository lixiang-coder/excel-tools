package com.xzy.panel;

import com.xzy.CrossSheetFiller;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
import com.xzy.util.ExcelUtils;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;

/**
 * Tab4: 跨表主键填充工具面板
 */
public class CrossSheetPanel {
    private final AppContext ctx;

    public CrossSheetPanel(AppContext ctx) {
        this.ctx = ctx;
    }

    public JPanel createPanel() {
        Font uiFont = new Font("Microsoft YaHei", Font.PLAIN, 16);
        JPanel filePanel = new JPanel(new GridBagLayout());
        filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1;
        JTextField aPath = new JTextField();
        JTextField bPath = new JTextField();
        aPath.setEditable(false);
        bPath.setEditable(false);
        JButton selA = new JButton("选择源表");
        JButton selB = new JButton("选择目标表");
        Font btnFont = new Font("Microsoft YaHei", Font.BOLD, 16);
        selA.setFont(btnFont);
        selB.setFont(btnFont);
        selA.setBorderPainted(false);
        selB.setBorderPainted(false);
        selA.setFocusPainted(false);
        selB.setFocusPainted(false);
        selA.setOpaque(true);
        selB.setOpaque(true);
        Color blue = new Color(66, 133, 244);
        Color grayBg = new Color(220, 220, 220);
        Color grayFont = new Color(120, 120, 120);
        java.util.function.BiConsumer<JButton, Boolean> setButtonVisualEnabled = (btn, enabled) -> {
            if (enabled) {
                btn.setBackground(blue);
                btn.setForeground(Color.WHITE);
            } else {
                btn.setBackground(grayBg);
                btn.setForeground(grayFont);
            }
            btn.setBorder(BorderFactory.createEmptyBorder());
            btn.setFocusPainted(false);
            btn.setOpaque(true);
        };
        setButtonVisualEnabled.accept(selA, true);
        setButtonVisualEnabled.accept(selB, true);
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(new JLabel("源表路径:"), gbc);
        gbc.gridx = 1; gbc.gridy = 0; gbc.gridwidth = 2;
        filePanel.add(aPath, gbc);
        gbc.gridx = 3; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(selA, gbc);
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1;
        filePanel.add(new JLabel("目标表路径:"), gbc);
        gbc.gridx = 1; gbc.gridy = 1; gbc.gridwidth = 2;
        filePanel.add(bPath, gbc);
        gbc.gridx = 3; gbc.gridy = 1; gbc.gridwidth = 1;
        filePanel.add(selB, gbc);

        JPanel headerPanel = new JPanel(new GridLayout(2, 1, 5, 5));
        headerPanel.setBorder(BorderFactory.createTitledBorder("表头展示区"));
        JTextArea aHeadersArea = new JTextArea(2, 40);
        aHeadersArea.setEditable(false);
        aHeadersArea.setLineWrap(true);
        aHeadersArea.setBorder(BorderFactory.createTitledBorder("源表表头"));
        JTextArea bHeadersArea = new JTextArea(2, 40);
        bHeadersArea.setEditable(false);
        bHeadersArea.setLineWrap(true);
        bHeadersArea.setBorder(BorderFactory.createTitledBorder("目标表表头"));
        headerPanel.add(new JScrollPane(aHeadersArea));
        headerPanel.add(new JScrollPane(bHeadersArea));

        JPanel mapPanel = new JPanel(new GridBagLayout());
        mapPanel.setBorder(BorderFactory.createTitledBorder("表头映射区"));
        JComboBox<String> aKeyCombo = new JComboBox<>();
        JComboBox<String> aValCombo = new JComboBox<>();
        JComboBox<String> bKeyCombo = new JComboBox<>();
        JComboBox<String> bValCombo = new JComboBox<>();
        aKeyCombo.setFont(uiFont); aValCombo.setFont(uiFont);
        bKeyCombo.setFont(uiFont); bValCombo.setFont(uiFont);
        aKeyCombo.setPreferredSize(new Dimension(180, 32));
        aValCombo.setPreferredSize(new Dimension(180, 32));
        bKeyCombo.setPreferredSize(new Dimension(180, 32));
        bValCombo.setPreferredSize(new Dimension(180, 32));
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
        mapPanel.add(new JLabel("源表匹配表头:"), gbc);
        gbc.gridx = 1; gbc.gridy = 0;
        mapPanel.add(aKeyCombo, gbc);
        gbc.gridx = 2; gbc.gridy = 0;
        mapPanel.add(new JLabel("源表数据表头:"), gbc);
        gbc.gridx = 3; gbc.gridy = 0;
        mapPanel.add(aValCombo, gbc);
        gbc.gridx = 0; gbc.gridy = 1;
        mapPanel.add(new JLabel("目标表匹配表头:"), gbc);
        gbc.gridx = 1; gbc.gridy = 1;
        mapPanel.add(bKeyCombo, gbc);
        gbc.gridx = 2; gbc.gridy = 1;
        mapPanel.add(new JLabel("目标表填充表头:"), gbc);
        gbc.gridx = 3; gbc.gridy = 1;
        mapPanel.add(bValCombo, gbc);

        JPanel opPanel = new JPanel(new BorderLayout(10, 10));
        JButton runBtn = new JButton("开始填充");
        opPanel.add(runBtn, BorderLayout.CENTER);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        mainPanel.add(filePanel);
        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(headerPanel);
        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(mapPanel);
        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(opPanel);

        // 事件
        selA.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                aPath.setText(fc.getSelectedFile().getAbsolutePath());
                ctx.log("已选择源表: " + aPath.getText());
                ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "加载源表表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try { headers = ExcelUtils.readNamedHeaders(new File(aPath.getText())); }
                        catch (Exception ex) { headers = null; }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        if (headers != null) {
                            aHeadersArea.setText(String.join(", ", headers));
                            aKeyCombo.removeAllItems(); aValCombo.removeAllItems();
                            for (String h : headers) { aKeyCombo.addItem(h); aValCombo.addItem(h); }
                            if (aKeyCombo.getItemCount() > 0) aKeyCombo.setSelectedIndex(0);
                            if (aValCombo.getItemCount() > 0) aValCombo.setSelectedIndex(0);
                        } else {
                            aHeadersArea.setText("读取源表表头失败");
                            aKeyCombo.removeAllItems(); aValCombo.removeAllItems();
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
            }
        });
        selB.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                bPath.setText(fc.getSelectedFile().getAbsolutePath());
                ctx.log("已选择目标表: " + bPath.getText());
                ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "加载目标表表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try { headers = ExcelUtils.readNamedHeaders(new File(bPath.getText())); }
                        catch (Exception ex) { headers = null; }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        if (headers != null) {
                            bHeadersArea.setText(String.join(", ", headers));
                            bKeyCombo.removeAllItems(); bValCombo.removeAllItems();
                            for (String h : headers) { bKeyCombo.addItem(h); bValCombo.addItem(h); }
                            if (bKeyCombo.getItemCount() > 0) bKeyCombo.setSelectedIndex(0);
                            if (bValCombo.getItemCount() > 0) bValCombo.setSelectedIndex(0);
                        } else {
                            bHeadersArea.setText("读取目标表表头失败");
                            bKeyCombo.removeAllItems(); bValCombo.removeAllItems();
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
            }
        });
        runBtn.addActionListener(e -> {
            String aPathVal = aPath.getText();
            String bPathVal = bPath.getText();
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fc.setDialogTitle("选择输出文件");
            if (fc.showSaveDialog(mainPanel) != JFileChooser.APPROVE_OPTION) {
                ctx.log("未选择输出文件，操作取消");
                return;
            }
            String outPathRaw = fc.getSelectedFile().getAbsolutePath();
            final String outPathVal = outPathRaw.toLowerCase().endsWith(".xlsx") ? outPathRaw : outPathRaw + ".xlsx";
            ctx.log("已选择输出: " + outPathVal);
            String aKey = (String) aKeyCombo.getSelectedItem();
            String aVal = (String) aValCombo.getSelectedItem();
            String bKey = (String) bKeyCombo.getSelectedItem();
            String bVal = (String) bValCombo.getSelectedItem();
            if (aPathVal == null || aPathVal.isEmpty() || bPathVal == null || bPathVal.isEmpty() ||
                aKey == null || aVal == null || bKey == null || bVal == null) {
                ctx.log("请先选择所有文件和表头映射项！");
                JOptionPane.showMessageDialog(mainPanel, "请先选择所有文件和表头映射项！", "提示", JOptionPane.WARNING_MESSAGE);
                return;
            }
            ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "处理中");
            SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                @Override
                protected Void doInBackground() {
                    try {
                        new CrossSheetFiller().fill(aPathVal, bPathVal, outPathVal, aKey, aVal, bKey, bVal,
                                (cur, total) -> publish((int) (cur * 100.0 / total)));
                    } catch (Exception ex) { throw new RuntimeException(ex); }
                    return null;
                }
                @Override
                protected void process(java.util.List<Integer> chunks) {
                    if (!chunks.isEmpty()) {
                        int val = chunks.get(chunks.size() - 1);
                        progressDialog.progressBar.setIndeterminate(false);
                        progressDialog.progressBar.setValue(val);
                        progressDialog.label.setText("正在处理数据... " + val + "%");
                    }
                }
                @Override
                protected void done() { progressDialog.dialog.dispose(); }
            };
            worker.execute();
            progressDialog.dialog.setVisible(true);
            try {
                worker.get();
                ctx.log("处理完成: " + outPathVal);
                JOptionPane.showMessageDialog(mainPanel, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ctx.log("失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }
}

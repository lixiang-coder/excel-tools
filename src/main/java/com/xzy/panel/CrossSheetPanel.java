package com.xzy.panel;

import com.xzy.CrossSheetFiller;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
import com.xzy.ui.Workbench;

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
        JTextField aPath = new JTextField();
        JTextField bPath = new JTextField();
        JButton selA = new JButton("选择源表");
        JButton selB = new JButton("选择目标表");
        JTextArea aHeadersArea = new JTextArea(2, 16);
        JTextArea bHeadersArea = new JTextArea(2, 16);
        JComboBox<String> aKeyCombo = new JComboBox<>();
        JComboBox<String> aValCombo = new JComboBox<>();
        JComboBox<String> bKeyCombo = new JComboBox<>();
        JComboBox<String> bValCombo = new JComboBox<>();
        JButton runBtn = Workbench.primaryButton("开始填充");

        JPanel mainPanel = Workbench.formPage(runBtn,
                workbookSection("源工作表", aPath, selA, aHeadersArea, aKeyCombo, "读取数据的列", aValCombo),
                workbookSection("目标工作表", bPath, selB, bHeadersArea, bKeyCombo, "填入数据的列", bValCombo));

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
                    String error;
                    @Override
                    protected Void doInBackground() {
                        try { headers = CrossSheetFiller.readMatchingHeaders(new File(aPath.getText())); }
                        catch (Exception ex) { error = ex.getMessage(); }
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
                            aHeadersArea.setText("读取源表表头失败：" + error);
                            ctx.log(aHeadersArea.getText());
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
                    String error;
                    @Override
                    protected Void doInBackground() {
                        try { headers = CrossSheetFiller.readMatchingHeaders(new File(bPath.getText())); }
                        catch (Exception ex) { error = ex.getMessage(); }
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
                            bHeadersArea.setText("读取目标表表头失败：" + error);
                            ctx.log(bHeadersArea.getText());
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
            File output = Workbench.chooseFilledWorkbook(mainPanel, new File(bPathVal), "保存跨表填充后的工作簿");
            if (output == null) {
                ctx.log("未选择输出文件，操作取消");
                return;
            }
            String outPathVal = output.getAbsolutePath();
            ctx.log("已选择输出: " + outPathVal);
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

    private static JPanel workbookSection(String title, JTextField path, JButton select, JTextArea headers,
                                          JComboBox<String> key, String valueLabel, JComboBox<String> value) {
        path.setEditable(false);
        path.getAccessibleContext().setAccessibleName(title + "路径");
        JPanel fileRow = new JPanel(new BorderLayout(12, 0));
        fileRow.add(path, BorderLayout.CENTER);
        fileRow.add(select, BorderLayout.LINE_END);

        headers.setEditable(false);
        headers.setLineWrap(true);
        headers.setWrapStyleWord(true);
        JLabel headersLabel = new JLabel("表头预览");
        headersLabel.setLabelFor(headers);
        JPanel preview = new JPanel(new BorderLayout(12, 0));
        preview.add(headersLabel, BorderLayout.LINE_START);
        preview.add(new JScrollPane(headers), BorderLayout.CENTER);

        key.setPrototypeDisplayValue("用于匹配的表头");
        value.setPrototypeDisplayValue("待填充的表头");
        JPanel columns = new JPanel(new GridLayout(1, 2, 16, 0));
        columns.add(columnSelector("主键列", key));
        columns.add(columnSelector(valueLabel, value));

        JPanel content = new JPanel(new BorderLayout(0, 10));
        content.add(fileRow, BorderLayout.NORTH);
        content.add(preview, BorderLayout.CENTER);
        content.add(columns, BorderLayout.SOUTH);
        return Workbench.formSection(title, content);
    }

    private static JPanel columnSelector(String name, JComboBox<String> combo) {
        JLabel label = new JLabel(name);
        label.setLabelFor(combo);
        JPanel field = new JPanel(new BorderLayout(8, 0));
        field.add(label, BorderLayout.LINE_START);
        field.add(combo, BorderLayout.CENTER);
        return field;
    }
}

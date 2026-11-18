package com.xzy.panel;

import com.xzy.DuplicateColumnMerger;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
import com.xzy.ui.Workbench;
import com.xzy.util.ExcelUtils;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;

/**
 * Tab5: 重复列合并工具面板
 */
public class DuplicateColPanel {
    private final AppContext ctx;

    public DuplicateColPanel(AppContext ctx) {
        this.ctx = ctx;
    }

    public JPanel createPanel() {
        JTextField inputPath = new JTextField();
        inputPath.setEditable(false);
        inputPath.getAccessibleContext().setAccessibleName("输入文件路径");
        JButton selectInput = new JButton("选择输入文件");
        JPanel filePanel = new JPanel(new BorderLayout(12, 0));
        filePanel.add(inputPath, BorderLayout.CENTER);
        filePanel.add(selectInput, BorderLayout.LINE_END);

        JComboBox<String> customerIdCombo = new JComboBox<>();
        customerIdCombo.setPrototypeDisplayValue("用于判断重复记录的字段");
        JLabel keyLabel = new JLabel("用于判断重复的列");
        keyLabel.setLabelFor(customerIdCombo);
        JPanel keyField = new JPanel(new BorderLayout(0, 6));
        keyField.add(keyLabel, BorderLayout.NORTH);
        keyField.add(customerIdCombo, BorderLayout.CENTER);
        JPanel keyColumn = new JPanel(new BorderLayout());
        keyColumn.add(keyField, BorderLayout.NORTH);

        DefaultListModel<String> mergeColsModel = new DefaultListModel<>();
        JList<String> mergeColsList = new JList<>(mergeColsModel);
        mergeColsList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        mergeColsList.setVisibleRowCount(5);
        JLabel mergeLabel = new JLabel("要合并的列（按住 Ctrl 可多选）");
        mergeLabel.setLabelFor(mergeColsList);
        JPanel mergeColumn = new JPanel(new BorderLayout(0, 6));
        mergeColumn.add(mergeLabel, BorderLayout.NORTH);
        mergeColumn.add(new JScrollPane(mergeColsList), BorderLayout.CENTER);

        JPanel headerPanel = new JPanel(new GridLayout(1, 2, 16, 0));
        headerPanel.add(keyColumn);
        headerPanel.add(mergeColumn);

        JButton runBtn = Workbench.primaryButton("开始合并");
        JPanel mainPanel = Workbench.formPage(runBtn,
                Workbench.formSection("输入文件", filePanel),
                Workbench.formSection("字段设置", headerPanel));

        // 事件
        selectInput.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                inputPath.setText(fc.getSelectedFile().getAbsolutePath());
                ctx.log("已选择输入: " + inputPath.getText());
                ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "加载输入文件表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try { headers = ExcelUtils.readNamedHeaders(new File(inputPath.getText())); }
                        catch (Exception ex) { headers = null; }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        if (headers != null) {
                            customerIdCombo.removeAllItems();
                            mergeColsModel.clear();
                            for (String h : headers) {
                                customerIdCombo.addItem(h);
                                mergeColsModel.addElement(h);
                            }
                            if (customerIdCombo.getItemCount() > 0) customerIdCombo.setSelectedIndex(0);
                        } else {
                            customerIdCombo.removeAllItems();
                            mergeColsModel.clear();
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
            }
        });
        runBtn.addActionListener(e -> {
            String inPath = inputPath.getText();
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fc.setDialogTitle("选择输出文件");
            if (fc.showSaveDialog(mainPanel) != JFileChooser.APPROVE_OPTION) {
                ctx.log("未选择输出文件，操作取消");
                return;
            }
            String outPathRaw = fc.getSelectedFile().getAbsolutePath();
            final String outPath = outPathRaw.toLowerCase().endsWith(".xlsx") ? outPathRaw : outPathRaw + ".xlsx";
            ctx.log("已选择输出: " + outPath);
            if (inPath == null || inPath.isEmpty()) {
                ctx.log("请先选择输入文件");
                return;
            }
            if (customerIdCombo.getSelectedItem() == null || mergeColsList.getSelectedValuesList().isEmpty()) {
                ctx.log("请选择CustomerId列和要合并的列");
                return;
            }
            String customerIdCol = (String) customerIdCombo.getSelectedItem();
            java.util.List<String> mergeCols = mergeColsList.getSelectedValuesList();
            ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "处理中");
            SwingWorker<DuplicateColumnMerger.MergeReport, Integer> worker =
                    new SwingWorker<DuplicateColumnMerger.MergeReport, Integer>() {
                @Override
                protected DuplicateColumnMerger.MergeReport doInBackground() {
                    try {
                        return new DuplicateColumnMerger().mergeWithReport(inPath, outPath, customerIdCol, mergeCols,
                                (cur, total) -> publish((int) (cur * 100.0 / total)));
                    } catch (Exception ex) { throw new RuntimeException(ex); }
                }
                @Override
                protected void process(java.util.List<Integer> chunks) {
                    if (!chunks.isEmpty()) {
                        int val = chunks.get(chunks.size() - 1);
                        progressDialog.progressBar.setIndeterminate(false);
                        progressDialog.progressBar.setValue(val);
                        progressDialog.label.setText("正在合并数据... " + val + "%");
                    }
                }
                @Override
                protected void done() { progressDialog.dialog.dispose(); }
            };
            worker.execute();
            progressDialog.dialog.setVisible(true);
            try {
                DuplicateColumnMerger.MergeReport report = worker.get();
                String summary = report.formatSummary();
                ctx.log("处理完成: " + outPath);
                ctx.log(summary);
                JOptionPane.showMessageDialog(mainPanel, "处理完成！\n" + summary, "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ctx.log("失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }
}

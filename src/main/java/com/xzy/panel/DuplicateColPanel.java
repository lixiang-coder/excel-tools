package com.xzy.panel;

import com.xzy.DuplicateColumnMerger;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
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
        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        JPanel filePanel = new JPanel(new GridBagLayout());
        filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1;
        JTextField inputPath = new JTextField();
        inputPath.setEditable(false);
        JButton selectInput = new JButton("选择输入文件");
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(new JLabel("输入文件路径:"), gbc);
        gbc.gridx = 1; gbc.gridy = 0; gbc.gridwidth = 2;
        filePanel.add(inputPath, gbc);
        gbc.gridx = 3; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(selectInput, gbc);
        mainPanel.add(filePanel);

        JPanel headerPanel = new JPanel(new GridBagLayout());
        headerPanel.setBorder(BorderFactory.createTitledBorder("表头选择区"));
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
        headerPanel.add(new JLabel("重复列:"), gbc);
        JComboBox<String> customerIdCombo = new JComboBox<>();
        customerIdCombo.setPreferredSize(new Dimension(180, 28));
        gbc.gridx = 1; gbc.gridy = 0; gbc.gridwidth = 1;
        headerPanel.add(customerIdCombo, gbc);
        gbc.gridx = 2; gbc.gridy = 0; gbc.gridwidth = 1;
        headerPanel.add(new JLabel("要合并的列(可多选):"), gbc);
        DefaultListModel<String> mergeColsModel = new DefaultListModel<>();
        JList<String> mergeColsList = new JList<>(mergeColsModel);
        mergeColsList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        JScrollPane mergeColsScroll = new JScrollPane(mergeColsList);
        mergeColsScroll.setPreferredSize(new Dimension(220, 150));
        gbc.gridx = 3; gbc.gridy = 0; gbc.gridwidth = 1;
        headerPanel.add(mergeColsScroll, gbc);
        mainPanel.add(headerPanel);

        JPanel opPanel = new JPanel(new BorderLayout(10, 10));
        JButton runBtn = new JButton("开始合并");
        opPanel.add(runBtn, BorderLayout.CENTER);
        mainPanel.add(opPanel);

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
            SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                @Override
                protected Void doInBackground() {
                    try {
                        new DuplicateColumnMerger().mergeByCustomerIdWithProgress(inPath, outPath, customerIdCol, mergeCols,
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
                        progressDialog.label.setText("正在合并数据... " + val + "%");
                    }
                }
                @Override
                protected void done() { progressDialog.dialog.dispose(); }
            };
            worker.execute();
            progressDialog.dialog.setVisible(true);
            try {
                worker.get();
                ctx.log("处理完成: " + outPath);
                JOptionPane.showMessageDialog(mainPanel, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ctx.log("失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }
}

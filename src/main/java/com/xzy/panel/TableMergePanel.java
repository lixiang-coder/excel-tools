package com.xzy.panel;

import com.xzy.TableMerger;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Tab6: 表格合并工具面板
 */
public class TableMergePanel {
    private final AppContext ctx;
    private final List<File> selectedFiles = new ArrayList<>();

    public TableMergePanel(AppContext ctx) {
        this.ctx = ctx;
    }

    public JPanel createPanel() {
        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JRadioButton sheetMode = new JRadioButton("单文件多Sheet合并", true);
        JRadioButton fileMode = new JRadioButton("多文件合并");
        ButtonGroup modeGroup = new ButtonGroup();
        modeGroup.add(sheetMode);
        modeGroup.add(fileMode);
        JComboBox<MergeStrategyOption> strategyCombo = new JComboBox<>(new MergeStrategyOption[]{
                new MergeStrategyOption("自动识别（推荐）", TableMerger.MergeStrategy.AUTO),
                new MergeStrategyOption("有表头合并", TableMerger.MergeStrategy.HEADER_BASED),
                new MergeStrategyOption("无表头合并（WPS样式）", TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER)
        });

        JPanel modePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 16, 5));
        modePanel.setBorder(BorderFactory.createTitledBorder("合并模式"));
        modePanel.add(sheetMode);
        modePanel.add(fileMode);
        modePanel.add(new JLabel("数据结构:"));
        modePanel.add(strategyCombo);
        mainPanel.add(modePanel);

        JPanel filePanel = new JPanel(new GridBagLayout());
        filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1;

        JTextArea selectedFilesArea = new JTextArea(6, 60);
        selectedFilesArea.setEditable(false);
        selectedFilesArea.setLineWrap(true);
        selectedFilesArea.setWrapStyleWord(true);

        JButton selectInputButton = new JButton("选择输入文件");
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        filePanel.add(new JLabel("输入文件:"), gbc);
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        filePanel.add(new JScrollPane(selectedFilesArea), gbc);
        gbc.gridx = 3;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        filePanel.add(selectInputButton, gbc);

        JTextField outputPath = new JTextField();
        outputPath.setEditable(false);
        JButton selectOutputButton = new JButton("选择输出文件");
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        filePanel.add(new JLabel("输出文件:"), gbc);
        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        filePanel.add(outputPath, gbc);
        gbc.gridx = 3;
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        filePanel.add(selectOutputButton, gbc);
        mainPanel.add(filePanel);

        JPanel infoPanel = new JPanel(new BorderLayout());
        infoPanel.setBorder(BorderFactory.createTitledBorder("合并规则"));
        JTextArea infoArea = new JTextArea();
        infoArea.setEditable(false);
        infoArea.setLineWrap(true);
        infoArea.setWrapStyleWord(true);
        infoArea.setText("自动识别会判断数据是否有表头；无表头时按WPS样式输出“报告”和“总表”，并在总表第1列添加来源标识。有表头时仍以第一张有效表的表头为准。");
        infoPanel.add(infoArea, BorderLayout.CENTER);
        mainPanel.add(infoPanel);

        JPanel opPanel = new JPanel(new BorderLayout(10, 10));
        JButton runButton = new JButton("开始合并");
        opPanel.add(runButton, BorderLayout.CENTER);
        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(opPanel);

        sheetMode.addActionListener(e -> clearInputs(selectedFilesArea));
        fileMode.addActionListener(e -> clearInputs(selectedFilesArea));

        selectInputButton.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fc.setMultiSelectionEnabled(fileMode.isSelected());
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                selectedFiles.clear();
                if (fileMode.isSelected()) {
                    File[] files = fc.getSelectedFiles();
                    for (File file : files) {
                        selectedFiles.add(file);
                    }
                } else {
                    selectedFiles.add(fc.getSelectedFile());
                }
                selectedFilesArea.setText(formatSelectedFiles());
                ctx.log("已选择待合并文件: " + selectedFiles.size() + " 个");
            }
        });

        selectOutputButton.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
            fc.setDialogTitle("选择输出文件");
            if (fc.showSaveDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                String path = fc.getSelectedFile().getAbsolutePath();
                if (!path.toLowerCase().endsWith(".xlsx")) {
                    path += ".xlsx";
                }
                outputPath.setText(path);
                ctx.log("已选择输出文件: " + path);
            }
        });

        runButton.addActionListener(e -> {
            if (selectedFiles.isEmpty()) {
                showWarning(mainPanel, "请先选择输入文件");
                return;
            }
            if (sheetMode.isSelected() && selectedFiles.size() != 1) {
                showWarning(mainPanel, "单文件多Sheet合并只能选择一个输入文件");
                return;
            }
            if (outputPath.getText() == null || outputPath.getText().trim().isEmpty()) {
                showWarning(mainPanel, "请先选择输出文件");
                return;
            }
            MergeStrategyOption option = (MergeStrategyOption) strategyCombo.getSelectedItem();
            TableMerger.MergeStrategy strategy = option == null ? TableMerger.MergeStrategy.AUTO : option.strategy;
            runMerge(mainPanel, sheetMode.isSelected(), outputPath.getText(), strategy);
        });

        return mainPanel;
    }

    private void runMerge(JPanel mainPanel, boolean sheetMode, String outputPath,
                          TableMerger.MergeStrategy strategy) {
        ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "表格合并中");
        SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
            @Override
            protected Void doInBackground() {
                try {
                    TableMerger merger = new TableMerger();
                    if (sheetMode) {
                        merger.mergeSheetsInWorkbook(selectedFiles.get(0).getAbsolutePath(), outputPath, strategy,
                                (cur, total) -> publish((int) (cur * 100.0 / total)));
                    } else {
                        merger.mergeFiles(new ArrayList<>(selectedFiles), outputPath, strategy,
                                (cur, total) -> publish((int) (cur * 100.0 / total)));
                    }
                } catch (Exception ex) {
                    throw new RuntimeException(ex);
                }
                return null;
            }

            @Override
            protected void process(List<Integer> chunks) {
                if (!chunks.isEmpty()) {
                    int value = chunks.get(chunks.size() - 1);
                    progressDialog.progressBar.setIndeterminate(false);
                    progressDialog.progressBar.setValue(value);
                    progressDialog.label.setText("正在合并数据... " + value + "%");
                }
            }

            @Override
            protected void done() {
                progressDialog.dialog.dispose();
            }
        };
        worker.execute();
        progressDialog.dialog.setVisible(true);
        try {
            worker.get();
            ctx.log("表格合并完成: " + outputPath);
            JOptionPane.showMessageDialog(mainPanel, "表格合并完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            String message = rootMessage(ex);
            ctx.log("表格合并失败: " + message);
            JOptionPane.showMessageDialog(mainPanel, "表格合并失败: " + message, "错误", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void clearInputs(JTextArea selectedFilesArea) {
        selectedFiles.clear();
        selectedFilesArea.setText("");
    }

    private String formatSelectedFiles() {
        StringBuilder builder = new StringBuilder();
        for (File file : selectedFiles) {
            if (builder.length() > 0) {
                builder.append('\n');
            }
            builder.append(file.getAbsolutePath());
        }
        return builder.toString();
    }

    private void showWarning(JPanel mainPanel, String message) {
        ctx.log(message);
        JOptionPane.showMessageDialog(mainPanel, message, "提示", JOptionPane.WARNING_MESSAGE);
    }

    private String rootMessage(Throwable throwable) {
        Throwable current = throwable;
        while (current.getCause() != null) {
            current = current.getCause();
        }
        return current.getMessage() == null ? current.toString() : current.getMessage();
    }

    private static class MergeStrategyOption {
        private final String label;
        private final TableMerger.MergeStrategy strategy;

        private MergeStrategyOption(String label, TableMerger.MergeStrategy strategy) {
            this.label = label;
            this.strategy = strategy;
        }

        @Override
        public String toString() {
            return label;
        }
    }
}

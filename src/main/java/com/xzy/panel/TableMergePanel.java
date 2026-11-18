package com.xzy.panel;

import com.xzy.TableMerger;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
import com.xzy.ui.Workbench;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

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
        JRadioButton sheetMode = new JRadioButton("单个文件的多个工作表", true);
        JRadioButton fileMode = new JRadioButton("多个文件的首个工作表");
        ButtonGroup modeGroup = new ButtonGroup();
        modeGroup.add(sheetMode);
        modeGroup.add(fileMode);
        JPanel modePanel = new JPanel(new FlowLayout(FlowLayout.LEADING, 0, 0));
        modePanel.add(sheetMode);
        modePanel.add(Box.createHorizontalStrut(16));
        modePanel.add(fileMode);

        JComboBox<MergeStrategyOption> strategyCombo = new JComboBox<>(new MergeStrategyOption[]{
                new MergeStrategyOption("自动识别（推荐）", TableMerger.MergeStrategy.AUTO),
                new MergeStrategyOption("有表头合并", TableMerger.MergeStrategy.HEADER_BASED),
                new MergeStrategyOption("无表头合并（WPS 样式）", TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER)
        });
        JLabel strategyLabel = new JLabel("表头处理");
        strategyLabel.setLabelFor(strategyCombo);
        JPanel strategyPanel = new JPanel(new FlowLayout(FlowLayout.LEADING, 0, 0));
        strategyPanel.add(strategyLabel);
        strategyPanel.add(Box.createHorizontalStrut(12));
        strategyPanel.add(strategyCombo);

        JPanel optionsPanel = new JPanel(new BorderLayout(0, 10));
        optionsPanel.add(modePanel, BorderLayout.NORTH);
        optionsPanel.add(strategyPanel, BorderLayout.CENTER);
        optionsPanel.add(new JLabel("<html>有表头时，以首张有效表的表头为准。<br>"
                + "无表头时，生成报告和总表，并在总表首列标记来源。</html>"), BorderLayout.SOUTH);

        JTextArea selectedFilesArea = new JTextArea(3, 16);
        selectedFilesArea.setEditable(false);
        selectedFilesArea.setLineWrap(true);
        selectedFilesArea.setWrapStyleWord(true);
        JTextField outputPath = new JTextField();
        outputPath.setEditable(false);
        JButton selectInputButton = new JButton("选择输入文件");
        JButton selectOutputButton = new JButton("选择输出文件");

        JPanel filePanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.anchor = GridBagConstraints.LINE_START;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(0, 0, 8, 12);
        JLabel inputLabel = new JLabel("输入文件");
        inputLabel.setLabelFor(selectedFilesArea);
        gbc.gridx = 0;
        gbc.gridy = 0;
        filePanel.add(inputLabel, gbc);
        gbc.gridx = 1;
        gbc.weightx = 1;
        filePanel.add(new JScrollPane(selectedFilesArea), gbc);
        gbc.gridx = 2;
        gbc.weightx = 0;
        gbc.insets = new Insets(0, 0, 8, 0);
        filePanel.add(selectInputButton, gbc);

        JLabel outputLabel = new JLabel("输出文件");
        outputLabel.setLabelFor(outputPath);
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.insets = new Insets(0, 0, 0, 12);
        filePanel.add(outputLabel, gbc);
        gbc.gridx = 1;
        gbc.weightx = 1;
        filePanel.add(outputPath, gbc);
        gbc.gridx = 2;
        gbc.weightx = 0;
        gbc.insets = new Insets(0, 0, 0, 0);
        filePanel.add(selectOutputButton, gbc);

        JButton runButton = Workbench.primaryButton("开始合并");
        JPanel mainPanel = Workbench.formPage(runButton,
                Workbench.formSection("合并设置", optionsPanel),
                Workbench.formSection("输入与输出", filePanel));

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
            Throwable cause = rootCause(ex);
            if (cause instanceof TableMerger.AmbiguousHeadersException) {
                Object[] choices = {"无表头合并（保留首行）", "有表头合并（按列名对齐）", "取消"};
                int choice = JOptionPane.showOptionDialog(mainPanel, cause.getMessage(), "请选择表头处理方式",
                        JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, choices, choices[0]);
                if (choice == 0 || choice == 1) {
                    runMerge(mainPanel, sheetMode, outputPath, choice == 0
                            ? TableMerger.MergeStrategy.WPS_STYLE_NO_HEADER : TableMerger.MergeStrategy.HEADER_BASED);
                }
                return;
            }
            String message = cause.getMessage() == null ? cause.toString() : cause.getMessage();
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

    private Throwable rootCause(Throwable throwable) {
        Throwable current = throwable;
        // 跳过后台任务包装，保留业务异常提供的源文件及单元格位置。
        while (current.getCause() != null &&
                (current instanceof ExecutionException || current.getClass() == RuntimeException.class)) {
            current = current.getCause();
        }
        return current;
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

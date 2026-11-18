package com.xzy;

import com.xzy.util.ExcelUtils;
import com.xzy.util.SafeExport;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Color;
import java.awt.Font;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

/**
 * Excel工具类：按指定重复列合并多行数据。
 * @deprecated 已弃用独立运行模式，功能已整合到 ExcelTools 的 Tab 中
 */
@Deprecated
public class DuplicateColumnMerger {
    private static final Logger logger = LogManager.getLogger(DuplicateColumnMerger.class);
    private static final String SEPARATOR = "|";

    /**
     * 合并Excel中指定列（重复列）相同的多行数据
     */
    public void mergeByKeyColumn(String inputFilePath, String outputFilePath, String keyColumn) throws IOException {
        File inputFile = new File(inputFilePath);
        File outputFile = new File(outputFilePath);
        if (!inputFile.exists() || !inputFile.canRead()) {
            throw new IOException("无法读取源文件: " + inputFilePath + "，请检查文件是否存在、权限是否正确，或是否被其他程序占用。");
        }
        try (SafeExport export = SafeExport.begin(outputFile, inputFile)) {
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(fis);
                 SXSSFWorkbook outputWorkbook = new SXSSFWorkbook();
                 FileOutputStream fos = new FileOutputStream(export.temporaryFile())) {
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) throw new IOException("源文件没有数据，请检查文件内容。");
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) throw new IOException("源文件没有表头，请检查文件格式。");
                int colCount = headerRow.getLastCellNum();
                int keyCol = -1;
                java.util.List<String> headers = new java.util.ArrayList<>();
                for (int i = 0; i < colCount; i++) {
                    Cell cell = headerRow.getCell(i);
                    String h = (cell != null) ? cell.getStringCellValue() : null;
                    headers.add(h);
                    if (keyColumn.equalsIgnoreCase(h)) keyCol = i;
                }
                if (keyCol == -1) throw new IOException("未找到重复列 '" + keyColumn + "'，请检查表头。");
                Map<String, java.util.List<java.util.List<String>>> merged = new LinkedHashMap<>();
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    Cell keyCell = row.getCell(keyCol);
                    String key = (keyCell != null) ? ExcelUtils.getCellValueAsString(keyCell) : null;
                    if (key == null || key.isEmpty()) continue;
                    java.util.List<java.util.List<String>> colVals = merged.computeIfAbsent(key, k -> {
                        java.util.List<java.util.List<String>> l = new java.util.ArrayList<>();
                        for (int j = 0; j < colCount; j++) l.add(new java.util.ArrayList<>());
                        return l;
                    });
                    for (int j = 0; j < colCount; j++) {
                        Cell cell = row.getCell(j);
                        String v = (cell != null) ? ExcelUtils.getCellValueAsString(cell) : null;
                        if (v != null && !v.trim().isEmpty()) colVals.get(j).add(v.trim());
                    }
                }
                Sheet outSheet = outputWorkbook.createSheet("Sheet1");
                Row outHeader = outSheet.createRow(0);
                for (int i = 0; i < headers.size(); i++) {
                    outHeader.createCell(i).setCellValue(headers.get(i));
                }
                int rowNum = 1;
                for (Map.Entry<String, java.util.List<java.util.List<String>>> entry : merged.entrySet()) {
                    Row outRow = outSheet.createRow(rowNum++);
                    java.util.List<java.util.List<String>> colVals = entry.getValue();
                    for (int j = 0; j < colVals.size(); j++) {
                        Set<String> uniq = new LinkedHashSet<>(colVals.get(j));
                        String mergedVal = String.join(SEPARATOR, uniq);
                        outRow.createCell(j).setCellValue(mergedVal);
                    }
                }
                outputWorkbook.write(fos);
                outputWorkbook.dispose();
            }
            export.commit();
        } catch (Exception e) {
            throw new IOException("处理Excel文件时发生错误: " + e.getMessage(), e);
        }
    }

    private static JDialog createProgressDialog(JFrame parent, String title, String message, boolean indeterminate, JProgressBar[] progressBarHolder, JLabel[] percentLabelHolder) {
        JDialog dialog = new JDialog(parent, title, true);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.setResizable(false);
        dialog.setUndecorated(false);
        JPanel panel = new JPanel(new BorderLayout(10, 5));
        panel.setBorder(BorderFactory.createEmptyBorder(8, 12, 8, 12));
        JPanel msgPanel = new JPanel(new BorderLayout());
        JLabel msgLabel = new JLabel(message);
        msgLabel.setFont(new Font("Microsoft YaHei", Font.PLAIN, 15));
        JLabel percentLabel = new JLabel(indeterminate ? "" : "0%");
        percentLabel.setFont(new Font("Microsoft YaHei", Font.PLAIN, 15));
        percentLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        msgPanel.add(msgLabel, BorderLayout.WEST);
        msgPanel.add(percentLabel, BorderLayout.EAST);
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(indeterminate);
        progressBar.setStringPainted(false);
        progressBar.setFont(new Font("Microsoft YaHei", Font.PLAIN, 15));
        progressBar.setForeground(Color.GREEN);
        panel.add(msgPanel, BorderLayout.NORTH);
        panel.add(progressBar, BorderLayout.CENTER);
        dialog.add(panel);
        dialog.pack();
        dialog.setLocationRelativeTo(parent);
        if (progressBarHolder != null && progressBarHolder.length > 0) progressBarHolder[0] = progressBar;
        if (percentLabelHolder != null && percentLabelHolder.length > 0) percentLabelHolder[0] = percentLabel;
        return dialog;
    }

    public static void main(String[] args) {
        IOUtils.setByteArrayMaxOverride(200_000_000);
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                logger.error("设置外观失败", e);
            }
            JFrame frame = new JFrame("重复列合并器");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(700, 420);
            frame.setLocationRelativeTo(null);
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
            JTextField outputPath = new JTextField();
            inputPath.setEditable(false);
            outputPath.setEditable(false);
            JButton selectInput = new JButton("选择输入文件");
            JButton selectOutput = new JButton("选择输出文件");
            gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
            filePanel.add(new JLabel("输入文件路径:"), gbc);
            gbc.gridx = 1; gbc.gridy = 0; gbc.gridwidth = 2;
            filePanel.add(inputPath, gbc);
            gbc.gridx = 3; gbc.gridy = 0; gbc.gridwidth = 1;
            filePanel.add(selectInput, gbc);
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1;
            filePanel.add(new JLabel("输出文件路径:"), gbc);
            gbc.gridx = 1; gbc.gridy = 1; gbc.gridwidth = 2;
            filePanel.add(outputPath, gbc);
            gbc.gridx = 3; gbc.gridy = 1; gbc.gridwidth = 1;
            filePanel.add(selectOutput, gbc);
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
            JLabel status = new JLabel("就绪");
            status.setHorizontalAlignment(SwingConstants.CENTER);
            opPanel.add(runBtn, BorderLayout.CENTER);
            opPanel.add(status, BorderLayout.SOUTH);
            mainPanel.add(opPanel);

            frame.add(mainPanel);
            frame.setVisible(true);

            selectInput.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    inputPath.setText(fc.getSelectedFile().getAbsolutePath());
                    status.setText("已选择输入: " + inputPath.getText());
                    JProgressBar[] progressBarHolder = new JProgressBar[1];
                    JLabel[] percentLabelHolder = new JLabel[1];
                    JDialog progressDialog = createProgressDialog(frame, "处理中", "正在加载输入文件表头，请稍候...", true, progressBarHolder, percentLabelHolder);
                    SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                        java.util.List<String> headers = null;
                        @Override
                        protected Void doInBackground() throws Exception {
                            try {
                                headers = ExcelUtils.readHeadersOnly(new File(inputPath.getText()));
                            } catch (Exception ex) {
                                headers = null;
                            }
                            return null;
                        }
                        @Override
                        protected void done() {
                            progressDialog.dispose();
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
                    progressDialog.setVisible(true);
                }
            });
            selectOutput.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    String outPath = fc.getSelectedFile().getAbsolutePath();
                    if (!outPath.toLowerCase().endsWith(".xlsx")) outPath += ".xlsx";
                    outputPath.setText(outPath);
                    status.setText("已选择输出: " + outputPath.getText());
                }
            });
            runBtn.addActionListener(e -> {
                String inPath = inputPath.getText();
                String outPath = outputPath.getText();
                if (inPath == null || inPath.isEmpty() || outPath == null || outPath.isEmpty()) {
                    status.setText("请先选择输入和输出文件");
                    return;
                }
                if (customerIdCombo.getSelectedItem() == null || mergeColsList.getSelectedValuesList().isEmpty()) {
                    status.setText("请选择CustomerId列和要合并的列");
                    return;
                }
                String customerIdCol = (String) customerIdCombo.getSelectedItem();
                java.util.List<String> mergeCols = mergeColsList.getSelectedValuesList();
                JProgressBar[] progressBarHolder = new JProgressBar[1];
                JLabel[] percentLabelHolder = new JLabel[1];
                JDialog progressDialog = createProgressDialog(frame, "处理中", "正在合并数据...", false, progressBarHolder, percentLabelHolder);
                JProgressBar progressBar = progressBarHolder[0];
                JLabel percentLabel = percentLabelHolder[0];
                SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                    @Override
                    protected Void doInBackground() {
                        try {
                            new DuplicateColumnMerger().mergeByCustomerIdWithProgress(inPath, outPath, customerIdCol, mergeCols, (cur, total) -> publish((int) (cur * 100.0 / total)));
                        } catch (Exception ex) {
                            throw new RuntimeException(ex);
                        }
                        return null;
                    }
                    @Override
                    protected void process(java.util.List<Integer> chunks) {
                        if (!chunks.isEmpty()) {
                            int val = chunks.get(chunks.size() - 1);
                            progressBar.setValue(val);
                            if (percentLabel != null) percentLabel.setText(val + "%");
                        }
                    }
                    @Override
                    protected void done() {
                        progressDialog.dispose();
                    }
                };
                worker.execute();
                progressDialog.setVisible(true);
                try {
                    worker.get();
                    status.setText("处理完成: " + outPath);
                    JOptionPane.showMessageDialog(frame, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    status.setText("失败: " + ex.getMessage());
                    JOptionPane.showMessageDialog(frame, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            });
        });
    }

    public void mergeByCustomerIdWithProgress(String inputFilePath, String outputFilePath, String customerIdCol, java.util.List<String> mergeCols, ProgressListener progressListener) throws IOException {
        File inputFile = new File(inputFilePath);
        File outputFile = new File(outputFilePath);
        if (!inputFile.exists() || !inputFile.canRead()) {
            throw new IOException("无法读取源文件: " + inputFilePath);
        }
        try (SafeExport export = SafeExport.begin(outputFile, inputFile)) {
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(fis);
                 Workbook outputWorkbook = new SXSSFWorkbook(100);
                 FileOutputStream fos = new FileOutputStream(export.temporaryFile())) {
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) throw new IOException("源文件没有数据");
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) throw new IOException("源文件没有表头");
                int colCount = headerRow.getLastCellNum();
                int customerIdColIdx = -1;
                java.util.List<Integer> mergeColIdxs = new java.util.ArrayList<>();
                java.util.List<String> headers = new java.util.ArrayList<>();
                for (int i = 0; i < colCount; i++) {
                    Cell cell = headerRow.getCell(i);
                    String h = (cell != null) ? cell.getStringCellValue() : null;
                    headers.add(h);
                    if (customerIdCol.equalsIgnoreCase(h)) customerIdColIdx = i;
                    if (mergeCols.contains(h)) mergeColIdxs.add(i);
                }
                if (customerIdColIdx == -1) throw new IOException("未找到CustomerId列");
                Map<String, java.util.List<java.util.List<String>>> merged = new LinkedHashMap<>();
                int totalRows = sheet.getLastRowNum();
                for (int i = 1; i <= totalRows; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    Cell cidCell = row.getCell(customerIdColIdx);
                    String cid = (cidCell != null) ? ExcelUtils.getCellValueAsString(cidCell) : null;
                    if (cid == null || cid.isEmpty()) continue;
                    java.util.List<java.util.List<String>> colVals = merged.computeIfAbsent(cid, k -> {
                        java.util.List<java.util.List<String>> l = new java.util.ArrayList<>();
                        for (int j = 0; j < colCount; j++) l.add(new java.util.ArrayList<>());
                        return l;
                    });
                    for (int j = 0; j < colCount; j++) {
                        Cell cell = row.getCell(j);
                        String v = (cell != null) ? ExcelUtils.getCellValueAsString(cell) : null;
                        if (v != null && !v.trim().isEmpty()) colVals.get(j).add(v.trim());
                    }
                    if (progressListener != null && (i % 100 == 0 || i == totalRows)) {
                        progressListener.onProgress(i, totalRows);
                    }
                }
                Sheet outSheet = outputWorkbook.createSheet("Sheet1");
                Row outHeader = outSheet.createRow(0);
                for (int i = 0; i < headers.size(); i++) {
                    outHeader.createCell(i).setCellValue(headers.get(i));
                }
                int rowNum = 1;
                for (Map.Entry<String, java.util.List<java.util.List<String>>> entry : merged.entrySet()) {
                    Row outRow = outSheet.createRow(rowNum++);
                    java.util.List<java.util.List<String>> colVals = entry.getValue();
                    for (int j = 0; j < colVals.size(); j++) {
                        Set<String> uniq = new LinkedHashSet<>(colVals.get(j));
                        String mergedVal;
                        if (mergeColIdxs.contains(j)) {
                            mergedVal = String.join(SEPARATOR, uniq);
                        } else {
                            mergedVal = uniq.isEmpty() ? "" : uniq.iterator().next();
                        }
                        outRow.createCell(j).setCellValue(mergedVal);
                    }
                }
                outputWorkbook.write(fos);
                ((SXSSFWorkbook) outputWorkbook).dispose();
            }
            export.commit();
        }
    }

    public interface ProgressListener {
        void onProgress(int current, int total);
    }
}

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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;

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
        try {
            merge(inputFilePath, outputFilePath, keyColumn, header -> true, null);
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
                                headers = ExcelUtils.readNamedHeaders(new File(inputPath.getText()));
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
                SwingWorker<MergeReport, Integer> worker = new SwingWorker<MergeReport, Integer>() {
                    @Override
                    protected MergeReport doInBackground() {
                        try {
                            return new DuplicateColumnMerger().mergeWithReport(inPath, outPath, customerIdCol, mergeCols,
                                    (cur, total) -> publish((int) (cur * 100.0 / total)));
                        } catch (Exception ex) {
                            throw new RuntimeException(ex);
                        }
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
                    MergeReport report = worker.get();
                    String summary = report.formatSummary();
                    status.setText("处理完成: " + summary);
                    JOptionPane.showMessageDialog(frame, "处理完成！\n" + summary + "\n结果已保存到: " + outPath,
                            "成功", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    status.setText("失败: " + ex.getMessage());
                    JOptionPane.showMessageDialog(frame, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            });
        });
    }

    public void mergeByCustomerIdWithProgress(String inputFilePath, String outputFilePath, String keyColumn,
                                               List<String> mergeCols, ProgressListener progress) throws IOException {
        mergeWithReport(inputFilePath, outputFilePath, keyColumn, mergeCols, progress);
    }

    public MergeReport mergeWithReport(String inputFilePath, String outputFilePath, String keyColumn,
                                       List<String> mergeCols, ProgressListener progress) throws IOException {
        return merge(inputFilePath, outputFilePath, keyColumn, mergeCols::contains, progress);
    }

    private MergeReport merge(String inputPath, String outputPath, String keyColumn,
                              Predicate<String> shouldMerge, ProgressListener progress) throws IOException {
        File input = new File(inputPath);
        if (!input.isFile() || !input.canRead()) {
            throw new IOException("无法读取源文件: " + inputPath);
        }
        try (SafeExport export = SafeExport.begin(new File(outputPath), input)) {
            MergeReport report;
            try (FileInputStream in = new FileInputStream(input);
                 Workbook source = WorkbookFactory.create(in);
                 SXSSFWorkbook result = new SXSSFWorkbook(100)) {
                try {
                    Sheet sourceSheet = source.getSheetAt(0);
                    Row header = sourceSheet.getRow(0);
                    if (header == null) throw new IOException("源文件没有表头");
                    int keyIndex = -1;
                    boolean[] mergeColumns = new boolean[Math.max(0, header.getLastCellNum())];
                    Sheet output = result.createSheet("Sheet1");
                    Row outputHeader = output.createRow(0);
                    for (int column = 0; column < mergeColumns.length; column++) {
                        Cell headerCell = header.getCell(column);
                        String name = headerCell == null ? "" : headerCell.getStringCellValue();
                        if (!name.isEmpty() && keyColumn.equalsIgnoreCase(name)) keyIndex = column;
                        mergeColumns[column] = shouldMerge.test(name);
                        if (headerCell != null) outputHeader.createCell(column).setCellValue(name);
                    }
                    if (keyIndex == -1) throw new IOException("未找到重复列 '" + keyColumn + "'");
                    report = mergeRows(sourceSheet, output, keyIndex, mergeColumns, progress);
                    try (FileOutputStream out = new FileOutputStream(export.temporaryFile())) {
                        result.write(out);
                    }
                } finally {
                    result.dispose();
                }
            }
            export.commit();
            return report;
        }
    }

    private static MergeReport mergeRows(Sheet source, Sheet output, int keyIndex,
                                         boolean[] mergeColumns, ProgressListener progress) {
        Map<String, List<Row>> keyedGroups = new HashMap<>();
        // 每个正常主键仅登记首行，空主键的数据行各自登记，二者保持输入顺序。
        List<Row> outputOrder = new ArrayList<>();
        int inputRows = 0;
        int blankKeyRows = 0;
        int lastRow = source.getLastRowNum();
        for (Row row : source) {
            int rowIndex = row.getRowNum();
            if (rowIndex == 0) continue;
            Cell keyCell = row.getCell(keyIndex);
            String key = ExcelUtils.getCellValueAsString(keyCell);
            if (!isBlank(key)) {
                List<Row> group = keyedGroups.get(key);
                if (group == null) {
                    group = new ArrayList<>();
                    keyedGroups.put(key, group);
                    outputOrder.add(row);
                }
                group.add(row);
                inputRows++;
            } else if (hasData(row)) {
                outputOrder.add(row);
                inputRows++;
                // 错误值也不参加分组，但不是空主键。
                if (!isErrorCell(keyCell)) blankKeyRows++;
            }
            if (progress != null && (rowIndex % 100 == 0 || rowIndex == lastRow)) {
                progress.onProgress(rowIndex, lastRow);
            }
        }
        int outputRows = 0;
        for (Row first : outputOrder) {
            Row target = output.createRow(++outputRows);
            List<Row> group = keyedGroups.get(ExcelUtils.getCellValueAsString(first.getCell(keyIndex)));
            if (group == null) {
                writeIndependentRow(first, target, keyIndex);
            } else {
                writeGroup(group, target, mergeColumns);
            }
        }
        return new MergeReport(inputRows, outputRows, blankKeyRows);
    }

    private static void writeGroup(List<Row> rows, Row target, boolean[] mergeColumns) {
        for (int column = 0; column < mergeColumns.length; column++) {
            Set<String> values = new LinkedHashSet<>();
            for (Row row : rows) {
                String text = ExcelUtils.getCellValueAsString(row.getCell(column));
                if (!text.isEmpty()) {
                    values.add(text);
                    if (!mergeColumns[column]) break;
                }
            }
            target.createCell(column).setCellValue(String.join(SEPARATOR, values));
        }
    }

    private static void writeIndependentRow(Row source, Row target, int keyIndex) {
        for (Cell cell : source) {
            String value = independentCellText(cell);
            if (value.isEmpty() || (cell.getColumnIndex() == keyIndex && isBlank(value))) continue;
            target.createCell(cell.getColumnIndex()).setCellValue(value);
        }
    }

    private static boolean hasData(Row row) {
        for (Cell cell : row) {
            if (!isBlank(independentCellText(cell))) return true;
        }
        return false;
    }

    private static String independentCellText(Cell cell) {
        if (isErrorCell(cell)) return FormulaError.forInt(cell.getErrorCellValue()).getString();
        if (cell != null && cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
        return ExcelUtils.getCellValueAsString(cell);
    }

    private static boolean isErrorCell(Cell cell) {
        return cell != null && (cell.getCellType() == CellType.ERROR ||
                (cell.getCellType() == CellType.FORMULA && cell.getCachedFormulaResultType() == CellType.ERROR));
    }

    private static boolean isBlank(String value) {
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (!Character.isWhitespace(c) && !Character.isSpaceChar(c)) return false;
        }
        return true;
    }

    public static final class MergeReport {
        private final int inputRows;
        private final int outputRows;
        private final int blankKeyRows;

        private MergeReport(int inputRows, int outputRows, int blankKeyRows) {
            this.inputRows = inputRows;
            this.outputRows = outputRows;
            this.blankKeyRows = blankKeyRows;
        }

        public int getInputRows() { return inputRows; }
        public int getOutputRows() { return outputRows; }
        public int getBlankKeyRows() { return blankKeyRows; }

        public String formatSummary() {
            return "有效输入 " + inputRows + " 行，输出 " + outputRows + " 行，其中空主键逐条保留 " + blankKeyRows + " 行。";
        }
    }

    public interface ProgressListener {
        void onProgress(int current, int total);
    }
}

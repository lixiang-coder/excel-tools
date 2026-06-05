package com.xzy;

import com.xzy.util.ExcelUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Font;
import java.awt.Dimension;
import java.awt.Insets;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.BorderLayout;
import java.awt.GridBagConstraints;
import java.io.*;
import java.util.*;
import java.awt.Color;

/**
 * Excel跨表主键数据填充工具
 * @deprecated 已弃用独立运行模式，功能已整合到 ExcelTools 的 Tab 中
 */
@Deprecated
public class CrossSheetFiller {
    private static final Logger logger = LogManager.getLogger(CrossSheetFiller.class);
    /**
     * 执行跨表主键数据填充（带进度监听）
     */
    public void fill(String aPath, String bPath, String outputPath,
                     String aKeyHeader, String aValueHeader,
                     String bKeyHeader, String bValueHeader,
                     ProgressListener progressListener) throws IOException {
        File aFile = new File(aPath);
        File bFile = new File(bPath);
        File outFile = new File(outputPath);
        if (!aFile.exists() || !aFile.canRead()) throw new IOException("无法读取源表: " + aPath);
        if (!bFile.exists() || !bFile.canRead()) throw new IOException("无法读取目标表: " + bPath);
        File outDir = outFile.getParentFile();
        if (outDir != null && !outDir.exists()) {
            if (!outDir.mkdirs()) throw new IOException("无法创建输出目录: " + outDir.getAbsolutePath());
        }
        if (outDir != null && !outDir.canWrite())
            throw new IOException("无法写入输出目录: " + outDir.getAbsolutePath());
        try (FileInputStream aFis = new FileInputStream(aFile);
             FileInputStream bFis = new FileInputStream(bFile);
             Workbook aWb = WorkbookFactory.create(aFis);
             Workbook bWb = WorkbookFactory.create(bFis);
             Workbook outWb = new SXSSFWorkbook(100);
             FileOutputStream fos = new FileOutputStream(outFile)) {
            Sheet aSheet = aWb.getSheetAt(0);
            Sheet bSheet = bWb.getSheetAt(0);
            if (aSheet == null || bSheet == null) throw new IOException("源表或目标表无数据");
            Row aHeaderRow = aSheet.getRow(0);
            Row bHeaderRow = bSheet.getRow(0);
            if (aHeaderRow == null || bHeaderRow == null) throw new IOException("源表或目标表无表头");
            // ---增强容错：建立表头容错Map---
            Map<String, Integer> aHeaderMap = new HashMap<>();
            for (int i = 0; i < aHeaderRow.getLastCellNum(); i++) {
                String h = ExcelUtils.getCellValueAsString(aHeaderRow.getCell(i));
                if (h != null) aHeaderMap.put(ExcelUtils.cleanString(h).toLowerCase(), i);
            }
            Map<String, Integer> bHeaderMap = new HashMap<>();
            for (int i = 0; i < bHeaderRow.getLastCellNum(); i++) {
                String h = ExcelUtils.getCellValueAsString(bHeaderRow.getCell(i));
                if (h != null) bHeaderMap.put(ExcelUtils.cleanString(h).toLowerCase(), i);
            }
            int aKeyIdx = aHeaderMap.getOrDefault(ExcelUtils.cleanString(aKeyHeader).toLowerCase(), -1);
            int aValIdx = aHeaderMap.getOrDefault(ExcelUtils.cleanString(aValueHeader).toLowerCase(), -1);
            int bKeyIdx = bHeaderMap.getOrDefault(ExcelUtils.cleanString(bKeyHeader).toLowerCase(), -1);
            int bValIdx = bHeaderMap.getOrDefault(ExcelUtils.cleanString(bValueHeader).toLowerCase(), -1);
            // ---友好报错---
            if (aKeyIdx == -1)
                throw new IOException("源表未找到主键表头: " + aKeyHeader + "（已自动忽略空格和大小写）");
            if (aValIdx == -1)
                throw new IOException("源表未找到数据表头: " + aValueHeader + "（已自动忽略空格和大小写）");
            if (bKeyIdx == -1)
                throw new IOException("目标表未找到主键表头: " + bKeyHeader + "（已自动忽略空格和大小写）");
            if (bValIdx == -1)
                throw new IOException("目标表未找到目标表头: " + bValueHeader + "（已自动忽略空格和大小写）");
            // 复制目标表到输出表，始终创建每一行
            Sheet outSheet = outWb.createSheet("Sheet1");
            for (int i = 0; i <= bSheet.getLastRowNum(); i++) {
                Row src = bSheet.getRow(i);
                Row dst = outSheet.getRow(i);
                if (dst == null) dst = outSheet.createRow(i);
                if (src != null) {
                    for (int j = 0; j < src.getLastCellNum(); j++) {
                        Cell srcCell = src.getCell(j);
                        Cell dstCell = dst.createCell(j);
                        if (srcCell != null) ExcelUtils.copyCellValue(srcCell, dstCell);
                    }
                }
            }
            // 优化：建立目标表主键到所有行的映射
            Map<String, java.util.List<Row>> bKeyToRows = new HashMap<>();
            for (int j = 1; j <= outSheet.getLastRowNum(); j++) {
                Row bRow = outSheet.getRow(j);
                if (bRow == null) bRow = outSheet.createRow(j);
                String bKey = ExcelUtils.cleanString(ExcelUtils.getCellValueAsString(bRow.getCell(bKeyIdx)));
                if (bKey == null) continue;
                bKeyToRows.computeIfAbsent(bKey, k -> new java.util.ArrayList<>()).add(bRow);
            }
            // 遍历源表，直接查找目标表所有匹配行
            int total = aSheet.getLastRowNum();
            int cur = 0;
            for (int i = 1; i <= total; i++) {
                cur++;
                Row aRow = aSheet.getRow(i);
                if (aRow == null) continue;
                String aKey = ExcelUtils.cleanString(ExcelUtils.getCellValueAsString(aRow.getCell(aKeyIdx)));
                String aVal = ExcelUtils.getCellValueAsString(aRow.getCell(aValIdx));
                if (aKey == null || aKey.isEmpty()) continue;
                java.util.List<Row> bRows = bKeyToRows.get(aKey);
                if (bRows != null) {
                    for (Row bRow : bRows) {
                        Cell cell = bRow.getCell(bValIdx);
                        if (cell == null) cell = bRow.createCell(bValIdx);
                        cell.setCellValue(aVal);
                        System.out.println("填充: 源表主键:" + aKey + " -> 目标表行:" + (bRow.getRowNum() + 1) + "，值:" + aVal);
                    }
                }
                if (progressListener != null && (cur % 100 == 0 || cur == total)) {
                    progressListener.onProgress(cur, total);
                }
            }
            outWb.write(fos);
            ((SXSSFWorkbook) outWb).dispose();
        }
    }

    // 进度监听接口
    public interface ProgressListener {
        void onProgress(int current, int total);
    }

    /**
     * 命令行/交互入口
     */
    public static void main(String[] args) {
        IOUtils.setByteArrayMaxOverride(200_000_000);
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                logger.error("设置外观失败", e);
            }
            // 设置全局字体
            java.awt.Font uiFont = new java.awt.Font("Microsoft YaHei", java.awt.Font.PLAIN, 16);
            java.util.Enumeration<Object> keys = UIManager.getDefaults().keys();
            while (keys.hasMoreElements()) {
                Object key = keys.nextElement();
                Object value = UIManager.get(key);
                if (value instanceof Font) {
                    UIManager.put(key, uiFont);
                }
            }
            JFrame frame = new JFrame("跨表头填充工具");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(900, 600);
            frame.setLocationRelativeTo(null);
            frame.setLayout(new BorderLayout(10, 10));

            // 文件选择区
            JPanel filePanel = new JPanel(new GridBagLayout());
            filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(8, 8, 8, 8);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weightx = 1;

            JTextField aPath = new JTextField();
            JTextField bPath = new JTextField();
            JTextField outPath = new JTextField();
            aPath.setEditable(false);
            bPath.setEditable(false);
            outPath.setEditable(false);
            JButton selA = new JButton("选择源表");
            JButton selB = new JButton("选择目标表");
            JButton selOut = new JButton("选择输出");
            // 优化按钮字体和边框
            java.awt.Font btnFont = new java.awt.Font("Microsoft YaHei", java.awt.Font.BOLD, 16);
            selA.setFont(btnFont);
            selB.setFont(btnFont);
            selOut.setFont(btnFont);
            // 去除蓝色边框，使用无边框风格
            selA.setBorderPainted(false);
            selB.setBorderPainted(false);
            selOut.setBorderPainted(false);
            selA.setFocusPainted(false);
            selB.setFocusPainted(false);
            selOut.setFocusPainted(false);
            // 让按钮文字始终可见
            selA.setOpaque(true);
            selB.setOpaque(true);
            selOut.setOpaque(true);
            // 按钮可用/不可用时颜色
            java.awt.Color blue = new java.awt.Color(66, 133, 244);
            java.awt.Color grayBg = new java.awt.Color(220, 220, 220);
            java.awt.Color grayFont = new java.awt.Color(120, 120, 120);
            // 工具方法：视觉禁用/启用按钮
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
            // 按钮视觉和逻辑禁用控制
            final boolean[] isSelAEnabled = {true};
            final boolean[] isSelBEnabled = {true};
            final boolean[] isSelOutEnabled = {true};
            setButtonVisualEnabled.accept(selA, true);
            setButtonVisualEnabled.accept(selB, true);
            setButtonVisualEnabled.accept(selOut, true);

            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.gridwidth = 1;
            filePanel.add(new JLabel("源表路径:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 0;
            gbc.gridwidth = 2;
            filePanel.add(aPath, gbc);
            gbc.gridx = 3;
            gbc.gridy = 0;
            gbc.gridwidth = 1;
            filePanel.add(selA, gbc);
            gbc.gridx = 0;
            gbc.gridy = 1;
            gbc.gridwidth = 1;
            filePanel.add(new JLabel("目标表路径:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 1;
            gbc.gridwidth = 2;
            filePanel.add(bPath, gbc);
            gbc.gridx = 3;
            gbc.gridy = 1;
            gbc.gridwidth = 1;
            filePanel.add(selB, gbc);
            gbc.gridx = 0;
            gbc.gridy = 2;
            gbc.gridwidth = 1;
            filePanel.add(new JLabel("输出路径:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 2;
            gbc.gridwidth = 2;
            filePanel.add(outPath, gbc);
            gbc.gridx = 3;
            gbc.gridy = 2;
            gbc.gridwidth = 1;
            filePanel.add(selOut, gbc);

            // 表头展示区
            JPanel headerPanel = new JPanel(new GridLayout(2, 1, 5, 5));
            headerPanel.setBorder(BorderFactory.createTitledBorder("表头展示区"));
            JTextArea aHeadersArea = new JTextArea(2, 40);
            aHeadersArea.setEditable(false);
            aHeadersArea.setLineWrap(true);
            aHeadersArea.setBorder(BorderFactory.createTitledBorder("源表表头"));
            JScrollPane aHeadersScroll = new JScrollPane(aHeadersArea);
            JTextArea bHeadersArea = new JTextArea(2, 40);
            bHeadersArea.setEditable(false);
            bHeadersArea.setLineWrap(true);
            bHeadersArea.setBorder(BorderFactory.createTitledBorder("目标表表头"));
            JScrollPane bHeadersScroll = new JScrollPane(bHeadersArea);
            headerPanel.add(aHeadersScroll);
            headerPanel.add(bHeadersScroll);

            // 表头映射区
            JPanel mapPanel = new JPanel(new GridBagLayout());
            mapPanel.setBorder(BorderFactory.createTitledBorder("表头映射区"));
            JComboBox<String> aKeyCombo = new JComboBox<>();
            JComboBox<String> aValCombo = new JComboBox<>();
            JComboBox<String> bKeyCombo = new JComboBox<>();
            JComboBox<String> bValCombo = new JComboBox<>();
            aKeyCombo.setFont(uiFont);
            aValCombo.setFont(uiFont);
            bKeyCombo.setFont(uiFont);
            bValCombo.setFont(uiFont);
            aKeyCombo.setPreferredSize(new Dimension(180, 32));
            aValCombo.setPreferredSize(new Dimension(180, 32));
            bKeyCombo.setPreferredSize(new Dimension(180, 32));
            bValCombo.setPreferredSize(new Dimension(180, 32));
            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.gridwidth = 1;
            mapPanel.add(new JLabel("源表匹配表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 0;
            mapPanel.add(aKeyCombo, gbc);
            gbc.gridx = 2;
            gbc.gridy = 0;
            mapPanel.add(new JLabel("源表数据表头:"), gbc);
            gbc.gridx = 3;
            gbc.gridy = 0;
            mapPanel.add(aValCombo, gbc);
            gbc.gridx = 0;
            gbc.gridy = 1;
            mapPanel.add(new JLabel("目标表匹配表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 1;
            mapPanel.add(bKeyCombo, gbc);
            gbc.gridx = 2;
            gbc.gridy = 1;
            mapPanel.add(new JLabel("目标表填充表头:"), gbc);
            gbc.gridx = 3;
            gbc.gridy = 1;
            mapPanel.add(bValCombo, gbc);

            // 操作区
            JPanel opPanel = new JPanel(new BorderLayout(10, 10));
            JButton runBtn = new JButton("开始填充");
            runBtn.setFont(new java.awt.Font("Microsoft YaHei", java.awt.Font.BOLD, 18));
            runBtn.setBackground(new java.awt.Color(66, 133, 244));
            runBtn.setForeground(java.awt.Color.WHITE);
            runBtn.setFocusPainted(false);
            JLabel status = new JLabel("就绪");
            status.setHorizontalAlignment(SwingConstants.CENTER);
            status.setFont(new java.awt.Font("Microsoft YaHei", java.awt.Font.PLAIN, 16));
            opPanel.add(runBtn, BorderLayout.CENTER);
            opPanel.add(status, BorderLayout.SOUTH);

            // 主布局
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
            frame.add(mainPanel, BorderLayout.CENTER);

            // 事件
            selA.addActionListener(e -> {
                if (!isSelAEnabled[0]) return;
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    aPath.setText(fc.getSelectedFile().getAbsolutePath());
                    JProgressBar[] progressBarHolder = new JProgressBar[1];
                    JLabel[] percentLabelHolder = new JLabel[1];
                    JDialog progressDialog = createProgressDialog(frame, "处理中", "正在加载源表表头，请稍候...", true, progressBarHolder, percentLabelHolder);
                    SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                        @Override
                        protected Void doInBackground() throws Exception {
                            try {
                                java.util.List<String> headers = ExcelUtils.readHeadersOnly(new File(aPath.getText()));
                                SwingUtilities.invokeLater(() -> {
                                    aHeadersArea.setText(String.join(", ", headers));
                                    aKeyCombo.removeAllItems();
                                    aValCombo.removeAllItems();
                                    for (String h : headers) {
                                        aKeyCombo.addItem(h);
                                        aValCombo.addItem(h);
                                    }
                                    if (aKeyCombo.getItemCount() > 0) aKeyCombo.setSelectedIndex(0);
                                    if (aValCombo.getItemCount() > 0) aValCombo.setSelectedIndex(0);
                                });
                            } catch (Exception ex) {
                                SwingUtilities.invokeLater(() -> {
                                    aHeadersArea.setText("读取源表表头失败: " + ex.getMessage());
                                    aKeyCombo.removeAllItems();
                                    aValCombo.removeAllItems();
                                });
                            }
                            return null;
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
                    } catch (Exception ex) {
                        // ignore
                    }
                }
            });
            selB.addActionListener(e -> {
                if (!isSelBEnabled[0]) return;
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    bPath.setText(fc.getSelectedFile().getAbsolutePath());
                    JProgressBar[] progressBarHolder = new JProgressBar[1];
                    JLabel[] percentLabelHolder = new JLabel[1];
                    JDialog progressDialog = createProgressDialog(frame, "处理中", "正在加载目标表表头，请稍候...", true, progressBarHolder, percentLabelHolder);
                    SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                        @Override
                        protected Void doInBackground() throws Exception {
                            try {
                                java.util.List<String> headers = ExcelUtils.readHeadersOnly(new File(bPath.getText()));
                                SwingUtilities.invokeLater(() -> {
                                    bHeadersArea.setText(String.join(", ", headers));
                                    bKeyCombo.removeAllItems();
                                    bValCombo.removeAllItems();
                                    for (String h : headers) {
                                        bKeyCombo.addItem(h);
                                        bValCombo.addItem(h);
                                    }
                                    if (bKeyCombo.getItemCount() > 0) bKeyCombo.setSelectedIndex(0);
                                    if (bValCombo.getItemCount() > 0) bValCombo.setSelectedIndex(0);
                                });
                            } catch (Exception ex) {
                                SwingUtilities.invokeLater(() -> {
                                    bHeadersArea.setText("读取目标表表头失败: " + ex.getMessage());
                                    bKeyCombo.removeAllItems();
                                    bValCombo.removeAllItems();
                                });
                            }
                            return null;
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
                    } catch (Exception ex) {
                        // ignore
                    }
                }
            });
            selOut.addActionListener(e -> {
                if (!isSelOutEnabled[0]) return;
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    String p = fc.getSelectedFile().getAbsolutePath();
                    if (!p.toLowerCase().endsWith(".xlsx")) p += ".xlsx";
                    outPath.setText(p);
                }
            });
            runBtn.addActionListener(e -> {
                JProgressBar[] progressBarHolder = new JProgressBar[1];
                JLabel[] percentLabelHolder = new JLabel[1];
                JDialog progressDialog = createProgressDialog(frame, "处理中", "正在处理数据...", false, progressBarHolder, percentLabelHolder);
                JProgressBar progressBar = progressBarHolder[0];
                JLabel percentLabel = percentLabelHolder[0];
                SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                    @Override
                    protected Void doInBackground() {
                        try {
                            new CrossSheetFiller().fill(
                                    aPath.getText(), bPath.getText(), outPath.getText(),
                                    (String)aKeyCombo.getSelectedItem(), (String)aValCombo.getSelectedItem(),
                                    (String)bKeyCombo.getSelectedItem(), (String)bValCombo.getSelectedItem(),
                                    (cur, total) -> publish((int) (cur * 100.0 / total))
                            );
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
                    status.setText("处理完成: " + outPath.getText());
                    JOptionPane.showMessageDialog(frame, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    status.setText("失败: " + ex.getMessage());
                    JOptionPane.showMessageDialog(frame, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            });
            frame.setVisible(true);
        });
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
}

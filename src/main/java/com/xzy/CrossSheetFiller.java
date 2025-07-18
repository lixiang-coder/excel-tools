package com.xzy;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.*;

/**
 * Curosr 实现
 * Excel跨表主键数据填充工具
 * 支持A表主键与B表主键逐行比对，数据填充，UI交互，进度条，日志
 */
public class CrossSheetFiller {
    /**
     * 执行跨表主键数据填充（带进度监听）
     */
    public void fill(String aPath, String bPath, String outputPath,
                     String aKeyHeader, String bKeyHeader,
                     String aValueHeader, String bValueHeader,
                     ProgressListener progressListener) throws IOException {
        File aFile = new File(aPath);
        File bFile = new File(bPath);
        File outFile = new File(outputPath);
        if (!aFile.exists() || !aFile.canRead()) throw new IOException("无法读取A表: " + aPath);
        if (!bFile.exists() || !bFile.canRead()) throw new IOException("无法读取B表: " + bPath);
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
             Workbook outWb = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outFile)) {
            Sheet aSheet = aWb.getSheetAt(0);
            Sheet bSheet = bWb.getSheetAt(0);
            if (aSheet == null || bSheet == null) throw new IOException("A表或B表无数据");
            Row aHeaderRow = aSheet.getRow(0);
            Row bHeaderRow = bSheet.getRow(0);
            if (aHeaderRow == null || bHeaderRow == null) throw new IOException("A表或B表无表头");
            int aKeyIdx = -1, aValIdx = -1;
            int bKeyIdx = -1, bValIdx = -1;
            for (int i = 0; i < aHeaderRow.getLastCellNum(); i++) {
                String h = getCellValueAsString(aHeaderRow.getCell(i));
                if (h != null && cleanString(h).equalsIgnoreCase(cleanString(aKeyHeader))) aKeyIdx = i;
                if (h != null && cleanString(h).equalsIgnoreCase(cleanString(aValueHeader))) aValIdx = i;
            }
            for (int i = 0; i < bHeaderRow.getLastCellNum(); i++) {
                String h = getCellValueAsString(bHeaderRow.getCell(i));
                if (h != null && cleanString(h).equalsIgnoreCase(cleanString(bKeyHeader))) bKeyIdx = i;
                if (h != null && cleanString(h).equalsIgnoreCase(cleanString(bValueHeader))) bValIdx = i;
            }
            if (aKeyIdx == -1 || aValIdx == -1 || bKeyIdx == -1 || bValIdx == -1)
                throw new IOException("表头映射有误，请检查");
            // 复制B表到输出表，始终创建每一行
            Sheet outSheet = outWb.createSheet("Sheet1");
            for (int i = 0; i <= bSheet.getLastRowNum(); i++) {
                Row src = bSheet.getRow(i);
                Row dst = outSheet.getRow(i);
                if (dst == null) dst = outSheet.createRow(i);
                if (src != null) {
                    for (int j = 0; j < src.getLastCellNum(); j++) {
                        Cell srcCell = src.getCell(j);
                        Cell dstCell = dst.createCell(j);
                        if (srcCell != null) copyCellValue(srcCell, dstCell);
                    }
                }
            }
            // 优化：建立B表主键到所有行的映射
            Map<String, java.util.List<Row>> bKeyToRows = new HashMap<>();
            for (int j = 1; j <= outSheet.getLastRowNum(); j++) {
                Row bRow = outSheet.getRow(j);
                if (bRow == null) bRow = outSheet.createRow(j);
                String bKey = cleanString(getCellValueAsString(bRow.getCell(bKeyIdx)));
                if (bKey == null) continue;
                bKeyToRows.computeIfAbsent(bKey, k -> new java.util.ArrayList<>()).add(bRow);
            }
            // 遍历A表，直接查找B表所有匹配行
            int total = aSheet.getLastRowNum();
            int cur = 0;
            for (int i = 1; i <= total; i++) {
                cur++;
                Row aRow = aSheet.getRow(i);
                if (aRow == null) continue;
                String aKey = cleanString(getCellValueAsString(aRow.getCell(aKeyIdx)));
                String aVal = getCellValueAsString(aRow.getCell(aValIdx));
                if (aKey == null || aKey.isEmpty()) continue;
                java.util.List<Row> bRows = bKeyToRows.get(aKey);
                if (bRows != null) {
                    for (Row bRow : bRows) {
                        Cell cell = bRow.getCell(bValIdx);
                        if (cell == null) cell = bRow.createCell(bValIdx);
                        cell.setCellValue(aVal);
                        System.out.println("填充: A主键:" + aKey + " -> B表行:" + (bRow.getRowNum() + 1) + "，值:" + aVal);
                    }
                }
                if (progressListener != null && (cur % 100 == 0 || cur == total)) {
                    progressListener.onProgress(cur, total);
                }
            }
            outWb.write(fos);
            // 不再需要dispose()
        }
    }

    // 进度监听接口
    public interface ProgressListener {
        void onProgress(int current, int total);
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        DataFormatter formatter = new DataFormatter();
                        return formatter.formatCellValue(cell);
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        try {
                            return String.valueOf(cell.getStringCellValue());
                        } catch (Exception e2) {
                            return cell.getCellFormula();
                        }
                    }
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 复制单元格内容
     */
    private void copyCellValue(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(src)) dst.setCellValue(src.getDateCellValue());
                else dst.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dst.setCellFormula(src.getCellFormula());
                break;
            case BLANK:
                dst.setBlank();
                break;
            case ERROR:
                dst.setCellErrorValue(src.getErrorCellValue());
                break;
        }
    }

    // 工具方法：去除所有不可见字符
    private String cleanString(String s) {
        if (s == null) return null;
        // 去除所有空白字符（空格、制表符、换行等）
        return s.replaceAll("\\s+", "").replace("\uFEFF", "");
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
            }
            JFrame frame = new JFrame("跨表头填充工具");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(600, 400);
            frame.setLocationRelativeTo(null);
            frame.setLayout(new BorderLayout(10, 10));

            // 文件选择区
            JPanel filePanel = new JPanel(new GridBagLayout());
            filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(5, 5, 5, 5);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weightx = 1;

            JTextField aPath = new JTextField();
            JTextField bPath = new JTextField();
            JTextField outPath = new JTextField();
            JButton selA = new JButton("选择A表");
            JButton selB = new JButton("选择B表");
            JButton selOut = new JButton("选择输出");

            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.gridwidth = 1;
            filePanel.add(new JLabel("A表路径:"), gbc);
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
            filePanel.add(new JLabel("B表路径:"), gbc);
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

            // 表头映射区
            JPanel mapPanel = new JPanel(new GridBagLayout());
            mapPanel.setBorder(BorderFactory.createTitledBorder("表头映射区"));
            JTextField aKey = new JTextField();
            JTextField bKey = new JTextField();
            JTextField aVal = new JTextField();
            JTextField bVal = new JTextField();
            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.gridwidth = 1;
            mapPanel.add(new JLabel("A表主键表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 0;
            gbc.gridwidth = 3;
            mapPanel.add(aKey, gbc);
            gbc.gridx = 0;
            gbc.gridy = 1;
            gbc.gridwidth = 1;
            mapPanel.add(new JLabel("B表主键表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 1;
            gbc.gridwidth = 3;
            mapPanel.add(bKey, gbc);
            gbc.gridx = 0;
            gbc.gridy = 2;
            gbc.gridwidth = 1;
            mapPanel.add(new JLabel("A表数据表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 2;
            gbc.gridwidth = 3;
            mapPanel.add(aVal, gbc);
            gbc.gridx = 0;
            gbc.gridy = 3;
            gbc.gridwidth = 1;
            mapPanel.add(new JLabel("B表目标表头:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 3;
            gbc.gridwidth = 3;
            mapPanel.add(bVal, gbc);

            // 操作区
            JPanel opPanel = new JPanel(new BorderLayout(10, 10));
            JButton runBtn = new JButton("开始填充");
            JLabel status = new JLabel("就绪");
            status.setHorizontalAlignment(SwingConstants.CENTER);
            opPanel.add(runBtn, BorderLayout.CENTER);
            opPanel.add(status, BorderLayout.SOUTH);

            // 主布局
            JPanel centerPanel = new JPanel();
            centerPanel.setLayout(new BoxLayout(centerPanel, BoxLayout.Y_AXIS));
            centerPanel.add(filePanel);
            centerPanel.add(Box.createVerticalStrut(10));
            centerPanel.add(mapPanel);
            frame.add(centerPanel, BorderLayout.CENTER);
            frame.add(opPanel, BorderLayout.SOUTH);

            // 事件
            selA.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION)
                    aPath.setText(fc.getSelectedFile().getAbsolutePath());
            });
            selB.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION)
                    bPath.setText(fc.getSelectedFile().getAbsolutePath());
            });
            selOut.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    String p = fc.getSelectedFile().getAbsolutePath();
                    if (!p.toLowerCase().endsWith(".xlsx")) p += ".xlsx";
                    outPath.setText(p);
                }
            });
            runBtn.addActionListener(e -> {
                JDialog progressDialog = new JDialog(frame, "处理中", true);
                JProgressBar progressBar = new JProgressBar();
                progressBar.setStringPainted(true);
                progressBar.setIndeterminate(false);
                progressDialog.setLayout(new BorderLayout(10, 10));
                progressDialog.add(new JLabel("正在填充数据，请稍候..."), BorderLayout.NORTH);
                progressDialog.add(progressBar, BorderLayout.CENTER);
                progressDialog.setSize(350, 100);
                progressDialog.setLocationRelativeTo(frame);

                SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                    @Override
                    protected Void doInBackground() {
                        try {
                            new CrossSheetFiller().fill(
                                    aPath.getText(), bPath.getText(), outPath.getText(),
                                    aKey.getText(), bKey.getText(), aVal.getText(), bVal.getText(),
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
}

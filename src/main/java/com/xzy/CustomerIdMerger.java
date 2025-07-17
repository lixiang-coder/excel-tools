package com.xzy;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.*;

import org.apache.poi.util.IOUtils;

/**
 * Cursor实现
 * Excel工具类：按CustomerId合并多行数据
 * 每个CustomerId只保留一行，其他列内容用英文逗号拼接
 */
public class CustomerIdMerger {
    private static final String SEPARATOR = ",";

    /**
     * 合并Excel中CustomerId相同的多行数据
     *
     * @param inputFilePath  输入文件路径
     * @param outputFilePath 输出文件路径
     * @throws IOException 读写异常
     */
    public void mergeByCustomerId(String inputFilePath, String outputFilePath) throws IOException {
        File inputFile = new File(inputFilePath);
        File outputFile = new File(outputFilePath);
        if (!inputFile.exists() || !inputFile.canRead()) {
            throw new IOException("无法读取源文件: " + inputFilePath);
        }
        File outputDir = outputFile.getParentFile();
        if (outputDir != null && !outputDir.exists()) {
            if (!outputDir.mkdirs()) {
                throw new IOException("无法创建输出目录: " + outputDir.getAbsolutePath());
            }
        }
        if (outputDir != null && !outputDir.canWrite()) {
            throw new IOException("无法写入输出目录: " + outputDir.getAbsolutePath());
        }
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis);
             Workbook outputWorkbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) throw new IOException("源文件没有数据");
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new IOException("源文件没有表头");
            int colCount = headerRow.getLastCellNum();
            int customerIdCol = -1;
            java.util.List<String> headers = new java.util.ArrayList<>();
            for (int i = 0; i < colCount; i++) {
                Cell cell = headerRow.getCell(i);
                String h = (cell != null) ? cell.getStringCellValue() : null;
                headers.add(h);
                if ("CustomerId".equalsIgnoreCase(h)) customerIdCol = i;
            }
            if (customerIdCol == -1) throw new IOException("未找到CustomerId列");
            // Map<CustomerId, List<每列内容List>>
            Map<String, java.util.List<java.util.List<String>>> merged = new LinkedHashMap<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell cidCell = row.getCell(customerIdCol);
                String cid = (cidCell != null) ? getCellValueAsString(cidCell) : null;
                if (cid == null || cid.isEmpty()) continue;
                java.util.List<java.util.List<String>> colVals = merged.computeIfAbsent(cid, k -> {
                    java.util.List<java.util.List<String>> l = new java.util.ArrayList<>();
                    for (int j = 0; j < colCount; j++) l.add(new java.util.ArrayList<>());
                    return l;
                });
                for (int j = 0; j < colCount; j++) {
                    Cell cell = row.getCell(j);
                    String v = (cell != null) ? getCellValueAsString(cell) : null;
                    if (v != null && !v.trim().isEmpty()) colVals.get(j).add(v.trim());
                }
            }
            // 写入新表
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
        }
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
     * 命令行/交互入口
     */
    public static void main(String[] args) {
        IOUtils.setByteArrayMaxOverride(200_000_000); // 允许最大200MB的单元格内容
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
            }
            JFrame frame = new JFrame("CustomerId合并工具");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(500, 200);
            frame.setLocationRelativeTo(null);
            JPanel panel = new JPanel(new GridLayout(3, 1, 10, 10));
            JButton selectInput = new JButton("选择输入文件");
            JButton selectOutput = new JButton("选择输出文件");
            JButton runBtn = new JButton("开始合并");
            JLabel status = new JLabel("就绪");
            panel.add(selectInput);
            panel.add(selectOutput);
            panel.add(runBtn);
            frame.add(panel, BorderLayout.CENTER);
            frame.add(status, BorderLayout.SOUTH);
            final String[] inPath = {null};
            final String[] outPath = {null};
            selectInput.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    inPath[0] = fc.getSelectedFile().getAbsolutePath();
                    status.setText("已选择输入: " + inPath[0]);
                }
            });
            selectOutput.addActionListener(e -> {
                JFileChooser fc = new JFileChooser();
                fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                if (fc.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
                    outPath[0] = fc.getSelectedFile().getAbsolutePath();
                    if (!outPath[0].toLowerCase().endsWith(".xlsx")) outPath[0] += ".xlsx";
                    status.setText("已选择输出: " + outPath[0]);
                }
            });
            runBtn.addActionListener(e -> {
                if (inPath[0] == null || outPath[0] == null) {
                    status.setText("请先选择输入和输出文件");
                    return;
                }
                try {
                    new CustomerIdMerger().mergeByCustomerId(inPath[0], outPath[0]);
                    status.setText("处理完成: " + outPath[0]);
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

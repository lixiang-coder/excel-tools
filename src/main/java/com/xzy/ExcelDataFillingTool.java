package com.xzy;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * 豆包实现
 * Excel数据填充工具
 */
public class ExcelDataFillingTool extends JFrame {
    private JTextField fileAPathField, fileBPathField, outputPathField;
    private JButton selectFileAButton, selectFileBButton, selectOutputButton, processButton;
    private JComboBox<String> tableAPrimaryKeyComboBox, tableBPrimaryKeyComboBox;
    private JComboBox<String> tableADataColumnComboBox, tableBTargetColumnComboBox;
    private JTextArea logTextArea;
    private JProgressBar progressBar;
    private List<String> tableAHeaders = new ArrayList<>();
    private List<String> tableBHeaders = new ArrayList<>();

    public ExcelDataFillingTool() {
        initUI();
    }

    private void initUI() {
        setTitle("Excel跨表主键数据填充工具");
        setSize(900, 700);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BorderLayout());
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 文件选择区
        JPanel fileSelectionPanel = new JPanel();
        fileSelectionPanel.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        JLabel fileALabel = new JLabel("表A (源数据):");
        gbc.gridx = 0;
        gbc.gridy = 0;
        fileSelectionPanel.add(fileALabel, gbc);

        fileAPathField = new JTextField(50);
        gbc.gridx = 1;
        gbc.gridy = 0;
        fileSelectionPanel.add(fileAPathField, gbc);

        selectFileAButton = new JButton("选择文件");
        gbc.gridx = 2;
        gbc.gridy = 0;
        fileSelectionPanel.add(selectFileAButton, gbc);

        JLabel fileBLabel = new JLabel("表B (目标数据):");
        gbc.gridx = 0;
        gbc.gridy = 1;
        fileSelectionPanel.add(fileBLabel, gbc);

        fileBPathField = new JTextField(50);
        gbc.gridx = 1;
        gbc.gridy = 1;
        fileSelectionPanel.add(fileBPathField, gbc);

        selectFileBButton = new JButton("选择文件");
        gbc.gridx = 2;
        gbc.gridy = 1;
        fileSelectionPanel.add(selectFileBButton, gbc);

        JLabel outputLabel = new JLabel("输出文件:");
        gbc.gridx = 0;
        gbc.gridy = 2;
        fileSelectionPanel.add(outputLabel, gbc);

        outputPathField = new JTextField(50);
        gbc.gridx = 1;
        gbc.gridy = 2;
        fileSelectionPanel.add(outputPathField, gbc);

        selectOutputButton = new JButton("选择路径");
        gbc.gridx = 2;
        gbc.gridy = 2;
        fileSelectionPanel.add(selectOutputButton, gbc);

        // 表头映射区
        JPanel mappingPanel = new JPanel();
        mappingPanel.setLayout(new GridBagLayout());
        gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        JLabel tableAPrimaryKeyLabel = new JLabel("表A主键列:");
        gbc.gridx = 0;
        gbc.gridy = 0;
        mappingPanel.add(tableAPrimaryKeyLabel, gbc);

        tableAPrimaryKeyComboBox = new JComboBox<>();
        gbc.gridx = 1;
        gbc.gridy = 0;
        mappingPanel.add(tableAPrimaryKeyComboBox, gbc);

        JLabel tableBPrimaryKeyLabel = new JLabel("表B主键列:");
        gbc.gridx = 0;
        gbc.gridy = 1;
        mappingPanel.add(tableBPrimaryKeyLabel, gbc);

        tableBPrimaryKeyComboBox = new JComboBox<>();
        gbc.gridx = 1;
        gbc.gridy = 1;
        mappingPanel.add(tableBPrimaryKeyComboBox, gbc);

        JLabel tableADataColumnLabel = new JLabel("表A数据列:");
        gbc.gridx = 0;
        gbc.gridy = 2;
        mappingPanel.add(tableADataColumnLabel, gbc);

        tableADataColumnComboBox = new JComboBox<>();
        gbc.gridx = 1;
        gbc.gridy = 2;
        mappingPanel.add(tableADataColumnComboBox, gbc);

        JLabel tableBTargetColumnLabel = new JLabel("表B目标列:");
        gbc.gridx = 0;
        gbc.gridy = 3;
        mappingPanel.add(tableBTargetColumnLabel, gbc);

        tableBTargetColumnComboBox = new JComboBox<>();
        gbc.gridx = 1;
        gbc.gridy = 3;
        mappingPanel.add(tableBTargetColumnComboBox, gbc);

        // 操作区
        JPanel operationPanel = new JPanel();
        operationPanel.setLayout(new FlowLayout());

        processButton = new JButton("开始处理");
        operationPanel.add(processButton);

        // 进度条
        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);
        progressBar.setString("准备就绪");

        // 日志区
        logTextArea = new JTextArea(15, 80);
        logTextArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logTextArea);

        // 添加所有面板到主面板
        mainPanel.add(fileSelectionPanel, BorderLayout.NORTH);
        mainPanel.add(mappingPanel, BorderLayout.CENTER);
        mainPanel.add(operationPanel, BorderLayout.SOUTH);

        JSplitPane splitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, mainPanel, scrollPane);
        splitPane.setDividerLocation(300);

        getContentPane().add(splitPane, BorderLayout.CENTER);
        getContentPane().add(progressBar, BorderLayout.SOUTH);

        // 添加事件监听器
        addActionListeners();
    }

    private void addActionListeners() {
        selectFileAButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel文件", "xlsx"));
            int result = fileChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                fileAPathField.setText(selectedFile.getAbsolutePath());
                loadTableAHeaders(selectedFile.getAbsolutePath());
            }
        });

        selectFileBButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel文件", "xlsx"));
            int result = fileChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                fileBPathField.setText(selectedFile.getAbsolutePath());
                loadTableBHeaders(selectedFile.getAbsolutePath());
            }
        });

        selectOutputButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel文件", "xlsx"));
            fileChooser.setDialogTitle("保存输出文件");
            fileChooser.setSelectedFile(new File("填充结果.xlsx"));
            int result = fileChooser.showSaveDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                String filePath = selectedFile.getAbsolutePath();
                if (!filePath.toLowerCase().endsWith(".xlsx")) {
                    filePath += ".xlsx";
                }
                outputPathField.setText(filePath);
            }
        });

        processButton.addActionListener(e -> {
            String fileAPath = fileAPathField.getText().trim();
            String fileBPath = fileBPathField.getText().trim();
            String outputPath = outputPathField.getText().trim();

            if (fileAPath.isEmpty() || fileBPath.isEmpty() || outputPath.isEmpty()) {
                JOptionPane.showMessageDialog(this, "请选择所有文件路径", "错误", JOptionPane.ERROR_MESSAGE);
                return;
            }

            if (tableAPrimaryKeyComboBox.getSelectedItem() == null ||
                    tableBPrimaryKeyComboBox.getSelectedItem() == null ||
                    tableADataColumnComboBox.getSelectedItem() == null ||
                    tableBTargetColumnComboBox.getSelectedItem() == null) {
                JOptionPane.showMessageDialog(this, "请选择所有表头映射", "错误", JOptionPane.ERROR_MESSAGE);
                return;
            }

            String tableAPrimaryKey = tableAPrimaryKeyComboBox.getSelectedItem().toString();
            String tableBPrimaryKey = tableBPrimaryKeyComboBox.getSelectedItem().toString();
            String tableADataColumn = tableADataColumnComboBox.getSelectedItem().toString();
            String tableBTargetColumn = tableBTargetColumnComboBox.getSelectedItem().toString();

            processButton.setEnabled(false);
            progressBar.setValue(0);
            progressBar.setString("处理中...");
            logTextArea.setText("开始处理...\n");

            ExecutorService executor = Executors.newSingleThreadExecutor();
            executor.submit(() -> {
                try {
                    processExcelFiles(fileAPath, fileBPath, outputPath,
                            tableAPrimaryKey, tableBPrimaryKey, tableADataColumn, tableBTargetColumn);
                    SwingUtilities.invokeLater(() -> {
                        progressBar.setValue(100);
                        progressBar.setString("处理完成");
                        processButton.setEnabled(true);
                        JOptionPane.showMessageDialog(ExcelDataFillingTool.this, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
                    });
                } catch (Exception ex) {
                    SwingUtilities.invokeLater(() -> {
                        logTextArea.append("错误: " + ex.getMessage() + "\n");
                        progressBar.setString("处理失败");
                        processButton.setEnabled(true);
                        JOptionPane.showMessageDialog(ExcelDataFillingTool.this, "处理失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                    });
                }
                executor.shutdown();
            });
        });
    }

    private void loadTableAHeaders(String filePath) {
        try {
            tableAHeaders.clear();
            tableAPrimaryKeyComboBox.removeAllItems();
            tableADataColumnComboBox.removeAllItems();

            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                        Cell cell = headerRow.getCell(i);
                        if (cell != null) {
                            String header = cleanString(cell.toString());
                            tableAHeaders.add(header);
                            tableAPrimaryKeyComboBox.addItem(header);
                            tableADataColumnComboBox.addItem(header);
                        }
                    }
                }
                logTextArea.append("已加载表A表头，共" + tableAHeaders.size() + "列\n");
            }
        } catch (Exception e) {
            logTextArea.append("加载表A表头失败: " + e.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "加载表A表头失败: " + e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void loadTableBHeaders(String filePath) {
        try {
            tableBHeaders.clear();
            tableBPrimaryKeyComboBox.removeAllItems();
            tableBTargetColumnComboBox.removeAllItems();

            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                        Cell cell = headerRow.getCell(i);
                        if (cell != null) {
                            String header = cleanString(cell.toString());
                            tableBHeaders.add(header);
                            tableBPrimaryKeyComboBox.addItem(header);
                            tableBTargetColumnComboBox.addItem(header);
                        }
                    }
                }
                logTextArea.append("已加载表B表头，共" + tableBHeaders.size() + "列\n");
            }
        } catch (Exception e) {
            logTextArea.append("加载表B表头失败: " + e.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "加载表B表头失败: " + e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void processExcelFiles(String fileAPath, String fileBPath, String outputPath,
                                   String tableAPrimaryKey, String tableBPrimaryKey,
                                   String tableADataColumn, String tableBTargetColumn) throws Exception {
        // 查找表头索引
        int tableAPrimaryKeyIndex = tableAHeaders.indexOf(tableAPrimaryKey);
        int tableADataColumnIndex = tableAHeaders.indexOf(tableADataColumn);
        int tableBPrimaryKeyIndex = tableBHeaders.indexOf(tableBPrimaryKey);
        int tableBTargetColumnIndex = tableBHeaders.indexOf(tableBTargetColumn);

        if (tableAPrimaryKeyIndex == -1 || tableADataColumnIndex == -1 ||
                tableBPrimaryKeyIndex == -1 || tableBTargetColumnIndex == -1) {
            throw new Exception("找不到指定的表头列");
        }

        logTextArea.append("开始读取表A数据...\n");
        Map<String, List<String>> primaryKeyToDataMap = new HashMap<>();

        // 读取表A数据
        try (FileInputStream fis = new FileInputStream(fileAPath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int totalRows = sheet.getLastRowNum();
            logTextArea.append("表A总行数: " + totalRows + "\n");

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell primaryKeyCell = row.getCell(tableAPrimaryKeyIndex);
                Cell dataCell = row.getCell(tableADataColumnIndex);

                if (primaryKeyCell != null) {
                    String primaryKey = cleanString(primaryKeyCell.toString());
                    String data = dataCell != null ? cleanString(dataCell.toString()) : "";

                    primaryKeyToDataMap.computeIfAbsent(primaryKey, k -> new ArrayList<>()).add(data);
                }

                if (i % 10000 == 0) {
                    int progress = (int) ((double) i / totalRows * 20); // 表A读取占20%进度
                    updateProgress(progress);
                }
            }
        }

        logTextArea.append("表A数据读取完成，共" + primaryKeyToDataMap.size() + "个唯一主键\n");
        logTextArea.append("开始处理表B数据...\n");

        // 处理表B数据
        try (FileInputStream fis = new FileInputStream(fileBPath);
             Workbook workbookB = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(outputPath)) {
            Sheet sheetB = workbookB.getSheetAt(0);
            int totalRowsB = sheetB.getLastRowNum();
            logTextArea.append("表B总行数: " + totalRowsB + "\n");

            // 使用SXSSFWorkbook处理大文件，增加内存窗口大小以避免错误
            SXSSFWorkbook outputWorkbook = new SXSSFWorkbook((XSSFWorkbook) workbookB, totalRowsB);
            Sheet outputSheet = outputWorkbook.getSheetAt(0);

            int matchedCount = 0;
            int totalProcessed = 0;

            for (int i = 1; i <= totalRowsB; i++) {
                Row rowB = outputSheet.getRow(i);
                if (rowB == null) {
                    // 如果行不存在，尝试从原始表中获取
                    Row originalRow = sheetB.getRow(i);
                    if (originalRow != null) {
                        // 复制原始行的内容到新行
                        rowB = outputSheet.createRow(i);
                        for (int j = 0; j < originalRow.getLastCellNum(); j++) {
                            Cell originalCell = originalRow.getCell(j);
                            if (originalCell != null) {
                                Cell newCell = rowB.createCell(j);
                                copyCell(originalCell, newCell);
                            }
                        }
                    } else {
                        // 如果原始行也不存在，创建一个新行
                        rowB = outputSheet.createRow(i);
                    }
                }

                Cell primaryKeyCellB = rowB.getCell(tableBPrimaryKeyIndex);
                if (primaryKeyCellB != null) {
                    String primaryKeyB = cleanString(primaryKeyCellB.toString());
                    if (!primaryKeyB.isEmpty() && primaryKeyToDataMap.containsKey(primaryKeyB)) {
                        List<String> dataValues = primaryKeyToDataMap.get(primaryKeyB);
                        if (!dataValues.isEmpty()) {
                            // 获取第一个匹配的数据值
                            String dataValue = dataValues.get(0);
                            // 移除已使用的值，以便下次匹配使用下一个值
                            if (dataValues.size() > 1) {
                                dataValues.remove(0);
                            }

                            Cell targetCell = rowB.getCell(tableBTargetColumnIndex);
                            if (targetCell == null) {
                                targetCell = rowB.createCell(tableBTargetColumnIndex);
                            }
                            targetCell.setCellValue(dataValue);
                            matchedCount++;
                        }
                    }
                }

                totalProcessed++;
                if (totalProcessed % 10000 == 0) {
                    int progress = 20 + (int) ((double) totalProcessed / totalRowsB * 80); // 表B处理占80%进度
                    updateProgress(progress);
                }
            }

            outputWorkbook.write(fos);
            outputWorkbook.dispose(); // 释放临时文件

            logTextArea.append("处理完成，共匹配 " + matchedCount + " 行数据\n");
        }
    }

    // 复制单元格内容和样式
    private void copyCell(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                // 保持空白
                break;
            default:
                // 其他类型不处理
                break;
        }

        // 复制单元格样式
        if (sourceCell.getCellStyle() != null) {
            targetCell.setCellStyle(sourceCell.getCellStyle());
        }
    }

    private String cleanString(String input) {
        if (input == null) return "";
        // 去除BOM
        if (input.startsWith("\uFEFF")) {
            input = input.substring(1);
        }
        // 去除所有不可见字符，包括空格、制表符、换行等
        return input.replaceAll("\\s+", "").toLowerCase();
    }

    private void updateProgress(int progress) {
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(progress);
            progressBar.setString("处理中: " + progress + "%");
        });
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelDataFillingTool tool = new ExcelDataFillingTool();
            tool.setVisible(true);
        });
    }
}

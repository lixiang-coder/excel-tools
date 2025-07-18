package com.xzy.oldFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 第二版
 * 表头映射工具
 * 用于处理Excel表头之间的映射关系
 */
public class TableMappingTool {
    private Map<String, String> mappingRelations = new HashMap<>();
    private JFrame mainFrame;
    private JTextArea logArea;
    private JLabel statusLabel;
    private JList<String> mappingList;
    private DefaultListModel<String> mappingListModel;
    private JPanel mappingPanel;
    private JPanel applyPanel;
    private JTabbedPane tabbedPane;
    private File currentMappingFile;
    private File currentTargetFile;

    public TableMappingTool() {
        initializeUI();
    }

    private void initializeUI() {
        mainFrame = new JFrame("表头映射工具");
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setSize(800, 600);
        mainFrame.setLocationRelativeTo(null);

        // 创建选项卡面板
        tabbedPane = new JTabbedPane();

        // 创建映射关系面板
        mappingPanel = createMappingPanel();
        tabbedPane.addTab("建立映射关系", mappingPanel);

        // 创建应用映射面板
        applyPanel = createApplyPanel();
        tabbedPane.addTab("应用映射", applyPanel);

        // 创建日志区域
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder("操作记录"));

        logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logScrollPane.setPreferredSize(new Dimension(0, 100)); // 限制高度为100像素
        logPanel.add(logScrollPane, BorderLayout.CENTER);

        // 创建状态标签
        statusLabel = new JLabel("就绪");

        // 布局
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        mainPanel.add(logPanel, BorderLayout.SOUTH);
        mainPanel.add(statusLabel, BorderLayout.NORTH);

        mainFrame.add(mainPanel);
        mainFrame.setVisible(true);
    }

    private JPanel createMappingPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 创建按钮面板，使用GridLayout使两个按钮各占一半宽度
        JPanel buttonPanel = new JPanel(new GridLayout(1, 2, 10, 0));

        // 创建上传按钮
        JButton uploadButton = new JButton("上传关系表");
        uploadButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(uploadButton);

        // 创建保存按钮
        JButton saveButton = new JButton("保存映射关系");
        saveButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(saveButton);

        panel.add(buttonPanel, BorderLayout.NORTH);

        // 创建表头选择面板
        JPanel headerPanel = new JPanel(new GridLayout(2, 2, 10, 10));
        headerPanel.setBorder(BorderFactory.createTitledBorder("选择映射字段"));

        // 源字段选择
        JPanel sourcePanel = new JPanel(new BorderLayout(5, 5));
        sourcePanel.setBorder(BorderFactory.createTitledBorder("源字段"));
        JList<String> sourceList = new JList<>(new DefaultListModel<>());
        sourceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        JScrollPane sourceScrollPane = new JScrollPane(sourceList);
        sourcePanel.add(sourceScrollPane, BorderLayout.CENTER);

        // 目标字段选择
        JPanel targetPanel = new JPanel(new BorderLayout(5, 5));
        targetPanel.setBorder(BorderFactory.createTitledBorder("目标字段"));
        JList<String> targetList = new JList<>(new DefaultListModel<>());
        targetList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        JScrollPane targetScrollPane = new JScrollPane(targetList);
        targetPanel.add(targetScrollPane, BorderLayout.CENTER);

        headerPanel.add(sourcePanel);
        headerPanel.add(targetPanel);

        panel.add(headerPanel, BorderLayout.CENTER);

        // 创建映射关系列表（可伸缩）
        mappingListModel = new DefaultListModel<>();
        mappingList = new JList<>(mappingListModel);
        JScrollPane mappingScrollPane = new JScrollPane(mappingList);
        mappingScrollPane.setBorder(BorderFactory.createTitledBorder("已保存的映射关系"));

        // 创建可伸缩面板
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, headerPanel, mappingScrollPane);
        splitPane.setDividerLocation(500);
        splitPane.setDividerSize(5);
        splitPane.setOneTouchExpandable(true);

        panel.add(splitPane, BorderLayout.CENTER);

        // 添加上传按钮事件
        uploadButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

            int result = fileChooser.showOpenDialog(mainFrame);
            if (result == JFileChooser.APPROVE_OPTION) {
                currentMappingFile = fileChooser.getSelectedFile();
                loadHeaders(currentMappingFile, sourceList, targetList);
            }
        });

        // 添加保存按钮事件
        saveButton.addActionListener(e -> {
            String sourceField = sourceList.getSelectedValue();
            String targetField = targetList.getSelectedValue();

            if (sourceField != null && targetField != null) {
                saveMapping(currentMappingFile, sourceField, targetField);
            } else {
                JOptionPane.showMessageDialog(mainFrame,
                        "请选择源字段和目标字段",
                        "提示",
                        JOptionPane.INFORMATION_MESSAGE);
            }
        });

        return panel;
    }

    private JPanel createApplyPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 创建按钮面板，使用GridLayout使两个按钮各占一半宽度
        JPanel buttonPanel = new JPanel(new GridLayout(1, 2, 10, 0));

        // 创建上传按钮
        JButton uploadButton = new JButton("上传目标表");
        uploadButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(uploadButton);

        // 创建应用按钮
        JButton applyButton = new JButton("应用映射");
        applyButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(applyButton);

        panel.add(buttonPanel, BorderLayout.NORTH);

        // 创建表头选择面板
        JPanel headerPanel = new JPanel(new GridLayout(2, 2, 10, 10));
        headerPanel.setBorder(BorderFactory.createTitledBorder("选择映射字段"));

        // 查找字段选择（用于查找映射关系的字段）
        JPanel lookupPanel = new JPanel(new BorderLayout(5, 5));
        lookupPanel.setBorder(BorderFactory.createTitledBorder("查找字段（如：用户姓名）"));
        JList<String> lookupList = new JList<>(new DefaultListModel<>());
        lookupList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        JScrollPane lookupScrollPane = new JScrollPane(lookupList);
        lookupPanel.add(lookupScrollPane, BorderLayout.CENTER);

        // 填充字段选择（用于填充映射值的字段）
        JPanel fillPanel = new JPanel(new BorderLayout(5, 5));
        fillPanel.setBorder(BorderFactory.createTitledBorder("填充字段（如：用户编号）"));
        JList<String> fillList = new JList<>(new DefaultListModel<>());
        fillList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        JScrollPane fillScrollPane = new JScrollPane(fillList);
        fillPanel.add(fillScrollPane, BorderLayout.CENTER);

        headerPanel.add(lookupPanel);
        headerPanel.add(fillPanel);

        panel.add(headerPanel, BorderLayout.CENTER);

        // 添加上传按钮事件
        uploadButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

            int result = fileChooser.showOpenDialog(mainFrame);
            if (result == JFileChooser.APPROVE_OPTION) {
                currentTargetFile = fileChooser.getSelectedFile();
                loadHeaders(currentTargetFile, lookupList, fillList);
            }
        });

        // 添加应用按钮事件
        applyButton.addActionListener(e -> {
            String lookupField = lookupList.getSelectedValue();
            String fillField = fillList.getSelectedValue();

            if (lookupField != null && fillField != null) {
                applyMapping(currentTargetFile, lookupField, fillField);
            } else {
                JOptionPane.showMessageDialog(mainFrame,
                        "请选择查找字段和填充字段",
                        "提示",
                        JOptionPane.INFORMATION_MESSAGE);
            }
        });

        return panel;
    }

    private void loadHeaders(File file, JList<String> sourceList, JList<String> targetList) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            DefaultListModel<String> sourceModel = new DefaultListModel<>();
            DefaultListModel<String> targetModel = new DefaultListModel<>();

            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String header = cell.getStringCellValue().trim();
                    sourceModel.addElement(header);
                    targetModel.addElement(header);
                }
            }

            sourceList.setModel(sourceModel);
            targetList.setModel(targetModel);

            statusLabel.setText("已加载 " + sourceModel.size() + " 个表头");
            log("成功加载表头: " + file.getName());

        } catch (IOException e) {
            log("加载表头时出错: " + e.getMessage());
            JOptionPane.showMessageDialog(mainFrame,
                    "加载表头时出错: " + e.getMessage(),
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void saveMapping(File file, String sourceField, String targetField) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 获取表头索引
            Map<String, Integer> headerIndexMap = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    headerIndexMap.put(cell.getStringCellValue().trim(), i);
                }
            }

            // 获取源字段和目标字段的索引
            Integer sourceIndex = headerIndexMap.get(sourceField);
            Integer targetIndex = headerIndexMap.get(targetField);

            if (sourceIndex == null || targetIndex == null) {
                throw new IOException("找不到指定的字段");
            }

            // 处理数据行
            final int totalRows = sheet.getLastRowNum();
            int processedRows = 0;
            int savedRelations = 0;

            // 显示开始处理的消息
            SwingUtilities.invokeLater(() -> {
                statusLabel.setText("正在处理映射关系...");
                log("开始处理映射关系，共 " + totalRows + " 行数据");
            });

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell sourceCell = row.getCell(sourceIndex);
                Cell targetCell = row.getCell(targetIndex);

                if (sourceCell != null && targetCell != null) {
                    String sourceValue = getCellValueAsString(sourceCell);
                    String targetValue = getCellValueAsString(targetCell);

                    if (!sourceValue.isEmpty() && !targetValue.isEmpty()) {
                        mappingRelations.put(sourceValue, targetValue);
                        savedRelations++;
                    }
                }

                processedRows++;
                // 只在处理完成时更新状态
                if (processedRows == totalRows) {
                    final int currentSaved = savedRelations;
                    SwingUtilities.invokeLater(() -> {
                        statusLabel.setText("已保存 " + currentSaved + " 条映射关系");
                        log("处理完成，共保存 " + currentSaved + " 条映射关系");
                    });
                }
            }

            // 更新映射列表
            mappingListModel.clear();
            for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                mappingListModel.addElement(entry.getKey() + " -> " + entry.getValue());
            }

            statusLabel.setText("已保存 " + mappingRelations.size() + " 条映射关系");
            log("成功保存映射关系: " + sourceField + " -> " + targetField);

        } catch (IOException e) {
            log("保存映射关系时出错: " + e.getMessage());
            JOptionPane.showMessageDialog(mainFrame,
                    "保存映射关系时出错: " + e.getMessage(),
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void applyMapping(File file, String lookupField, String fillField) {
        if (mappingRelations.isEmpty()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "请先建立映射关系",
                    "提示",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        fileChooser.setDialogTitle("选择最终文件的保存位置");

        int result = fileChooser.showSaveDialog(mainFrame);
        if (result == JFileChooser.APPROVE_OPTION) {
            File outputFile = new File(
                    fileChooser.getSelectedFile().getAbsolutePath() +
                            (!fileChooser.getSelectedFile().getName().toLowerCase().endsWith(".xlsx") ? ".xlsx" : "")
            );

            try (FileInputStream fis = new FileInputStream(file);
                 Workbook workbook = WorkbookFactory.create(fis);
                 FileOutputStream fos = new FileOutputStream(outputFile)) {

                Sheet sheet = workbook.getSheetAt(0);

                // 获取表头索引
                Map<String, Integer> headerIndexMap = new HashMap<>();
                Row headerRow = sheet.getRow(0);
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        headerIndexMap.put(cell.getStringCellValue().trim(), i);
                    }
                }

                // 获取查找字段和填充字段的索引
                Integer lookupIndex = headerIndexMap.get(lookupField);
                Integer fillIndex = headerIndexMap.get(fillField);

                if (lookupIndex == null || fillIndex == null) {
                    throw new IOException("找不到指定的字段");
                }

                // 创建新的工作簿
                Workbook outputWorkbook = new XSSFWorkbook();
                Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

                // 复制表头
                Row outputHeaderRow = outputSheet.createRow(0);
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        Cell outputCell = outputHeaderRow.createCell(i);
                        copyCellValue(cell, outputCell);
                    }
                }

                // 处理数据行
                final int totalRows = sheet.getLastRowNum();
                int processedRows = 0;
                int updatedCount = 0;

                // 显示开始处理的消息
                SwingUtilities.invokeLater(() -> {
                    statusLabel.setText("正在应用映射关系...");
                    log("开始应用映射关系，共 " + totalRows + " 行数据");
                });

                for (int i = 1; i <= totalRows; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    // 创建新行
                    Row outputRow = outputSheet.createRow(i);

                    // 复制所有单元格
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            Cell outputCell = outputRow.createCell(j);
                            copyCellValue(cell, outputCell);
                        }
                    }

                    // 获取查找字段的值
                    Cell lookupCell = row.getCell(lookupIndex);
                    if (lookupCell != null) {
                        String lookupValue = getCellValueAsString(lookupCell);

                        // 创建填充单元格
                        Cell outputFillCell = outputRow.getCell(fillIndex);
                        if (outputFillCell == null) {
                            outputFillCell = outputRow.createCell(fillIndex);
                        }

                        // 在映射关系中查找对应的值
                        String mappedValue = null;
                        for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                            if (entry.getValue().equals(lookupValue)) {
                                mappedValue = entry.getKey();
                                break;
                            }
                        }

                        // 如果找到对应的映射值，则填充到填充字段
                        if (mappedValue != null) {
                            outputFillCell.setCellValue(mappedValue);
                            updatedCount++;
                        } else {
                            // 如果找不到映射关系，则直接将查找字段的值复制到填充字段
                            outputFillCell.setCellValue(lookupValue);
                        }
                    }

                    processedRows++;
                    // 只在处理完成时更新状态
                    if (processedRows == totalRows) {
                        final int currentUpdated = updatedCount;
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("已更新 " + currentUpdated + " 条记录");
                            log("处理完成，共更新 " + currentUpdated + " 条记录");
                        });
                    }
                }

                // 保存文件
                outputWorkbook.write(fos);

                statusLabel.setText("已更新 " + updatedCount + " 条记录");
                log("成功应用映射关系，更新了 " + updatedCount + " 条记录");

            } catch (IOException e) {
                log("应用映射关系时出错: " + e.getMessage());
                JOptionPane.showMessageDialog(mainFrame,
                        "应用映射关系时出错: " + e.getMessage(),
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf((long)cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            case ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
        }
    }

    private void log(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new TableMappingTool();
        });
    }
}

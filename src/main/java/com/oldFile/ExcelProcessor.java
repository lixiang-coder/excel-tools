package com.xzy.oldFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.*;

/**
 * 第一版
 * Excel处理工具类
 * 用于处理任意表头的Excel文件之间的数据转换
 */
public class ExcelProcessor {

    private static final Logger logger = LogManager.getLogger(ExcelProcessor.class);
    private static final String OTHER_REMARK_SEPARATOR = " , ";
    private JFrame mainFrame;
    private JTextArea logArea;
    private JLabel statusLabel;
    private JPanel mappingPanel;
    private DefaultListModel<String> sourceListModel;
    private DefaultListModel<String> templateListModel;
    private DefaultListModel<String> mappingListModel;
    private JList<String> sourceList;
    private JList<String> templateList;
    private JList<String> mappingList;
    private Map<String, String> mappings;
    private List<String> mappingOrder;
    private String inputFilePath;
    private String templateFilePath;
    private String outputFilePath;

    /**
     * 表头映射配置类
     */
    public static class HeaderMapping {
        private final List<String> sourceHeaders;  // 源表头（表A）
        private final String targetHeader;         // 目标表头（表B）

        public HeaderMapping(List<String> sourceHeaders, String targetHeader) {
            this.sourceHeaders = new ArrayList<>(sourceHeaders);
            this.targetHeader = targetHeader;
        }

        public List<String> getSourceHeaders() {
            return new ArrayList<>(sourceHeaders);
        }

        public String getTargetHeader() {
            return targetHeader;
        }

        public boolean isMerge() {
            return sourceHeaders.size() > 1;
        }
    }

    /**
     * 处理Excel文件
     * 将源表的数据根据映射规则转换为目标表的格式
     *
     * @param inputFilePath    输入文件路径
     * @param templateFilePath 模板文件路径
     * @param outputFilePath   输出文件路径
     * @param mappings         映射关系
     * @throws IOException 如果文件读写过程中发生错误
     */
    public void processExcel(String inputFilePath, String templateFilePath, String outputFilePath, Map<String, String> mappings) throws IOException {
        // 检查文件是否存在且可访问
        File inputFile = new File(inputFilePath);
        File templateFile = new File(templateFilePath);
        File outputFile = new File(outputFilePath);

        if (!inputFile.exists()) {
            throw new IOException("源数据文件不存在: " + inputFilePath);
        }

        if (!templateFile.exists()) {
            throw new IOException("模板文件不存在: " + templateFilePath);
        }

        if (!inputFile.canRead()) {
            throw new IOException("无法读取源数据文件，请检查文件权限: " + inputFilePath);
        }

        if (!templateFile.canRead()) {
            throw new IOException("无法读取模板文件，请检查文件权限: " + templateFilePath);
        }

        // 检查输出目录是否可写
        File outputDir = outputFile.getParentFile();
        if (outputDir != null && !outputDir.exists()) {
            if (!outputDir.mkdirs()) {
                throw new IOException("无法创建输出目录: " + outputDir.getAbsolutePath());
            }
        }

        if (outputDir != null && !outputDir.canWrite()) {
            throw new IOException("无法写入输出目录，请检查权限: " + outputDir.getAbsolutePath());
        }

        // 创建进度对话框
        JDialog progressDialog = new JDialog();
        progressDialog.setTitle("处理中");
        progressDialog.setModal(true);
        progressDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        progressDialog.setResizable(false);
        progressDialog.setUndecorated(false);

        JPanel progressPanel = new JPanel(new BorderLayout(10, 10));
        progressPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JLabel progressLabel = new JLabel("正在处理数据...");
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);

        progressPanel.add(progressLabel, BorderLayout.NORTH);
        progressPanel.add(progressBar, BorderLayout.CENTER);

        progressDialog.add(progressPanel);
        progressDialog.pack();
        progressDialog.setLocationRelativeTo(null);

        // 在后台线程中处理数据
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                try (FileInputStream fis = new FileInputStream(inputFile);
                     Workbook workbook = WorkbookFactory.create(fis);
                     FileInputStream templateFis = new FileInputStream(templateFile);
                     Workbook templateWorkbook = WorkbookFactory.create(templateFis);
                     FileOutputStream fos = new FileOutputStream(outputFile)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    Sheet templateSheet = templateWorkbook.getSheetAt(0);

                    // 获取模板文件的表头
                    List<String> templateHeaders = new ArrayList<>();
                    Row templateHeaderRow = templateSheet.getRow(0);
                    if (templateHeaderRow != null) {
                        for (int i = 0; i < templateHeaderRow.getLastCellNum(); i++) {
                            Cell cell = templateHeaderRow.getCell(i);
                            if (cell != null) {
                                templateHeaders.add(cell.getStringCellValue());
                            }
                        }
                    }

                    // 获取源文件的表头
                    Map<String, Integer> sourceHeaderMap = new LinkedHashMap<>();
                    Row headerRow = sheet.getRow(0);
                    if (headerRow != null) {
                        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                            Cell cell = headerRow.getCell(i);
                            if (cell != null) {
                                sourceHeaderMap.put(cell.getStringCellValue(), i);
                            }
                        }
                    }

                    // 将简单映射转换为HeaderMapping对象
                    List<HeaderMapping> headerMappings = new ArrayList<>();
                    Map<String, List<String>> targetToSourceMap = new HashMap<>();

                    // 首先按目标表头分组
                    for (Map.Entry<String, String> entry : mappings.entrySet()) {
                        String sourceHeader = entry.getKey();
                        String targetHeader = entry.getValue();

                        targetToSourceMap.computeIfAbsent(targetHeader, k -> new ArrayList<>()).add(sourceHeader);
                    }

                    // 然后创建HeaderMapping对象
                    for (Map.Entry<String, List<String>> entry : targetToSourceMap.entrySet()) {
                        String targetHeader = entry.getKey();
                        List<String> sourceHeaders = entry.getValue();

                        headerMappings.add(new HeaderMapping(sourceHeaders, targetHeader));
                    }

                    // 创建新的工作簿
                    Workbook outputWorkbook = new XSSFWorkbook();
                    Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

                    // 写入模板表头
                    Row outputHeaderRow = outputSheet.createRow(0);
                    for (int i = 0; i < templateHeaders.size(); i++) {
                        Cell cell = outputHeaderRow.createCell(i);
                        cell.setCellValue(templateHeaders.get(i));
                    }

                    // 处理数据行
                    int rowNum = 1;
                    int totalRows = sheet.getLastRowNum();
                    int processedRows = 0;

                    // 显示开始处理的消息
                    SwingUtilities.invokeLater(() -> {
                        statusLabel.setText("正在处理数据...");
                        log("开始处理数据，共 " + totalRows + " 行");
                    });

                    for (int i = 1; i <= totalRows; i++) {
                        Row sourceRow = sheet.getRow(i);
                        if (sourceRow == null) continue;

                        Row outputRow = outputSheet.createRow(rowNum++);

                        // 根据映射规则处理每一行数据
                        for (HeaderMapping mapping : headerMappings) {
                            if (mapping.isMerge()) {
                                // 处理需要合并的多个源表头
                                StringBuilder mergedValue = new StringBuilder();
                                for (String sourceHeader : mapping.getSourceHeaders()) {
                                    Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                    if (sourceIndex != null) {
                                        Cell sourceCell = sourceRow.getCell(sourceIndex);
                                        if (sourceCell != null) {
                                            String value = getCellValueAsString(sourceCell);
                                            if (!value.isEmpty()) {
                                                if (mergedValue.length() > 0) {
                                                    mergedValue.append(OTHER_REMARK_SEPARATOR);
                                                }
                                                mergedValue.append(value);
                                            }
                                        }
                                    }
                                }
                                if (mergedValue.length() > 0) {
                                    int targetColumnIndex = templateHeaders.indexOf(mapping.getTargetHeader());
                                    if (targetColumnIndex >= 0) {
                                        Cell targetCell = outputRow.createCell(targetColumnIndex);
                                        targetCell.setCellValue(mergedValue.toString());
                                    }
                                }
                            } else {
                                // 处理单个源表头
                                String sourceHeader = mapping.getSourceHeaders().get(0);
                                Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                if (sourceIndex != null) {
                                    Cell sourceCell = sourceRow.getCell(sourceIndex);
                                    if (sourceCell != null) {
                                        String value = getCellValueAsString(sourceCell);
                                        if (!value.isEmpty()) {
                                            int targetColumnIndex = templateHeaders.indexOf(mapping.getTargetHeader());
                                            if (targetColumnIndex >= 0) {
                                                Cell targetCell = outputRow.createCell(targetColumnIndex);
                                                targetCell.setCellValue(value);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        processedRows++;
                        // 更新进度
                        final int progress = (int) ((i * 100.0) / totalRows);
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setIndeterminate(false);
                            progressBar.setValue(progress);
                            progressLabel.setText("正在处理数据... " + progress + "%");
                        });
                    }

                    // 保存文件
                    outputWorkbook.write(fos);

                    // 显示处理完成的消息
                    final int finalProcessedRows = processedRows;
                    SwingUtilities.invokeLater(() -> {
                        statusLabel.setText("处理完成，共处理 " + finalProcessedRows + " 行数据");
                        log("处理完成，共处理 " + finalProcessedRows + " 行数据");
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
        } catch (Exception e) {
            logger.error("处理Excel文件时发生错误", e);
            throw new RuntimeException("处理Excel文件失败: " + e.getMessage());
        }
    }

    /**
     * 获取单元格的字符串值
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        String result = "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    result = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = cell.getDateCellValue().toString();
                    } else {
                        DataFormatter formatter = new DataFormatter();
                        result = formatter.formatCellValue(cell);
                    }
                    break;
                case BOOLEAN:
                    result = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    try {
                        result = String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        try {
                            result = String.valueOf(cell.getStringCellValue());
                        } catch (Exception e2) {
                            result = cell.getCellFormula();
                        }
                    }
                    break;
                default:
                    result = "";
            }
        } catch (Exception e) {
            result = "";
        }

        return result != null ? result.trim() : "";
    }

    /**
     * 加载表头
     */
    private void loadHeaders(String inputFilePath, String templateFilePath) {
        // 检查文件是否存在且可访问
        File inputFile = new File(inputFilePath);
        File templateFile = new File(templateFilePath);

        if (!inputFile.exists()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "源数据文件不存在: " + inputFilePath,
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (!templateFile.exists()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "模板文件不存在: " + templateFilePath,
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (!inputFile.canRead()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "无法读取源数据文件，请检查文件权限: " + inputFilePath,
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (!templateFile.canRead()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "无法读取模板文件，请检查文件权限: " + templateFilePath,
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 创建进度对话框
        JDialog progressDialog = new JDialog();
        progressDialog.setTitle("加载中");
        progressDialog.setModal(true);
        progressDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        progressDialog.setResizable(false);
        progressDialog.setUndecorated(false);

        JPanel progressPanel = new JPanel(new BorderLayout(10, 10));
        progressPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JLabel progressLabel = new JLabel("正在加载表头...");
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);

        progressPanel.add(progressLabel, BorderLayout.NORTH);
        progressPanel.add(progressBar, BorderLayout.CENTER);

        progressDialog.add(progressPanel);
        progressDialog.pack();
        progressDialog.setLocationRelativeTo(null);

        // 在后台线程中加载表头
        SwingWorker<Map<String, List<String>>, Void> worker = new SwingWorker<Map<String, List<String>>, Void>() {
            @Override
            protected Map<String, List<String>> doInBackground() throws Exception {
                Map<String, List<String>> result = new HashMap<>();

                try (FileInputStream fis = new FileInputStream(inputFile);
                     Workbook workbook = WorkbookFactory.create(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    Row headerRow = sheet.getRow(0);

                    if (headerRow != null) {
                        List<String> sourceHeaders = new ArrayList<>();
                        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                            Cell cell = headerRow.getCell(i);
                            if (cell != null) {
                                sourceHeaders.add(cell.getStringCellValue());
                            }
                        }
                        result.put("sourceHeaders", sourceHeaders);
                    }
                }

                try (FileInputStream fis = new FileInputStream(templateFile);
                     Workbook workbook = WorkbookFactory.create(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    Row headerRow = sheet.getRow(0);

                    if (headerRow != null) {
                        List<String> templateHeaders = new ArrayList<>();
                        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                            Cell cell = headerRow.getCell(i);
                            if (cell != null) {
                                templateHeaders.add(cell.getStringCellValue());
                            }
                        }
                        result.put("templateHeaders", templateHeaders);
                    }
                }

                return result;
            }

            @Override
            protected void done() {
                progressDialog.dispose();
            }
        };

        worker.execute();
        progressDialog.setVisible(true);

        try {
            Map<String, List<String>> headers = worker.get();
            List<String> sourceHeaders = headers.get("sourceHeaders");
            List<String> templateHeaders = headers.get("templateHeaders");

            if (sourceHeaders == null || sourceHeaders.isEmpty()) {
                JOptionPane.showMessageDialog(mainFrame,
                        "源数据文件没有表头",
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
                return;
            }

            if (templateHeaders == null || templateHeaders.isEmpty()) {
                JOptionPane.showMessageDialog(mainFrame,
                        "模板文件没有表头",
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
                return;
            }

            // 更新表头列表
            sourceListModel.clear();
            templateListModel.clear();
            mappingListModel.clear();
            mappings.clear();
            mappingOrder.clear();

            for (String header : sourceHeaders) {
                sourceListModel.addElement(header);
            }

            for (String header : templateHeaders) {
                templateListModel.addElement(header);
            }

            // 启用映射面板
            mappingPanel.setEnabled(true);
            log("已加载源表头和模板表头");
            statusLabel.setText("已加载表头，请配置映射关系");

        } catch (Exception e) {
            logger.error("加载表头时发生错误", e);
            JOptionPane.showMessageDialog(mainFrame,
                    "加载表头失败: " + e.getMessage(),
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * 更新映射列表显示
     */
    private void updateMappingList() {
        mappingListModel.clear();
        // 按照添加顺序显示映射
        for (String sourceHeader : mappingOrder) {
            if (mappings.containsKey(sourceHeader)) {
                mappingListModel.addElement(sourceHeader + " -> " + mappings.get(sourceHeader));
            }
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

    /**
     * 初始化UI
     */
    private void initializeUI() {
        mainFrame = new JFrame("Excel处理工具");
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setSize(1000, 700);
        mainFrame.setLocationRelativeTo(null);

        // 创建主面板
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 创建顶部面板
        JPanel topPanel = new JPanel(new BorderLayout(10, 10));

        // 创建状态标签
        statusLabel = new JLabel("就绪");
        topPanel.add(statusLabel, BorderLayout.NORTH);

        // 创建按钮面板
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));

        // 创建选择文件按钮
        JButton selectFilesButton = new JButton("选择文件");
        selectFilesButton.setPreferredSize(new Dimension(100, 30));
        buttonPanel.add(selectFilesButton);

        // 创建保存按钮
        JButton saveButton = new JButton("保存文件");
        saveButton.setPreferredSize(new Dimension(100, 30));
        saveButton.setEnabled(false);
        buttonPanel.add(saveButton);

        topPanel.add(buttonPanel, BorderLayout.CENTER);
        mainPanel.add(topPanel, BorderLayout.NORTH);

        // 创建映射面板
        mappingPanel = new JPanel(new BorderLayout(10, 10));
        mappingPanel.setBorder(BorderFactory.createTitledBorder("配置字段映射关系"));
        mappingPanel.setEnabled(false);

        // 创建表头面板，使用GridLayout布局，分为三列：源表头、映射关系、模板表头
        JPanel headerPanel = new JPanel(new GridLayout(1, 3, 10, 0));

        // 创建源表头列表
        JPanel sourcePanel = new JPanel(new BorderLayout());
        sourcePanel.setBorder(BorderFactory.createTitledBorder("源表表头"));
        sourceListModel = new DefaultListModel<>();
        sourceList = new JList<>(sourceListModel);
        sourceList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        sourcePanel.add(new JScrollPane(sourceList), BorderLayout.CENTER);

        // 创建模板表头列表
        JPanel templatePanel = new JPanel(new BorderLayout());
        templatePanel.setBorder(BorderFactory.createTitledBorder("模板表头"));
        templateListModel = new DefaultListModel<>();
        templateList = new JList<>(templateListModel);
        templateList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        templatePanel.add(new JScrollPane(templateList), BorderLayout.CENTER);

        // 创建映射关系面板
        JPanel mappingListPanel = new JPanel(new BorderLayout());
        mappingListPanel.setBorder(BorderFactory.createTitledBorder("已配置的映射关系"));
        mappingListModel = new DefaultListModel<>();
        mappingList = new JList<>(mappingListModel);
        mappingListPanel.add(new JScrollPane(mappingList), BorderLayout.CENTER);

        // 创建映射操作按钮面板
        JPanel mappingButtonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
        JButton addMappingButton = new JButton("添加映射");
        addMappingButton.setPreferredSize(new Dimension(100, 30));
        JButton removeMappingButton = new JButton("删除映射");
        removeMappingButton.setPreferredSize(new Dimension(100, 30));
        mappingButtonPanel.add(addMappingButton);
        mappingButtonPanel.add(removeMappingButton);
        mappingListPanel.add(mappingButtonPanel, BorderLayout.SOUTH);

        // 添加组件到表头面板
        headerPanel.add(sourcePanel);
        headerPanel.add(mappingListPanel);
        headerPanel.add(templatePanel);

        // 添加组件到映射面板
        mappingPanel.add(headerPanel, BorderLayout.CENTER);

        // 添加组件到主面板
        mainPanel.add(mappingPanel, BorderLayout.CENTER);

        // 创建日志区域
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder("操作记录"));
        logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logScrollPane.setPreferredSize(new Dimension(0, 100));
        logPanel.add(logScrollPane, BorderLayout.CENTER);
        mainPanel.add(logPanel, BorderLayout.SOUTH);

        // 初始化映射数据
        mappings = new HashMap<>();
        mappingOrder = new ArrayList<>();

        // 添加选择文件按钮事件
        selectFilesButton.addActionListener(e -> {
            try {
                // 创建文件选择器
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

                // 第一步：选择源数据文件
                fileChooser.setDialogTitle("选择源数据文件");
                int result = fileChooser.showOpenDialog(mainFrame);
                if (result != JFileChooser.APPROVE_OPTION) {
                    return;
                }
                inputFilePath = fileChooser.getSelectedFile().getAbsolutePath();
                log("已选择源数据文件: " + inputFilePath);

                // 第二步：选择模板文件
                fileChooser.setDialogTitle("选择模板文件");
                result = fileChooser.showOpenDialog(mainFrame);
                if (result != JFileChooser.APPROVE_OPTION) {
                    return;
                }
                templateFilePath = fileChooser.getSelectedFile().getAbsolutePath();
                log("已选择模板文件: " + templateFilePath);

                // 加载表头
                loadHeaders(inputFilePath, templateFilePath);

            } catch (Exception ex) {
                logger.error("选择文件时发生错误", ex);
                log("选择文件失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainFrame,
                        "选择文件失败: " + ex.getMessage(),
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
            }
        });

        // 添加映射按钮事件
        addMappingButton.addActionListener(e -> {
            List<String> selectedSources = sourceList.getSelectedValuesList();
            String selectedTemplate = templateList.getSelectedValue();

            if (!selectedSources.isEmpty() && selectedTemplate != null) {
                // 检查是否已经存在映射
                for (String sourceHeader : selectedSources) {
                    if (mappings.containsKey(sourceHeader)) {
                        int confirm = JOptionPane.showConfirmDialog(mainFrame,
                                "源表头 '" + sourceHeader + "' 已经映射到 '" + mappings.get(sourceHeader) +
                                        "'，是否覆盖?", "确认覆盖", JOptionPane.YES_NO_OPTION);
                        if (confirm != JOptionPane.YES_OPTION) {
                            continue;
                        }
                        // 如果确认覆盖，从顺序列表中移除旧的映射
                        mappingOrder.remove(sourceHeader);
                    }

                    // 添加映射
                    mappings.put(sourceHeader, selectedTemplate);
                    // 添加到顺序列表
                    mappingOrder.add(sourceHeader);
                }

                // 更新映射列表显示
                updateMappingList();

                log("已添加映射: " + String.join(", ", selectedSources) + " -> " + selectedTemplate);
                statusLabel.setText("已添加映射关系");

                // 启用保存按钮
                saveButton.setEnabled(true);
            } else {
                JOptionPane.showMessageDialog(mainFrame,
                        "请选择源表头和模板表头", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        });

        // 删除映射按钮事件
        removeMappingButton.addActionListener(e -> {
            String selectedMapping = mappingList.getSelectedValue();
            if (selectedMapping != null) {
                // 从映射字符串中提取源表头
                String sourceHeader = selectedMapping.split(" -> ")[0];

                // 删除映射
                mappings.remove(sourceHeader);
                // 从顺序列表中移除
                mappingOrder.remove(sourceHeader);

                // 更新映射列表显示
                updateMappingList();

                log("已删除映射: " + selectedMapping);
                statusLabel.setText("已删除映射关系");

                // 如果没有映射关系，禁用保存按钮
                if (mappings.isEmpty()) {
                    saveButton.setEnabled(false);
                }
            } else {
                JOptionPane.showMessageDialog(mainFrame,
                        "请选择要删除的映射", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        });

        // 添加保存按钮事件
        saveButton.addActionListener(e -> {
            try {
                // 如果用户取消了映射或没有配置任何映射，则退出
                if (mappings == null || mappings.isEmpty()) {
                    log("未配置任何映射关系");
                    JOptionPane.showMessageDialog(mainFrame,
                            "未配置任何映射关系",
                            "提示",
                            JOptionPane.INFORMATION_MESSAGE);
                    return;
                }

                // 选择输出文件
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                fileChooser.setDialogTitle("选择保存位置");
                int result = fileChooser.showSaveDialog(mainFrame);
                if (result != JFileChooser.APPROVE_OPTION) {
                    return;
                }
                outputFilePath = fileChooser.getSelectedFile().getAbsolutePath();

                // 确保输出文件有.xlsx扩展名
                if (!outputFilePath.toLowerCase().endsWith(".xlsx")) {
                    outputFilePath += ".xlsx";
                }
                log("已选择输出文件: " + outputFilePath);

                // 处理Excel文件
                processExcel(inputFilePath, templateFilePath, outputFilePath, mappings);

                // 显示成功消息
                log("处理完成！结果已保存到: " + outputFilePath);
                JOptionPane.showMessageDialog(mainFrame,
                        "处理完成！结果已保存到: " + outputFilePath,
                        "成功",
                        JOptionPane.INFORMATION_MESSAGE);

                // 禁用保存按钮
                saveButton.setEnabled(false);

                // 清空映射关系
                mappings.clear();
                mappingOrder.clear();
                updateMappingList();

                // 禁用映射面板
                mappingPanel.setEnabled(false);

                // 更新状态
                statusLabel.setText("就绪");
            } catch (Exception ex) {
                logger.error("保存文件时发生错误", ex);
                log("保存文件失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainFrame,
                        "保存文件失败: " + ex.getMessage(),
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
            }
        });

        mainFrame.add(mainPanel);
        mainFrame.setVisible(true);
    }

    /**
     * 记录日志
     */
    private void log(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    /**
     * 主方法
     */
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            ExcelProcessor processor = new ExcelProcessor();
            processor.initializeUI();
        });
    }
}

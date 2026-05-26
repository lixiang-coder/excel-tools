package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Color;
import java.io.*;
import java.util.List;
import java.util.*;

/**
 * 最终版
 * Excel工具集
 * 整合了Excel处理器和表头映射工具的功能
 */
public class ExcelTools {
    private static final Logger logger = LogManager.getLogger(ExcelTools.class);
    private static final String OTHER_REMARK_SEPARATOR = "|";

    // 使用逗号分隔的表头集合（不显示源表头名称）
    private static final Set<String> COMMA_SEPARATED_HEADERS = new HashSet<>(Arrays.asList("手机号码"));

    // 配置常量
    private static final int MAX_FILE_SIZE_MB = 1000; // 最大文件大小（MB）
    private static final int BATCH_SIZE = 1000; // 批处理大小
    private static final int MEMORY_WINDOW_SIZE = 100; // SXSSF内存窗口大小
    private static final String[] SUPPORTED_EXTENSIONS = {"xlsx", "xls"};
    private static final String ERROR_TITLE = "错误";
    private static final String SUCCESS_TITLE = "成功";
    private static final String INFO_TITLE = "提示";

    // 主窗口组件
    private JFrame mainFrame;
    private JTextArea logArea;
    private JLabel statusLabel;
    private JTabbedPane tabbedPane;

    // 表头映射工具相关组件
    private Map<String, String> mappingRelations;  // 用于建立映射关系标签页
    private JList<String> mappingList;            // 用于建立映射关系标签页
    private DefaultListModel<String> mappingListModel;  // 用于建立映射关系标签页

    // Excel处理器相关组件
    private DefaultListModel<String> sourceListModel;
    private DefaultListModel<String> templateListModel;
    private JList<String> sourceList;
    private JList<String> templateList;
    private Map<String, String> processorMappings;  // 用于Excel处理器标签页
    private List<String> processorMappingOrder;    // 用于Excel处理器标签页
    private JList<String> processorMappingList;    // 用于Excel处理器标签页
    private DefaultListModel<String> processorMappingListModel;  // 用于Excel处理器标签页
    private File currentMappingFile;
    private File currentTargetFile;

    private String inputFilePath;
    private String templateFilePath;
    private String outputFilePath;

    public ExcelTools() {
        // 设置POI的大文件处理参数
        IOUtils.setByteArrayMaxOverride(1000000000); // 设置更大的记录长度限制
        initializeUI();
    }

    private void initializeUI() {
        // 初始化数据
        mappingRelations = new HashMap<>();
        processorMappings = new HashMap<>();
        processorMappingOrder = new ArrayList<>();

        // 创建主窗口
        mainFrame = new JFrame("Excel工具集");
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setSize(1000, 700);
        mainFrame.setLocationRelativeTo(null);

        // 创建选项卡面板
        tabbedPane = new JTabbedPane();

        // 创建Excel处理器面板
        JPanel processorPanel = createProcessorPanel();
        tabbedPane.addTab("Excel处理器", processorPanel);

        // 创建映射关系面板
        JPanel mappingPanel = createMappingPanel();
        tabbedPane.addTab("建立映射关系", mappingPanel);

        // 创建应用映射面板
        JPanel applyPanel = createApplyPanel();
        tabbedPane.addTab("应用映射", applyPanel);

        // 创建日志区域
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder("操作记录"));
        logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logScrollPane.setPreferredSize(new Dimension(0, 100));
        logPanel.add(logScrollPane, BorderLayout.CENTER);

        // 创建状态标签
        statusLabel = new JLabel("就绪");

        // 布局
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        mainPanel.add(statusLabel, BorderLayout.NORTH);
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        mainPanel.add(logPanel, BorderLayout.SOUTH);

        mainFrame.add(mainPanel);
        mainFrame.setVisible(true);

        // 集成跨表主键填充工具
        JPanel crossSheetPanel = createCrossSheetPanel(logArea);
        tabbedPane.addTab("跨表主键填充工具", crossSheetPanel);
        // 集成重复列工具
        JPanel duplicateColPanel = createDuplicateColPanel(logArea);
        tabbedPane.addTab("重复列工具", duplicateColPanel);
    }

    private JPanel createProcessorPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 创建顶部面板
        JPanel topPanel = new JPanel(new BorderLayout(10, 10));

        // 创建按钮面板
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));

        // 创建选择文件按钮
        JButton selectFilesButton = new JButton("选择文件");
        selectFilesButton.setPreferredSize(new Dimension(100, 30));
        buttonPanel.add(selectFilesButton);

        // 新增重新导入文件按钮
        JButton reloadFilesButton = new JButton("重新导入文件");
        reloadFilesButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(reloadFilesButton);

        // 创建保存按钮
        JButton saveButton = new JButton("保存文件");
        saveButton.setPreferredSize(new Dimension(100, 30));
        saveButton.setEnabled(false);
        buttonPanel.add(saveButton);

        topPanel.add(buttonPanel, BorderLayout.CENTER);
        panel.add(topPanel, BorderLayout.NORTH);

        // 创建映射面板
        JPanel mappingConfigPanel = new JPanel(new BorderLayout(10, 10));
        mappingConfigPanel.setBorder(BorderFactory.createTitledBorder("配置字段映射关系"));

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
        JPanel mappingListPanel = new JPanel(new BorderLayout(5, 5));
        mappingListPanel.setBorder(BorderFactory.createTitledBorder("已配置的映射关系"));

        // 初始化Excel处理器的映射列表模型和列表
        processorMappingListModel = new DefaultListModel<>();
        processorMappingList = new JList<>(processorMappingListModel);
        processorMappingList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        processorMappingList.setVisibleRowCount(10);
        processorMappingList.setFixedCellHeight(25);

        // 创建滚动面板并设置首选大小
        JScrollPane mappingScrollPane = new JScrollPane(processorMappingList);
        mappingScrollPane.setPreferredSize(new Dimension(200, 250));
        mappingListPanel.add(mappingScrollPane, BorderLayout.CENTER);

        // 创建映射操作按钮面板
        JPanel mappingButtonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
        JButton addMappingButton = new JButton("添加映射");
        addMappingButton.setPreferredSize(new Dimension(100, 30));
        JButton removeMappingButton = new JButton("删除映射");
        removeMappingButton.setPreferredSize(new Dimension(100, 30));
        mappingButtonPanel.add(addMappingButton);
        mappingButtonPanel.add(removeMappingButton);
        mappingListPanel.add(mappingButtonPanel, BorderLayout.SOUTH);

        // 设置映射列表面板的首选大小
        mappingListPanel.setPreferredSize(new Dimension(250, 300));

        // 添加组件到表头面板
        headerPanel.add(sourcePanel);
        headerPanel.add(mappingListPanel);
        headerPanel.add(templatePanel);

        // 添加组件到映射面板
        mappingConfigPanel.add(headerPanel, BorderLayout.CENTER);

        // 添加组件到主面板
        panel.add(mappingConfigPanel, BorderLayout.CENTER);

        // 添加事件处理
        addProcessorPanelEventHandlers(selectFilesButton, saveButton, addMappingButton, removeMappingButton, reloadFilesButton);

        // 初始化Excel处理器的映射数据结构
        processorMappings = new HashMap<>();
        processorMappingOrder = new ArrayList<>();

        log("初始化映射面板完成");
        return panel;
    }

    private JPanel createMappingPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 创建按钮面板，使用GridLayout使三个按钮各占一半宽度
        JPanel buttonPanel = new JPanel(new GridLayout(1, 3, 10, 0));

        // 创建上传按钮
        JButton uploadButton = new JButton("上传关系表");
        uploadButton.setPreferredSize(new Dimension(100, 30));
        buttonPanel.add(uploadButton);

        // 创建保存按钮
        JButton saveButton = new JButton("保存映射关系");
        saveButton.setPreferredSize(new Dimension(100, 30));
        buttonPanel.add(saveButton);

        // 新增重置按钮
        JButton resetButton = new JButton("重置映射关系");
        resetButton.setPreferredSize(new Dimension(120, 30));
        buttonPanel.add(resetButton);

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

        // 添加事件处理
        addMappingPanelEventHandlers(uploadButton, saveButton, sourceList, targetList);

        // 新增重置按钮事件
        resetButton.addActionListener(e -> {
            mappingRelations.clear();
            mappingListModel.clear();
            statusLabel.setText("已重置映射关系");
            log("已重置所有已保存的映射关系");
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
        uploadButton.setPreferredSize(new Dimension(100, 30));
        buttonPanel.add(uploadButton);

        // 创建应用按钮
        JButton applyButton = new JButton("应用映射");
        applyButton.setPreferredSize(new Dimension(100, 30));
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

        // 添加事件处理
        addApplyPanelEventHandlers(uploadButton, applyButton, lookupList, fillList);

        return panel;
    }

    private void log(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private void addProcessorPanelEventHandlers(JButton selectFilesButton, JButton saveButton,
                                                JButton addMappingButton, JButton removeMappingButton,
                                                JButton reloadFilesButton) {
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

        // 添加重新导入文件按钮事件
        reloadFilesButton.addActionListener(e -> {
            try {
                // 重新选择文件
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

                // 选择源数据文件
                fileChooser.setDialogTitle("选择源数据文件");
                int result = fileChooser.showOpenDialog(mainFrame);
                if (result != JFileChooser.APPROVE_OPTION) {
                    return;
                }
                inputFilePath = fileChooser.getSelectedFile().getAbsolutePath();
                log("已重新选择源数据文件: " + inputFilePath);

                // 选择模板文件
                fileChooser.setDialogTitle("选择模板文件");
                result = fileChooser.showOpenDialog(mainFrame);
                if (result != JFileChooser.APPROVE_OPTION) {
                    return;
                }
                templateFilePath = fileChooser.getSelectedFile().getAbsolutePath();
                log("已重新选择模板文件: " + templateFilePath);

                // 加载表头
                loadHeaders(inputFilePath, templateFilePath);

                // 清空映射关系
                processorMappings.clear();
                processorMappingOrder.clear();
                updateProcessorMappingList();

                // 禁用保存按钮
                saveButton.setEnabled(false);

                statusLabel.setText("已重新导入文件，请重新配置映射关系");
            } catch (Exception ex) {
                logger.error("重新导入文件时发生错误", ex);
                log("重新导入文件失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainFrame,
                        "重新导入文件失败: " + ex.getMessage(),
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
                    if (processorMappings.containsKey(sourceHeader)) {
                        int confirm = JOptionPane.showConfirmDialog(mainFrame,
                                "源表头 '" + sourceHeader + "' 已经映射到 '" + processorMappings.get(sourceHeader) +
                                        "'，是否覆盖?", "确认覆盖", JOptionPane.YES_NO_OPTION);
                        if (confirm != JOptionPane.YES_OPTION) {
                            continue;
                        }
                        // 如果确认覆盖，从顺序列表中移除旧的映射
                        processorMappingOrder.remove(sourceHeader);
                    }

                    // 添加映射
                    processorMappings.put(sourceHeader, selectedTemplate);
                    // 添加到顺序列表
                    processorMappingOrder.add(sourceHeader);
                }

                // 更新映射列表显示
                updateProcessorMappingList();

                log("已添加Excel处理器映射: " + String.join(", ", selectedSources) + " -> " + selectedTemplate);
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
            String selectedMapping = processorMappingList.getSelectedValue();
            if (selectedMapping != null) {
                // 从映射字符串中提取源表头
                String sourceHeader = selectedMapping.split(" -> ")[0];

                // 删除映射
                processorMappings.remove(sourceHeader);
                // 从顺序列表中移除
                processorMappingOrder.remove(sourceHeader);

                // 更新映射列表显示
                updateProcessorMappingList();

                log("已删除Excel处理器映射: " + selectedMapping);
                statusLabel.setText("已删除映射关系");

                // 如果没有映射关系，禁用保存按钮
                if (processorMappings.isEmpty()) {
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
                if (processorMappings == null || processorMappings.isEmpty()) {
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
                processExcel(inputFilePath, templateFilePath, outputFilePath, processorMappings);

                // 显示成功消息
                log("处理完成！结果已保存到: " + outputFilePath);
                JOptionPane.showMessageDialog(mainFrame,
                        "处理完成！结果已保存到: " + outputFilePath,
                        "成功",
                        JOptionPane.INFORMATION_MESSAGE);

                // 禁用保存按钮
                saveButton.setEnabled(false);

                // 清空映射关系
                processorMappings.clear();
                processorMappingOrder.clear();
                updateProcessorMappingList();

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
    }

    private void updateProcessorMappingList() {
        if (processorMappingListModel == null) {
            log("错误：processorMappingListModel 未初始化");
            return;
        }
        if (processorMappingList == null) {
            log("错误：processorMappingList 未初始化");
            return;
        }

        SwingUtilities.invokeLater(() -> {
            processorMappingListModel.clear();
            log("开始更新Excel处理器映射列表，当前映射数量: " + processorMappings.size());

            // 按照添加顺序显示映射
            for (String sourceHeader : processorMappingOrder) {
                if (processorMappings.containsKey(sourceHeader)) {
                    String mappingText = sourceHeader + " -> " + processorMappings.get(sourceHeader);
                    processorMappingListModel.addElement(mappingText);
                    log("添加Excel处理器映射到列表: " + mappingText);
                }
            }

            // 强制刷新列表显示
            processorMappingList.repaint();
            log("Excel处理器映射列表更新完成，当前显示项数: " + processorMappingListModel.getSize());
        });
    }

    /**
     * 使用SAX模式只读取Excel文件的表头行
     */
    private List<String> readHeadersOnly(File file) throws IOException {
        if (file.getName().toLowerCase().endsWith(".xlsx")) {
            return readXLSXHeaders(file);
        } else {
            return readXLSHeaders(file);
        }
    }

    /**
     * 使用SAX模式读取XLSX文件的表头
     */
    private List<String> readXLSXHeaders(File file) throws IOException {
        List<String> headers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             OPCPackage pkg = OPCPackage.open(fis)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();

            if (it.hasNext()) {
                InputStream sheet = it.next();
                try {
                    XMLReader parser = javax.xml.parsers.SAXParserFactory.newInstance()
                            .newSAXParser().getXMLReader();
                    try {
                        parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                        parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
                        parser.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
                    } catch (SAXNotRecognizedException | SAXNotSupportedException e) {
                        logger.error("设置XML解析器特性时出错: " + e.getMessage());
                    }

                    HeaderHandler handler = new HeaderHandler(sst);
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheet));
                    headers = handler.getHeaders();
                } catch (SAXException e) {
                    throw new IOException("解析Excel文件时出错: " + e.getMessage(), e);
                } finally {
                    sheet.close();
                }
            }
        } catch (OpenXML4JException | javax.xml.parsers.ParserConfigurationException e) {
            throw new IOException("打开Excel文件时出错: " + e.getMessage(), e);
        }
        return headers;
    }

    /**
     * 读取XLS文件的表头
     */
    private List<String> readXLSHeaders(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            List<String> headers = new ArrayList<>();
            if (headerRow != null) {
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        headers.add(cell.getStringCellValue().trim());
                    }
                }
            }
            return headers;
        }
    }

    /**
     * SAX处理器，用于解析XLSX文件的表头，支持共享字符串
     */
    private static class HeaderHandler extends DefaultHandler {
        private final List<String> headers = new ArrayList<>();
        private final SharedStringsTable sst;
        private boolean isHeaderRow = true;
        private String cellType = null;
        private StringBuilder currentValue = new StringBuilder();
        private boolean isCellValue = false;

        public HeaderHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (isHeaderRow) {
                if ("c".equals(qName)) {
                    cellType = attributes.getValue("t");
                    isCellValue = true;
                }
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCellValue) {
                currentValue.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (isHeaderRow) {
                if ("c".equals(qName)) {
                    isCellValue = false;
                    String value = currentValue.toString().trim();
                    if (!value.isEmpty()) {
                        if ("s".equals(cellType) && sst != null) {
                            int idx = Integer.parseInt(value);
                            value = sst.getItemAt(idx).getString();
                        }
                        headers.add(value);
                    }
                    currentValue.setLength(0);
                    cellType = null;
                } else if ("row".equals(qName)) {
                    isHeaderRow = false;
                }
            }
        }

        public List<String> getHeaders() {
            return headers;
        }
    }

    /**
     * SAX处理器，用于流式解析XLSX文件的全部数据（表头+数据行）
     * 通过回调逐行处理，避免将全部数据加载到内存
     */
    private static class SheetDataHandler extends DefaultHandler {
        private final SharedStringsTable sst;
        private final List<CellStyle> cellStyles;
        private final java.util.function.Consumer<List<String>> rowCallback;

        private List<String> headers;
        private boolean isFirstRow = true;
        private String cellType = null;
        private int cellStyleIndex = 0;
        private int currentColIndex = 0;
        private final TreeMap<Integer, String> currentRowData = new TreeMap<>();
        private StringBuilder currentValue = new StringBuilder();
        private boolean isCellValue = false;
        private int totalRows = 0;

        public SheetDataHandler(SharedStringsTable sst, List<CellStyle> cellStyles,
                                java.util.function.Consumer<List<String>> rowCallback) {
            this.sst = sst;
            this.cellStyles = cellStyles;
            this.rowCallback = rowCallback;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("c".equals(qName)) {
                String ref = attributes.getValue("r");
                if (ref != null) {
                    currentColIndex = parseColumnRef(ref);
                }
                cellType = attributes.getValue("t");
                String styleStr = attributes.getValue("s");
                cellStyleIndex = (styleStr != null) ? Integer.parseInt(styleStr) : 0;
                isCellValue = true;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCellValue) {
                currentValue.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if ("c".equals(qName)) {
                isCellValue = false;
                String value = currentValue.toString().trim();
                if ("s".equals(cellType) && sst != null) {
                    int idx = Integer.parseInt(value);
                    value = sst.getItemAt(idx).getString();
                } else if ("inlineStr".equals(cellType)) {
                    // 内联字符串，直接使用
                } else if (value != null && !value.isEmpty()) {
                    // 数值类型，检查是否为日期
                    if (cellStyleIndex > 0 && cellStyleIndex < cellStyles.size()) {
                        CellStyle style = cellStyles.get(cellStyleIndex);
                        if (style != null && DateUtil.isADateFormat(style.getDataFormat(), style.getDataFormatString())) {
                            try {
                                double excelDate = Double.parseDouble(value);
                                value = DateUtil.getJavaDate(excelDate).toString();
                            } catch (NumberFormatException ignored) {
                            }
                        }
                    }
                }
                currentRowData.put(currentColIndex, (value != null) ? value : "");
                currentValue.setLength(0);
                cellType = null;
                cellStyleIndex = 0;
            } else if ("row".equals(qName)) {
                totalRows++;
                List<String> row = new ArrayList<>();
                if (!currentRowData.isEmpty()) {
                    int maxCol = currentRowData.lastKey();
                    for (int i = 0; i <= maxCol; i++) {
                        row.add(currentRowData.getOrDefault(i, ""));
                    }
                }
                currentRowData.clear();

                if (isFirstRow) {
                    headers = row;
                    isFirstRow = false;
                } else {
                    rowCallback.accept(row);
                }
            }
        }

        public List<String> getHeaders() {
            return headers;
        }

        public int getTotalRows() {
            return totalRows;
        }

        private int parseColumnRef(String ref) {
            int col = 0;
            for (int i = 0; i < ref.length(); i++) {
                char c = ref.charAt(i);
                if (c >= 'A' && c <= 'Z') {
                    col = col * 26 + (c - 'A' + 1);
                } else {
                    break;
                }
            }
            return col - 1;
        }
    }

    /**
     * 使用SAX流式解析XLSX文件，逐行处理数据，避免OOM
     */
    private void processSheetDataWithSAX(File file, SharedStringsTable sst, List<CellStyle> cellStyles,
                                          java.util.function.Consumer<List<String>> rowCallback) throws Exception {
        try (FileInputStream fis = new FileInputStream(file);
             OPCPackage pkg = OPCPackage.open(fis)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();

            if (it.hasNext()) {
                InputStream sheet = it.next();
                try {
                    XMLReader parser = javax.xml.parsers.SAXParserFactory.newInstance()
                            .newSAXParser().getXMLReader();
                    try {
                        parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                        parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
                        parser.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
                    } catch (SAXNotRecognizedException | SAXNotSupportedException e) {
                        logger.error("设置XML解析器特性时出错: " + e.getMessage());
                    }

                    SheetDataHandler handler = new SheetDataHandler(sst, cellStyles, rowCallback);
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheet));
                } finally {
                    sheet.close();
                }
            }
        }
    }

    private void loadHeaders(String inputFilePath, String templateFilePath) {
        // 创建进度对话框
        JDialog progressDialog = new JDialog(mainFrame, "加载表头", true);
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
        progressDialog.setLocationRelativeTo(mainFrame);

        // 在后台线程中加载表头
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                try {
                    // 加载源数据文件表头
                    SwingUtilities.invokeLater(() -> {
                        progressLabel.setText("正在加载源数据文件表头...");
                    });

                    List<String> sourceHeaders = readHeadersOnly(new File(inputFilePath));

                    // 更新源数据表头列表
                    SwingUtilities.invokeLater(() -> {
                        sourceListModel.clear();
                        for (String header : sourceHeaders) {
                            sourceListModel.addElement(header);
                        }
                        sourceList.repaint();
                    });

                    // 加载模板文件表头
                    SwingUtilities.invokeLater(() -> {
                        progressLabel.setText("正在加载模板文件表头...");
                    });

                    List<String> templateHeaders = readHeadersOnly(new File(templateFilePath));

                    // 更新模板表头列表
                    SwingUtilities.invokeLater(() -> {
                        templateListModel.clear();
                        for (String header : templateHeaders) {
                            templateListModel.addElement(header);
                        }
                        templateList.repaint();

                        // 更新状态
                        statusLabel.setText("已加载表头，请配置映射关系");
                        log("成功加载表头");
                    });

                } catch (Exception e) {
                    throw new IOException("加载表头失败: " + e.getMessage(), e);
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
            logger.error("加载表头时发生错误", e);
            JOptionPane.showMessageDialog(mainFrame,
                    "加载表头失败: " + e.getMessage(),
                    "错误",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void addMappingPanelEventHandlers(JButton uploadButton, JButton saveButton,
                                              JList<String> sourceList, JList<String> targetList) {
        // 添加上传按钮事件
        uploadButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

            int result = fileChooser.showOpenDialog(mainFrame);
            if (result == JFileChooser.APPROVE_OPTION) {
                currentMappingFile = fileChooser.getSelectedFile();
                loadMappingHeaders(currentMappingFile, sourceList, targetList);
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
    }

    private void loadMappingHeaders(File file, JList<String> sourceList, JList<String> targetList) {
        try (FileInputStream fis = new FileInputStream(file)) {
            // 使用XSSFWorkbook来处理大文件
            Workbook workbook = new XSSFWorkbook(fis);
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
            workbook.close();

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

    private void addApplyPanelEventHandlers(JButton uploadButton, JButton applyButton,
                                            JList<String> lookupList, JList<String> fillList) {
        // 添加上传按钮事件
        uploadButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));

            int result = fileChooser.showOpenDialog(mainFrame);
            if (result == JFileChooser.APPROVE_OPTION) {
                currentTargetFile = fileChooser.getSelectedFile();
                loadMappingHeaders(currentTargetFile, lookupList, fillList);
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

            ProgressDialog progressDialog = createProgressDialog("处理中");

            // 构建反向映射（value -> key），避免每行都遍历
            Map<String, String> reverseMapping = new HashMap<>();
            for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                reverseMapping.putIfAbsent(entry.getValue(), entry.getKey());
            }

            // 在后台线程中处理数据
            SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                @Override
                protected Void doInBackground() throws Exception {
                    if (file.getName().toLowerCase().endsWith(".xlsx")) {
                        // XLSX文件：SAX流式读取 + SXSSFWorkbook流式写入
                        processApplyMappingXlsx(file, outputFile, lookupField, fillField, reverseMapping,
                                progressDialog.label, progressDialog.progressBar);
                    } else {
                        // XLS文件：POI读取 + SXSSFWorkbook写入
                        processApplyMappingXls(file, outputFile, lookupField, fillField, reverseMapping,
                                progressDialog.label, progressDialog.progressBar);
                    }
                    return null;
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
            } catch (Exception e) {
                logger.error("应用映射关系时出错", e);
                JOptionPane.showMessageDialog(mainFrame,
                        "应用映射关系时出错: " + e.getMessage(),
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * XLSX文件：SAX流式读取 + SXSSFWorkbook流式写入
     */
    private void processApplyMappingXlsx(File inputFile, File outputFile, String lookupField, String fillField,
                                          Map<String, String> reverseMapping, JLabel progressLabel,
                                          JProgressBar progressBar) throws Exception {
        // 先读取表头和共享字符串
        List<String> headers;
        SharedStringsTable sst;
        List<CellStyle> cellStyles = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(inputFile);
             OPCPackage pkg = OPCPackage.open(fis)) {
            XSSFReader reader = new XSSFReader(pkg);
            sst = (SharedStringsTable) reader.getSharedStringsTable();

            try {
                StylesTable stylesTable = reader.getStylesTable();
                for (int i = 0; i < stylesTable.getNumCellStyles(); i++) {
                    cellStyles.add(stylesTable.getStyleAt(i));
                }
            } catch (Exception e) {
                logger.warn("无法读取样式表: " + e.getMessage());
            }

            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();
            if (it.hasNext()) {
                InputStream sheet = it.next();
                try {
                    XMLReader parser = javax.xml.parsers.SAXParserFactory.newInstance()
                            .newSAXParser().getXMLReader();
                    HeaderHandler handler = new HeaderHandler(sst);
                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(sheet));
                    headers = handler.getHeaders();
                } finally {
                    sheet.close();
                }
            } else {
                throw new IOException("源文件中没有找到工作表");
            }
        }

        // 获取查找字段和填充字段的索引
        int lookupIndex = -1;
        int fillIndex = -1;
        for (int i = 0; i < headers.size(); i++) {
            if (headers.get(i).trim().equals(lookupField)) lookupIndex = i;
            if (headers.get(i).trim().equals(fillField)) fillIndex = i;
        }
        if (lookupIndex < 0 || fillIndex < 0) {
            throw new IOException("找不到指定的字段: " + lookupField + " 或 " + fillField);
        }

        SwingUtilities.invokeLater(() -> {
            statusLabel.setText("正在应用映射关系...");
            log("开始应用映射关系（SAX流式模式）");
        });

        final int lkIdx = lookupIndex;
        final int flIdx = fillIndex;
        final int[] stats = {0, 0}; // [processedRows, updatedCount]
        final int[] rowNum = {1};

        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

            // 写入表头
            Row outputHeaderRow = outputSheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                outputHeaderRow.createCell(i).setCellValue(headers.get(i));
            }

            // SAX流式读取并逐行处理
            processSheetDataWithSAX(inputFile, sst, cellStyles, sourceRow -> {
                if (sourceRow.isEmpty()) return;

                Row outRow = outputSheet.createRow(rowNum[0]++);

                // 复制所有单元格
                for (int j = 0; j < sourceRow.size(); j++) {
                    String val = sourceRow.get(j);
                    if (val != null && !val.isEmpty()) {
                        outRow.createCell(j).setCellValue(val);
                    }
                }

                // 获取查找字段的值并应用映射
                if (lkIdx < sourceRow.size()) {
                    String lookupValue = sourceRow.get(lkIdx);
                    if (lookupValue != null && !lookupValue.isEmpty()) {
                        String mappedValue = mappingRelations.get(lookupValue);
                        if (mappedValue == null) {
                            mappedValue = reverseMapping.get(lookupValue);
                        }

                        Cell fillCell = outRow.getCell(flIdx);
                        if (fillCell == null) {
                            fillCell = outRow.createCell(flIdx);
                        }

                        if (mappedValue != null) {
                            fillCell.setCellValue(mappedValue);
                            stats[1]++;
                        } else {
                            fillCell.setCellValue(lookupValue);
                        }
                    }
                }

                stats[0]++;
                if (stats[0] % BATCH_SIZE == 0) {
                    try {
                        ((SXSSFSheet) outputSheet).flushRows();
                    } catch (IOException e) {
                        throw new RuntimeException("刷新行数据失败", e);
                    }
                    final int processed = stats[0];
                    SwingUtilities.invokeLater(() -> {
                        progressBar.setIndeterminate(false);
                        progressBar.setValue(Math.min(99, processed / 1000));
                        progressLabel.setText("正在处理数据... 已处理 " + processed + " 行");
                    });
                }
            });

            // 保存文件
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }

        final int totalProcessed = stats[0];
        final int totalUpdated = stats[1];
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(100);
            statusLabel.setText("已更新 " + totalUpdated + " 条记录");
            log("处理完成，共处理 " + totalProcessed + " 行，更新 " + totalUpdated + " 条记录");
        });
    }

    /**
     * XLS文件：POI读取 + SXSSFWorkbook写入
     */
    private void processApplyMappingXls(File inputFile, File outputFile, String lookupField, String fillField,
                                         Map<String, String> reverseMapping, JLabel progressLabel,
                                         JProgressBar progressBar) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Integer> headerIndexMap = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    headerIndexMap.put(cell.getStringCellValue().trim(), i);
                }
            }

            Integer lookupIndex = headerIndexMap.get(lookupField);
            Integer fillIndex = headerIndexMap.get(fillField);
            if (lookupIndex == null || fillIndex == null) {
                throw new IOException("找不到指定的字段");
            }

            SwingUtilities.invokeLater(() -> {
                statusLabel.setText("正在应用映射关系...");
                log("开始应用映射关系");
            });

            try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
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

                final int totalRows = sheet.getLastRowNum();
                int processedRows = 0;
                int updatedCount = 0;

                for (int i = 1; i <= totalRows; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    Row outputRow = outputSheet.createRow(i);

                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            Cell outputCell = outputRow.createCell(j);
                            copyCellValue(cell, outputCell);
                        }
                    }

                    Cell lookupCell = row.getCell(lookupIndex);
                    if (lookupCell != null) {
                        String lookupValue = getCellValueAsString(lookupCell);

                        Cell outputFillCell = outputRow.getCell(fillIndex);
                        if (outputFillCell == null) {
                            outputFillCell = outputRow.createCell(fillIndex);
                        }

                        String mappedValue = mappingRelations.get(lookupValue);
                        if (mappedValue == null) {
                            mappedValue = reverseMapping.get(lookupValue);
                        }

                        if (mappedValue != null) {
                            outputFillCell.setCellValue(mappedValue);
                            updatedCount++;
                        } else {
                            outputFillCell.setCellValue(lookupValue);
                        }
                    }

                    processedRows++;
                    if (processedRows % BATCH_SIZE == 0) {
                        ((SXSSFSheet) outputSheet).flushRows();
                        final int progress = (int) ((i * 100.0) / totalRows);
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setIndeterminate(false);
                            progressBar.setValue(progress);
                            progressLabel.setText("正在处理数据... " + progress + "%");
                        });
                    }
                }

                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }

                final int finalUpdatedCount = updatedCount;
                SwingUtilities.invokeLater(() -> {
                    progressBar.setValue(100);
                    statusLabel.setText("已更新 " + finalUpdatedCount + " 条记录");
                    log("处理完成，共更新 " + finalUpdatedCount + " 条记录");
                });
            }
        }
    }

    private void processExcel(String inputFilePath, String templateFilePath, String outputFilePath,
                              Map<String, String> mappings) throws IOException {
        // 验证文件
        File inputFile = new File(inputFilePath);
        File templateFile = new File(templateFilePath);
        File outputFile = new File(outputFilePath);

        validateFile(inputFile, "源数据");
        validateFile(templateFile, "模板");

        // 检查输出目录
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
        ProgressDialog progressDialog = createProgressDialog("处理中");
        JLabel progressLabel = progressDialog.label;
        JProgressBar progressBar = progressDialog.progressBar;

        // 在后台线程中处理数据
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                // 读取模板文件的表头（模板文件通常较小，用POI读取即可）
                List<String> templateHeaders;
                try (FileInputStream templateFis = new FileInputStream(templateFile);
                     Workbook templateWorkbook = WorkbookFactory.create(templateFis)) {
                    Sheet templateSheet = templateWorkbook.getSheetAt(0);
                    templateHeaders = new ArrayList<>();
                    Row templateHeaderRow = templateSheet.getRow(0);
                    if (templateHeaderRow != null) {
                        for (int i = 0; i < templateHeaderRow.getLastCellNum(); i++) {
                            Cell cell = templateHeaderRow.getCell(i);
                            if (cell != null) {
                                templateHeaders.add(cell.getStringCellValue());
                            }
                        }
                    }
                }

                // 将映射转换为HeaderMapping对象
                List<HeaderMapping> headerMappings = new ArrayList<>();
                Map<String, List<String>> targetToSourceMap = new HashMap<>();
                for (Map.Entry<String, String> entry : mappings.entrySet()) {
                    targetToSourceMap.computeIfAbsent(entry.getValue(), k -> new ArrayList<>()).add(entry.getKey());
                }
                for (Map.Entry<String, List<String>> entry : targetToSourceMap.entrySet()) {
                    headerMappings.add(new HeaderMapping(entry.getValue(), entry.getKey()));
                }

                // 预建模板表头索引Map
                Map<String, Integer> templateHeaderIndexMap = new HashMap<>();
                for (int i = 0; i < templateHeaders.size(); i++) {
                    templateHeaderIndexMap.put(templateHeaders.get(i), i);
                }

                // 使用SXSSFWorkbook流式写入输出
                try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
                    Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

                    // 写入模板表头
                    Row outputHeaderRow = outputSheet.createRow(0);
                    for (int i = 0; i < templateHeaders.size(); i++) {
                        Cell cell = outputHeaderRow.createCell(i);
                        cell.setCellValue(templateHeaders.get(i));
                    }

                    if (inputFile.getName().toLowerCase().endsWith(".xlsx")) {
                        // XLSX文件：使用SAX流式读取，避免OOM
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("正在使用SAX流式处理数据...");
                            log("开始流式处理XLSX文件");
                        });

                        // 先读取表头和单元格样式
                        List<String> sourceHeaders;
                        SharedStringsTable sst;
                        List<CellStyle> cellStyles = new ArrayList<>();
                        try (FileInputStream fis = new FileInputStream(inputFile);
                             OPCPackage pkg = OPCPackage.open(fis)) {
                            XSSFReader reader = new XSSFReader(pkg);
                            sst = (SharedStringsTable) reader.getSharedStringsTable();

                            // 读取样式表
                            try {
                                StylesTable stylesTable = reader.getStylesTable();
                                for (int i = 0; i < stylesTable.getNumCellStyles(); i++) {
                                    cellStyles.add(stylesTable.getStyleAt(i));
                                }
                            } catch (Exception e) {
                                logger.warn("无法读取样式表: " + e.getMessage());
                            }

                            // 读取表头
                            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();
                            if (it.hasNext()) {
                                InputStream sheet = it.next();
                                try {
                                    XMLReader parser = javax.xml.parsers.SAXParserFactory.newInstance()
                                            .newSAXParser().getXMLReader();
                                    HeaderHandler handler = new HeaderHandler(sst);
                                    parser.setContentHandler(handler);
                                    parser.parse(new InputSource(sheet));
                                    sourceHeaders = handler.getHeaders();
                                } finally {
                                    sheet.close();
                                }
                            } else {
                                throw new IOException("源文件中没有找到工作表");
                            }
                        }

                        // 构建源表头索引
                        Map<String, Integer> sourceHeaderMap = new LinkedHashMap<>();
                        for (int i = 0; i < sourceHeaders.size(); i++) {
                            sourceHeaderMap.put(sourceHeaders.get(i), i);
                        }

                        SwingUtilities.invokeLater(() -> log("源表表头: " + sourceHeaders));

                        // 使用SAX流式解析并逐行写入输出
                        final int[] processedRows = {0};
                        final int[] rowNum = {1};

                        processSheetDataWithSAX(inputFile, sst, cellStyles, sourceRow -> {
                            if (sourceRow.isEmpty()) return;

                            Row outRow = outputSheet.createRow(rowNum[0]++);

                            for (HeaderMapping mapping : headerMappings) {
                                if (mapping.isMerge()) {
                                    StringBuilder mergedValue = new StringBuilder();
                                    for (String sourceHeader : mapping.getSourceHeaders()) {
                                        Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                        if (sourceIndex != null && sourceIndex < sourceRow.size()) {
                                            String value = sourceRow.get(sourceIndex);
                                            if (value != null && !value.isEmpty()) {
                                                if (mergedValue.length() > 0) {
                                                    mergedValue.append(mapping.getFormatType() == HeaderMapping.FormatType.COMMA_SEPARATED ?
                                                            "," : OTHER_REMARK_SEPARATOR);
                                                }
                                                if (mapping.getFormatType() == HeaderMapping.FormatType.PIPE_SEPARATED) {
                                                    mergedValue.append(sourceHeader).append("：");
                                                }
                                                mergedValue.append(value);
                                            }
                                        }
                                    }
                                    if (mergedValue.length() > 0) {
                                        int targetIdx = templateHeaderIndexMap.getOrDefault(mapping.getTargetHeader(), -1);
                                        if (targetIdx >= 0) {
                                            outRow.createCell(targetIdx).setCellValue(mergedValue.toString());
                                        }
                                    }
                                } else {
                                    String sourceHeader = mapping.getSourceHeaders().get(0);
                                    Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                    if (sourceIndex != null && sourceIndex < sourceRow.size()) {
                                        String value = sourceRow.get(sourceIndex);
                                        if (value != null && !value.isEmpty()) {
                                            int targetIdx = templateHeaderIndexMap.getOrDefault(mapping.getTargetHeader(), -1);
                                            if (targetIdx >= 0) {
                                                outRow.createCell(targetIdx).setCellValue(value);
                                            }
                                        }
                                    }
                                }
                            }

                            processedRows[0]++;
                            if (processedRows[0] % BATCH_SIZE == 0) {
                                try {
                                    ((SXSSFSheet) outputSheet).flushRows();
                                } catch (IOException e) {
                                    throw new RuntimeException("刷新行数据失败", e);
                                }
                                final int pct = processedRows[0];
                                SwingUtilities.invokeLater(() -> {
                                    progressBar.setIndeterminate(false);
                                    progressBar.setValue(Math.min(99, pct / 100));
                                    progressLabel.setText("正在处理数据... 已处理 " + pct + " 行");
                                });
                            }
                        });

                        // 保存文件
                        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                            outputWorkbook.write(fos);
                        }

                        final int total = processedRows[0];
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(100);
                            statusLabel.setText("处理完成，共处理 " + total + " 行数据");
                            log("处理完成，共处理 " + total + " 行数据");
                        });

                    } else {
                        // XLS文件：使用POI读取（XLS文件通常较小）
                        try (FileInputStream fis = new FileInputStream(inputFile);
                             Workbook workbook = WorkbookFactory.create(fis)) {
                            Sheet sheet = workbook.getSheetAt(0);

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

                            int totalRows = sheet.getLastRowNum();
                            int processedRows = 0;
                            int rowNum = 1;

                            SwingUtilities.invokeLater(() -> {
                                statusLabel.setText("正在处理数据...");
                                log("开始处理数据，共 " + totalRows + " 行");
                            });

                            for (int i = 1; i <= totalRows; i += BATCH_SIZE) {
                                int endRow = Math.min(i + BATCH_SIZE - 1, totalRows);
                                for (int j = i; j <= endRow; j++) {
                                    Row sourceRow = sheet.getRow(j);
                                    if (sourceRow == null) continue;
                                    Row outRow = outputSheet.createRow(rowNum++);

                                    for (HeaderMapping mapping : headerMappings) {
                                        if (mapping.isMerge()) {
                                            StringBuilder mergedValue = new StringBuilder();
                                            for (String sourceHeader : mapping.getSourceHeaders()) {
                                                Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                                if (sourceIndex != null) {
                                                    Cell sourceCell = sourceRow.getCell(sourceIndex);
                                                    if (sourceCell != null) {
                                                        String value = getCellValueAsString(sourceCell);
                                                        if (!value.isEmpty()) {
                                                            if (mergedValue.length() > 0) {
                                                                mergedValue.append(mapping.getFormatType() == HeaderMapping.FormatType.COMMA_SEPARATED ?
                                                                        "," : OTHER_REMARK_SEPARATOR);
                                                            }
                                                            if (mapping.getFormatType() == HeaderMapping.FormatType.PIPE_SEPARATED) {
                                                                mergedValue.append(sourceHeader).append("：");
                                                            }
                                                            mergedValue.append(value);
                                                        }
                                                    }
                                                }
                                            }
                                            if (mergedValue.length() > 0) {
                                                int targetIdx = templateHeaderIndexMap.getOrDefault(mapping.getTargetHeader(), -1);
                                                if (targetIdx >= 0) {
                                                    outRow.createCell(targetIdx).setCellValue(mergedValue.toString());
                                                }
                                            }
                                        } else {
                                            String sourceHeader = mapping.getSourceHeaders().get(0);
                                            Integer sourceIndex = sourceHeaderMap.get(sourceHeader);
                                            if (sourceIndex != null) {
                                                Cell sourceCell = sourceRow.getCell(sourceIndex);
                                                if (sourceCell != null) {
                                                    String value = getCellValueAsString(sourceCell);
                                                    if (!value.isEmpty()) {
                                                        int targetIdx = templateHeaderIndexMap.getOrDefault(mapping.getTargetHeader(), -1);
                                                        if (targetIdx >= 0) {
                                                            outRow.createCell(targetIdx).setCellValue(value);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    processedRows++;
                                }

                                final int progress = (int) ((processedRows * 100.0) / totalRows);
                                SwingUtilities.invokeLater(() -> {
                                    progressBar.setIndeterminate(false);
                                    progressBar.setValue(progress);
                                    progressLabel.setText("正在处理数据... " + progress + "%");
                                });
                                ((SXSSFSheet) outputSheet).flushRows();
                            }

                            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                                outputWorkbook.write(fos);
                            }

                            final int finalProcessedRows = processedRows;
                            SwingUtilities.invokeLater(() -> {
                                statusLabel.setText("处理完成，共处理 " + finalProcessedRows + " 行数据");
                                log("处理完成，共处理 " + finalProcessedRows + " 行数据");
                            });
                        }
                    }
                }
                return null;
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
        } catch (Exception e) {
            logger.error("处理Excel文件时发生错误", e);
            throw new RuntimeException("处理Excel文件失败: " + e.getMessage());
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
                // 对于数值类型，直接返回字符串形式
                double value = cell.getNumericCellValue();
                // 检查是否为整数
                if (value == Math.floor(value)) {
                    return String.format("%.0f", value);
                }
                return String.valueOf(value);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // 对于公式，获取其计算结果
                try {
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return cell.getDateCellValue().toString();
                            }
                            double formulaValue = cell.getNumericCellValue();
                            if (formulaValue == Math.floor(formulaValue)) {
                                return String.format("%.0f", formulaValue);
                            }
                            return String.valueOf(formulaValue);
                        case BOOLEAN:
                            return String.valueOf(cell.getBooleanCellValue());
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    logger.warn("获取公式计算结果时出错: " + e.getMessage());
                    return "";
                }
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

    public static class HeaderMapping {
        private final List<String> sourceHeaders;  // 源表头（表A）
        private final String targetHeader;         // 目标表头（表B）
        private final FormatType formatType;       // 格式类型

        public enum FormatType {
            COMMA_SEPARATED,    // 使用逗号分隔，不显示源表头
            PIPE_SEPARATED      // 使用竖线分隔，显示源表头
        }

        public HeaderMapping(List<String> sourceHeaders, String targetHeader) {
            this.sourceHeaders = new ArrayList<>(sourceHeaders);
            this.targetHeader = targetHeader;
            // 根据目标表头名称确定格式类型
            this.formatType = COMMA_SEPARATED_HEADERS.contains(targetHeader) ?
                FormatType.COMMA_SEPARATED : FormatType.PIPE_SEPARATED;
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

        public FormatType getFormatType() {
            return formatType;
        }
    }

    /**
     * 验证文件是否有效
     */
    private void validateFile(File file, String fileType) throws IOException {
        if (!file.exists()) {
            throw new IOException(fileType + "文件不存在: " + file.getAbsolutePath());
        }
        if (!file.canRead()) {
            throw new IOException("无法读取" + fileType + "文件，请检查文件权限: " + file.getAbsolutePath());
        }
        if (file.length() > MAX_FILE_SIZE_MB * 1024 * 1024) {
            throw new IOException(fileType + "文件过大，超过" + MAX_FILE_SIZE_MB + "MB限制");
        }
        String extension = getFileExtension(file.getName());
        if (!Arrays.asList(SUPPORTED_EXTENSIONS).contains(extension.toLowerCase())) {
            throw new IOException("不支持的文件格式: " + extension);
        }
    }

    /**
     * 获取文件扩展名
     */
    private String getFileExtension(String fileName) {
        int lastDotIndex = fileName.lastIndexOf('.');
        return lastDotIndex > 0 ? fileName.substring(lastDotIndex + 1) : "";
    }

    /**
     * 进度对话框包装类，避免脆弱的 getComponent() 层级引用
     */
    private static class ProgressDialog {
        final JDialog dialog;
        final JLabel label;
        final JProgressBar progressBar;

        ProgressDialog(JDialog dialog, JLabel label, JProgressBar progressBar) {
            this.dialog = dialog;
            this.label = label;
            this.progressBar = progressBar;
        }
    }

    /**
     * 创建进度对话框
     */
    private ProgressDialog createProgressDialog(String title) {
        JDialog dialog = new JDialog(mainFrame, title, true);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.setResizable(false);
        dialog.setUndecorated(false);

        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JLabel label = new JLabel("正在处理...");
        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);

        panel.add(label, BorderLayout.NORTH);
        panel.add(progressBar, BorderLayout.CENTER);

        dialog.add(panel);
        dialog.pack();
        dialog.setLocationRelativeTo(mainFrame);

        return new ProgressDialog(dialog, label, progressBar);
    }

    /**
     * 显示错误消息
     */
    private void showError(String message) {
        logger.error(message);
        JOptionPane.showMessageDialog(mainFrame, message, ERROR_TITLE, JOptionPane.ERROR_MESSAGE);
    }

    /**
     * 显示成功消息
     */
    private void showSuccess(String message) {
        logger.info(message);
        JOptionPane.showMessageDialog(mainFrame, message, SUCCESS_TITLE, JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * 显示提示消息
     */
    private void showInfo(String message) {
        logger.info(message);
        JOptionPane.showMessageDialog(mainFrame, message, INFO_TITLE, JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * 创建文件选择器
     */
    private JFileChooser createFileChooser(String title, boolean isSave) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", SUPPORTED_EXTENSIONS));
        chooser.setDialogTitle(title);
        return chooser;
    }

    // 跨表主键填充工具 tab
    private JPanel createCrossSheetPanel(JTextArea logArea) {
        java.awt.Font uiFont = new java.awt.Font("Microsoft YaHei", java.awt.Font.PLAIN, 16);
        JPanel filePanel = new JPanel(new GridBagLayout());
        filePanel.setBorder(BorderFactory.createTitledBorder("文件选择区"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1;
        JTextField aPath = new JTextField();
        JTextField bPath = new JTextField();
        aPath.setEditable(false);
        bPath.setEditable(false);
        JButton selA = new JButton("选择源表");
        JButton selB = new JButton("选择目标表");
        java.awt.Font btnFont = new java.awt.Font("Microsoft YaHei", java.awt.Font.BOLD, 16);
        selA.setFont(btnFont);
        selB.setFont(btnFont);
        selA.setBorderPainted(false);
        selB.setBorderPainted(false);
        selA.setFocusPainted(false);
        selB.setFocusPainted(false);
        selA.setOpaque(true);
        selB.setOpaque(true);
        java.awt.Color blue = new java.awt.Color(66, 133, 244);
        java.awt.Color grayBg = new java.awt.Color(220, 220, 220);
        java.awt.Color grayFont = new java.awt.Color(120, 120, 120);
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
        setButtonVisualEnabled.accept(selA, true);
        setButtonVisualEnabled.accept(selB, true);
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

        JPanel opPanel = new JPanel(new BorderLayout(10, 10));
        // 修改开始填充按钮样式为默认
        JButton runBtn = new JButton("开始填充");
        // 不设置 setFont、setBackground、setForeground、setFocusPainted、setOpaque 等，仅保留默认样式
        opPanel.add(runBtn, BorderLayout.CENTER);

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

        // 事件
        selA.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                aPath.setText(fc.getSelectedFile().getAbsolutePath());
                logArea.append("已选择源表: " + aPath.getText() + "\n");
                ProgressDialog progressDialog = createProgressDialog("加载源表表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try {
                            headers = new com.xzy.CrossSheetFiller().readHeadersOnly(new File(aPath.getText()));
                        } catch (Exception ex) {
                            headers = null;
                        }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        if (headers != null) {
                            aHeadersArea.setText(String.join(", ", headers));
                            aKeyCombo.removeAllItems();
                            aValCombo.removeAllItems();
                            for (String h : headers) {
                                aKeyCombo.addItem(h);
                                aValCombo.addItem(h);
                            }
                            if (aKeyCombo.getItemCount() > 0) aKeyCombo.setSelectedIndex(0);
                            if (aValCombo.getItemCount() > 0) aValCombo.setSelectedIndex(0);
                        } else {
                            aHeadersArea.setText("读取源表表头失败");
                            aKeyCombo.removeAllItems();
                            aValCombo.removeAllItems();
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
            }
        });
        selB.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                bPath.setText(fc.getSelectedFile().getAbsolutePath());
                logArea.append("已选择目标表: " + bPath.getText() + "\n");
                ProgressDialog progressDialog = createProgressDialog("加载目标表表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try {
                            headers = new com.xzy.CrossSheetFiller().readHeadersOnly(new File(bPath.getText()));
                        } catch (Exception ex) {
                            headers = null;
                        }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        if (headers != null) {
                            bHeadersArea.setText(String.join(", ", headers));
                            bKeyCombo.removeAllItems();
                            bValCombo.removeAllItems();
                            for (String h : headers) {
                                bKeyCombo.addItem(h);
                                bValCombo.addItem(h);
                            }
                            if (bKeyCombo.getItemCount() > 0) bKeyCombo.setSelectedIndex(0);
                            if (bValCombo.getItemCount() > 0) bValCombo.setSelectedIndex(0);
                        } else {
                            bHeadersArea.setText("读取目标表表头失败");
                            bKeyCombo.removeAllItems();
                            bValCombo.removeAllItems();
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
            }
        });
        runBtn.addActionListener(e -> {
            String aPathVal = aPath.getText();
            String bPathVal = bPath.getText();
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fc.setDialogTitle("选择输出文件");
            int result = fc.showSaveDialog(mainPanel);
            if (result != JFileChooser.APPROVE_OPTION) {
                logArea.append("未选择输出文件，操作取消\n");
                return;
            }
            final String outPathVal;
            String outPathRaw = fc.getSelectedFile().getAbsolutePath();
            if (!outPathRaw.toLowerCase().endsWith(".xlsx")) outPathVal = outPathRaw + ".xlsx";
            else outPathVal = outPathRaw;
            logArea.append("已选择输出: " + outPathVal + "\n");
            String aKey = (String) aKeyCombo.getSelectedItem();
            String aVal = (String) aValCombo.getSelectedItem();
            String bKey = (String) bKeyCombo.getSelectedItem();
            String bVal = (String) bValCombo.getSelectedItem();
            if (aPathVal == null || aPathVal.isEmpty() ||
                bPathVal == null || bPathVal.isEmpty() ||
                outPathVal == null || outPathVal.isEmpty() ||
                aKey == null || aVal == null || bKey == null || bVal == null) {
                logArea.append("请先选择所有文件和表头映射项！\n");
                JOptionPane.showMessageDialog(mainPanel, "请先选择所有文件和表头映射项！", "提示", JOptionPane.WARNING_MESSAGE);
                return;
            }
            ProgressDialog progressDialog = createProgressDialog("处理中");
            final JLabel progressLabel = progressDialog.label;
            final JProgressBar progressBar = progressDialog.progressBar;
            SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                @Override
                protected Void doInBackground() {
                    try {
                        new com.xzy.CrossSheetFiller().fill(
                                aPathVal, bPathVal, outPathVal,
                                aKey, aVal, bKey, bVal,
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
                        progressBar.setIndeterminate(false);
                        progressBar.setValue(val);
                        progressLabel.setText("正在处理数据... " + val + "%");
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
                logArea.append("处理完成: " + outPathVal + "\n");
                JOptionPane.showMessageDialog(mainPanel, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                logArea.append("失败: " + ex.getMessage() + "\n");
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }

    // 重复列合并工具 tab
    private JPanel createDuplicateColPanel(JTextArea logArea) {
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
        inputPath.setEditable(false);
        JButton selectInput = new JButton("选择输入文件");
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(new JLabel("输入文件路径:"), gbc);
        gbc.gridx = 1; gbc.gridy = 0; gbc.gridwidth = 2;
        filePanel.add(inputPath, gbc);
        gbc.gridx = 3; gbc.gridy = 0; gbc.gridwidth = 1;
        filePanel.add(selectInput, gbc);

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
        opPanel.add(runBtn, BorderLayout.CENTER);
        mainPanel.add(opPanel);
        // 事件
        selectInput.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fc.showOpenDialog(mainPanel) == JFileChooser.APPROVE_OPTION) {
                inputPath.setText(fc.getSelectedFile().getAbsolutePath());
                logArea.append("已选择输入: " + inputPath.getText() + "\n");
                ProgressDialog progressDialog = createProgressDialog("加载输入文件表头");
                SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                    java.util.List<String> headers = null;
                    @Override
                    protected Void doInBackground() {
                        try {
                            headers = new com.xzy.DuplicateColumnMerger().readHeadersOnly(new File(inputPath.getText()));
                        } catch (Exception ex) {
                            headers = null;
                        }
                        return null;
                    }
                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
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
                progressDialog.dialog.setVisible(true);
            }
        });
        runBtn.addActionListener(e -> {
            String inPath = inputPath.getText();
            // 点击开始合并时弹出保存文件对话框
            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fc.setDialogTitle("选择输出文件");
            int result = fc.showSaveDialog(mainPanel);
            if (result != JFileChooser.APPROVE_OPTION) {
                logArea.append("未选择输出文件，操作取消\n");
                return;
            }
            final String outPath;
            String outPathRaw = fc.getSelectedFile().getAbsolutePath();
            if (!outPathRaw.toLowerCase().endsWith(".xlsx")) outPath = outPathRaw + ".xlsx";
            else outPath = outPathRaw;
            logArea.append("已选择输出: " + outPath + "\n");
            if (inPath == null || inPath.isEmpty()) {
                logArea.append("请先选择输入文件\n");
                return;
            }
            if (customerIdCombo.getSelectedItem() == null || mergeColsList.getSelectedValuesList().isEmpty()) {
                logArea.append("请选择CustomerId列和要合并的列\n");
                return;
            }
            String customerIdCol = (String) customerIdCombo.getSelectedItem();
            java.util.List<String> mergeCols = mergeColsList.getSelectedValuesList();
            ProgressDialog progressDialog = createProgressDialog("处理中");
            final JLabel progressLabel = progressDialog.label;
            final JProgressBar progressBar = progressDialog.progressBar;
            SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                @Override
                protected Void doInBackground() {
                    try {
                        new com.xzy.DuplicateColumnMerger().mergeByCustomerIdWithProgress(inPath, outPath, customerIdCol, mergeCols, (cur, total) -> publish((int) (cur * 100.0 / total)));
                    } catch (Exception ex) {
                        throw new RuntimeException(ex);
                    }
                    return null;
                }
                @Override
                protected void process(java.util.List<Integer> chunks) {
                    if (!chunks.isEmpty()) {
                        int val = chunks.get(chunks.size() - 1);
                        progressBar.setIndeterminate(false);
                        progressBar.setValue(val);
                        progressLabel.setText("正在合并数据... " + val + "%");
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
                logArea.append("处理完成: " + outPath + "\n");
                JOptionPane.showMessageDialog(mainPanel, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                logArea.append("失败: " + ex.getMessage() + "\n");
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new ExcelTools();
        });
    }
}

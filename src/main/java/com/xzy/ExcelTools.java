package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
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
        try (FileInputStream fis = new FileInputStream(file)) {
            OPCPackage pkg = OPCPackage.open(fis);
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetsData();
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();

            if (it.hasNext()) {
                InputStream sheet = it.next();
                try {
                    XMLReader parser = XMLReaderFactory.createXMLReader();
                    try {
                        parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                        parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
                        parser.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
                    } catch (SAXNotRecognizedException | SAXNotSupportedException e) {
                        logger.warn("设置XML解析器特性时出错: " + e.getMessage());
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
            pkg.close();
        } catch (OpenXML4JException e) {
            throw new IOException("打开Excel文件时出错: " + e.getMessage(), e);
        }
        return headers;
    }

    /**
     * 读取XLS文件的表头
     */
    private List<String> readXLSHeaders(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook = WorkbookFactory.create(fis);
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
            workbook.close();
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

            // 创建进度对话框
            JDialog progressDialog = new JDialog(mainFrame, "处理中", true);
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
            progressDialog.setLocationRelativeTo(mainFrame);

            // 在后台线程中处理数据
            SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                @Override
                protected Void doInBackground() throws Exception {
                    try (FileInputStream fis = new FileInputStream(file);
                         Workbook workbook = new XSSFWorkbook(fis);
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
                                // 优先用key查找（即查找字段是key）
                                if (mappingRelations.containsKey(lookupValue)) {
                                    mappedValue = mappingRelations.get(lookupValue);
                                } else {
                                    // 否则遍历value查找key（兼容旧逻辑）
                                    for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                                        if (entry.getValue().equals(lookupValue)) {
                                            mappedValue = entry.getKey();
                                            break;
                                        }
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
                        final int finalUpdatedCount = updatedCount;
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("已更新 " + finalUpdatedCount + " 条记录");
                            log("处理完成，共更新 " + finalUpdatedCount + " 条记录");
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
                logger.error("应用映射关系时出错", e);
                JOptionPane.showMessageDialog(mainFrame,
                        "应用映射关系时出错: " + e.getMessage(),
                        "错误",
                        JOptionPane.ERROR_MESSAGE);
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
        JDialog progressDialog = createProgressDialog("处理中");
        JLabel progressLabel = (JLabel) ((JPanel)progressDialog.getContentPane().getComponent(0)).getComponent(0);
        JProgressBar progressBar = (JProgressBar) ((JPanel)progressDialog.getContentPane().getComponent(0)).getComponent(1);

        // 在后台线程中处理数据
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                try (FileInputStream fis = new FileInputStream(inputFile);
                     Workbook workbook = WorkbookFactory.create(fis);
                     FileInputStream templateFis = new FileInputStream(templateFile);
                     Workbook templateWorkbook = WorkbookFactory.create(templateFis)) {

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

                    // 使用SXSSFWorkbook来处理大文件
                    try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
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
                        int batchSize = BATCH_SIZE; // 每批处理的行数

                        // 显示开始处理的消息
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("正在处理数据...");
                            log("开始处理数据，共 " + totalRows + " 行");
                        });

                        // 批量处理数据
                        for (int i = 1; i <= totalRows; i += batchSize) {
                            int endRow = Math.min(i + batchSize - 1, totalRows);

                            for (int j = i; j <= endRow; j++) {
                                Row sourceRow = sheet.getRow(j);
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
                                                            // 根据格式类型选择分隔符
                                                            mergedValue.append(mapping.getFormatType() == HeaderMapping.FormatType.COMMA_SEPARATED ?
                                                                "," : OTHER_REMARK_SEPARATOR);
                                                        }
                                                        // 根据格式类型决定是否添加源表头名称
                                                        if (mapping.getFormatType() == HeaderMapping.FormatType.PIPE_SEPARATED) {
                                                            mergedValue.append(sourceHeader).append("：");
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
                            }

                            // 更新进度
                            final int progress = (int) ((processedRows * 100.0) / totalRows);
                            SwingUtilities.invokeLater(() -> {
                                progressBar.setIndeterminate(false);
                                progressBar.setValue(progress);
                                progressLabel.setText("正在处理数据... " + progress + "%");
                            });

                            // 清理临时文件
                            ((SXSSFSheet)outputSheet).flushRows();
                        }

                        // 保存文件
                        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                            outputWorkbook.write(fos);
                        }

                        // 显示处理完成的消息
                        final int finalProcessedRows = processedRows;
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("处理完成，共处理 " + finalProcessedRows + " 行数据");
                            log("处理完成，共处理 " + finalProcessedRows + " 行数据");
                        });
                    }
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
            this.formatType = targetHeader.equals("手机号码") ?
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
     * 创建进度对话框
     */
    private JDialog createProgressDialog(String title) {
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

        return dialog;
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

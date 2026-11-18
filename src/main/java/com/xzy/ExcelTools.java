package com.xzy;

import com.xzy.model.AppContext;
import com.xzy.model.HeaderMapping;
import com.xzy.model.ProgressDialog;
import com.xzy.panel.CrossSheetPanel;
import com.xzy.panel.DuplicateColPanel;
import com.xzy.panel.TableMergePanel;
import com.xzy.util.ExcelUtils;
import com.xzy.util.SafeExport;
import com.xzy.util.XlsxSheetReader;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

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
    private final Set<File> mappingSourceFiles = new LinkedHashSet<>();
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
        AppContext ctx = new AppContext(mainFrame, logArea, statusLabel);
        JPanel crossSheetPanel = new CrossSheetPanel(ctx).createPanel();
        tabbedPane.addTab("跨表主键填充工具", crossSheetPanel);
        // 集成重复列工具
        JPanel duplicateColPanel = new DuplicateColPanel(ctx).createPanel();
        tabbedPane.addTab("重复列工具", duplicateColPanel);
        // 集成表格合并工具
        JPanel tableMergePanel = new TableMergePanel(ctx).createPanel();
        tabbedPane.addTab("表格合并", tableMergePanel);
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
            mappingSourceFiles.clear();
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

                    List<String> sourceHeaders = ExcelUtils.readNamedHeaders(new File(inputFilePath));

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

                    List<String> templateHeaders = ExcelUtils.readNamedHeaders(new File(templateFilePath));

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
        try {
            DefaultListModel<String> sourceModel = new DefaultListModel<>();
            DefaultListModel<String> targetModel = new DefaultListModel<>();
            for (String header : ExcelUtils.readNamedHeaders(file)) {
                sourceModel.addElement(header);
                targetModel.addElement(header);
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
        try {
            statusLabel.setText("正在处理映射关系...");
            log("开始处理映射关系: " + file.getName());
            int savedRelations = readMappingRelations(file, sourceField, targetField,
                    mappingRelations, mappingSourceFiles);
            mappingListModel.clear();
            for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                mappingListModel.addElement(entry.getKey() + " -> " + entry.getValue());
            }
            statusLabel.setText("已保存 " + mappingRelations.size() + " 条映射关系");
            log("处理完成，共保存 " + savedRelations + " 条映射关系");
            log("成功保存映射关系: " + sourceField + " -> " + targetField);
        } catch (IOException e) {
            log("保存映射关系时出错: " + e.getMessage());
            JOptionPane.showMessageDialog(mainFrame,
                    "保存映射关系时出错: " + e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
        }
    }

    static int readMappingRelations(File file, String sourceField, String targetField,
                                    Map<String, String> mappingRelations, Set<File> sourceFiles) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 获取表头索引
            Map<String, Integer> headerIndexMap = indexNamedHeaders(ExcelUtils.readHeaderRow(sheet.getRow(0)));

            // 获取源字段和目标字段的索引
            Integer sourceIndex = headerIndexMap.get(sourceField.trim());
            Integer targetIndex = headerIndexMap.get(targetField.trim());

            if (sourceIndex == null || targetIndex == null) {
                throw new IOException("找不到指定的字段");
            }

            // 处理数据行
            final int totalRows = sheet.getLastRowNum();
            int savedRelations = 0;

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell sourceCell = row.getCell(sourceIndex);
                Cell targetCell = row.getCell(targetIndex);

                if (sourceCell != null && targetCell != null) {
                    String sourceValue = ExcelUtils.getCellValueAsString(sourceCell);
                    String targetValue = ExcelUtils.getCellValueAsString(targetCell);

                    if (!sourceValue.isEmpty() && !targetValue.isEmpty()) {
                        mappingRelations.put(sourceValue, targetValue);
                        if (savedRelations == 0) {
                            sourceFiles.add(file);
                        }
                        savedRelations++;
                    }
                }
            }

            return savedRelations;
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
                // 使用SwingWorker + 进度对话框加载表头，避免UI冻结
                ProgressDialog progressDialog = ProgressDialog.create(mainFrame,"加载表头");
                SwingWorker<List<String>, Void> worker = new SwingWorker<List<String>, Void>() {
                    @Override
                    protected List<String> doInBackground() throws Exception {
                        return ExcelUtils.readNamedHeaders(currentTargetFile);
                    }

                    @Override
                    protected void done() {
                        progressDialog.dialog.dispose();
                        try {
                            List<String> headers = get();
                            DefaultListModel<String> lookupModel = new DefaultListModel<>();
                            DefaultListModel<String> fillModel = new DefaultListModel<>();
                            for (String header : headers) {
                                lookupModel.addElement(header);
                                fillModel.addElement(header);
                            }
                            lookupList.setModel(lookupModel);
                            fillList.setModel(fillModel);
                            statusLabel.setText("已加载 " + headers.size() + " 个表头");
                            log("成功加载表头: " + currentTargetFile.getName());
                        } catch (Exception ex) {
                            log("加载表头时出错: " + ex.getMessage());
                            JOptionPane.showMessageDialog(mainFrame,
                                    "加载表头时出错: " + ex.getMessage(),
                                    "错误",
                                    JOptionPane.ERROR_MESSAGE);
                        }
                    }
                };
                worker.execute();
                progressDialog.dialog.setVisible(true);
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

            ProgressDialog progressDialog = ProgressDialog.create(mainFrame,"处理中");

            // 在后台线程中处理数据，只有安全提交完成后才发布成功状态。
            SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                @Override
                protected Void doInBackground() throws Exception {
                    int[] stats = exportAppliedMapping(file, outputFile, lookupField, fillField,
                            mappingRelations, mappingSourceFiles,
                            (percent, message) -> updateExportProgress(progressDialog, percent, message), ExcelTools.this::log);
                    SwingUtilities.invokeLater(() -> {
                        progressDialog.progressBar.setValue(100);
                        statusLabel.setText("已更新 " + stats[1] + " 条记录");
                        log("处理完成，共处理 " + stats[0] + " 行，更新 " + stats[1] + " 条记录");
                    });
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
    private static int[] processApplyMappingXlsx(File inputFile, File outputFile, String lookupField, String fillField,
                                                 Map<String, String> mappingRelations, Map<String, String> reverseMapping,
                                                 BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        List<String> headers = XlsxSheetReader.readHeaders(inputFile);

        // 获取查找字段和填充字段的索引
        Map<String, Integer> indices = indexNamedHeaders(headers);
        Integer lookupIndex = indices.get(lookupField);
        Integer fillIndex = indices.get(fillField);
        if (lookupIndex == null || fillIndex == null) {
            throw new IOException("找不到指定的字段: " + lookupField + " 或 " + fillField);
        }

        progress.accept(-1, "正在应用映射关系...");
        log.accept("开始应用映射关系（SAX流式模式）");

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
            XlsxSheetReader.readDataRows(inputFile, sourceRow -> {
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
                if (lookupIndex < sourceRow.size()) {
                    String lookupValue = sourceRow.get(lookupIndex);
                    if (lookupValue != null && !lookupValue.isEmpty()) {
                        String mappedValue = mappingRelations.get(lookupValue);
                        if (mappedValue == null) {
                            mappedValue = reverseMapping.get(lookupValue);
                        }

                        Cell fillCell = outRow.getCell(fillIndex);
                        if (fillCell == null) {
                            fillCell = outRow.createCell(fillIndex);
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
                    progress.accept(Math.min(99, processed / 1000), "正在处理数据... 已处理 " + processed + " 行");
                }
            });

            // 保存文件
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }

        return stats;
    }

    /**
     * XLS文件：POI读取 + SXSSFWorkbook写入
     */
    private static int[] processApplyMappingXls(File inputFile, File outputFile, String lookupField, String fillField,
                                                Map<String, String> mappingRelations, Map<String, String> reverseMapping,
                                                BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        int[] stats = {0, 0};
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            Map<String, Integer> headerIndexMap = indexNamedHeaders(ExcelUtils.readHeaderRow(headerRow));

            Integer lookupIndex = headerIndexMap.get(lookupField);
            Integer fillIndex = headerIndexMap.get(fillField);
            if (lookupIndex == null || fillIndex == null) {
                throw new IOException("找不到指定的字段");
            }

            progress.accept(-1, "正在应用映射关系...");
            log.accept("开始应用映射关系");

            try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
                Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

                // 复制表头
                Row outputHeaderRow = outputSheet.createRow(0);
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        Cell outputCell = outputHeaderRow.createCell(i);
                        ExcelUtils.copyCellValue(cell, outputCell);
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
                            ExcelUtils.copyCellValue(cell, outputCell);
                        }
                    }

                    Cell lookupCell = row.getCell(lookupIndex);
                    if (lookupCell != null) {
                        String lookupValue = ExcelUtils.getCellValueAsString(lookupCell);

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
                        final int percent = Math.min(99, (int) ((i * 100.0) / totalRows));
                        progress.accept(percent, "正在处理数据... " + percent + "%");
                    }
                }

                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }

                stats[0] = processedRows;
                stats[1] = updatedCount;
            }
        }
        return stats;
    }

    private void processExcel(String inputFilePath, String templateFilePath, String outputFilePath,
                              Map<String, String> mappings) throws IOException {
        // 验证文件
        File inputFile = new File(inputFilePath);
        File templateFile = new File(templateFilePath);
        File outputFile = new File(outputFilePath);

        // 创建进度对话框
        ProgressDialog progressDialog = ProgressDialog.create(mainFrame,"处理中");
        // 在后台线程中处理数据
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                int totalProcessed = exportProcessedExcel(inputFile, templateFile, outputFile, mappings,
                        (percent, message) -> updateExportProgress(progressDialog, percent, message), ExcelTools.this::log);
                SwingUtilities.invokeLater(() -> {
                    progressDialog.progressBar.setValue(100);
                    statusLabel.setText("处理完成，共处理 " + totalProcessed + " 行数据");
                    log("处理完成，共处理 " + totalProcessed + " 行数据");
                });
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

    private void updateExportProgress(ProgressDialog dialog, int percent, String message) {
        SwingUtilities.invokeLater(() -> {
            if (percent >= 0) {
                dialog.progressBar.setIndeterminate(false);
                dialog.progressBar.setValue(percent);
            }
            dialog.label.setText(message);
            statusLabel.setText(message);
        });
    }

    static int[] exportAppliedMapping(File inputFile, File outputFile, String lookupField, String fillField,
                                      Map<String, String> mappingRelations, Collection<File> mappingSources,
                                      BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        List<File> protectedInputs = new ArrayList<>(mappingSources);
        protectedInputs.add(inputFile);
        try (SafeExport export = SafeExport.begin(outputFile, protectedInputs.toArray(new File[0]))) {
            Map<String, String> reverseMapping = new HashMap<>();
            for (Map.Entry<String, String> entry : mappingRelations.entrySet()) {
                reverseMapping.putIfAbsent(entry.getValue(), entry.getKey());
            }
            int[] stats;
            if (inputFile.getName().toLowerCase().endsWith(".xlsx")) {
                stats = processApplyMappingXlsx(inputFile, export.temporaryFile(), lookupField, fillField,
                        mappingRelations, reverseMapping, progress, log);
            } else {
                stats = processApplyMappingXls(inputFile, export.temporaryFile(), lookupField, fillField,
                        mappingRelations, reverseMapping, progress, log);
            }
            export.commit();
            return stats;
        }
    }

    static int exportProcessedExcel(File inputFile, File templateFile, File outputFile,
                                    Map<String, String> mappings, BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        validateFile(inputFile, "源数据");
        validateFile(templateFile, "模板");
        try (SafeExport export = SafeExport.begin(outputFile, inputFile, templateFile)) {
            int processedRows = writeProcessedExcel(inputFile, templateFile, export.temporaryFile(), mappings, progress, log);
            export.commit();
            return processedRows;
        }
    }

    private static Map<String, Integer> indexNamedHeaders(List<String> headers) {
        Map<String, Integer> indices = new HashMap<>();
        for (int column = 0; column < headers.size(); column++) {
            String name = headers.get(column);
            if (!name.trim().isEmpty()) {
                indices.put(name, column);
            }
        }
        return indices;
    }

    private static int writeProcessedExcel(File inputFile, File templateFile, File outputFile,
                                           Map<String, String> mappings, BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        int totalProcessed;
        List<String> templateHeaders = ExcelUtils.readHeadersOnly(templateFile);

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
        Map<String, Integer> templateHeaderIndexMap = indexNamedHeaders(templateHeaders);

        // 使用SXSSFWorkbook流式写入输出
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

            // 写入模板表头
            Row outputHeaderRow = outputSheet.createRow(0);
            for (int i = 0; i < templateHeaders.size(); i++) {
                if (!templateHeaders.get(i).isEmpty()) {
                    outputHeaderRow.createCell(i).setCellValue(templateHeaders.get(i));
                }
            }

            if (inputFile.getName().toLowerCase().endsWith(".xlsx")) {
                // XLSX文件：使用SAX流式读取，避免OOM
                progress.accept(-1, "正在使用SAX流式处理数据...");
                log.accept("开始流式处理XLSX文件");

                List<String> sourceHeaders = XlsxSheetReader.readHeaders(inputFile);

                // 构建源表头索引
                Map<String, Integer> sourceHeaderMap = indexNamedHeaders(sourceHeaders);

                log.accept("源表表头: " + sourceHeaders);

                // 使用SAX流式解析并逐行写入输出
                final int[] processedRows = {0};
                final int[] rowNum = {1};

                XlsxSheetReader.readDataRows(inputFile, sourceRow -> {
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
                        progress.accept(Math.min(99, pct / 100), "正在处理数据... 已处理 " + pct + " 行");
                    }
                });

                // 保存文件
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }

                totalProcessed = processedRows[0];

            } else {
                // XLS文件：使用POI读取（XLS文件通常较小）
                try (FileInputStream fis = new FileInputStream(inputFile);
                     Workbook workbook = WorkbookFactory.create(fis)) {
                    Sheet sheet = workbook.getSheetAt(0);

                    Map<String, Integer> sourceHeaderMap = indexNamedHeaders(ExcelUtils.readHeaderRow(sheet.getRow(0)));

                    int totalRows = sheet.getLastRowNum();
                    int processedRows = 0;
                    int rowNum = 1;

                    progress.accept(-1, "正在处理数据...");
                    log.accept("开始处理数据，共 " + totalRows + " 行");

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
                                                String value = ExcelUtils.getCellValueAsString(sourceCell);
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
                                            String value = ExcelUtils.getCellValueAsString(sourceCell);
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

                        final int percent = Math.min(99, (int) ((processedRows * 100.0) / totalRows));
                        progress.accept(percent, "正在处理数据... " + percent + "%");
                        ((SXSSFSheet) outputSheet).flushRows();
                    }

                    try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                        outputWorkbook.write(fos);
                    }

                    totalProcessed = processedRows;
                }
            }
        }
        return totalProcessed;
    }

    /**
     * 验证文件是否有效
     */
    private static void validateFile(File file, String fileType) throws IOException {
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
    private static String getFileExtension(String fileName) {
        int lastDotIndex = fileName.lastIndexOf('.');
        return lastDotIndex > 0 ? fileName.substring(lastDotIndex + 1) : "";
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

    // --- createCrossSheetPanel 和 createDuplicateColPanel 已移至 panel/ 包 ---

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                logger.error("设置外观失败", e);
            }
            new ExcelTools();
        });
    }
}

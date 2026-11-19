package com.xzy;

import com.xzy.ui.Workbench;
import com.xzy.ui.WorkbookArt;
import static com.formdev.flatlaf.util.UIScale.scale;
import com.xzy.model.AppContext;
import com.xzy.model.HeaderMapping;
import com.xzy.model.MappingRelations;
import com.xzy.model.ProgressDialog;
import com.xzy.panel.CrossSheetPanel;
import com.xzy.panel.DuplicateColPanel;
import com.xzy.panel.TableMergePanel;
import com.xzy.util.ExcelUtils;
import com.xzy.util.FormulaCache;
import com.xzy.util.SafeExport;
import com.xzy.util.XlsxSheetReader;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

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
    private final MappingRelations mappingRelations = new MappingRelations();
    private JList<String> mappingList;            // 用于建立映射关系标签页

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
        processorMappings = new HashMap<>();
        processorMappingOrder = new ArrayList<>();
        mainFrame = new JFrame("Excel 工具集");
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setIconImage(WorkbookArt.windowIcon());
        Rectangle screen = GraphicsEnvironment.getLocalGraphicsEnvironment().getMaximumWindowBounds();
        mainFrame.setSize(Math.min(scale(1080), screen.width), Math.min(scale(700), screen.height));
        mainFrame.setMinimumSize(new Dimension(Math.min(scale(660), screen.width), Math.min(scale(470), screen.height)));
        mainFrame.setLocationRelativeTo(null);

        statusLabel = new JLabel("就绪");
        logArea = new JTextArea(4, 0);
        logArea.setEditable(false);
        AppContext ctx = new AppContext(mainFrame, logArea, statusLabel);
        tabbedPane = new JTabbedPane(JTabbedPane.TOP, JTabbedPane.SCROLL_TAB_LAYOUT);
        tabbedPane.putClientProperty("JTabbedPane.leadingComponent", Workbench.brand());
        tabbedPane.addTab("Excel 处理器", Workbench.page("字段映射",
                "将源表数据按模板的列顺序导出，多个源字段可以合并到同一列。", createProcessorPanel()));
        tabbedPane.addTab("建立映射", Workbench.page("建立映射关系",
                "从关系表中选择一对字段，保存它们在每一行的对应值。", createMappingPanel()));
        tabbedPane.addTab("应用映射", Workbench.page("应用映射",
                "选择目标表中的查找字段和填充字段，使用已保存的对应关系填入数据。", createApplyPanel()));
        tabbedPane.addTab("跨表填充", Workbench.page("跨表主键填充",
                "按主键匹配两张表，将源表数据追加到目标表的指定列。", Workbench.scrollForm(new CrossSheetPanel(ctx).createPanel())));
        tabbedPane.addTab("重复项合并", Workbench.page("重复记录合并",
                "按指定列归组，对选中列的值去重后用 | 拼接。", Workbench.scrollForm(new DuplicateColPanel(ctx).createPanel())));
        tabbedPane.addTab("表格合并", Workbench.page("表格合并",
                "将多个工作表或文件合并为一张表，同时生成合并报告。", Workbench.scrollForm(new TableMergePanel(ctx).createPanel())));

        mainFrame.add(tabbedPane, BorderLayout.CENTER);
        mainFrame.add(createStatusBar(), BorderLayout.SOUTH);
        mainFrame.setVisible(true);
    }

    private JPanel createStatusBar() {
        JPanel footer = new JPanel(new BorderLayout());
        footer.setBorder(BorderFactory.createMatteBorder(1, 0, 0, 0, new java.awt.Color(0xD6E0DA)));
        JPanel bar = new JPanel(new BorderLayout(scale(12), 0));
        bar.setBorder(BorderFactory.createEmptyBorder(scale(3), scale(24), scale(3), scale(16)));
        statusLabel.setForeground(new java.awt.Color(0x426353));
        bar.add(statusLabel);
        JToggleButton toggle = new JToggleButton("展开操作记录");
        toggle.putClientProperty("JButton.buttonType", "borderless");
        bar.add(toggle, BorderLayout.LINE_END);
        JScrollPane history = new JScrollPane(logArea);
        history.setVisible(false);
        toggle.addActionListener(e -> {
            history.setVisible(toggle.isSelected());
            toggle.setText(toggle.isSelected() ? "收起操作记录" : "展开操作记录");
            footer.revalidate();
        });
        footer.add(bar, BorderLayout.NORTH);
        footer.add(history);
        return footer;
    }

    private JPanel createProcessorPanel() {
        CardLayout view = new CardLayout();
        JPanel cards = new JPanel(view);
        cards.setOpaque(false);
        JButton start = Workbench.primaryButton("选择文件");
        cards.add(Workbench.welcome(start), "welcome");

        JPanel editor = new JPanel(new BorderLayout(0, scale(14)));
        editor.setOpaque(false);
        JButton select = new JButton("选择文件");
        select.setToolTipText("选择源表和模板，保留已配置的映射");
        JButton reload = new JButton("重新导入文件");
        reload.setToolTipText("选择源表和模板，并清空已配置的映射");
        JButton save = Workbench.primaryButton("保存文件");
        save.setEnabled(false);
        JPanel toolbar = new JPanel(new BorderLayout());
        toolbar.setOpaque(false);
        toolbar.add(Workbench.actions(select, reload), BorderLayout.LINE_START);
        toolbar.add(save, BorderLayout.LINE_END);
        editor.add(toolbar, BorderLayout.NORTH);

        sourceListModel = new DefaultListModel<>();
        templateListModel = new DefaultListModel<>();
        processorMappingListModel = new DefaultListModel<>();
        sourceList = Workbench.fieldList(sourceListModel, "源表字段", "源表没有可用的字段");
        sourceList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        templateList = Workbench.fieldList(templateListModel, "模板字段", "模板没有可用的字段");
        processorMappingList = Workbench.fieldList(processorMappingListModel, "已配置的映射", "尚未添加映射");

        JLabel sourceFile = new JLabel();
        JLabel templateFile = new JLabel();
        JPanel source = Workbench.listSection("源表字段", sourceList);
        JPanel template = Workbench.listSection("模板字段", templateList);
        JPanel mappings = Workbench.listSection("已配置的映射", processorMappingList);
        source.add(sourceFile, BorderLayout.SOUTH);
        template.add(templateFile, BorderLayout.SOUTH);
        JPanel fields = new JPanel(new GridLayout(1, 3, scale(12), 0));
        fields.setOpaque(false);
        fields.add(source);
        fields.add(template);
        fields.add(mappings);
        editor.add(fields);

        JButton add = Workbench.primaryButton("添加映射");
        JButton remove = new JButton("删除映射");
        add.setEnabled(false);
        remove.setEnabled(false);
        javax.swing.event.ListSelectionListener selectionChanged = e ->
                add.setEnabled(!sourceList.isSelectionEmpty() && !templateList.isSelectionEmpty());
        sourceList.addListSelectionListener(selectionChanged);
        templateList.addListSelectionListener(selectionChanged);
        processorMappingList.addListSelectionListener(e -> remove.setEnabled(!processorMappingList.isSelectionEmpty()));
        JPanel controls = new JPanel(new BorderLayout(scale(8), 0));
        controls.setOpaque(false);
        controls.add(new JLabel("按住 Ctrl 可多选源字段，再选一个模板字段。"));
        controls.add(Workbench.actions(add, remove), BorderLayout.LINE_END);
        editor.add(controls, BorderLayout.SOUTH);
        cards.add(editor, "editor");

        start.addActionListener(e -> select.doClick(0));
        addProcessorPanelEventHandlers(select, save, add, remove, reload, () -> {
            sourceFile.setText(new File(inputFilePath).getName());
            sourceFile.setToolTipText(inputFilePath);
            templateFile.setText(new File(templateFilePath).getName());
            templateFile.setToolTipText(templateFilePath);
            view.show(cards, "editor");
            sourceList.requestFocusInWindow();
            String problem = updateProcessorMappingList(save);
            if (!problem.isEmpty()) {
                JOptionPane.showMessageDialog(mainFrame, problem, "请检查字段映射", JOptionPane.WARNING_MESSAGE);
            }
        });
        return cards;
    }

    private JPanel createMappingPanel() {
        JButton upload = Workbench.primaryButton("选择关系表");
        JButton save = new JButton("保存映射关系");
        JButton reset = new JButton("重置映射关系");
        JList<String> source = Workbench.fieldList(new DefaultListModel<>(), "源字段", "选择关系表后显示字段");
        JList<String> target = Workbench.fieldList(new DefaultListModel<>(), "目标字段", "选择关系表后显示字段");
        mappingList = Workbench.fieldList(new DefaultListModel<>(), "已保存的映射关系", "尚未保存映射");

        JPanel columns = new JPanel(new GridLayout(1, 2, scale(12), 0));
        columns.setOpaque(false);
        columns.add(Workbench.listSection("源字段", source));
        columns.add(Workbench.listSection("目标字段", target));
        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, columns,
                Workbench.listSection("已保存的映射关系", mappingList));
        split.setBorder(BorderFactory.createEmptyBorder());
        split.setBackground(tabbedPane.getBackground());
        split.setResizeWeight(0.65);
        split.setDividerLocation(scale(550));
        split.setOneTouchExpandable(true);
        JPanel panel = new JPanel(new BorderLayout(0, scale(14)));
        panel.setOpaque(false);
        panel.add(Workbench.actions(upload, save, reset), BorderLayout.NORTH);
        panel.add(split);
        addMappingPanelEventHandlers(upload, save, source, target);
        reset.addActionListener(e -> {
            mappingRelations.clear();
            mappingList.setModel(new DefaultListModel<>());
            statusLabel.setText("已重置映射关系");
            log("已重置所有已保存的映射关系");
        });
        return panel;
    }

    private JPanel createApplyPanel() {
        JButton upload = Workbench.primaryButton("选择目标表");
        JButton apply = new JButton("应用映射");
        JList<String> lookup = Workbench.fieldList(new DefaultListModel<>(), "查找字段", "选择目标表后显示字段");
        JList<String> fill = Workbench.fieldList(new DefaultListModel<>(), "填充字段", "选择目标表后显示字段");
        JPanel fields = new JPanel(new GridLayout(1, 2, scale(12), 0));
        fields.setOpaque(false);
        fields.add(Workbench.listSection("查找字段", lookup));
        fields.add(Workbench.listSection("填充字段", fill));
        JPanel panel = new JPanel(new BorderLayout(0, scale(14)));
        panel.setOpaque(false);
        panel.add(Workbench.actions(upload, apply), BorderLayout.NORTH);
        panel.add(fields);
        addApplyPanelEventHandlers(upload, apply, lookup, fill);
        return panel;
    }

    private void log(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private void chooseProcessorFiles(Runnable onFilesLoaded) {
        try {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Excel 文件", "xlsx", "xls"));
            chooser.setDialogTitle("选择源数据文件");
            if (chooser.showOpenDialog(mainFrame) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            File sourceFile = chooser.getSelectedFile();
            chooser.setDialogTitle("选择模板文件");
            if (chooser.showOpenDialog(mainFrame) == JFileChooser.APPROVE_OPTION) {
                loadHeaders(sourceFile, chooser.getSelectedFile(), onFilesLoaded);
            }
        } catch (RuntimeException ex) {
            logger.error("选择文件时发生错误", ex);
            log("选择文件失败：" + ex.getMessage());
            JOptionPane.showMessageDialog(mainFrame, "选择文件失败：" + ex.getMessage(),
                    ERROR_TITLE, JOptionPane.ERROR_MESSAGE);
        }
    }

    private void addProcessorPanelEventHandlers(JButton selectFilesButton, JButton saveButton,
                                                JButton addMappingButton, JButton removeMappingButton,
                                                JButton reloadFilesButton, Runnable onFilesLoaded) {
        selectFilesButton.addActionListener(e -> chooseProcessorFiles(onFilesLoaded));
        reloadFilesButton.addActionListener(e -> chooseProcessorFiles(() -> {
            processorMappings.clear();
            processorMappingOrder.clear();
            statusLabel.setText("已重新导入文件，请重新配置映射关系");
            onFilesLoaded.run();
        }));

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

                log("已添加Excel处理器映射: " + String.join(", ", selectedSources) + " -> " + selectedTemplate);
                statusLabel.setText("已添加映射关系");
                updateProcessorMappingList(saveButton);
            } else {
                JOptionPane.showMessageDialog(mainFrame,
                        "请选择源表头和模板表头", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        });

        // 删除映射按钮事件
        removeMappingButton.addActionListener(e -> {
            int selected = processorMappingList.getSelectedIndex();
            if (selected >= 0) {
                String sourceHeader = processorMappingOrder.get(selected);
                String selectedMapping = processorMappingList.getSelectedValue();

                // 删除映射
                processorMappings.remove(sourceHeader);
                // 从顺序列表中移除
                processorMappingOrder.remove(sourceHeader);

                log("已删除Excel处理器映射: " + selectedMapping);
                statusLabel.setText("已删除映射关系");
                updateProcessorMappingList(saveButton);
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

                // 清空映射关系
                processorMappings.clear();
                processorMappingOrder.clear();
                updateProcessorMappingList(saveButton);

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

    private String updateProcessorMappingList(JButton saveButton) {
        Set<String> sources = new HashSet<>(Collections.list(sourceListModel.elements()));
        Set<String> targets = new HashSet<>(Collections.list(templateListModel.elements()));
        processorMappingListModel.clear();
        for (String source : processorMappingOrder) {
            String target = processorMappings.get(source);
            String text = source + " -> " + target;
            if (!sources.contains(source) || !targets.contains(target)) text += "（字段缺失）";
            processorMappingListModel.addElement(text);
        }
        String problem = missingProcessorFields(processorMappings, sources, targets);
        saveButton.setEnabled(!processorMappings.isEmpty() && problem.isEmpty());
        saveButton.setToolTipText(problem.isEmpty() ? null : problem);
        if (!problem.isEmpty()) {
            statusLabel.setText(problem);
            log(problem);
        }
        return problem;
    }

    private void loadHeaders(File sourceFile, File templateFile, Runnable onFilesLoaded) {
        ProgressDialog progress = ProgressDialog.create(mainFrame, "加载表头");
        progress.label.setText("正在读取源表表头…");
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            private List<String> sourceHeaders;
            private List<String> templateHeaders;

            @Override
            protected Void doInBackground() throws IOException {
                sourceHeaders = ExcelUtils.readNamedHeaders(sourceFile);
                SwingUtilities.invokeLater(() -> progress.label.setText("正在读取模板表头…"));
                templateHeaders = ExcelUtils.readNamedHeaders(templateFile);
                return null;
            }

            @Override
            protected void done() {
                progress.dialog.dispose();
                if (isCancelled()) {
                    return;
                }
                try {
                    get();
                } catch (Exception ex) {
                    Throwable cause = ex.getCause() == null ? ex : ex.getCause();
                    logger.error("加载表头失败", cause);
                    log("加载表头失败：" + cause.getMessage());
                    JOptionPane.showMessageDialog(mainFrame, "加载表头失败：" + cause.getMessage(),
                            ERROR_TITLE, JOptionPane.ERROR_MESSAGE);
                    return;
                }
                inputFilePath = sourceFile.getAbsolutePath();
                templateFilePath = templateFile.getAbsolutePath();
                sourceListModel.clear();
                sourceHeaders.forEach(sourceListModel::addElement);
                templateListModel.clear();
                templateHeaders.forEach(templateListModel::addElement);
                log("已加载源数据文件：" + inputFilePath);
                log("已加载模板文件：" + templateFilePath);
                statusLabel.setText("已加载表头，请配置映射关系");
                onFilesLoaded.run();
            }
        };
        progress.dialog.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent event) {
                worker.cancel(true);
            }
        });
        worker.execute();
        progress.dialog.setVisible(true);
    }

    private void addMappingPanelEventHandlers(JButton uploadButton, JButton saveButton,
                                              JList<String> sourceList, JList<String> targetList) {
        uploadButton.addActionListener(e -> chooseMappingFile(sourceList, targetList, currentMappingFile,
                file -> currentMappingFile = file));

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

    private void saveMapping(File file, String sourceField, String targetField) {
        ProgressDialog progress = ProgressDialog.create(mainFrame, "保存映射关系");
        progress.dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        progress.label.setText("正在读取关系表并检查冲突...");
        SwingWorker<DefaultListModel<String>, Void> worker = new SwingWorker<DefaultListModel<String>, Void>() {
            private int importedRows;
            private int conflictGroups;

            @Override
            protected DefaultListModel<String> doInBackground() throws IOException {
                importedRows = readMappingRelations(file, sourceField, targetField, mappingRelations);
                DefaultListModel<String> rows = new DefaultListModel<>();
                mappingRelations.forEachConflict((key, match) -> {
                    conflictGroups++;
                    String direction = match.isReverse() ? "反向" : "正向";
                    rows.addElement("【" + direction + "冲突】查找值 " + key);
                    match.getCandidates().forEach((candidate, origins) -> {
                        for (MappingRelations.Origin origin : origins) {
                            rows.addElement("    候选 " + candidate + "    " + origin.getFile().getPath()
                                    + "，第 " + origin.getRowNumber() + " 行");
                        }
                    });
                });
                mappingRelations.forEachPair((source, target) -> rows.addElement(source + " → " + target));
                return rows;
            }

            @Override
            protected void done() {
                progress.dialog.dispose();
                try {
                    mappingList.setModel(get());
                    String summary = "已保存 " + mappingRelations.size() + " 条关系，"
                            + (conflictGroups == 0 ? "无冲突" : conflictGroups + " 组冲突，详见列表顶部");
                    statusLabel.setText(summary);
                    log(file.getName() + " 读取 " + importedRows + " 行有效关系。" + summary);
                } catch (Exception e) {
                    String message = "保存映射关系失败：" + (e.getCause() == null ? e : e.getCause()).getMessage();
                    statusLabel.setText(message);
                    log(message);
                    JOptionPane.showMessageDialog(mainFrame, message, "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
        progress.dialog.setVisible(true);
    }

    static int readMappingRelations(File file, String sourceField, String targetField,
                                    MappingRelations mappingRelations) throws IOException {
        MappingRelations imported = new MappingRelations();
        int savedRelations = 0;
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 获取表头索引
            Map<String, Integer> headerIndexMap = ExcelUtils.indexNamedHeaders(
                    ExcelUtils.readHeaderRow(sheet.getRow(0)), file.getName() + " / " + sheet.getSheetName());

            // 获取源字段和目标字段的索引
            Integer sourceIndex = headerIndexMap.get(sourceField.trim());
            Integer targetIndex = headerIndexMap.get(targetField.trim());

            if (sourceIndex == null || targetIndex == null) {
                throw new IOException("找不到指定的字段");
            }

            // 处理数据行
            final int totalRows = sheet.getLastRowNum();
            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell sourceCell = row.getCell(sourceIndex);
                Cell targetCell = row.getCell(targetIndex);

                if (sourceCell != null && targetCell != null) {
                    String sourceValue = ExcelUtils.getCheckedCellText(file, sourceCell);
                    String targetValue = ExcelUtils.getCheckedCellText(file, targetCell);

                    if (!sourceValue.isEmpty() && !targetValue.isEmpty()) {
                        imported.add(sourceValue, targetValue, file, i + 1);
                        savedRelations++;
                    }
                }
            }

        }
        mappingRelations.addAll(imported);
        return savedRelations;
    }

    private void addApplyPanelEventHandlers(JButton uploadButton, JButton applyButton,
                                            JList<String> lookupList, JList<String> fillList) {
        uploadButton.addActionListener(e -> chooseMappingFile(lookupList, fillList, currentTargetFile,
                file -> currentTargetFile = file));

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

    private void chooseMappingFile(JList<String> firstList, JList<String> secondList,
                                   File currentFile, Consumer<File> activate) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        if (chooser.showOpenDialog(mainFrame) != JFileChooser.APPROVE_OPTION) return;
        File selectedFile = chooser.getSelectedFile();
        ProgressDialog progress = ProgressDialog.create(mainFrame, "加载表头");
        SwingWorker<List<String>, Void> worker = new SwingWorker<List<String>, Void>() {
            @Override
            protected List<String> doInBackground() throws IOException {
                return ExcelUtils.readNamedHeaders(selectedFile);
            }

            @Override
            protected void done() {
                progress.dialog.dispose();
                if (isCancelled()) return;
                try {
                    List<String> headers = get();
                    String[] fields = headers.toArray(new String[0]);
                    firstList.setListData(fields);
                    secondList.setListData(fields);
                    activate.accept(selectedFile);
                    statusLabel.setText("当前文件：" + selectedFile.getName() + "（" + fields.length + " 个表头）");
                    log("成功加载表头：" + selectedFile.getName());
                } catch (Exception ex) {
                    Throwable cause = ex.getCause() == null ? ex : ex.getCause();
                    String message = "加载表头失败：" + cause.getMessage();
                    statusLabel.setText(currentFile == null ? "请选择有效的 Excel 文件" : "当前文件：" + currentFile.getName());
                    log(message);
                    showError(message);
                }
            }
        };
        progress.dialog.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent event) {
                worker.cancel(true);
            }
        });
        worker.execute();
        progress.dialog.setVisible(true);
    }

    private void applyMapping(File file, String lookupField, String fillField) {
        if (mappingRelations.isEmpty()) {
            JOptionPane.showMessageDialog(mainFrame,
                    "请先建立映射关系",
                    "提示",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        File outputFile = Workbench.chooseFilledWorkbook(mainFrame, file, "保存应用映射后的工作簿");
        if (outputFile != null) {
            ProgressDialog progressDialog = ProgressDialog.create(mainFrame,"处理中");

            // 在后台线程中处理数据，只有安全提交完成后才发布成功状态。
            SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
                @Override
                protected Void doInBackground() throws Exception {
                    int[] stats = exportAppliedMapping(file, outputFile, lookupField, fillField,
                            mappingRelations,
                            (percent, message) -> updateExportProgress(progressDialog, percent, message), ExcelTools.this::log);
                    SwingUtilities.invokeLater(() -> {
                        progressDialog.progressBar.setValue(100);
                        String summary = "已匹配更新 " + stats[1] + " 行"
                                + (stats[2] == 0 ? "" : "，" + stats[2] + " 行因冲突保留原值，详见输出文件末尾新增的映射冲突表");
                        statusLabel.setText(summary);
                        log("共处理 " + stats[0] + " 行。" + summary);
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

    private static int[] applyToWorkbook(File file, Workbook workbook, String lookupField, String fillField,
                                         MappingRelations relations, BiConsumer<Integer, String> progress,
                                         Consumer<String> log) throws IOException {
        if (workbook.getNumberOfSheets() == 0) {
            throw new IOException("目标文件中没有工作表：" + file.getName());
        }
        Sheet sheet = workbook.getSheetAt(0);
        boolean xlsx = !(workbook instanceof HSSFWorkbook);
        int headerRow = xlsx ? sheet.getFirstRowNum() : 0;
        Map<String, Integer> indices = new HashMap<>();
        Row header = headerRow < 0 ? null : sheet.getRow(headerRow);
        if (header != null) {
            String location = file.getName() + " / " + sheet.getSheetName();
            for (Cell cell : header) {
                String name = ExcelUtils.getCheckedCellText(file, cell);
                if (cell instanceof XSSFCell) name = xlsxLookupText((XSSFCell) cell, name);
                ExcelUtils.putUniqueHeader(indices, name, cell.getColumnIndex(), location);
            }
        }
        Integer lookup = indices.get(lookupField);
        Integer fill = indices.get(fillField);
        if (lookup == null || fill == null) {
            throw new IOException("找不到指定的字段：" + lookupField + " 或 " + fillField);
        }

        progress.accept(-1, "正在应用映射关系...");
        log.accept("正在处理工作表：" + sheet.getSheetName());
        MappingConflictWriter conflicts = new MappingConflictWriter(workbook);
        List<CellRangeAddress> merged = new ArrayList<>();
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.getFirstColumn() <= fill && fill <= region.getLastColumn()) merged.add(region);
        }
        merged.sort(Comparator.comparingInt(CellRangeAddress::getFirstRow));
        int regionIndex = 0;
        boolean filled = false;
        int[] stats = new int[3]; // 处理行、匹配更新行、冲突行
        for (Row row : sheet) {
            int index = row.getRowNum();
            if (index <= headerRow || (xlsx && row.getPhysicalNumberOfCells() == 0)) continue;
            stats[0]++;
            if (stats[0] % BATCH_SIZE == 0) {
                progress.accept(Math.min(99, (int) (index * 100.0 / sheet.getLastRowNum())), "已处理 " + stats[0] + " 行");
            }
            Cell key = row.getCell(lookup);
            if (key == null) continue;
            String lookupKey = ExcelUtils.getCheckedCellText(file, key);
            CellType keyType = key.getCellType() == CellType.FORMULA ? key.getCachedFormulaResultType() : key.getCellType();
            if (keyType == CellType.BLANK || (keyType == CellType.STRING && lookupKey.chars()
                    .allMatch(c -> Character.isWhitespace(c) || Character.isSpaceChar(c)))) continue;
            String fallback = lookupKey;
            if (key instanceof XSSFCell) {
                fallback = xlsxLookupText((XSSFCell) key, lookupKey);
                if (fallback == null || fallback.isEmpty()) continue;
                if (lookupKey.isEmpty()) lookupKey = fallback.trim();
            }
            MappingRelations.Match match = relations.resolve(lookupKey);
            Cell cell = row.getCell(fill);
            if (match.isAmbiguous()) {
                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    freezeMappingFormula(file, cell);
                }
                conflicts.record(index + 1, lookupKey, match);
                stats[2]++;
                continue;
            }
            while (regionIndex < merged.size() && merged.get(regionIndex).getLastRow() < index) regionIndex++;
            if (regionIndex < merged.size()) {
                CellRangeAddress region = merged.get(regionIndex);
                if (index >= region.getFirstRow() && (index != region.getFirstRow() || fill != region.getFirstColumn())) {
                    throw new IOException("工作表“" + sheet.getSheetName() + "”第 " + (index + 1)
                            + " 行的填充格位于合并区域 " + region.formatAsString() + " 内，只有左上角可以写入。");
                }
            }
            if (cell == null) {
                cell = row.createCell(fill);
            } else {
                // 清除旧公式或 inlineStr 内容，保留该格的样式后再写入文本。
                cell.setBlank();
            }
            String value = match.getValue();
            cell.setCellValue(value == null ? fallback : value);
            filled = true;
            if (value != null) stats[1]++;
        }
        if (filled) {
            ExcelUtils.markForRecalculation(workbook);
        }
        return stats;
    }

    /** .xlsx 未命中时回填读取到的原文本，例如数字 7.0；查找时使用与关系表相同的键 7。 */
    private static String xlsxLookupText(XSSFCell cell, String key) {
        switch (cell.getCellType()) {
            case NUMERIC:
                return DateUtil.isCellDateFormatted(cell)
                        ? ExcelUtils.formatExcelDate(cell.getNumericCellValue(), cell.getSheet().getWorkbook().isDate1904())
                        : cell.getRawValue();
            case BOOLEAN: return cell.getRawValue();
            case STRING: return "inlineStr".equals(cell.getCTCell().getT().toString()) ? key : cell.getStringCellValue();
            case ERROR: return FormulaError.forInt(cell.getErrorCellValue()).getString();
            case FORMULA:
                return cell.getCachedFormulaResultType() == CellType.ERROR
                        ? FormulaError.forInt(cell.getErrorCellValue()).getString() : key;
            default: return key;
        }
    }

    private static void freezeMappingFormula(File file, Cell cell) throws IOException {
        CellValue cached;
        try {
            cached = FormulaCache.read(cell);
        } catch (FormulaCache.FormulaCacheException e) {
            throw e.inFile(file);
        }
        cell.setBlank();
        switch (cached.getCellType()) {
            case NUMERIC: cell.setCellValue(cached.getNumberValue()); break;
            case STRING: cell.setCellValue(cached.getStringValue()); break;
            case BOOLEAN: cell.setCellValue(cached.getBooleanValue()); break;
            case ERROR: cell.setCellErrorValue(cached.getErrorValue()); break;
            default: throw new IOException("无法读取公式结果：" + cell.getAddress());
        }
    }

    private static final class MappingConflictWriter {
        private final Workbook workbook;
        private Sheet sheet;
        private int sheetNumber;
        private int rowNumber = 1;

        private MappingConflictWriter(Workbook workbook) {
            this.workbook = workbook;
        }

        private void record(int targetRowNumber, String lookupValue, MappingRelations.Match match) {
            for (Map.Entry<String, List<MappingRelations.Origin>> candidate : match.getCandidates().entrySet()) {
                for (MappingRelations.Origin origin : candidate.getValue()) {
                    if (sheet == null || rowNumber > workbook.getSpreadsheetVersion().getLastRowIndex()) {
                        newSheet();
                    }
                    Row row = sheet.createRow(rowNumber++);
                    row.createCell(0).setCellValue(targetRowNumber);
                    row.createCell(1).setCellValue(lookupValue);
                    row.createCell(2).setCellValue(match.isReverse() ? "反向" : "正向");
                    row.createCell(3).setCellValue(candidate.getKey());
                    row.createCell(4).setCellValue(origin.getFile().getAbsolutePath());
                    row.createCell(5).setCellValue(origin.getRowNumber());
                }
            }
        }

        private void newSheet() {
            String name;
            do {
                sheetNumber++;
                name = sheetNumber == 1 ? "映射冲突" : "映射冲突" + sheetNumber;
            } while (workbook.getSheet(name) != null);
            sheet = workbook.createSheet(name);
            rowNumber = 1;
            Row header = sheet.createRow(0);
            String[] columns = {"目标行号", "查找值", "匹配方向", "候选值", "关系文件", "关系行号"};
            for (int i = 0; i < columns.length; i++) {
                header.createCell(i).setCellValue(columns[i]);
            }
        }
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
                                      MappingRelations mappingRelations,
                                      BiConsumer<Integer, String> progress, Consumer<String> log) throws Exception {
        List<File> protectedInputs = new ArrayList<>(mappingRelations.getSourceFiles());
        protectedInputs.add(inputFile);
        try (SafeExport export = SafeExport.begin(outputFile, protectedInputs.toArray(new File[0]))) {
            int[] stats;
            try (FileInputStream input = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(input)) {
                boolean xlsOutput = outputFile.getName().toLowerCase(Locale.ROOT).endsWith(".xls");
                if (xlsOutput && !(workbook instanceof HSSFWorkbook)) {
                    throw new IOException("请将此工作簿保存为 .xlsx，以保留原格式。");
                }
                stats = applyToWorkbook(inputFile, workbook, lookupField, fillField, mappingRelations, progress, log);
                try (FileOutputStream stream = new FileOutputStream(export.temporaryFile())) {
                    ExcelUtils.writeFilledWorkbook(workbook, outputFile, stream);
                }
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

    private static String missingProcessorFields(Map<String, String> mappings, Set<String> sources, Set<String> targets) {
        Set<String> missingSources = new LinkedHashSet<>(mappings.keySet());
        missingSources.removeAll(sources);
        Set<String> missingTargets = new LinkedHashSet<>(mappings.values());
        missingTargets.removeAll(targets);
        List<String> problems = new ArrayList<>();
        if (!missingSources.isEmpty()) problems.add("源表缺少字段：" + String.join("、", missingSources));
        if (!missingTargets.isEmpty()) problems.add("模板缺少字段：" + String.join("、", missingTargets));
        return problems.isEmpty() ? "" : String.join("。", problems) + "。请调整映射或重新选择文件。";
    }

    private static Map<String, Integer> indexProcessorSource(List<String> headers, String location,
                                                             Map<String, String> mappings, Set<String> targets) throws IOException {
        Map<String, Integer> sources = ExcelUtils.indexNamedHeaders(headers, location);
        String problem = missingProcessorFields(mappings, sources.keySet(), targets);
        if (!problem.isEmpty()) throw new IOException(problem);
        return sources;
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
        Map<String, Integer> templateHeaderIndexMap = ExcelUtils.indexNamedHeaders(
                templateHeaders, templateFile.getName() + " / 首个工作表");

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
                Map<String, Integer> sourceHeaderMap = indexProcessorSource(sourceHeaders,
                        inputFile.getName() + " / 首个工作表", mappings, templateHeaderIndexMap.keySet());

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

                    Map<String, Integer> sourceHeaderMap = indexProcessorSource(ExcelUtils.readHeaderRow(sheet.getRow(0)),
                            inputFile.getName() + " / " + sheet.getSheetName(), mappings, templateHeaderIndexMap.keySet());

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
            Workbench.install();
            new ExcelTools();
        });
    }
}

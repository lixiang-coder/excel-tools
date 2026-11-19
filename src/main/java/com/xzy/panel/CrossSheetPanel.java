package com.xzy.panel;

import com.xzy.CrossSheetFiller;
import com.xzy.model.AppContext;
import com.xzy.model.ProgressDialog;
import com.xzy.ui.Workbench;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.function.Consumer;
import java.util.function.Predicate;

/**
 * Tab4: 跨表主键填充工具面板
 */
public class CrossSheetPanel {
    private final AppContext ctx;

    public CrossSheetPanel(AppContext ctx) {
        this.ctx = ctx;
    }

    public JPanel createPanel() {
        JTextField aPath = new JTextField();
        JTextField bPath = new JTextField();
        JButton selA = new JButton("选择源表");
        JButton selB = new JButton("选择目标表");
        JTextArea aHeadersArea = new JTextArea(2, 16);
        JTextArea bHeadersArea = new JTextArea(2, 16);
        JComboBox<String> aKeyCombo = new JComboBox<>();
        JComboBox<String> aValCombo = new JComboBox<>();
        JComboBox<String> bKeyCombo = new JComboBox<>();
        JComboBox<String> bValCombo = new JComboBox<>();
        JButton runBtn = Workbench.primaryButton("开始填充");

        JPanel mainPanel = Workbench.formPage(runBtn,
                workbookSection("源工作表", aPath, selA, aHeadersArea, aKeyCombo, "读取数据的列", aValCombo),
                workbookSection("目标工作表", bPath, selB, bHeadersArea, bKeyCombo, "填入数据的列", bValCombo));

        // 事件
        WorkbookFields source = new WorkbookFields(aPath, aHeadersArea, aKeyCombo, aValCombo, false);
        WorkbookFields target = new WorkbookFields(bPath, bHeadersArea, bKeyCombo, bValCombo, true);
        selA.addActionListener(e -> source.choose(ctx.getMainFrame(), ctx::log));
        selB.addActionListener(e -> target.choose(ctx.getMainFrame(), ctx::log));
        runBtn.addActionListener(e -> {
            String aPathVal = aPath.getText();
            String bPathVal = bPath.getText();
            String aKey = (String) aKeyCombo.getSelectedItem();
            String aVal = (String) aValCombo.getSelectedItem();
            String bKey = (String) bKeyCombo.getSelectedItem();
            String bVal = (String) bValCombo.getSelectedItem();
            if (aPathVal == null || aPathVal.isEmpty() || bPathVal == null || bPathVal.isEmpty() ||
                aKey == null || aVal == null || bKey == null) {
                ctx.log("请先选择所有文件和表头映射项！");
                JOptionPane.showMessageDialog(mainPanel, "请先选择所有文件和表头映射项！", "提示", JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (!confirmTargetColumn(mainPanel, bKey, bVal)) return;
            File output = Workbench.chooseFilledWorkbook(mainPanel, new File(bPathVal), "保存跨表填充后的工作簿");
            if (output == null) {
                ctx.log("未选择输出文件，操作取消");
                return;
            }
            String outPathVal = output.getAbsolutePath();
            ctx.log("已选择输出: " + outPathVal);
            ProgressDialog progressDialog = ProgressDialog.create(ctx.getMainFrame(), "处理中");
            SwingWorker<Void, Integer> worker = new SwingWorker<Void, Integer>() {
                @Override
                protected Void doInBackground() {
                    try {
                        new CrossSheetFiller().fill(aPathVal, bPathVal, outPathVal, aKey, aVal, bKey, bVal,
                                (cur, total) -> publish((int) (cur * 100.0 / total)));
                    } catch (Exception ex) { throw new RuntimeException(ex); }
                    return null;
                }
                @Override
                protected void process(java.util.List<Integer> chunks) {
                    if (!chunks.isEmpty()) {
                        int val = chunks.get(chunks.size() - 1);
                        progressDialog.progressBar.setIndeterminate(false);
                        progressDialog.progressBar.setValue(val);
                        progressDialog.label.setText("正在处理数据... " + val + "%");
                    }
                }
                @Override
                protected void done() { progressDialog.dialog.dispose(); }
            };
            worker.execute();
            progressDialog.dialog.setVisible(true);
            try {
                worker.get();
                ctx.log("处理完成: " + outPathVal);
                JOptionPane.showMessageDialog(mainPanel, "处理完成!", "成功", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ctx.log("失败: " + ex.getMessage());
                JOptionPane.showMessageDialog(mainPanel, "失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
            }
        });
        return mainPanel;
    }

    /** 主窗口和独立窗口使用同一套表头加载及列选择规则。 */
    public static final class WorkbookFields {
        private final JTextField path;
        private final JTextArea preview;
        private final JComboBox<String> key;
        private final JComboBox<String> value;
        private final boolean target;
        private SwingWorker<List<String>, Void> currentLoad;

        public WorkbookFields(JTextField path, JTextArea preview, JComboBox<String> key, JComboBox<String> value,
                              boolean target) {
            this.path = path;
            this.preview = preview;
            this.key = key;
            this.value = value;
            this.target = target;
        }

        public boolean choose(JFrame owner, Consumer<String> log) {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (chooser.showOpenDialog(owner) != JFileChooser.APPROVE_OPTION) return false;
            ProgressDialog progress = ProgressDialog.create(owner, target ? "加载目标表表头" : "加载源表表头");
            load(chooser.getSelectedFile(), log, progress.dialog::dispose);
            progress.dialog.setVisible(true);
            return true;
        }

        void load(File file, Consumer<String> log, Runnable doneLoading) {
            key.removeAllItems();
            value.removeAllItems();
            path.setText(file.getAbsolutePath());
            preview.setText("正在读取表头…");
            String role = target ? "目标表" : "源表";
            log.accept("已选择" + role + "：" + path.getText());
            currentLoad = new SwingWorker<List<String>, Void>() {
                @Override
                protected List<String> doInBackground() throws IOException {
                    return CrossSheetFiller.readMatchingHeaders(file);
                }

                @Override
                protected void done() {
                    try {
                        // 进度窗可以提前关闭；新选择的文件可能已经开始加载。
                        if (currentLoad != this) return;
                        String[] names = get().toArray(new String[0]);
                        key.setModel(new DefaultComboBoxModel<>(names));
                        DefaultComboBoxModel<String> values = new DefaultComboBoxModel<>(names);
                        if (target) values.setSelectedItem(null);
                        value.setModel(values);
                        preview.setText(String.join(", ", names));
                    } catch (InterruptedException ex) {
                        Thread.currentThread().interrupt();
                        preview.setText("读取" + role + "表头已中断");
                        log.accept(preview.getText());
                    } catch (ExecutionException ex) {
                        preview.setText("读取" + role + "表头失败：" + ex.getCause().getMessage());
                        log.accept(preview.getText());
                    } finally {
                        if (currentLoad == this) currentLoad = null;
                        doneLoading.run();
                    }
                }
            };
            currentLoad.execute();
        }
    }

    public static boolean confirmTargetColumn(Component owner, String key, String value) {
        return allowTargetColumn(key, value,
                message -> JOptionPane.showMessageDialog(owner, message, "请选择填入列", JOptionPane.WARNING_MESSAGE),
                message -> {
                    String[] options = {"取消", "确认写入主键列"};
                    return JOptionPane.showOptionDialog(owner, message, "确认修改目标主键",
                            JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null, options, options[0]) == 1;
                });
    }

    static boolean allowTargetColumn(String key, String value, Consumer<String> warn, Predicate<String> confirm) {
        if (value == null) {
            warn.accept("请选择目标表中要填入数据的列。");
            return false;
        }
        return !value.equals(key) || confirm.test("填入列和目标主键列都是“" + key
                + "”。继续将把数据追加到匹配行的主键，导出后的编号会改变。是否继续？");
    }

    private static JPanel workbookSection(String title, JTextField path, JButton select, JTextArea headers,
                                          JComboBox<String> key, String valueLabel, JComboBox<String> value) {
        path.setEditable(false);
        path.getAccessibleContext().setAccessibleName(title + "路径");
        JPanel fileRow = new JPanel(new BorderLayout(12, 0));
        fileRow.add(path, BorderLayout.CENTER);
        fileRow.add(select, BorderLayout.LINE_END);

        headers.setEditable(false);
        headers.setLineWrap(true);
        headers.setWrapStyleWord(true);
        JLabel headersLabel = new JLabel("表头预览");
        headersLabel.setLabelFor(headers);
        JPanel preview = new JPanel(new BorderLayout(12, 0));
        preview.add(headersLabel, BorderLayout.LINE_START);
        preview.add(new JScrollPane(headers), BorderLayout.CENTER);

        key.setPrototypeDisplayValue("用于匹配的表头");
        value.setPrototypeDisplayValue("待填充的表头");
        JPanel columns = new JPanel(new GridLayout(1, 2, 16, 0));
        columns.add(columnSelector("主键列", key));
        columns.add(columnSelector(valueLabel, value));

        JPanel content = new JPanel(new BorderLayout(0, 10));
        content.add(fileRow, BorderLayout.NORTH);
        content.add(preview, BorderLayout.CENTER);
        content.add(columns, BorderLayout.SOUTH);
        return Workbench.formSection(title, content);
    }

    private static JPanel columnSelector(String name, JComboBox<String> combo) {
        JLabel label = new JLabel(name);
        label.setLabelFor(combo);
        JPanel field = new JPanel(new BorderLayout(8, 0));
        field.add(label, BorderLayout.LINE_START);
        field.add(combo, BorderLayout.CENTER);
        return field;
    }
}

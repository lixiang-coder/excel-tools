package com.xzy.ui;

import com.formdev.flatlaf.FlatLaf;
import com.formdev.flatlaf.FlatLightLaf;
import com.formdev.flatlaf.ui.FlatEmptyBorder;
import com.formdev.flatlaf.ui.FlatLineBorder;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.plaf.FontUIResource;
import java.awt.*;
import java.io.File;
import java.util.Locale;

import static com.formdev.flatlaf.util.UIScale.scale;

public final class Workbench {
    private static final Color CANVAS = new Color(0xF4F7F5);
    private static final Color LINE = new Color(0xD6E0DA);
    private static final Color MUTED = new Color(0x53695F);

    private Workbench() {
    }

    public static File chooseFilledWorkbook(Component parent, File target, String title) {
        File directory = target.getAbsoluteFile().getParentFile();
        String name = target.getName();
        boolean legacy = name.toLowerCase(Locale.ROOT).endsWith(".xls");
        String extension = legacy ? ".xls" : ".xlsx";
        int dot = name.lastIndexOf('.');
        JFileChooser chooser = new JFileChooser(directory);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setFileFilter(legacy ? new FileNameExtensionFilter("Excel 工作簿 (*.xls, *.xlsx)", "xls", "xlsx")
                : new FileNameExtensionFilter("Excel 工作簿 (*.xlsx)", "xlsx"));
        chooser.setSelectedFile(new File(directory, (dot > 0 ? name.substring(0, dot) : name) + "-已填充" + extension));
        chooser.setDialogTitle(title);
        if (chooser.showSaveDialog(parent) != JFileChooser.APPROVE_OPTION) return null;
        File chosen = chooser.getSelectedFile();
        String suffix = chosen.getName().toLowerCase(Locale.ROOT);
        return suffix.endsWith(".xls") || suffix.endsWith(".xlsx") ? chosen : new File(chosen.getPath() + extension);
    }

    public static void install() {
        FlatLaf.registerCustomDefaultsSource("com.xzy.ui");
        FlatLightLaf.setup();
        Font system = UIManager.getFont("Label.font");
        Font chinese = new Font("Microsoft YaHei UI", Font.PLAIN, Math.max(14, system.getSize()));
        if (chinese.getFamily().equalsIgnoreCase("Microsoft YaHei UI")) {
            UIManager.put("defaultFont", new FontUIResource(chinese));
        }
    }

    public static JButton primaryButton(String text) {
        JButton button = new JButton(text);
        button.putClientProperty("FlatLaf.styleClass", "primary");
        return button;
    }

    private static JLabel heading(String text, float factor) {
        JLabel label = new JLabel(text);
        Font font = label.getFont();
        label.setFont(font.deriveFont(Font.BOLD, font.getSize2D() * factor));
        return label;
    }

    public static JLabel brand() {
        JLabel label = heading("Excel 工具集", 1.22f);
        label.setIcon(WorkbookArt.mark());
        label.setIconTextGap(scale(10));
        label.setBorder(new FlatEmptyBorder(0, 22, 0, 28));
        return label;
    }

    public static JPanel page(String title, String hint, JComponent body) {
        JPanel header = new JPanel(new BorderLayout(0, scale(5)));
        header.setOpaque(false);
        header.add(heading(title, 1.65f), BorderLayout.NORTH);
        JLabel description = new JLabel(hint);
        description.setForeground(MUTED);
        header.add(description);
        JPanel page = new JPanel(new BorderLayout(0, scale(18)));
        page.setBorder(new FlatEmptyBorder(22, 24, 20, 24));
        page.setBackground(CANVAS);
        page.add(header, BorderLayout.NORTH);
        page.add(body);
        return page;
    }

    private static JPanel card(LayoutManager layout) {
        JPanel card = new JPanel(layout);
        card.putClientProperty("FlatLaf.style", "arc: 12");
        card.setBorder(new FlatLineBorder(new Insets(14, 16, 14, 16), LINE, 1, 12));
        return card;
    }

    public static JPanel formSection(String title, JComponent content) {
        JPanel section = card(new BorderLayout(0, scale(10)));
        JLabel label = heading(title, 1.08f);
        label.setLabelFor(content);
        section.add(label, BorderLayout.NORTH);
        section.add(content);
        return section;
    }

    public static JPanel formPage(JButton action, JPanel... sections) {
        JPanel page = new JPanel(new GridBagLayout());
        page.setBackground(CANVAS);
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0;
        c.weightx = 1;
        c.fill = GridBagConstraints.HORIZONTAL;
        c.insets = new Insets(0, 0, scale(12), 0);
        for (JPanel section : sections) {
            page.add(section, c);
        }
        c.fill = GridBagConstraints.NONE;
        c.anchor = GridBagConstraints.FIRST_LINE_START;
        c.weighty = 1;
        page.add(action, c);
        return page;
    }

    public static JScrollPane scrollForm(JPanel form) {
        JScrollPane scroll = new JScrollPane(form);
        scroll.setBorder(BorderFactory.createEmptyBorder());
        scroll.getVerticalScrollBar().setUnitIncrement(scale(18));
        return scroll;
    }

    public static JPanel actions(JButton... buttons) {
        JPanel bar = new JPanel();
        bar.setOpaque(false);
        bar.setLayout(new BoxLayout(bar, BoxLayout.LINE_AXIS));
        for (JButton button : buttons) {
            if (bar.getComponentCount() > 0) {
                bar.add(Box.createHorizontalStrut(scale(8)));
            }
            bar.add(button);
        }
        return bar;
    }

    public static JPanel welcome(JButton select) {
        JPanel card = card(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0;
        c.insets = new Insets(0, 0, scale(14), 0);
        card.add(new JLabel(WorkbookArt.workflow()), c);
        c.insets.bottom = scale(6);
        card.add(heading("选择源表和模板", 1.4f), c);
        JLabel hint = new JLabel("先选择包含数据的源表，再选择决定输出列的模板。");
        hint.setForeground(MUTED);
        c.insets.bottom = scale(19);
        card.add(hint, c);
        c.insets.bottom = scale(10);
        card.add(select, c);
        JLabel formats = new JLabel("支持 .xlsx 和 .xls");
        formats.setForeground(MUTED);
        c.insets.bottom = 0;
        card.add(formats, c);
        return card;
    }

    public static JList<String> fieldList(DefaultListModel<String> model, String title, String emptyHint) {
        JList<String> list = new JList<String>(model) {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                if (getModel().getSize() == 0) {
                    FontMetrics metrics = g.getFontMetrics();
                    g.setColor(MUTED);
                    g.drawString(emptyHint, Math.max(scale(4), (getWidth() - metrics.stringWidth(emptyHint)) / 2),
                            (getHeight() + metrics.getAscent() - metrics.getDescent()) / 2);
                }
            }
        };
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        list.getAccessibleContext().setAccessibleName(title);
        list.getAccessibleContext().setAccessibleDescription(emptyHint);
        return list;
    }

    public static JPanel listSection(String title, JList<String> list) {
        JScrollPane scroll = new JScrollPane(list);
        scroll.setBorder(BorderFactory.createEmptyBorder());
        return formSection(title, scroll);
    }
}

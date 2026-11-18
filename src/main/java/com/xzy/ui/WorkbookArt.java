package com.xzy.ui;

import com.formdev.flatlaf.icons.FlatAbstractIcon;

import javax.swing.Icon;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Component;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;

public final class WorkbookArt {
    private static final Color GREEN = new Color(0x176B50);
    private static final Color MINT = new Color(0xCEE7DA);

    private WorkbookArt() {
    }

    public static Image windowIcon() {
        Icon icon = mark();
        BufferedImage image = new BufferedImage(icon.getIconWidth(), icon.getIconHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = image.createGraphics();
        icon.paintIcon(null, graphics, 0, 0);
        graphics.dispose();
        return image;
    }

    public static Icon mark() {
        return new FlatAbstractIcon(30, 30, GREEN) {
            @Override
            protected void paintIcon(Component c, Graphics2D g) {
                g.fillRoundRect(0, 0, 30, 30, 8, 8);
                g.setColor(Color.WHITE);
                g.setStroke(new BasicStroke(1.6f));
                g.drawRoundRect(7, 7, 16, 16, 3, 3);
                g.drawLine(7, 12, 23, 12);
                g.drawLine(7, 17, 23, 17);
                g.drawLine(13, 12, 13, 23);
            }
        };
    }

    public static Icon workflow() {
        return new FlatAbstractIcon(210, 110, new Color(0xEFF7F2)) {
            @Override
            protected void paintIcon(Component c, Graphics2D g) {
                g.fillOval(9, 14, 192, 82);
                Graphics2D source = (Graphics2D) g.create();
                source.rotate(Math.toRadians(-8), 54, 56);
                paintSheet(source, 25, 18, new Color(0x679985));
                source.dispose();
                paintSheet(g, 126, 12, GREEN);

                g.setColor(GREEN);
                g.setStroke(new BasicStroke(2.4f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g.drawLine(94, 55, 115, 55);
                g.drawLine(109, 49, 115, 55);
                g.drawLine(109, 61, 115, 55);
            }
        };
    }

    private static void paintSheet(Graphics2D g, int x, int y, Color ink) {
        g.setColor(MINT);
        g.fillRoundRect(x + 3, y + 4, 59, 76, 10, 10);
        g.setColor(Color.WHITE);
        g.fillRoundRect(x, y, 59, 76, 10, 10);
        g.setColor(ink);
        g.setStroke(new BasicStroke(1.5f));
        g.drawRoundRect(x, y, 59, 76, 10, 10);
        g.fillRoundRect(x + 9, y + 11, 29, 4, 3, 3);
        g.setColor(MINT);
        g.fillRoundRect(x + 9, y + 19, 17, 3, 2, 2);
        g.fillRect(x + 9, y + 30, 41, 10);
        g.setColor(ink);
        g.setStroke(new BasicStroke(1));
        g.drawRect(x + 9, y + 30, 41, 33);
        g.drawLine(x + 24, y + 30, x + 24, y + 63);
        g.drawLine(x + 9, y + 40, x + 50, y + 40);
        g.drawLine(x + 9, y + 52, x + 50, y + 52);
    }
}

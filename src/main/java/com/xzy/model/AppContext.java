package com.xzy.model;

import javax.swing.*;

/**
 * 应用上下文，持有各Tab共享的组件引用
 */
public class AppContext {
    private final JFrame mainFrame;
    private final JTextArea logArea;
    private final JLabel statusLabel;

    public AppContext(JFrame mainFrame, JTextArea logArea, JLabel statusLabel) {
        this.mainFrame = mainFrame;
        this.logArea = logArea;
        this.statusLabel = statusLabel;
    }

    public JFrame getMainFrame() {
        return mainFrame;
    }

    public JTextArea getLogArea() {
        return logArea;
    }

    public JLabel getStatusLabel() {
        return statusLabel;
    }

    public void log(String message) {
        SwingUtilities.invokeLater(() -> logArea.append(message + "\n"));
    }
}

package org.ExcellSplitter.GUI;

import javax.swing.*;

public class ESFrame extends JFrame {
    private JPanel panel;
    private JButton start, choose, direct;
    private JLabel columnL, typeL;
    private JTextField columnT;
    private JRadioButton xls, csv;
    private JFileChooser fileChooser;

    public ESFrame () {
        super("Excel Splitter");
        initiate();
    }

    private void initiate() {
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        panel = new JPanel();
        this.getContentPane().add(panel);
        markUpPanel();
        this.pack();
        this.setVisible(true);
    }

    private void markUpPanel() {

    }
}

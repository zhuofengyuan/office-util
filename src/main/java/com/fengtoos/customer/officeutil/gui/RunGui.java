package com.fengtoos.customer.officeutil.gui;

import javax.swing.*;

public class RunGui {

    public static void main(String[] args) {

        //更改默认的选择文件的外观为windows的风格
        if (UIManager.getLookAndFeel().isSupportedLookAndFeel()) {
            final String platform = UIManager.getSystemLookAndFeelClassName();
            if (!UIManager.getLookAndFeel().getName().equals(platform)) {
                try {
                    UIManager.setLookAndFeel(platform);
                } catch (Exception exception) {
                    exception.printStackTrace();
                }
            }
        }
        BuildGui gui = new BuildGui();
    }
}

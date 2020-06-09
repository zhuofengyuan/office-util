package com.fengtoos.customer.officeutil.gui;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.fengtoos.customer.officeutil.doc.NumberDocument;
import com.fengtoos.customer.officeutil.resp.Result;
import com.fengtoos.customer.officeutil.util.ExcelUtil;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;

import static javax.swing.JFileChooser.DIRECTORIES_ONLY;


public class BuildGui {//实现监听器的接口
    private JFrame frame;
    private JPanel p0, p1, p2, p3, p4;

    private JTextField dataName;
    private JTextField wordOutName;
    private JTextField splitString;
    private JButton dataChoose;
    private JButton wordChoose;
    private JButton register;
    private JButton word2pdf;
    private JFileChooser dataChooser; //数据导入选择器
    private JFileChooser wordOutChooser; //word路径选择器
    private JFileChooser imgChooser; //图片路径选择器

    {
        dataChooser = new JFileChooser();
        FileFilter filter = new FileNameExtensionFilter("Excel表格（xlsx）", "XLSX");
        dataChooser.setFileFilter(filter);

        wordOutChooser = new JFileChooser();
        wordOutChooser.setFileSelectionMode(DIRECTORIES_ONLY);

        imgChooser = new JFileChooser();
        imgChooser.setFileSelectionMode(DIRECTORIES_ONLY);
    }

    public BuildGui() {
        frame = new JFrame("Excel拆分工具");
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);//设置窗口的点击右上角的x的处理方式，这里设置的是退出程序
        p0 = new JPanel();
        p0.setPreferredSize(new Dimension(550, 40));
        p0.add(new JLabel("Excel拆分工具"));
        frame.add(p0);

        //------------------------数据导入路径---------------------
        p1 = new JPanel();
        p1.setPreferredSize(new Dimension(550, 40));
        p1.add(new JLabel("\t数据路径:"));
        dataName = new JTextField(20);
        dataName.setEnabled(false);
        dataName.setPreferredSize(new Dimension(dataName.getWidth(), dataName.getHeight() + 23));
        p1.add(dataName);
        dataChoose = new JButton("选择文件");
        dataChoose.addActionListener(e -> {

            int i = dataChooser.showOpenDialog(frame.getContentPane());// 显示文件选择对话框

            // 判断用户单击的是否为“打开”按钮
            if (i == JFileChooser.APPROVE_OPTION) {

                File selectedFile = dataChooser.getSelectedFile();// 获得选中的文件对象
                dataName.setText(selectedFile.getPath());// 显示选中文件的名称
            }
        });
        p1.add(dataChoose);

        //------------------------数据输出目录---------------------
        p2 = new JPanel();
        p2.setPreferredSize(new Dimension(550, 40));
        p2.add(new JLabel("\t生成路径:"));
        wordOutName = new JTextField(20);
        wordOutName.setEnabled(false);
        wordOutName.setPreferredSize(new Dimension(wordOutName.getWidth(), wordOutName.getHeight() + 23));
        p2.add(wordOutName);
        wordChoose = new JButton("选择路径");
        wordChoose.addActionListener(e -> {

            int i = wordOutChooser.showOpenDialog(frame.getContentPane());// 显示文件选择对话框

            // 判断用户单击的是否为“打开”按钮
            if (i == JFileChooser.APPROVE_OPTION) {

                File selectedFile = wordOutChooser.getSelectedFile();// 获得选中的文件对象
                wordOutName.setText(selectedFile.getPath());// 显示选中文件的名称
            }
        });
        p2.add(wordChoose);

        //------------------------需要分割的列---------------
        p3 = new JPanel();
        p3.setPreferredSize(new Dimension(550, 40));
        p3.add(new JLabel("\t需要格式化的列序号"));
        splitString = new JTextField(20);
        splitString.setDocument(new NumberDocument());
        splitString.setPreferredSize(new Dimension(wordOutName.getWidth(), wordOutName.getHeight() + 23));
        p3.add(splitString);

        //------------------------操作列---------------------
        p4 = new JPanel();
        register = new JButton("拆分");
//        word2pdf = new JButton("转换PDF");
        register.addActionListener(e -> {
            try {
                long t1 = System.currentTimeMillis();
                String out = wordOutChooser.getSelectedFile().getPath()+ "\\" + dataChooser.getSelectedFile().getName();
                Result rs = ExcelUtil.readTable(dataChooser.getSelectedFile(), Integer.parseInt(splitString.getText()));
                if(rs.isSuccess()){
                    ExcelUtil.createDocument(rs.getData(), out);
                    saveProp();
                }
                JOptionPane.showMessageDialog(frame, "本次耗时：" + ((System.currentTimeMillis() - t1) / 1000) + "秒\n操作结果：" + rs.getMsg());
            } catch (IOException exception) {
                exception.printStackTrace();
                if(exception instanceof FileNotFoundException){
                    JOptionPane.showMessageDialog(frame, "拆分失败，请检查导入的excel数据是否存在");
                } else {
                    JOptionPane.showMessageDialog(frame, "拆分失败，请检查excel单元格格式或数据换行");
                }
            }
        });

        p4.add(register);
        p4.setPreferredSize(new Dimension(550, 40));

        frame.add(p1);
        frame.add(p2);
        frame.add(p3);
        frame.add(p4);

        frame.pack();
        frame.setVisible(true);
        show();

        //初始化配置
        JSONObject prop = loadProp();
        if(prop.containsKey("dataPath")){
            dataName.setText(prop.getString("dataPath"));
            dataChooser.setSelectedFile(new File(prop.getString("dataPath")));
        }

        if(prop.containsKey("outPath")){
            wordOutName.setText(prop.getString("outPath"));
            wordOutChooser.setSelectedFile(new File(prop.getString("outPath")));
        }

        if(prop.containsKey("index")){
            splitString.setText(prop.getString("index"));
        }
    }

    public void show() {
        frame.setBounds(500, 500, 550, 400);//设置大小
        frame.setLayout(new FlowLayout());//设置流式布局
    }

    public void saveProp() {
        try {
            File outFile = new File("conf/save2path.json");
            if (!outFile.getParentFile().exists()) {
                outFile.getParentFile().mkdirs();
            }
            Writer fw = new FileWriter("conf/save2path.json");
            JSONObject map = new JSONObject();
            map.put("dataPath", dataChooser.getSelectedFile().getPath());
            map.put("outPath", wordOutChooser.getSelectedFile().getPath());
            map.put("index", splitString.getText());
            fw.write(map.toString());
            fw.flush();
            fw.close();
        } catch (Exception e1) {
            JOptionPane.showMessageDialog(null, "创建历史配置文档失败");
        }
    }

    public JSONObject loadProp() {
        //获取JSON文件内容，转化为字符串类型
        StringBuilder json = new StringBuilder();
        try {
            String temp = "";
            Reader fw = new FileReader("conf/save2path.json");
            BufferedReader bfr = new BufferedReader(fw);
            while ((temp = bfr.readLine()) != null) {
                json.append(temp);
            }
            bfr.close();
            fw.close();
        } catch (Exception e) {
            return new JSONObject();
        }

        //将字符串转化为json
        return JSON.parseObject(json.toString());
    }
}

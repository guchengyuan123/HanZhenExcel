package com.lijia.excel;

import com.alibaba.fastjson.JSON;
import com.github.crab2died.ExcelUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.Version;
import org.apache.commons.collections4.CollectionUtils;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelToWord {
    private JPanel jpanel;
    private JButton button;
    private JTextField excelUrl;
    private JTextField wordSaveUrl;

    public ExcelToWord() {
        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String excelUrl = getExcelUrl().getText();
                String wordSaveUrl = getWordSaveUrl().getText();
                excelExcute(excelUrl, wordSaveUrl);
            }
        });
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame("转换鸭");
        frame.setContentPane(new ExcelToWord().jpanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }

    public String excelExcute(String path, String wordSaveUrl) {
        try {
            List<HanZheng> hanZhengs = ExcelUtils.getInstance().readExcel2Objects(path, HanZheng.class, 0, 0);
            System.out.println("读取Excel至对象数组(支持类型转换)：");
            for (int i = 0; i < hanZhengs.size(); i++) {
                HanZheng hanZheng = hanZhengs.get(i);
                if (hanZheng.getId() == null && i > 0) {
                    HanZheng pre = hanZhengs.get(i - 1);
                    hanZheng.setId(pre.getId());
                    hanZheng.setJjmc(pre.getJjmc());
                    hanZheng.setBtzgsmc(pre.getBtzgsmc());
                    hanZheng.setLxr(pre.getLxr());
                    hanZheng.setBtzgsdz(pre.getBtzgsdz());
                    hanZheng.setLxdh(pre.getLxdh());
                }
            }
            Map<String, List<HanZheng>> map = hanZhengs.stream().collect(Collectors.groupingBy(o -> o.getId()));
            System.out.println(JSON.toJSONString(map, true));
            String fileName = this.getClass().getClassLoader().getResource("hzmb.ftl").getPath();
            map.forEach((id, list) -> {
                try {
                    Map<String, Object> dataMap = new HashMap<String, Object>();

                    dataMap.put("被投资公司名称", list.get(0).getBtzgsmc());
                    dataMap.put("被投资公司地址", list.get(0).getBtzgsdz());
                    dataMap.put("联系人", list.get(0).getLxr());
                    dataMap.put("基金名称", list.get(0).getJjmc());

                    for (String type : getTypes()) {
                        if (type.equals("Convertible Promissory note")) {
                            List<HanZheng> cpn = list.stream().filter(o -> o.getTzlx().equals(type)).collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(cpn)) {
                                dataMap.put("可换票投资成本", "");
                                dataMap.put("可换票计息收入", "");
                            } else {
                                dataMap.put("可换票投资成本", formatString2(cpn.get(0).getCost()));
                                dataMap.put("可换票计息收入", formatString2(cpn.get(0).getAi()));
                            }
                        }
                        if (type.equals("Common Shares")) {
                            List<HanZheng> cs = list.stream().filter(o -> o.getTzlx().equals(type)).collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(cs)) {
                                dataMap.put("普股数量", "");
                                dataMap.put("普股持股比例", "");
                                dataMap.put("普股投资成本", "");
                            } else {
                                dataMap.put("普股数量", formatString2(cs.get(0).getNos()));
                                dataMap.put("普股持股比例", formatString(cs.get(0).getSp()));
                                dataMap.put("普股投资成本", formatString2(cs.get(0).getCost()));
                            }
                        }
                        if (type.equals("Preferred Shares")) {
                            List<HanZheng> ps = list.stream().filter(o -> o.getTzlx().equals(type)).collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(ps)) {
                                dataMap.put("优先股持股数", "");
                                dataMap.put("优先股数量", "");
                                dataMap.put("优先股持股比例", "");
                                dataMap.put("优先股投资成本", "");
                            } else {
                                dataMap.put("优先股持股数", ps.get(0).getClazz() == null ? "" : formatString2(ps.get(0).getClazz()));
                                dataMap.put("优先股数量", formatString2(ps.get(0).getNos()));
                                dataMap.put("优先股持股比例", formatString(ps.get(0).getSp()));
                                dataMap.put("优先股投资成本", formatString2(ps.get(0).getCost()));
                            }
                        }
                        if (type.equals("Warrants")) {
                            List<HanZheng> w = list.stream().filter(o -> o.getTzlx().equals(type)).collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(w)) {
                                dataMap.put("认权数量", "");
                                dataMap.put("认权投资成本", "");
                            } else {
                                dataMap.put("认权数量", formatString2(w.get(0).getNos()));
                                dataMap.put("认权投资成本", formatString2(w.get(0).getCost()));
                            }
                        }
                    }

                    File outFile = new File(wordSaveUrl + "/hanzheng" + id + ".doc");
                    Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"), 10240);
                    Configuration configuration = new Configuration(new Version("2.3.0"));
                    configuration.setDefaultEncoding("utf-8");
                    configuration.setDirectoryForTemplateLoading(new File(fileName.replace("hzmb.ftl", "")));

                    Template template = configuration.getTemplate("hzmb.ftl", "utf-8");
                    template.process(dataMap, out);

                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
            //处理word
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }

    private List<String> getTypes() {
        List<String> list = new ArrayList<>();
        list.add("Convertible Promissory note");
        list.add("Common Shares");
        list.add("Preferred Shares");
        list.add("Warrants");
        return list;
    }
    public static String formatString(String data) {
        float dataf=Float.parseFloat(data);
        DecimalFormat df = new DecimalFormat("0.00%");
        return df.format(dataf);
    }

    public static String formatString2(String data) {

        float dataf=Float.parseFloat(data);
        DecimalFormat df = new DecimalFormat("#,###");
        return dataf == 0f ? "-" : df.format(dataf);
    }

    public JPanel getJpanel() {
        return jpanel;
    }

    public void setJpanel(JPanel jpanel) {
        this.jpanel = jpanel;
    }

    public JButton getButton() {
        return button;
    }

    public void setButton(JButton button) {
        this.button = button;
    }

    public JTextField getExcelUrl() {
        return excelUrl;
    }

    public void setExcelUrl(JTextField excelUrl) {
        this.excelUrl = excelUrl;
    }

    public JTextField getWordSaveUrl() {
        return wordSaveUrl;
    }

    public void setWordSaveUrl(JTextField wordSaveUrl) {
        this.wordSaveUrl = wordSaveUrl;
    }
}

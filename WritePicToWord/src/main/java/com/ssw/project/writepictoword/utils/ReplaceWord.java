package com.ssw.project.writepictoword.utils;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 文件数据替换
 *
 * @author 23  *
 */

public class ReplaceWord {

    public static String path = "D:\\tp\\cs\\test.docx";

    public static void main(String[] args) throws Exception {
        Map<String, Object> data = new HashMap<>();
        Map<String, Object> pic = new HashMap<>();
        List<String> list = new ArrayList<>();
        list.add("D:\\tp\\pie.png");
        list.add("D:\\tp\\pie.png");
        pic.put("${pics}$", list);
        data.put("$projectName$", "项目名称");
        data.put("$major$", "专业");
        data.put("$suoBie$", "所别");
        data.put("$firstFounded$", "□");
        data.put("$oneEdition$", "□");
        data.put("$twoEdition$", "□");
        data.put("$others$", "☑");
        data.put("$productionDrawing$", "施工图");
        data.put("$|BIM|$", "施工图");
        data.put("$problemDescription$", "问题描述");
        data.put("$category$", "类别");
        data.put("$classification$", "分级");
        data.put("$suggestionReply$", "意见回复");
        data.put("$problemsOne$", "1");
        data.put("$problemsTwo$", "2");
        data.put("$problemsThree$", "3");
        data.put("$problemsFour$", "4");
        data.put("$superior$", "☑");
        data.put("$good$", "□");
        data.put("$pass$", "□");
        data.put("$fail$", "□");

        // 列表(List)是对象的有序集合
        List<List<String[]>> tabledataList = new ArrayList<>();
        getWord(data, tabledataList, pic);
    }

    public static void getWord(Map<String, Object> data, List<List<String[]>> tabledataList, Map<String, Object> picmap)
            throws Exception {
        try (FileInputStream is = new FileInputStream(path); XWPFDocument document = new XWPFDocument(is)) {
            // 替换掉表格之外的文本(仅限文本)
            changeText(document, data);

            // 替换表格内的文本对象
            changeTableText(document, data);

            // 替换图片
//            changePic(document, picmap);

            // 替换表格内的图片对象
            changeTablePic(document, picmap);

            long time = System.currentTimeMillis();// 获取系统时间
            System.out.println(time); // 打印时间
            // 使用try和catch关键字捕获异常
            try (FileOutputStream out = new FileOutputStream("D:\\tp\\cs\\test2" + ".docx")) {
                document.write(out);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, Object> textMap) {
        // 获取段落集合
        // 返回包含页眉或页脚文本的段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        // 增强型for循环语句，前面一个为声明语句，后一个为表达式
        for (XWPFParagraph paragraph : paragraphs) {
            // 判断此段落是否需要替换
            String text = paragraph.getText();// 检索文档中的所有文本
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    // 替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    if (ob instanceof String) {
                        if (textMap.containsKey(run.toString())) {
                            run.setText((String) ob, 0);
                        }
                    }
                }
            }
        }
    }

    /* 检查文本中是否包含指定的字符(此处为“$”)，并返回值 */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.contains("$")) {
            check = true;
        }
        return check;
    }

    /**
     * 替换图片
     *
     * @param document
     * @param textMap
     * @throws Exception
     */

    public static void changePic(XWPFDocument document, Map<String, Object> textMap) throws Exception {
        // 获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            // 判断此段落是否需要替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    // 替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    if (ob instanceof String) {
                        if (textMap.containsKey(run.toString())) {
                            run.setText("", 0);
                            try (FileInputStream is = new FileInputStream((String) ob)) {
                                run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, (String) ob, Units.toEMU(100),
                                        Units.toEMU(100));
                            }
                        }
                    }
                }
            }
        }
    }

    public static void changeTableText(XWPFDocument document, Map<String, Object> data) {
        // 获取文件的表格
        List<XWPFTable> tableList = document.getTables();

        // 循环所有需要进行替换的文本，进行替换
        for (int i = 0; i < tableList.size(); i++) {
            XWPFTable table = tableList.get(i);
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                // 遍历表格，并替换模板
                eachTable(document, rows, data);
            }
        }
    }

    public static void changeTablePic(XWPFDocument document, Map<String, Object> pic) throws Exception {
        List<XWPFTable> tableList = document.getTables();

        // 循环所有需要替换的文本，进行替换
        for (int i = 0; i < tableList.size(); i++) {
            XWPFTable table = tableList.get(i);
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                System.out.println("简单表格替换：" + rows);
                // 遍历表格，并替换模板
                eachTablePic(document, rows, pic);
            }
        }
    }

    public static void eachTablePic(XWPFDocument document, List<XWPFTableRow> rows, Map<String, Object> pic)
            throws Exception {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                // 判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
//                            if (ob instanceof String) {
                                System.out.println("run" + "'" + run.toString() + "'");
                                if (pic.containsKey(run.toString())) {
                                    List<String> list = (List<String>) changeValue(run.toString(), pic);
                                    System.out.println("run" + run.toString() + "替换为" + list);
                                    run.setText("", 0);
                                    for (String ob : list) {
                                        try (FileInputStream is = new FileInputStream(ob)) {
                                            run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, ob, Units.toEMU(420),
                                                    Units.toEMU(250));
                                        }
                                    }
                                } else {
                                    System.out.println("'" + run.toString() + "' 不匹配");
                                }
//                            }
                        }
                    }
                }
            }
        }
    }

    public static Object changeValue(String value, Map<String, Object> textMap) {
        Set<Map.Entry<String, Object>> textSets = textMap.entrySet();
        Object valu = "";
        for (Map.Entry<String, Object> textSet : textSets) {
            // 匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.contains(key)) {
                valu = textSet.getValue();
            }
        }
        return valu;
    }

    public static void eachTable(XWPFDocument document, List<XWPFTableRow> rows, Map<String, Object> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                // 判断单元格是否需要替换
                String text = cell.getText();
                if (checkText(text)) {
                    System.out.println("cell:" + text);
                    /**
                     * 动态替换表格中文本
                     */
                    int time = 0;
                    for (Map.Entry<String, Object> entry : textMap.entrySet()) {
                        if (text.contains(entry.getKey())) {
                            time ++;
                            System.out.println("repalce table's text");
                            text = text.replace(entry.getKey(), (String)entry.getValue());
                        }
                    }
                    if (time > 0) {
                        cell.removeParagraph(0);
                        cell.addParagraph();
                        cell.setText(text);
                    }

//                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
//                    for (XWPFParagraph paragraph : paragraphs) {
//                        List<XWPFRun> runs = paragraph.getRuns();
//                        for (XWPFRun run : runs) {

//                            Object ob = changeValue(run.toString(), textMap);
//                            if (ob instanceof String) {
//
//                                System.out.println("run:" + "'" + run.toString() + "'");
//                                if (textMap.containsKey(run.toString())) {
//                                    System.out.println("run:" + run.toString() + "替换为" + ob);
//                                    run.setText((String) ob, 0);
//                                } else {
//                                    System.out.println("'" + run.toString() + "'不匹配");
//                                }
//                            }
//                        }
//                    }
                }
            }
        }
    }
}

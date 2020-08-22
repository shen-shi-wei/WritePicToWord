package com.ssw.project.writepictoword.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 *
 * @Auther: ssw
 * @Date: 2020/08/21/17:31
 * @Description:
 */
public class replaceManyTableUtil {

    public static void main(String[] args) throws Exception {
        Map<String,String> header = new HashMap<>();
        header.put("$projectName$", "项目名称");
        header.put("$major$", "专业");
        header.put("$suoBie$", "所别");
        header.put("$firstFounded$", "□");
        header.put("$oneEdition$", "□");
        header.put("$twoEdition$", "□");
        header.put("$others$", "☑");
        header.put("$productionDrawing$", "施工图");
        header.put("$|BIM|$", "BIM");

        List<Map<String, String>> body = new ArrayList<>();
        Map<String, String> body1 = new HashMap<>();
        body1.put("$problemDescription$", "问题描述1");
        body1.put("$category$", "类别1");
        body1.put("${pics}$", "D:\\tp\\pie.png");
        body1.put("$classification$", "分级1");
        body1.put("$suggestionReply$", "意见回复1");
        body.add(body1);
        Map<String, String> body2 = new HashMap<>();
        body2.put("$problemDescription$", "问题描述2");
        body2.put("$category$", "类别2");
        body2.put("${pics}$", "D:\\tp\\pie.png");
        body2.put("$classification$", "分级2");
        body2.put("$suggestionReply$", "意见回复2");
        body.add(body2);

        Map<String,String> footer = new HashMap<>();
        footer.put("$problemsOne$", "1");
        footer.put("$problemsTwo$", "2");
        footer.put("$problemsThree$", "3");
        footer.put("$problemsFour$", "4");
        footer.put("$superior$", "☑");
        footer.put("$good$", "□");
        footer.put("$pass$", "□");
        footer.put("$fail$", "□");

        //创建word模板
        createWord(body.size(),"D:\\tp\\cs\\collision_header.docx","D:\\tp\\cs\\collision_body.docx","D:\\tp\\cs\\collision_footer.docx","D:\\tp\\cs\\test3.docx");

        //替换内容
        getWord(header, body, footer,"D:\\tp\\cs\\test3.docx","D:\\tp\\cs\\test3.docx");

    }

    /**
     * 创建word模板
     * @throws Exception
     */
    public static void createWord(int size, String header, String detail, String footer, String outPath) throws Exception{
        try (FileInputStream is = new FileInputStream(header); XWPFDocument document = new XWPFDocument(is)) {
            List<XWPFTable> tableList = document.getTables();
            XWPFTable table = tableList.get(0);
            for (int i = 0; i < size-1; i++) {
                XWPFParagraph xwpfParagraph = document.createParagraph();//设置分页
                xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                createDetail(detail, document);
            }
            createFooter(footer,document);

            long time = System.currentTimeMillis();// 获取系统时间
            System.out.println(time); // 打印时间
            // 使用try和catch关键字捕获异常
            try (FileOutputStream out = new FileOutputStream(outPath)) {
                document.write(out);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
    private static void createFooter(String footer, XWPFDocument document) throws Exception{
        List<XWPFTable> tables = document.getTables();
        System.out.println("all table's num is :"+ tables.size());
        try (FileInputStream is = new FileInputStream(footer); XWPFDocument doc = new XWPFDocument(is)) {
            List<XWPFTable> tableList = doc.getTables();
            XWPFTable table = tableList.get(0);
            XWPFTable xwpfTable = tables.get(tables.size() - 1);
            xwpfTable.addRow(table.getRow(0));
            xwpfTable.addRow(table.getRow(1));
            xwpfTable.addRow(table.getRow(2));
            xwpfTable.addRow(table.getRow(3));
            xwpfTable.addRow(table.getRow(4));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    private static void createDetail(String detail, XWPFDocument document) throws Exception{
        try (FileInputStream is = new FileInputStream(detail); XWPFDocument doc = new XWPFDocument(is)) {
            List<XWPFTable> tableList = doc.getTables();
            XWPFTable table = tableList.get(0);
            XWPFParagraph xwpfParagraph = document.createParagraph();//设置分页
            xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
            xwpfParagraph.setPageBreak(true);
            XWPFTable newTable = document.createTable();// 创建一个空的Table
            newTable.removeRow(0);
            newTable.addRow(table.getRow(0));
            newTable.addRow(table.getRow(1));
            newTable.addRow(table.getRow(2));
            newTable.addRow(table.getRow(3));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public static void getWord(Map<String, String> header, List<Map<String, String>> body, Map<String, String> footer, String path, String outPath) throws Exception{

        try (FileInputStream is = new FileInputStream(path); XWPFDocument document = new XWPFDocument(is)) {

            // 替换表头内的文本对象
            changeTableText(document, header);

            // 替换表格内的图片对象
            changeTablePic(document, body);

            // 替换表尾内的文本对象
            changeTableText(document, footer);

            long time = System.currentTimeMillis();// 获取系统时间
            System.out.println(time); // 打印时间
            // 使用try和catch关键字捕获异常
            try (FileOutputStream out = new FileOutputStream(outPath)) {
                document.write(out);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }


    }

    public static void changeTableText(XWPFDocument document, Map<String, String> textMap) {
        // 获取文件的表格
        List<XWPFTable> tableList = document.getTables();

        // 循环所有需要进行替换的文本，进行替换
        for (int i = 0; i < tableList.size(); i++) {
            XWPFTable table = tableList.get(i);
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                changeRow(textMap, row);
            }

        }
    }

    private static void changeRow(Map<String, String> textMap, XWPFTableRow row) {
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
                for (Map.Entry<String, String> entry : textMap.entrySet()) {
                    if (text.contains(entry.getKey())) {
                        time ++;
                        System.out.println("repalce table's text");
                        text = text.replace(entry.getKey(), entry.getValue());
                    }
                }
                if (time > 0) {
                    cell.removeParagraph(0);
                    cell.addParagraph();
                    cell.setText(text);
                }
            }
        }
    }

    public static void changeTablePic(XWPFDocument document, List<Map<String, String>> body) throws Exception{
        // 获取文件的表格
        List<XWPFTable> tableList = document.getTables();
        XWPFTable table = tableList.get(0);
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < body.size(); i++) {
            Map<String, String> map = body.get(i);
            if (i == 0 ) {
                changeDetail(rows, 3, map);
            }else {
                XWPFTable xwpfTable = tableList.get(i);
                List<XWPFTableRow> tableRows = xwpfTable.getRows();
                changeDetail(tableRows, 0, map);
            }
        }
    }

    private static void changeDetail(List<XWPFTableRow> rows, int i, Map<String, String> map) throws IOException, InvalidFormatException {
        changeRow(map, rows.get(i)); //问题描述
        changeRow(map, rows.get(i+1));  //类别
        changeRow(map, rows.get(i+3));  //意见回复
        XWPFTableRow xwpfTableRow5 = rows.get(i+2);
        List<XWPFTableCell> cells = xwpfTableRow5.getTableCells();
        for (XWPFTableCell cell : cells) {
            // 判断单元格是否需要替换
            if (checkText(cell.getText())) {
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    for (XWPFRun run : runs) {
//                            if (ob instanceof String) {
                        System.out.println("run" + "'" + run.toString() + "'");
                        if (map.containsKey(run.toString())) {
                            String ob = map.get(run.toString());
                            System.out.println("run" + run.toString() + "替换为" + ob);
                            run.setText("", 0);
                            try (FileInputStream is = new FileInputStream(ob)) {
                                run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, ob, Units.toEMU(420),
                                        Units.toEMU(250));
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

    /* 检查文本中是否包含指定的字符(此处为“$”)，并返回值 */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.contains("$")) {
            check = true;
        }
        return check;
    }

}

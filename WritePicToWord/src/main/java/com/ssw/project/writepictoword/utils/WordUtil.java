//package com.ssw.project.writepictoword.utils;
//
//import org.apache.poi.POIXMLDocument;
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.xwpf.usermodel.*;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.Iterator;
//import java.util.List;
//import java.util.Map;
//
///**
// * Created with IntelliJ IDEA.
// *
// * @Auther: ssw
// * @Date: 2020/08/12/15:15
// * @Description:
// */
//public class WordUtil {
//
//    /**
//     * 实现Java替换word中的文本图片或文字
//     */
//    public static void wordUtil(Map<String, String> map, Map<String, List<String>> pics, String template, String outPath) throws IOException, InvalidFormatException {
//        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(template));
//        Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
//        List<XWPFTable> tables = document.getTables();
//        for (XWPFTable table : tables) {
//            for (int i = 0; i < table.getNumberOfRows(); i++) {
//                XWPFTableRow row = table.getRow(i);
//                List<XWPFTableCell> tableCells = row.getTableCells();
//                for (int j = 0; j < tableCells.size(); j++) {
//                    XWPFTableCell cell = tableCells.get(j);
//                    String text = cell.getText();
//                    System.out.println("text = " + text);
//                    /**
//                     * 动态替换表格中文本
//                     */
//                    int time = 0;
//                    for (Map.Entry<String, String> entry : map.entrySet()) {
//                        if (text.contains(entry.getKey())) {
//                            time ++;
//                            System.out.println("repalce table's text");
//                            text = text.replace(entry.getKey(), entry.getValue());
//                        }
//                    }
//                    if (time > 0) {
//                        cell.removeParagraph(0);
//                        cell.addParagraph();
//                        cell.setText(text);
//                    }
//
//                    /**
//                     * 动态替换表格中的文本为图片
//                     */
//                    for (Map.Entry<String, List<String>> entry : pics.entrySet()) {
//                        if (text.contains(entry.getKey())) {
//                            cell.removeParagraph(0);
//                            //  cell.setText("aa");
//                            XWPFParagraph pargraph =    cell.addParagraph();
//                            List<String> paths = entry.getValue();
//                            for (String path : paths) {
//                                System.out.println("replace table'text to pic");
//                                OPCPackage pack = document.getPackage();
//                                CustomXWPFDocument doc = new CustomXWPFDocument(pack);
//                                File pic = new File(path);
//                                FileInputStream is = new FileInputStream(pic);
//                                int ind =   doc.addPicture(is, doc.PICTURE_TYPE_PNG);
//                                doc.createPicture(ind, 500, 250,pargraph);
//                            }
//                        }
//                    }
//                }
//            }
//        }
//
//        /**
//         * 动态替换段落里面的文本
//         */
//        while (itPara.hasNext()) {
//            XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
//            List<XWPFRun> runs = paragraph.getRuns();
//            for (int i = 0; i < runs.size(); i++) {
//                String oneparaString = runs.get(i).getText(runs.get(i).getTextPosition());
//                System.out.println("oneparaString = " + oneparaString);
//                for (Map.Entry<String, String> entry : map.entrySet()) {
//                    if (oneparaString.contains(entry.getKey())) {
//                        oneparaString = oneparaString.replace(entry.getKey(), entry.getValue());
//                    }
//                }
//                runs.get(i).setText(oneparaString, 0);
//            }
//        }
//        FileOutputStream outStream = null;
//        outStream = new FileOutputStream(outPath);
//        document.write(outStream);
//        outStream.close();
//    }
//}

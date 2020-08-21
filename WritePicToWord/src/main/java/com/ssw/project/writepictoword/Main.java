package com.ssw.project.writepictoword;

import com.google.common.collect.Maps;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 *
 * @Auther: ssw
 * @Date: 2020/08/12/13:19
 * @Description:
 */
public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, String> map = Maps.newHashMap();
        map.put("$1", "测试项目");
        map.put("$2", "软件工程");
        map.put("$3", "一级");
        map.put("$4", "□");
        map.put("$5", "□");
        map.put("$6", "□");
        map.put("$7", "☑");
        map.put("$8", "测试施工图");
        map.put("$9", "BIM测试1111111111111111111111");
        map.put("$10", "爱的人跟科比牛仔短裤");
        map.put("$11", "A");
        map.put("$12", "I");
        map.put("$13", "1德国和你激动个续费发的功能性和烦恼没吃过v迷惑敌人坦克圆满成功粉红色走人拿分发现一看就是每次v美女先发给女的和支付宝和自然风干阖家幸福Erfghjdsasghfhhfj fhgdm RYDTJU  hfgh地图开门红3");
        map.put("$14", "14");
        map.put("$15", "15");
        map.put("$16", "16");
        map.put("$17", "17");
        map.put("$18", "□");
        map.put("$19", "□");
        map.put("$20", "☑");
        map.put("$21", "☑");
        Map<String, List<String>> pics = Maps.newHashMap();
        List<String> list = new ArrayList<>();
        list.add("D:\\tp\\pie.png");
        list.add("D:\\tp\\pie.png");
        pics.put("${pics}",list);
//        WordUtil.wordUtil(map, pics,"D:\\tp\\cs\\test.docx","D:\\tp\\cs\\test1.docx");







    }
}

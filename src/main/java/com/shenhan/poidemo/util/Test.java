package com.shenhan.poidemo.util;

import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author shenhan
 * @Date: 2019/12/16 22:23
 * @Description:
 */
public class Test {
    public void method1() throws Exception {
        // 循环数据
        List<Object> list = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            Map<String, Object> data = new HashMap<>();
            data.put("a1", (int) (Math.random() * 100));
            data.put("a2", (int) (Math.random() * 100));
            data.put("a3", (int) (Math.random() * 100));
            data.put("a4", (int) (Math.random() * 100));
            data.put("a5", (int) (Math.random() * 100));
            data.put("a6", (int) (Math.random() * 100));
            list.add(data);
        }
        // 表格使用的数据
        Map map = new HashMap();
        map.put("data", list);
        map.put("projectName", "智能井盖");
        map.put("projectUrl", "http://www.baidu.com");
        map.put("projectId", "jkiajsduioqweh123");
        map.put("projectDesc", "智能井盖");
        map.put("masterKey", "hasdiuquwhe12312334");
        // 获取模板文件
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("template.xlsx");
        // 实例化 XLSTransformer 对象
        XLSTransformer xlsTransformer = new XLSTransformer();
        // 获取 Workbook ，传入 模板 和 数据
        Workbook workbook = xlsTransformer.transformXLS(is, map);
        // 写出文件
        OutputStream os = new BufferedOutputStream(new FileOutputStream("D://temp.xlsx"));
        // 输出
        workbook.write(os);
        // 关闭和刷新管道，不然可能会出现表格数据不齐，打不开之类的问题
        is.close();
        os.flush();
        os.close();
    }

    public static void main(String[] args) {
        Test test = new Test();
        try {
            test.method1();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

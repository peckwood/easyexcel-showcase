package com.example.easyexcel_showcase.demo.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.ListUtils;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class WriteDemo01 {
    /**
     * 不创建对象的写
     */
    @Test
    public void noModelWrite() {
        // 写法1
        String fileName = "output/noModelWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName).head(head()).sheet("模板").doWrite(dataList());
    }
    public static List<List<String>> head() {
        List<List<String>> list = new ArrayList<>();
        List<String> head0 = new ArrayList<>();
        head0.add("大标题");
        head0.add("字符串" + System.currentTimeMillis());
        List<String> head1 = new ArrayList<>();
        head1.add("大标题");
        head1.add("数字" + System.currentTimeMillis());
        List<String> head2 = new ArrayList<>();
        head2.add("大标题");
        head2.add("日期" + System.currentTimeMillis());
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }
    public static List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<Object> data = new ArrayList<>();
            data.add("字符串" + i);
            data.add(0.56);
            data.add(new Date());
            list.add(data);
        }
        return list;
    }
}

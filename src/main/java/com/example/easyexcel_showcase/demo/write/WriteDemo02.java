package com.example.easyexcel_showcase.demo.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.example.easyexcel_showcase.handler.WriteDemo02CustomCellWriteHandler;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class WriteDemo02 {
    @Test
    public void demo02() {
        //构建工资单标题
        List<String> titleList = new ArrayList<>();
        titleList.add("序号");
        titleList.add("姓名");
        titleList.add("职务工资");
        List<List<Object>> gzdDataList = new ArrayList<>();
        List<Object> row1 = new ArrayList<>();
        row1.add("1");
        row1.add("张三");
        row1.add("3400");
        List<Object> row2 = new ArrayList<>();
        row2.add("2");
        row2.add("李四");
        row2.add("4400");
        List<Object> row3 = new ArrayList<>();
        row3.add("合计");
        row3.add("");
        row3.add("9999");
        gzdDataList.add(row1);
        gzdDataList.add(row2);
        gzdDataList.add(row3);

        //最后一行
        String[] lastRowArr = {"制表：                 审核：                      部门负责人："};
        gzdDataList.add(Arrays.asList(lastRowArr));

        // 保险标题
        List<String> insuranceTitleList = new ArrayList<>();
        insuranceTitleList.add("序号");
        insuranceTitleList.add("姓名");
        insuranceTitleList.add("工资单时间");
        insuranceTitleList.add("基数");

        // 准备生成excel
//        String fileName = URLEncoder.encode("历史工资单.xlsx", "UTF-8");
        String fileName = "output/演示02-" + System.currentTimeMillis() + ".xlsx";

        List<List<Object>> insuranceDataList = new ArrayList<>();
        List<Object> insuranceRow1 = new ArrayList<>();
        insuranceRow1.add("3");
        insuranceRow1.add("张三2");
        insuranceRow1.add("2402");
        insuranceRow1.add("3000");
        List<Object> insuranceRow2 = new ArrayList<>();
        insuranceRow2.add("4");
        insuranceRow2.add("李四2");
        insuranceRow2.add("2403");
        insuranceRow2.add("3000");
        List<Object> insuranceRow3 = new ArrayList<>();
        insuranceRow3.add("合计");
        insuranceRow3.add("");
        insuranceRow3.add("2403");
        insuranceRow3.add("6000");
        insuranceDataList.add(insuranceRow1);
        insuranceDataList.add(insuranceRow2);
        insuranceDataList.add(insuranceRow3);
        insuranceDataList.add(Arrays.asList(lastRowArr));

        List<List<String>> gzdHead = new ArrayList<>();
        titleList.forEach(t -> {
            List<String> l = new ArrayList<>();
            l.add("太原市总工会机关工资单");
            l.add("账套号: 01，账套单位：太原市总工会       日期： ");
            l.add(t);
            gzdHead.add(l);
        });
        List<List<String>> insuranceHead = new ArrayList<>();
        insuranceTitleList.forEach(t -> {
            List<String> l = new ArrayList<>();
            l.add("太原市总工会各类保险缴纳明细");
            l.add("账套号: 01，账套单位：太原市总工会       日期： ");
            l.add(t);
            insuranceHead.add(l);
        });


        WriteDemo02CustomCellWriteHandler handler = new WriteDemo02CustomCellWriteHandler();
        handler.setDataSize(gzdDataList.size());
        handler.setTitleSize(titleList.size());
        handler.setInsuranceTitleSize(insuranceTitleList.size());

        // 生成excel
        try (ExcelWriter excelWriter =
                     EasyExcel.write(fileName)
                             .registerWriteHandler(handler)
                             //设置所有列的宽度
                             .registerWriteHandler(new SimpleColumnWidthStyleStrategy(30))
                             .build())
        {

            WriteSheet writeSheet1 = EasyExcel.writerSheet(0, "历史工资单信息").head(gzdHead).build();
            excelWriter.write(gzdDataList, writeSheet1);
            WriteSheet writeSheet2 = EasyExcel.writerSheet(1, "历史各类保险缴纳信息").head(insuranceHead).build();
            excelWriter.write(insuranceDataList, writeSheet2);
        }





    }

}

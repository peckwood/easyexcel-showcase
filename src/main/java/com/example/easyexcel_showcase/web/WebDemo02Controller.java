package com.example.easyexcel_showcase.web;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@RestController
public class WebDemo02Controller{
    private final ObjectMapper objectMapper = new ObjectMapper();

    @GetMapping("web-demo02")
    public void web1(HttpServletResponse response) throws IOException{

        try{
            //to manually create error
//            int a = 3/0;

            ServletOutputStream outputStream = response.getOutputStream();

            Map<String, Object> map = new HashMap<>();
            map.put("startYearText", "startYearText");
            map.put("startMonthText", "startMonthText");
            map.put("endYearText", "endYearText");
            map.put("endMonthText", "endMonthText");
            map.put("tbYearStartText", "tbYearStartText");
            map.put("tbYearEndText", "tbYearEndText");

            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode("下载文件名", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这种也可以, 但是不如上一种, 上一种不行在这样
//        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

            //注意: 一般取的是启动类所在的模块的resources文件夹, 可能与当前类所在模块不同
            String templateFileName = WebDemo02Controller.class.getResource("/").getPath() + "excel-template/WebDemo02ControllerTemplate.xlsx";
            ExcelWriter excelWriter = EasyExcel
                    .write(outputStream)
                    .withTemplate(templateFileName)
                        .registerWriteHandler(new CellWriteHandler() {
                            @Override
                            public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead){

                                Sheet sheet = writeSheetHolder.getSheet();
                                //第3个sheet
                                if (writeSheetHolder.getSheetNo() == 2) {
                                    //第1列
                                    if(cell.getColumnIndex() == 0){
                                        int rowIndex = cell.getRowIndex();
                                        //合并第1列4-6行
                                        if(rowIndex == 3){
                                            CellRangeAddress cellRangeAddress = new CellRangeAddress(3, 5, 0, 0);
                                            sheet.addMergedRegion(cellRangeAddress);
                                        }
                                        //合并第1列7-9行
                                        if(rowIndex == 6){
                                            CellRangeAddress cellRangeAddress = new CellRangeAddress(6, 8, 0, 0);
                                            sheet.addMergedRegion(cellRangeAddress);
                                        }
                                    }
                                }
                            }
                        })
                    .build();
            WriteSheet sheet1 = EasyExcel.writerSheet(0).build();
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
            excelWriter.fill(getData(), fillConfig, sheet1);
            excelWriter.fill(map, sheet1);

            WriteSheet sheet2 = EasyExcel.writerSheet(1).build();
            excelWriter.fill(getData2(), sheet2);
            excelWriter.fill(map, sheet2);

            WriteSheet sheet3 = EasyExcel.writerSheet(2).build();
            excelWriter.fill(getData3(), sheet3);
            excelWriter.fill(map, sheet3);

            excelWriter.finish();

        } catch (Exception e){
            //出错的时候返回JSON

            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, Object> resultMap = new HashMap<>();
            resultMap.put("success", false);
            resultMap.put("msg", e.getMessage());

            response.getWriter().println(objectMapper.writeValueAsString(resultMap));
        }
    }

    private List<Map<String, Object>> getData(){
        return IntStream.rangeClosed(1, 10).mapToObj(int0 -> {
            Map<String, Object> map = new HashMap<>();
            map.put("admdvsName", "admdvsName" + int0);
            map.put("setlPsntime", "setlPsntime" + int0);
            map.put("setlPsntimeCur", "setlPsntimeCur" + int0);
            map.put("setlPsntimeZz", "setlPsntimeZz" + int0);
            map.put("setlPsntimeZzlText", "setlPsntimeZzlText" + int0);
            map.put("sumamt", "sumamt" + int0);
            map.put("sumamtCur", "sumamtCur" + int0);
            map.put("sumamtZz", "sumamtZz" + int0);
            map.put("sumamtZzlText", "sumamtZzlText" + int0);
            return map;
        }).collect(Collectors.toList());
    }

    private List<Map<String, Object>> getData2(){
        List<Map<String, Object>> list = IntStream.rangeClosed(1, 6).mapToObj(int0 -> {
            Map<String, Object> map = new HashMap<>();
            map.put("admdvsName", "admdvsName" + int0);
            map.put("setlPsntime", "setlPsntime" + int0);
            map.put("setlPsntimeTb", "setlPsntimeTb" + int0);
            map.put("setlPsntimeZz", "setlPsntimeZz" + int0);
            map.put("setlPsntimeZzlText", "setlPsntimeZzlText" + int0);
            map.put("sumamt", "sumamt" + int0);
            map.put("sumamtTb", "sumamtTb" + int0);
            map.put("sumamtZz", "sumamtZz" + int0);
            map.put("sumamtZzlText", "sumamtZzlText" + int0);
            return map;
        }).collect(Collectors.toList());
        return list;
    }

    private List<Map<String, Object>> getData3(){
        List<Map<String, Object>> list = IntStream.rangeClosed(1, 6).mapToObj(int0 -> {
            Map<String, Object> map = new HashMap<>();
            map.put("admdvsName", "admdvsName" + int0);
            map.put("setlPsntime", "setlPsntime" + int0);
            map.put("setlPsntimeTb", "setlPsntimeTb" + int0);
            map.put("setlPsntimeZz", "setlPsntimeZz" + int0);
            map.put("setlPsntimeZzlText", "setlPsntimeZzlText" + int0);
            map.put("sumamt", "sumamt" + int0);
            map.put("sumamtTb", "sumamtTb" + int0);
            map.put("sumamtZz", "sumamtZz" + int0);
            map.put("sumamtZzlText", "sumamtZzlText" + int0);
            return map;
        }).collect(Collectors.toList());

        list.get(0).put("col0", "参保地");
        list.get(1).put("col0", "参保地");
        list.get(2).put("col0", "参保地");
        list.get(3).put("col0", "就医地");
        list.get(4).put("col0", "就医地");
        list.get(5).put("col0", "就医地");
        list.get(0).put("col1", "省内");
        list.get(1).put("col1", "跨省");
        list.get(2).put("col1", "小计");
        list.get(3).put("col1", "省内");
        list.get(4).put("col1", "跨省");
        list.get(5).put("col1", "小计");
        return list;
    }
}

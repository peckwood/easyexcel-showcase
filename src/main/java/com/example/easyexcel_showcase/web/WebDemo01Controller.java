package com.example.easyexcel_showcase.web;

import com.alibaba.excel.EasyExcel;
import com.example.easyexcel_showcase.demo.write.WriteDemo01;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLEncoder;

@RestController
public class WebDemo01Controller {

    // 通过 http://localhost:8080/web-demo01 访问
    @GetMapping("web-demo01")
    public void web1(HttpServletResponse response) throws IOException {
        String fileName = URLEncoder.encode("文件名", "UTF-8").replaceAll("\\+", "%20");

        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        // 这种也可以, 但是不如上一种, 上一种不行在这样
//        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        EasyExcel.write(response.getOutputStream()).head(WriteDemo01.head()).sheet("sheetName1").doWrite(WriteDemo01.dataList());
    }
}

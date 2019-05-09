package com.example.reportdownload.controller;

import com.example.reportdownload.service.ReportService;
import com.example.reportdownload.util.DateUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

@Controller
@RequestMapping("/api")
public class DemoController {

    @Autowired
    private ReportService reportService;

    @RequestMapping("/download/report")
    @ResponseBody
    public void createReport(HttpServletResponse response){
        String filename =
                "攻击报告" + DateUtil.dateFormat(System.currentTimeMillis(), "yyyyMMddHHmmss") + ".docx";
        byte[] bytes = reportService.createReport();
        try {
            filename = new String(filename.getBytes("UTF-8"), "ISO8859-1");
            response.setHeader("content-type", "application/octet-stream");
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=" + filename);
            response.getOutputStream().write(bytes);
            response.getOutputStream().flush();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}

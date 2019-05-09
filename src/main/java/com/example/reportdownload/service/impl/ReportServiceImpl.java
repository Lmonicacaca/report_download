package com.example.reportdownload.service.impl;

import com.example.reportdownload.service.ReportService;
import com.example.reportdownload.util.DocxUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Tbl;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class ReportServiceImpl implements ReportService {
    private String fileName = "attacker_report_v2.2.docx";
    @Override
    public byte[] createReport() {
        try {
            WordprocessingMLPackage wordMLPackage = DocxUtils.loadWordprocessingMLPackageFromResources(fileName);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            Map<String, String> map = new HashMap<>();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日HH时mm分");

            // 0. 首页日期
            map.put("dateM", sdf.format(System.currentTimeMillis()).substring(0, 8));
            // 1. 摘要及简报
            map.put("attackerIp", "1.1.1.1");
            // 摘要信息
            map.put("LAN", "中国-浙江-杭州");
            map.put("longitude_latitude", "");
            map.put("operator", "移动");
            map.put("startdate", sdf.format(new Date()));
            map.put("enddate", sdf.format(new Date()));
            // 4.1.	情报信息
            this.threatInfo(wordMLPackage, "#maliciousIP");

            // 将变量集合加入文档对象中
            mainDocumentPart.variableReplace(map);

            ByteArrayOutputStream ops = new ByteArrayOutputStream();
            wordMLPackage.save(ops);
            return ops.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private void threatInfo(WordprocessingMLPackage wordMLPackage,String replaceText){
        List<String> headers = new ArrayList<>(2);
        headers.add("    " + "1.1.1.1" + "  情报名称");
        headers.add("");
        // 情报信息
        List<List<String>> data = new ArrayList<>();
        List<String> ipInfoList = new ArrayList<>();
        ipInfoList.add("IP地址：");
        ipInfoList.add("1.1.1.1");
        List<String> threatIntelligenceType = new ArrayList<>();
        threatIntelligenceType.add("情报类型：");
        threatIntelligenceType.add("恶意");
        data.add(threatIntelligenceType);

        List<String> threatIntelligenceLibName = new ArrayList<>();
        threatIntelligenceLibName.add("情报源：");
        threatIntelligenceLibName.add("安恒数据大脑");
        data.add(threatIntelligenceLibName);

        Tbl table = DocxUtils.createTreatInfoTemplateTable(wordMLPackage, headers, data);
        DocxUtils.replaceTable(wordMLPackage, replaceText, table);
    }
}

package com.example.reportdownload.util;

import com.example.reportdownload.util.RPType;
import com.example.reportdownload.util.ReplaceContent;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.images.FileConversionImageHandler;
import org.docx4j.model.properties.table.tr.TrHeight;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.time.TimeSeries;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;
import java.io.*;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;


/**
 * Created by delin on 2017/7/6.
 */
public class DocxUtils {

    private static final Logger LOG = LoggerFactory.getLogger(DocxUtils.class);
    private static final String THREAT_DBAPPSECURITY_URL = "https://threat.dbappsecurity.com.cn";
    private static final String THREAT_DBAPPSECURITY_LOGIN_URL = "https://threat.dbappsecurity.com.cn/login";

    private static ObjectFactory factory = Context.getWmlObjectFactory();
    private static String simSunPath = SystemProperUtil.getSysPath() + File.separator + "phantom" + File.separator + "simsun.ttc";
    private static String msyHlPath = SystemProperUtil.getSysPath() + File.separator + "phantom" + File.separator + "msyhl.ttc";
    public static String reportDemoPath = SystemProperUtil.getSysPath() + File.separator + "phantom" + File.separator + "report" + File.separator + "demo.html";
    public static String reportPngPath = SystemProperUtil.getSysPath() + File.separator + "phantom" + File.separator + "report" + File.separator + "demo.png";

    /**
     * 创建文档处理包对象
     *
     * @return 返回值：返回文档处理包对象
     * @throws Exception
     * @author delinz
     */
    public static WordprocessingMLPackage createWordprocessingMLPackage() throws Exception {
        return WordprocessingMLPackage.createPackage();
    }

    /**
     * 加载已有的文档包对象
     *
     * @param filepath 路径
     * @return
     * @throws Exception
     */
    public static WordprocessingMLPackage loadWordprocessingMLPackage(String filepath) throws Exception {
        WordprocessingMLPackage wMLPackage = WordprocessingMLPackage.load(new FileInputStream(new File(filepath)));
        return wMLPackage;
    }

    /**
     * 从resources下加载文档包对象
     *
     * @param filepath 资源包下路径
     * @return
     * @throws Exception
     */
    public static WordprocessingMLPackage loadWordprocessingMLPackageFromResources(String filepath) throws Exception {
        String sysPath = SystemProperUtil.getSysPath();
        WordprocessingMLPackage wMLPackage = loadWordprocessingMLPackage(sysPath + "/" + filepath);
        return wMLPackage;
    }

    /**
     * 从resources下加载文档包对象
     *
     * @param filepath 资源包下路径
     * @return
     * @throws Exception
     */
    public static WordprocessingMLPackage loadWordprocessingMLPackageFromResources$(String filepath) throws Exception {
        String sysPath = SystemProperUtil.getSysPath();
        WordprocessingMLPackage wMLPackage = loadWordprocessingMLPackage(sysPath + "/" + filepath);
        //此句必加，word本身会用标签分割文本变量,用VariablePrepare可修复
        VariablePrepare.prepare(wMLPackage);
        return wMLPackage;
    }

    public static WordprocessingMLPackage loadWordprocessingMLPackageFromResourcesTest(String filepath) throws Exception {
        WordprocessingMLPackage wMLPackage = loadWordprocessingMLPackage(filepath);
        //此句必加，word本身会用标签分割文本变量,用VariablePrepare可修复
//        VariablePrepare.prepare(wMLPackage);
        return wMLPackage;
    }


    /**
     * 获取文档的可用宽度
     *
     * @param wordPackage 文档处理包对象
     * @return 返回值：返回值文档的可用宽度
     * @throws Exception
     * @author delinz
     */
    private static int getWritableWidth(WordprocessingMLPackage wordPackage) throws Exception {
        return wordPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
    }

    /**
     * 保存文档信息
     *
     * @param wordPackage 文档处理包对象
     * @param fileName    完整的输出文件名称，包括路径
     * @throws Exception
     * @author delinz
     */
    public static void saveWordPackage(WordprocessingMLPackage wordPackage, String fileName) throws Exception {
        saveWordPackage(wordPackage, new File(fileName));
    }

    public static void saveWordPackage(WordprocessingMLPackage wordPackage, OutputStream outputStream) throws Docx4JException {
        wordPackage.save(outputStream);
    }

    /**
     * 保存文档信息
     *
     * @param wordPackage 文档处理包对象
     * @param file        文件
     * @throws Exception
     * @author delinz
     */
    public static void saveWordPackage(WordprocessingMLPackage wordPackage, File file) throws Exception {
        wordPackage.save(file);
    }

    public static File saveToFile(WordprocessingMLPackage wordPackage, String filename) throws Exception {
        File file = new File(filename);
        wordPackage.save(file);
        return file;
    }

    /**
     * 遍历所有的Text，或者Table，或者R或者P等等
     *
     * @param obj
     * @param toSearch
     * @return
     */
    public static List getAllElementFromObject(Object obj, Class<?> toSearch) {
        List result = new ArrayList();
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }
        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    /**
     * 文本替换
     *
     * @param mlPackage
     * @param replace
     * @param str
     * @throws Exception
     */
    public static void replaceText(WordprocessingMLPackage mlPackage, String replace, String str) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                content.setValue(str);
            }
        }
    }

    /**
     * 文本整体替换-B2-第二部分
     *
     * @throws Exception
     */
    public static void replaceTextMap(WordprocessingMLPackage mlPackage, Map<String, Object> map) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (map.containsKey(content.getValue())) {
                if (map.get(content.getValue()) != null) {
                    content.setValue(map.get(content.getValue()).toString());
                } else {
                    R r = (R) content.getParent();
                    P p = (P) r.getParent();
                    ContentAccessor doc = (ContentAccessor) p.getParent();
                    doc.getContent().remove(p);
                }
            }
        }
    }

    /**
     * 文本整体替换-B2-第二部分
     *
     * @throws Exception
     */
    public static void replaceContainTextMap(WordprocessingMLPackage mlPackage, Map<String, Object> map) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            for (String key : map.keySet()) {
                if (content.getValue().contains(key)) {
                    content.setValue(content.getValue().replace(key, (String) map.get(key)));
                }
            }
        }
    }

    /**
     * 文本内容部分替换
     *
     * @param mlPackage
     * @param replace
     * @param str
     * @throws Exception
     */
    public static void replaceTextPart(WordprocessingMLPackage mlPackage, String replace, String str) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            String value = content.getValue();
            if (value.contains(replace)) {
                content.setValue(value.replace(replace, str));
            } else {
                String partRight = replace.substring(1, replace.length());
                String partLeft = replace.substring(0, replace.length() - 1);
                String part = replace.substring(1, replace.length() - 1); //避免将非目标字符替换成特定内容，部分内容会被分开，如name,srcUserName
                if (value.contains(partRight)) {
                    content.setValue(value.replace(partRight, str));
                } else {
                    if (value.contains(partLeft)) {
                        content.setValue(value.replace(partLeft, str));
                    } else {
                        if (value.contains(part)) {
                            content.setValue(value.replace(part, str));
                        }
                    }
                }
            }
        }
    }


    /**
     * 文本内容Map部分替换-B1-取证
     * DocxUtils.replaceTextPart(evidenceEntityMDP, "{" + key + "}", resultMap.get(key).toString());
     *
     * @param mlPackage
     * @param map
     * @param
     * @throws Exception
     */
    public static void replaceTextPartMap(WordprocessingMLPackage mlPackage, Map<String, Object> map) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            String value = content.getValue();
            mapKeyInStr(map, value, content);
        }
    }
    /*
    {}不能离开，避免将正常字符替换成目标字段,name和srcUserName会被单独识别，part有必要。
     */

    private static void mapKeyInStr(Map<String, Object> map, String value, Text content) {
        if (map.containsKey(value)) {
            content.setValue((String) map.get(value));
        } else {
            for (String key : map.keySet()) {
                String partRight = key.substring(1);
                String partLeft = key.substring(0, key.length() - 1);
                String part = key.substring(1, key.length() - 1); //避免将非目标字符替换成特定内容，部分内容会被分开，如name,srcUserName
                if (value.contains(partRight)) {
                    content.setValue(value.replace(partRight, (String) map.get(key)));
                } else {
                    if (value.contains(partLeft)) {
                        content.setValue(value.replace(partLeft, (String) map.get(key)));
                    } else {
                        if (value.contains(part)) {
                            content.setValue(value.replace(part, (String) map.get(key)));
                        } else {
                            R r = (R) content.getParent();
                            P p = (P) r.getParent();
                            List<Object> pList = p.getContent();
                            StringBuilder pStr = new StringBuilder();
                            for (Object i : pList
                            ) {
                                System.out.println(i.getClass().toString());
                                if (i.getClass().toString().contains("org.docx4j.wml.R")) {
                                    R r1 = (R) i;
                                    List<Object> r1List = r1.getContent();
                                    for (Object j : r1List) {
                                        JAXBElement itext = (JAXBElement) j;
                                        pStr.append(((Text) itext.getValue()).getValue());
                                    }
                                }
                            }
                            if (key.equals(pStr.toString())) {
                                Text nTxt = new Text();
                                nTxt.setValue((String) map.get(key));
                                R nr = (R) pList.get(0);
                                ((JAXBElement) (nr.getContent().get(0))).setValue(nTxt);
                                for (int k = 1; k < pList.size(); k++) {
                                    pList.remove(k);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * @param mlPackage
     * @param replace
     * @throws Exception
     */
    public static void clearPContent(WordprocessingMLPackage mlPackage, String replace) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                doc.getContent().remove(p);
            }
        }
    }

    public static void replaceParagraph(WordprocessingMLPackage mlPackage, String replace,
                                        List<Map<String, List<Object>>> contents, String numid) {
        replaceParagraph(mlPackage, replace, contents, "0", numid);
    }

    public static void replaceParagraph(WordprocessingMLPackage mlPackage, String replace,
                                        List<Map<String, List<Object>>> contents) {
        replaceParagraph(mlPackage, replace, contents, "0", "5");
    }

    /**
     * 替换关键字到 新增段落
     *
     * @param mlPackage
     * @param replace   待替换占位符
     * @param contents  一个map一个段落,map类型，键为title，值为内容[内容可以为:String文字,Tbl表格,Drawing图片]。分别建立段落，不同样式。
     */
    public static void replaceParagraphV2(WordprocessingMLPackage mlPackage, String replace,
                                          List<Map<String, List<Object>>> contents) {

        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                for (Map<String, List<Object>> lineMap : contents) {

                    for (Map.Entry<String, List<Object>> o : lineMap.entrySet()) {
                        //标题
                        P lp = factory.createP();
                        R lr = factory.createR();
                        lp.getContent().add(lr);
                        Text lt = factory.createText();
                        lt.setValue(converObjToStr(o.getKey(), ""));
                        lr.getContent().add(lt);
                        setParagraphIndInfo(lp, "0", "0", null, null, null, null, "0", "0");
                        setParagraphSpacing(lp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                        doc.getContent().add(doc.getContent().indexOf(p), lp);

                        List<Object> items = o.getValue();
                        for (Object item : items) {
                            if (item instanceof String) {//文字 String
                                P llp = factory.createP();
                                R llr = factory.createR();
                                llp.getContent().add(llr);
                                Text llt = factory.createText();
                                llt.setValue((String) item);
                                llr.getContent().add(llt);
                                setParagraphIndInfo(llp, "420", "200", null, null, null, null, "0", "0");
                                setParagraphSpacing(llp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                                item = llp;
                            }
                            if (item instanceof Text) {
                                P llp = factory.createP();
                                R llr = factory.createR();
                                llp.getContent().add(llr);
                                llr.getContent().add(item);
                                setParagraphIndInfo(llp, "420", "200", null, null, null, null, "0", "0");
                                setParagraphSpacing(llp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                                item = llp;

                            }
                            if (item instanceof Drawing) {//图片 Drawing
                                P llp = factory.createP();
                                R llr = factory.createR();
                                llr.getContent().add(item);
                                item = llp;
                            }
                            //table - Tbl
                            doc.getContent().add(doc.getContent().indexOf(p), item);
                        }
                    }
                }
            }
        }
    }


    /**
     * 替换关键字到 新增段落
     *
     * @param mlPackage
     * @param replace   待替换占位符
     * @param contents  一个map一个段落,map类型，键为title，值为内容[内容可以为:String文字,Tbl表格,Drawing图片]。分别建立段落，不同样式，title前加自动序号。
     * @param numid     序号id
     */
    public static void replaceParagraph(WordprocessingMLPackage mlPackage, String replace,
                                        List<Map<String, List<Object>>> contents, String ilvlStr, String numid) {
        //
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                for (Map<String, List<Object>> lineMap : contents) {

                    for (Map.Entry<String, List<Object>> o : lineMap.entrySet()) {
                        //标题
                        P lp = factory.createP();
                        R lr = factory.createR();
                        lp.getContent().add(lr);
                        Text lt = factory.createText();
                        lt.setValue(o.getKey());
                        lr.getContent().add(lt);
                        setParagraphNum(lp, ilvlStr, numid);
                        setParagraphIndInfo(lp, "0", "0", null, null, null, null, "0", "0");
                        setParagraphSpacing(lp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                        doc.getContent().add(doc.getContent().indexOf(p), lp);

                        List<Object> items = o.getValue();
                        for (Object item : items) {

                            //
                            if (item instanceof String) {//文字 String
                                P llp = factory.createP();
                                R llr = factory.createR();
                                llp.getContent().add(llr);
                                Text llt = factory.createText();
                                llt.setValue((String) item);
                                llr.getContent().add(llt);
                                setParagraphIndInfo(llp, "420", "200", null, null, null, null, "0", "0");
                                setParagraphSpacing(llp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                                item = llp;
                            }
                            if (item instanceof Drawing) {//图片 Drawing
                                P llp = factory.createP();
                                R llr = factory.createR();
                                llr.getContent().add(item);
                                item = llp;
                            }
                            //table - Tbl
                            doc.getContent().add(doc.getContent().indexOf(p), item);
                        }


//                        P llp = factory.createP();
//                        R llr = factory.createR();
//                        llp.getContent().add(llr);
//                        Text llt = factory.createText();
//                        llt.setValue(o.getValue());
//                        llr.getContent().add(llt);
//                        setParagraphIndInfo(llp, "420", "200", null, null, null, null, "0", "0");
//                        setParagraphSpacing(llp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
//                        doc.getContent().add(doc.getContent().indexOf(p), llp);
                    }
                }
            }
        }
    }

    /**
     * 替换关键字到 新增段落
     *
     * @param mlPackage
     * @param replace   待替换占位符
     * @param lines     map类型，键为title，值为内容。分别建立段落，不同样式，title前加自动序号。
     */
    public static void replaceTextToParagraph(WordprocessingMLPackage mlPackage, String replace,
                                              List<Map<String, String>> lines) {
        replaceTextToParagraph(mlPackage, replace, lines, "0", "7");
    }

    public static void replaceTextToParagraph(WordprocessingMLPackage mlPackage, String replace,
                                              List<Map<String, String>> lines, String ilvlStr, String id) {
        NumberingDefinitionsPart ndp = null;
        try {
            ndp = new NumberingDefinitionsPart();
            mlPackage.getMainDocumentPart().addTargetPart(ndp);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        long newNumId = ndp.restart(7, 0, 1);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();

                if (lines.isEmpty()) {
                    doc.getContent().add(getPText("GOOD JOB！无需处置建议，继续保持！"));
                }
                for (Map<String, String> lineMap : lines) {
                    for (Map.Entry<String, String> o : lineMap.entrySet()) {
                        P lp = factory.createP();
                        R lr = factory.createR();
                        lp.getContent().add(lr);
                        Text lt = factory.createText();
                        lt.setValue(o.getKey());
                        lr.getContent().add(lt);
                        setParagraphNum(lp, ilvlStr, newNumId + "");
                        setParagraphIndInfo(lp, "0", "0", null, "200", null, null, "0", "0");
                        setParagraphSpacing(lp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                        doc.getContent().add(doc.getContent().indexOf(p), lp);

                        if (StringUtils.isBlank(o.getValue())) {
                            continue;
                        }
                        P llp = factory.createP();
                        R llr = factory.createR();
                        llp.getContent().add(llr);
                        Text llt = factory.createText();

                        llt.setValue(o.getValue());
                        llr.getContent().add(llt);
                        setParagraphIndInfo(llp, "420", "200", null, null, null, null, "0", "0");
                        setParagraphSpacing(llp, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AUTO);
                        doc.getContent().add(doc.getContent().indexOf(p), llp);
                    }
                }
            }
        }
    }

    public static void addTable(WordprocessingMLPackage mlPackage, List th, List data, boolean isFirstColMerge) throws Exception {
        MainDocumentPart mp = mlPackage.getMainDocumentPart();
        ObjectFactory factory = Context.getWmlObjectFactory();

        int rowNum = data.size() + (th == null ? 0 : 1);
        int colsNum = th == null ? ((List) data.get(0)).size() : th.size();
        Tbl table = DocxUtils.createTable(mlPackage, rowNum, colsNum, th != null);
        DocxUtils.fillTableData(mlPackage, table, data, th, true, "宋体", "15", "ffffff", true, JcEnumeration.CENTER, "宋体", "15", "000000", false, JcEnumeration.CENTER);
        if (isFirstColMerge) {
            int mergeStart = 0;
            int mergeEnd = 0;
            String colValue = null;
            for (Object datum : data) {
                String colValue1 = (String) ((List) datum).get(0);
                if (colValue1.equals(colValue)) {
                    mergeEnd++;
                } else {
                    if (colValue != null) {
                        if (mergeStart < mergeEnd) {
                            mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                        }
                    }
                    colValue = colValue1;
                    mergeEnd++;
                    mergeStart = mergeEnd;
                }
            }
            if (mergeEnd > mergeStart) {
                mergeCellsVertically(table, 0, mergeStart, mergeEnd);
            }
        }
        mlPackage.getMainDocumentPart().addObject(table);
    }


    public static void replaceTable(WordprocessingMLPackage mlPackage, String replace, List th, List data) throws Exception {
        replaceTable(mlPackage, replace, th, data, false);
    }

    public static void replaceTableDT(WordprocessingMLPackage mlPackage, String replace, List th, List data, boolean isDiffColor) throws Exception {
        replaceTableDT(mlPackage, replace, th, data, false, isDiffColor);
    }

    /**
     * 报告表格前两列合并
     *
     * @param mlPackage
     * @param replace
     * @param th
     * @param data
     */
    public static void replaceTableM2(WordprocessingMLPackage mlPackage, String replace, List th, List data, boolean isFirstColMerge, boolean isSecondColMerge) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                int rowNum = data.size() + (th == null ? 0 : 1);
                int colsNum = th == null ? ((List) data.get(0)).size() : th.size();
                Tbl table = DocxUtils.createTableDT(mlPackage, rowNum, colsNum, th != null);
                DocxUtils.fillTableDataDT(mlPackage, table, data, th, false);
                if (isFirstColMerge) {
                    int mergeStart = 0;
                    int mergeEnd = 0;
                    String colValue = null;
                    for (Object datum : data) {
                        String colValue1 = (String) ((List) datum).get(0);
                        if (colValue1.equals(colValue)) {
                            mergeEnd++;
                        } else {
                            if (colValue != null) {
                                if (mergeStart < mergeEnd) {
                                    mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                                }
                            }
                            colValue = colValue1;
                            mergeEnd++;
                            mergeStart = mergeEnd;
                        }
                    }
                    if (mergeEnd > mergeStart) {
                        mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                    }
                }
                if (isSecondColMerge) {
                    int mergeStart = 0;
                    int mergeEnd = 0;
                    String colValue = null;
                    for (Object datum : data) {
                        String colValue1 = (String) ((List) datum).get(1);
                        if (colValue1.equals(colValue)) {
                            mergeEnd++;
                        } else {
                            if (colValue != null) {
                                if (mergeStart < mergeEnd) {
                                    mergeCellsVertically(table, 1, mergeStart, mergeEnd);
                                }
                            }
                            colValue = colValue1;
                            mergeEnd++;
                            mergeStart = mergeEnd;
                        }
                    }
                    if (mergeEnd > mergeStart) {
                        mergeCellsVertically(table, 1, mergeStart, mergeEnd);
                    }
                }


                ContentAccessor doc = (ContentAccessor) p.getParent();
                doc.getContent().set(doc.getContent().indexOf(p), table);
            }
        }
    }

    /**
     * 把关键词替换成表格
     *
     * @param mlPackage
     * @param replace
     * @param th
     * @param data
     */
    public static void replaceTable(WordprocessingMLPackage mlPackage, String replace, List th, List data, boolean isFirstColMerge) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                int rowNum = data.size() + (th == null ? 0 : 1);
                int colsNum = th == null ? ((List) data.get(0)).size() : th.size();
                Tbl table = DocxUtils.createTable(mlPackage, rowNum, colsNum, th != null);
                DocxUtils.fillTableData(mlPackage, table, data, th);
                if (isFirstColMerge) {
                    int mergeStart = 0;
                    int mergeEnd = 0;
                    String colValue = null;
                    for (Object datum : data) {
                        String colValue1 = (String) ((List) datum).get(0);
                        if (colValue1.equals(colValue)) {
                            mergeEnd++;
                        } else {
                            if (colValue != null) {
                                if (mergeStart < mergeEnd) {
                                    mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                                }
                            }
                            colValue = colValue1;
                            mergeEnd++;
                            mergeStart = mergeEnd;
                        }
                    }
                    if (mergeEnd > mergeStart) {
                        mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                    }
                }
                ContentAccessor doc = (ContentAccessor) p.getParent();
                doc.getContent().set(doc.getContent().indexOf(p), table);
            }
        }
    }

    /**
     * 把关键词替换成威胁报告表格
     *
     * @param mlPackage
     * @param replace
     * @param th
     * @param data
     */
    public static void replaceTableDT(WordprocessingMLPackage mlPackage, String replace, List th, List data, boolean isFirstColMerge, boolean isDiffColor) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                int rowNum = data.size() + (th == null ? 0 : 1);
                int colsNum = th == null ? ((List) data.get(0)).size() : th.size();
                Tbl table = DocxUtils.createTableDT(mlPackage, rowNum, colsNum, th != null);
                DocxUtils.fillTableDataDT(mlPackage, table, data, th, isDiffColor);
                if (isFirstColMerge) {
                    int mergeStart = 0;
                    int mergeEnd = 0;
                    String colValue = null;
                    for (Object datum : data) {
                        String colValue1 = (String) ((List) datum).get(0);
                        if (colValue1.equals(colValue)) {
                            mergeEnd++;
                        } else {
                            if (colValue != null) {
                                if (mergeStart < mergeEnd) {
                                    mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                                }
                            }
                            colValue = colValue1;
                            mergeEnd++;
                            mergeStart = mergeEnd;
                        }
                    }
                    if (mergeEnd > mergeStart) {
                        mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                    }
                }
                ContentAccessor doc = (ContentAccessor) p.getParent();
                doc.getContent().set(doc.getContent().indexOf(p), table);
            }
        }
    }


    public static void replaceTable(WordprocessingMLPackage mlPackage, Text content, List th, List list) {
        R r = (R) content.getParent();
        r.getContent().clear();
        P p = (P) r.getParent();
        Tbl table = createTable(mlPackage, th, list);
        ContentAccessor doc = (ContentAccessor) p.getParent();
        doc.getContent().set(doc.getContent().indexOf(p), table);
    }

    public static void replaceMultiTable(WordprocessingMLPackage mlPackage, String replace, List th,
                                         List<Map<String, Object>> list) throws Exception {
//        Set<Streing> ipSet = new HashSet<>();
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                // List data = (List) (list.get(0).get("table"));
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                for (int i = 0; i < list.size(); i++) {
                    Map<String, Object> map = list.get(i);
                    if (map.get("ip") != null) {
                        doc.getContent().add(doc.getContent().indexOf(p), factory.createP());
                        String ip = (String) map.get("ip");
                        P ipP = getPText(ip);
                        setParagraphNum(ipP, "0", "10");
                        setParagraphIndInfo(ipP, "0", null, null, null, null, null, "0", "0");
                        setParagraphSpacing(ipP, JcEnumeration.LEFT, false, null, null, "1", "1", false, "200", STLineSpacingRule.AT_LEAST);
                        doc.getContent().add(doc.getContent().indexOf(p), ipP);

                        P baseP = getPText("基本信息");
                        doc.getContent().add(doc.getContent().indexOf(p), baseP);

                        List data = (List) map.get("table");
                        List title = (List) map.get("title");
                        Tbl table = DocxUtils.createTable(mlPackage, data.size() + 1, title.size());
                        DocxUtils.fillTableData(mlPackage, table, data, title);
                        doc.getContent().add(doc.getContent().indexOf(p), table);
                    } else {
                        String title = (String) map.get("title");
                        P lp = factory.createP();
                        R lr = factory.createR();
                        Text lt = factory.createText();
                        lt.setValue(title);
                        lr.getContent().add(lt);
                        lp.getContent().add(lr);
                        List data = (List) (map.get("table"));
                        Tbl table = DocxUtils.createTable(mlPackage, data.size() + 1, th.size());
                        DocxUtils.fillTableData(mlPackage, table, data, th);
                        doc.getContent().add(doc.getContent().indexOf(p), lp);
                        doc.getContent().add(doc.getContent().indexOf(p), table);
                    }
                }
            }
        }
    }

    public static void replaceImage(WordprocessingMLPackage wordMLPackage,
                                    String replace,
                                    byte[] bytes,
                                    String filenameHint, String altText) throws Exception {
        List texts = getAllElementFromObject(wordMLPackage.getMainDocumentPart(), Text.class);
        for (Object text : texts) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                p.getContent().add(newImage(wordMLPackage, bytes, filenameHint, altText));
            }
        }

    }

    public static void replaceBarPic(WordprocessingMLPackage wordMLPackage, String placeholder, List<Map<String, Object>> attackerList, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarDcd(attackerList, legend);
        byte[] bytes = ChartUtils.createBarChart(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceBarPicDT(WordprocessingMLPackage wordMLPackage, String placeholder, List<Map<String, Object>> attackerList, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarDcd(attackerList, legend);
        byte[] bytes = ChartUtils.createBarChartDT(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceBarPicHorizental(WordprocessingMLPackage wordMLPackage, String placeholder, List<Map<String, Object>> attackerList, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarDcd(attackerList, legend);
        byte[] bytes = ChartUtils.createBarChart(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceBarPic(WordprocessingMLPackage wordMLPackage, String placeholder, List<Map<String, Object>> attackerList, String name, String value, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarDcd(attackerList, legend, name, value);
        byte[] bytes = ChartUtils.createBarChart(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replacePiePic(WordprocessingMLPackage wordMLPackage, String placeholder, Map<String, Integer> attDomain, String title, int width, int height) throws Exception {
        DefaultPieDataset pieDataset = ChartUtils.createDefaultPieDataset(attDomain);
        byte[] bytes = ChartUtils.createPieChart(title, pieDataset, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replacePiePicForLong(WordprocessingMLPackage wordMLPackage, String placeholder, Map<String, Long> valueMap, String title, int width, int height) throws Exception {
        DefaultPieDataset pieDataset = ChartUtils.createDefaultPieDatasetForLong(valueMap);
        byte[] bytes = ChartUtils.createPieChart(title, pieDataset, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceLinePic(WordprocessingMLPackage wordMLPackage, String placeholder, List<Map<String, Object>> attackerList, String name, String value, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarDcd(attackerList, legend, name, value);
//        TimeSeriesCollection tsc = ChartUtils.formatTimeLineTsc(attackerList, legend, name, value);
        byte[] bytes = ChartUtils.createLineChart(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
//        DateTickUnit dateTickUnit = new DateTickUnit(DateTickUnitType.HOUR, 12, new SimpleDateFormat("MM-dd HH:mm"));
//        byte[] bytes = ChartUtils.createLineChartForDateX(title, categoygAxisLable, valueAxisLable, dcd, width, height,true, dateTickUnit);

        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceLinePic(WordprocessingMLPackage wordMLPackage, String placeholder, Map<String, List<Map<String, Object>>> map, String name, String value, String title, String categoygAxisLable, String valueAxisLable, String legend, CategoryLabelPositions labelPositions, int width, int height) throws Exception {
        DefaultCategoryDataset dcd = ChartUtils.formatBarsDcd(map, name, value);
//        TimeSeriesCollection tsc = ChartUtils.formatTimeLineTsc(attackerList, legend, name, value);
        byte[] bytes = ChartUtils.createLineChart(title, categoygAxisLable, valueAxisLable, labelPositions, dcd, width, height);
//        DateTickUnit dateTickUnit = new DateTickUnit(DateTickUnitType.HOUR, 12, new SimpleDateFormat("MM-dd HH:mm"));
//        byte[] bytes = ChartUtils.createLineChartForDateX(title, categoygAxisLable, valueAxisLable, dcd, width, height,true, dateTickUnit);

        DocxUtils.replaceImage(wordMLPackage, placeholder, bytes, "pic", "pic");
    }

    public static void replaceTimeSeriesPic(WordprocessingMLPackage wordMLPackage, String placeHolder, String catalogy, ArrayList<Object[]> dateValues, String timeAxisLabel, String title, String valueAxisLabel, int width, int height, int dateType) throws Exception {
        TimeSeries timeSeries = ChartUtils.createTimeseries(catalogy, dateValues);
        byte[] bytes = ChartUtils.createTimeSeriesChart(title, timeAxisLabel, valueAxisLabel, timeSeries, width, height, dateType);
        DocxUtils.replaceImage(wordMLPackage, placeHolder, bytes, "pic", "pic");
    }

    /**
     * //曲线图
     *
     * @param wordMLPackage
     * @param placeHolder   替换字符串
     * @param xAxis         [1,2,3]
     * @param data          {series:[1,2,3]} key曲线名
     * @param title
     * @param xAxisLabel
     * @param yAxisLabel
     * @param width
     * @param height
     * @throws Exception
     */
    public static void replaceLinePic(WordprocessingMLPackage wordMLPackage, String placeHolder,
                                      List<String> xAxis, Map<String, List<Number>> data, String title,
                                      String xAxisLabel, String yAxisLabel, CategoryLabelPositions labelPositions,
                                      int width, int height) throws Exception {

        DefaultCategoryDataset dataset = ChartUtils.createLinePicDataset(xAxis, data);
        byte[] chart = ChartUtils.createLineChart(dataset, title, xAxisLabel, yAxisLabel, labelPositions, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeHolder, chart, "pic", "pic");
    }

    public static void replaceLinePic(WordprocessingMLPackage wordMLPackage, String placeHolder,
                                      List<String> xAxis, Map<String, List<Number>> data, String title,
                                      String xAxisLabel, String yAxisLabel, int width, int height)
            throws Exception {
        replaceLinePic(wordMLPackage, placeHolder, xAxis, data, title, xAxisLabel, yAxisLabel, null, width, height);
    }

    /**
     * @param wordMLPackage
     * @param placeHolder            替换字符串
     * @param chartTitle             标题
     * @param xAxis                  x轴标签
     * @param yAxis                  y轴标签
     * @param rowKeys                图例legend["1","2"]
     * @param columnKeys             相当于x轴坐标["北京", "上海", "广州", "成都", "深圳"]
     * @param data                   二维数组,x轴对应的y轴数据[[672, 766, 223, 540, 126],[325, 521, 210, 340, 106]]
     * @param categoryLabelPositions x坐标值倾斜角度
     * @param width
     * @param height
     * @throws Exception /**
     *                   |
     *                   data  | [[1,  2,  3],  [A, ->rowKeys
     *                   |  [4,  5,  6],   B,
     *                   |  [7,  8,  9]]   C]
     *                   -------------------->
     *                   [a,  b,  c] -> colKeys
     */
    public static void replaceStackedBarChart(WordprocessingMLPackage wordMLPackage, String placeHolder,
                                              String chartTitle, String xAxis, String yAxis, List<String> rowKeys,
                                              List<String> columnKeys, List<List<Number>> data,
                                              CategoryLabelPositions categoryLabelPositions,
                                              int width, int height) throws Exception {
        CategoryDataset dataset = ChartUtils.createCategoryDataset(rowKeys, columnKeys, data);
        byte[] chart = ChartUtils.createStackedBarChart(chartTitle, xAxis, yAxis, dataset, categoryLabelPositions, width, height);
        DocxUtils.replaceImage(wordMLPackage, placeHolder, chart, "pic", "pic");
    }

    public static void replaceStackedBarChart(WordprocessingMLPackage wordMLPackage, String placeHolder,
                                              String chartTitle, String xAxis, String yAxis,
                                              List<String> rowKeys, List<String> columnKeys,
                                              List<List<Number>> data, int width, int height) throws Exception {
        replaceStackedBarChart(wordMLPackage, placeHolder, chartTitle, xAxis, yAxis, rowKeys, columnKeys, data,
                null, width, height);
    }

    /**
     * 创建图片内容对象
     *
     * @param wordMLPackage
     * @param imagePath
     * @param filenameHint
     * @param altText
     * @return
     * @throws Exception
     */
    public static R newImage(WordprocessingMLPackage wordMLPackage,
                             String imagePath,
                             String filenameHint, String altText) throws Exception {
        InputStream is = new FileInputStream(imagePath);
        byte[] bytes = IOUtils.toByteArray(is);
        return newImage(wordMLPackage, bytes, filenameHint, altText);

    }

    /**
     * 创建图片内容对象
     *
     * @param wordMLPackage
     * @param bytes
     * @param filenameHint
     * @param altText
     * @return
     * @throws Exception
     */
    public static R newImage(WordprocessingMLPackage wordMLPackage,
                             byte[] bytes,
                             String filenameHint, String altText) throws Exception {
        int id1 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
        int id2 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline(filenameHint, altText,
                id1, id2, false);

// Now add the inline in w:p/w:r/w:drawing
        ObjectFactory factory = Context.getWmlObjectFactory();
        R run = factory.createR();
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return run;

    }

    /**
     * @param wordMLPackage
     * @param type          Title
     *                      Subtitle
     *                      Heading1
     *                      Heading2
     *                      Heading3
     *                      Normal
     * @param text
     * @throws Exception
     */
    public static void addsStyleedParagraph(WordprocessingMLPackage wordMLPackage, String type, String text) throws Exception {
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText(type, text);
    }

    public static void alterStyleSheet(WordprocessingMLPackage wordMLPackage) {
        StyleDefinitionsPart styleDefinitionsPart =
                wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart();
        Styles styles = styleDefinitionsPart.getJaxbElement();
        List<Style> stylesList = styles.getStyle();
        for (Style style : stylesList) {
            alterFontStyleToSimSun(style);
            if (style.getStyleId().equals("Normal")) {
                alterNormalStyle(style);
            } else if (style.getStyleId().equals("Title")) {
                alterTitleStyle(style);
            } else {

            }
        }
    }

    public static void alterDTStyleSheet(WordprocessingMLPackage wordMLPackage) {
        StyleDefinitionsPart styleDefinitionsPart =
                wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart();
        Styles styles = styleDefinitionsPart.getJaxbElement();
        List<Style> stylesList = styles.getStyle();
        for (Style style : stylesList) {
            alterFontStyleToMsyHl(style);
            if (style.getStyleId().equals("Normal")) {
                alterNormalStyle(style);
            } else if (style.getStyleId().equals("Title")) {
                alterTitleStyle(style);
            } else {

            }
        }
    }

    public static void replaceTOC(WordprocessingMLPackage mlPackage, String replace) throws Exception {
        List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            String value = content.getValue();
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                P paragraph = mlPackage.getMainDocumentPart().createStyledParagraphOfText("TOC", "目录");
                addTableOfContent(paragraph);
                doc.getContent().add(doc.getContent().indexOf(p), paragraph);
                break;
            }
        }
    }

    public static void addTableOfContent(WordprocessingMLPackage wordMLPackage) {
        ObjectFactory factory = Context.getWmlObjectFactory();

        P paragraph = factory.createP();

        addFieldBegin(paragraph);
        addTableOfContentField(paragraph);
        addFieldEnd(paragraph);
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        documentPart.getJaxbElement().getBody().getContent().add(paragraph);
    }

    private static void alterTitleStyle(Style style) {
        RPr rpr = style.getRPr();
        removeUnderline(rpr);
    }

    private static void removeUnderline(RPr runProperties) {
        U underline = new U();
        underline.setVal(UnderlineEnumeration.NONE);
        runProperties.setU(underline);
    }

    private static void changeFontSize(RPr runProperties, int fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(fontSize));
        runProperties.setSz(size);
    }

    private static void alterNormalStyle(Style style) {
        RPr rpr = (style.getRPr() == null) ? new RPr() : style.getRPr();
        changeFontSize(rpr, 20);
        style.setRPr(rpr);
    }

    private static void addTableOfContent(P paragraph) {

        addFieldBegin(paragraph);
        addTableOfContentField(paragraph);
        addFieldEnd(paragraph);
    }

    private static void addTableOfContentField(P paragraph) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        R run = factory.createR();
        Text txt = new Text();
        txt.setSpace("preserve");
        txt.setValue("TOC \\o \"1-3\" \\h \\z \\u");
        run.getContent().add(factory.createRInstrText(txt));
        paragraph.getContent().add(run);
    }

    /**
     * 每个域都需要用复杂的域字符来确定界限. 本方法向给定段落添加在真正域之前的界定符.
     * <p>
     * 再一次以创建一个可运行块开始, 然后创建一个域字符来标记域的起始并标记域是'脏的'因为我们想要
     * 在整个文档生成之后进行内容更新.
     * 最后将域字符转换成JAXB元素并将其添加到可运行块, 然后将可运行块添加到段落中.
     *
     * @param paragraph
     */
    private static void addFieldBegin(P paragraph) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        R run = factory.createR();
        FldChar fldchar = factory.createFldChar();
        fldchar.setFldCharType(STFldCharType.BEGIN);
        fldchar.setDirty(true);
        run.getContent().add(getWrappedFldChar(fldchar));
        paragraph.getContent().add(run);
    }

    public static JAXBElement getWrappedFldChar(FldChar fldchar) {
        return new JAXBElement(new QName(Namespaces.NS_WORD12, "fldChar"), FldChar.class, fldchar);
    }

    /**
     * 每个域都需要用复杂的域字符来确定界限. 本方法向给定段落添加在真正域之后的界定符.
     * <p>
     * 跟前面一样, 从创建可运行块开始, 然后创建域字符标记域的结束, 最后将域字符转换成JAXB元素并
     * 将其添加到可运行块, 可运行块再添加到段落中.
     *
     * @param paragraph
     */
    private static void addFieldEnd(P paragraph) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        R run = factory.createR();
        FldChar fldcharend = factory.createFldChar();
        fldcharend.setFldCharType(STFldCharType.END);
        run.getContent().add(getWrappedFldChar(fldcharend));
        paragraph.getContent().add(run);
    }

    private static void alterFontStyleToSimSun(Style style) {
        RPr rpr = (style.getRPr() == null) ? new RPr() : style.getRPr();
        RFonts runFont = new RFonts();
        runFont.setAscii("宋体");
        runFont.setHAnsi("宋体");
        rpr.setRFonts(runFont);
        style.setRPr(rpr);
    }

    /*
     * 微软雅黑字体
     */
    private static void alterFontStyleToMsyHl(Style style) {
        RPr rpr = (style.getRPr() == null) ? new RPr() : style.getRPr();
        RFonts runFont = new RFonts();
        runFont.setAscii("微软雅黑");
        runFont.setHAnsi("微软雅黑");
        runFont.setEastAsia("微软雅黑");
        rpr.setRFonts(runFont);
        style.setRPr(rpr);
    }

    public static void addPageBreak(WordprocessingMLPackage wordMLPackage) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
        mp.getContent().add(paragraph);

        Br breakObj = new Br();
        breakObj.setType(STBrType.PAGE);
        paragraph.getContent().add(breakObj);
    }

    public static void appendStyleedParaRContent(WordprocessingMLPackage wordMLPackage, String type, String content) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P paragraph = factory.createP();
        MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
        mp.getContent().add(paragraph);
        RPr runProperties = factory.createRPr();
        if (content != null) {
            R run = new R();
            paragraph.getContent().add(run);
            String[] contentArr = content.split("\n");
            Text text = new Text();
            text.setSpace("preserve");
            text.setValue(contentArr[0]);
            run.setRPr(runProperties);
            run.getContent().add(text);

            for (int i = 1, len = contentArr.length; i < len; i++) {
                Br br = new Br();
                run.getContent().add(br);// 换行
                text = new Text();
                text.setSpace("preserve");
                text.setValue(contentArr[i]);
                run.setRPr(runProperties);
                run.getContent().add(text);
            }
            if (mp.getPropertyResolver().activateStyle(type)) {
                // Style is available
                PPr pPr = factory.createPPr();
                paragraph.setPPr(pPr);
                PPrBase.PStyle pStyle = factory.createPPrBasePStyle();
                pPr.setPStyle(pStyle);
                pStyle.setVal(type);
            }
        }
    }

    public static void addTextParagraph(WordprocessingMLPackage wordMLPackage, String text, String hpsMeasureSize) throws Exception {
        MainDocumentPart mp = wordMLPackage.getMainDocumentPart();
        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr contentRpr = getRPr("宋体", hpsMeasureSize, false);
        P paragraph = factory.createP();
        Text txt = factory.createText();
        txt.setValue(text);
        R run = factory.createR();
        run.getContent().add(txt);
        run.setRPr(contentRpr);
        paragraph.getContent().add(run);
        mp.getContent().add(paragraph);
    }

//    public static void toDTPDF(WordprocessingMLPackage wordPackage, ByteArrayOutputStream reportBytes, List paths) throws Exception {
//        Mapper fontMapper = new IdentityPlusMapper();
//        wordPackage.setFontMapper(fontMapper);
//
//        String fontFamily = "微软雅黑";
//        URL simSunUrl = new File(msyHlPath).toURI().toURL();
//        ; //加载字体文件（解决linux环境下无中文字体问题）
//        PhysicalFonts.addPhysicalFonts(fontFamily, simSunUrl);
//        PhysicalFont simSunFont = PhysicalFonts.get(fontFamily);
//        fontMapper.put(fontFamily, simSunFont);
//
//        RFonts rfonts = Context.getWmlObjectFactory().createRFonts(); //设置文件默认字体
//        rfonts.setAsciiTheme(null);
//        rfonts.setAscii(fontFamily);
//        rfonts.setHAnsi(fontFamily);
//        rfonts.setHint(STHint.EAST_ASIA);
//        wordPackage.getMainDocumentPart().getPropertyResolver()
//                .getDocumentDefaultRPr().setRFonts(rfonts);
//        FOSettings foSettings = Docx4J.createFOSettings();
//        foSettings.setWmlPackage(wordPackage);
//        foSettings.setApacheFopMime("application/pdf");
//        foSettings.setImageHandler(new FileConversionImageHandler(foSettings.getImageDirPath(), foSettings.isImageIncludeUUID()) {
//            @Override
//            protected String storeImage(BinaryPart binaryPart, byte[] bytes, File folder, String filename) throws Docx4JException {
//                String uri = null;
//                File imageFile = new File(folder, filename);
//                FileOutputStream out = null;
//
//                if (imageFile.exists()) {
//                    log.warn("Overwriting (!) existing file!");
//                }
//                try {
//                    out = new FileOutputStream(imageFile);
//                    out.write(bytes);
//
//                    // return the uri
//                    uri = setImageUri(imageFile);
//                    log.info("Wrote @src='" + uri);
//                } catch (IOException ioe) {
//                    throw new Docx4JException("Exception storing '" + filename + "', " + ioe.toString(), ioe);
//                } finally {
//                    try {
//                        out.close();
//                    } catch (IOException ioe) {
//                        ioe.printStackTrace();
//                    }
//                }
//                paths.add(uri);
//                return uri;
//            }
//
//            protected String setImageUri(File imageFile) {
//                try {
//                    return imageFile.toURI().toURL().toString();
//                } catch (MalformedURLException var3) {
//                    log.error(var3.getMessage(), var3);
//                    return imageFile.getName();
//                }
//            }
//        });
//        Docx4J.toFO(foSettings, reportBytes, Docx4J.FLAG_NONE);
//    }


//    public static void toPDF(WordprocessingMLPackage wordPackage, ByteArrayOutputStream reportBytes, List paths) throws Exception {
//        Mapper fontMapper = new IdentityPlusMapper();
//        wordPackage.setFontMapper(fontMapper);
//
//        String fontFamily = "宋体";
//        URL simSunUrl = new File(simSunPath).toURI().toURL();
//        ; //加载字体文件（解决linux环境下无中文字体问题）
//        PhysicalFonts.addPhysicalFonts(fontFamily, simSunUrl);
//        PhysicalFont simSunFont = PhysicalFonts.get(fontFamily);
//        fontMapper.put(fontFamily, simSunFont);
//
//        RFonts rfonts = Context.getWmlObjectFactory().createRFonts(); //设置文件默认字体
//        rfonts.setAsciiTheme(null);
//        rfonts.setAscii(fontFamily);
//        rfonts.setHAnsi(fontFamily);
//        rfonts.setHint(STHint.EAST_ASIA);
//        wordPackage.getMainDocumentPart().getPropertyResolver()
//                .getDocumentDefaultRPr().setRFonts(rfonts);
//        FOSettings foSettings = Docx4J.createFOSettings();
//        foSettings.setWmlPackage(wordPackage);
//        foSettings.setApacheFopMime("application/pdf");
//        foSettings.setImageHandler(new FileConversionImageHandler(foSettings.getImageDirPath(), foSettings.isImageIncludeUUID()) {
//            @Override
//            protected String storeImage(BinaryPart binaryPart, byte[] bytes, File folder, String filename) throws Docx4JException {
//                String uri = null;
//                File imageFile = new File(folder, filename);
//                FileOutputStream out = null;
//
//                if (imageFile.exists()) {
//                    log.warn("Overwriting (!) existing file!");
//                }
//                try {
//                    out = new FileOutputStream(imageFile);
//                    out.write(bytes);
//
//                    // return the uri
//                    uri = setImageUri(imageFile);
//                    log.info("Wrote @src='" + uri);
//                } catch (IOException ioe) {
//                    throw new Docx4JException("Exception storing '" + filename + "', " + ioe.toString(), ioe);
//                } finally {
//                    try {
//                        out.close();
//                    } catch (IOException ioe) {
//                        ioe.printStackTrace();
//                    }
//                }
//                paths.add(uri);
//                return uri;
//            }
//
//            protected String setImageUri(File imageFile) {
//                try {
//                    return imageFile.toURI().toURL().toString();
//                } catch (MalformedURLException var3) {
//                    log.error(var3.getMessage(), var3);
//                    return imageFile.getName();
//                }
//            }
//        });
//        Docx4J.toFO(foSettings, reportBytes, Docx4J.FLAG_NONE);
//    }

    public static void addImageparagraph(WordprocessingMLPackage wordMLPackage, String filePath, String filenameHint, String altText) throws Exception {
        File image = new File(filePath);
        if (image.exists()) {
            int id1 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
            int id2 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
            byte[] fileContent = Files.readAllBytes(image.toPath());
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, fileContent);
            Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
            ObjectFactory factory = Context.getWmlObjectFactory();
            P p = factory.createP();
            R r = factory.createR();
            p.getContent().add(r);
            Drawing drawing = factory.createDrawing();
            r.getContent().add(drawing);
            drawing.getAnchorOrInline().add(inline);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            mainDocumentPart.getContent().add(p);
        }
    }

    /**
     * 创建图片内容对象
     *
     * @param wordMLPackage
     * @param bytes
     * @param filenameHint
     * @param altText
     * @return
     * @throws Exception
     */
    public static Drawing drawingImage(WordprocessingMLPackage wordMLPackage,
                                       byte[] bytes,
                                       String filenameHint, String altText) throws Exception {
        int id1 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
        int id2 = (int) ((Math.random() * 1000) * (Math.random() * 1000));
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline(filenameHint, altText,
                id1, id2, false);

        // Now add the inline in w:p/w:r/w:drawing
        ObjectFactory factory = Context.getWmlObjectFactory();
        Drawing drawing = factory.createDrawing();
        drawing.getAnchorOrInline().add(inline);
        return drawing;
    }

    /**
     * 创建表格，有格式
     *
     * @param wordMLPackage
     * @param tableTh
     * @param data
     * @return
     */
    public static Tbl createTableFormat(WordprocessingMLPackage wordMLPackage, List<String> tableTh, List<List<String>> data) throws Exception {

        Tbl table = DocxUtils.createTable(wordMLPackage, data.size() + 1, tableTh.size());
        DocxUtils.fillTableData(wordMLPackage, table, data, tableTh);

        return table;
    }

    /**
     * 创建表格，无格式
     *
     * @param tableTh
     * @param data
     * @return
     */
    public static Tbl createTable(WordprocessingMLPackage wordMLPackage, List<String> tableTh, List<List<String>> data) {
        int writableWidthTwips =
                wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
        Tbl table =
                TblFactory.createTable(data.size() + 1, tableTh.size(), writableWidthTwips / tableTh.size());

        Tr headerRow = (Tr) table.getContent().get(0);
        for (int i = 0; i < tableTh.size(); i++) {
            Tc tc = (Tc) headerRow.getContent().get(i);
            tc.getContent().add(getPText(tableTh.get(i)));
        }

        for (int i = 1; i < data.size() + 1; i++) {
            Tr row = (Tr) table.getContent().get(i);
            for (int j = 0; j < data.get(i - 1).size(); j++) {
                Tc tc = (Tc) row.getContent().get(j);
                tc.getContent().add(getPText(converObjToStr(data.get(i - 1).get(j), "")));
            }
        }

        return table;
    }

    /**
     * 创建情报信息表格模版
     *
     * @param wordMLPackage docx
     * @param tableTh       表头: size=列数
     * @param data          数据
     * @return
     */
    public static Tbl createTreatInfoTemplateTable(WordprocessingMLPackage wordMLPackage, List<String> tableTh, List<List<String>> data) {

        int writableWidthTwips =
                wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
        Tbl table =
                TblFactory.createTable(data.size() + 1, tableTh.size(), writableWidthTwips / tableTh.size());
        Tr headerRow = (Tr) table.getContent().get(0);
        // 写入表头内容
        Tc headerCell = (Tc) headerRow.getContent().get(0);
        P headerP = getPText(tableTh.get(0));
        setCellContentStyle(headerP, JcEnumeration.LEFT);
        headerCell.getContent().remove(0);
        headerCell.getContent().add(headerP);
        // 边框样式
        CTBorder border = getCTBorder(1, "d2d2d2", "1", "0");
        // 表头下边框
        TcPrInner.TcBorders tcBorders = factory.createTcPrInnerTcBorders();
        tcBorders.setBottom(border);
        TcPr headerTcPr = factory.createTcPr();
        headerTcPr.setTcBorders(tcBorders);
        headerCell.setTcPr(headerTcPr);
        // 合并表头
        mergeCellsHorizontal(table, 0, 0, tableTh.size() - 1);
        // 写入表格数据
        for (int i = 1; i < data.size() + 1; i++) {
            Tr row = (Tr) table.getContent().get(i);
            for (int j = 0; j < data.get(i - 1).size(); j++) {
                Tc tc = (Tc) row.getContent().get(j);
                String originalContent = data.get(i - 1).get(j);
                String content = (StringUtils.isEmpty(originalContent) || "".equals(originalContent.trim())) ? "暂无数据" : originalContent;
                P contentP = getPText(content, "4a505e");
                if ("暂无数据".equals(content)) {
                    contentP = getPText(content, "d2d2d2");
                }
                BigInteger bigInteger;
                if (j == 0) {
                    setCellContentStyle(contentP, JcEnumeration.RIGHT);
                    bigInteger = BigInteger.valueOf(writableWidthTwips * 2 / 10);
                } else {
                    setCellContentStyle(contentP, JcEnumeration.LEFT);
                    bigInteger = BigInteger.valueOf(writableWidthTwips * 8 / 10);
                }
                TblWidth cellWidth = factory.createTblWidth();
                cellWidth.setType("dxa");
                cellWidth.setW(bigInteger);
                TcPr tcPr = factory.createTcPr();
                tcPr.setTcW(cellWidth);
                tc.setTcPr(tcPr);
                tc.getContent().remove(0);
                tc.getContent().add(contentP);
            }
        }
        // 表格外边框
        TblBorders tblBorders = factory.createTblBorders();
        tblBorders.setTop(border);
        tblBorders.setLeft(border);
        tblBorders.setRight(border);
        tblBorders.setBottom(border);
        TblPr tblPr = factory.createTblPr();
        tblPr.setTblBorders(tblBorders);
        table.setTblPr(tblPr);

        return table;
    }

    /**
     * 在P元素上塞入文本
     *
     * @param str
     * @return
     */
    public static P getPText(String str, String color) {
        Text text = factory.createText();
        text.setValue(str);
        R run = factory.createR();
        run.getContent().add(text);
        RPr rpR = getRPr("宋体", color, "20", STHint.EAST_ASIA, false);
        run.setRPr(rpR);
        P p = factory.createP();
        p.getContent().add(run);
        return p;
    }

    /**
     * 在P元素上塞入文本
     *
     * @param str
     * @return
     */
    public static P getPText(String str) {
        Text text = factory.createText();
        text.setValue(str);
        R run = factory.createR();
        run.getContent().add(text);
        P p = factory.createP();
        p.getContent().add(run);
        return p;
    }

    /**
     * 设置字体的样式
     *
     * @param fontFamily     字体类型
     * @param colorVal       字体颜色
     * @param hpsMeasureSize 字号大小
     * @param sTHint         字体格式
     * @param isBlod         是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author delinz
     */
    private static RPr getRPr(String fontFamily, String colorVal, String hpsMeasureSize, STHint sTHint, boolean isBlod) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr rPr = factory.createRPr();

        RFonts rf = new RFonts();
        rf.setHint(sTHint);
        rf.setAscii(fontFamily);
        rf.setHAnsi(fontFamily);
        rPr.setRFonts(rf);

        BooleanDefaultTrue bdt = Context.getWmlObjectFactory().createBooleanDefaultTrue();
        rPr.setBCs(bdt);
        if (isBlod) {
            rPr.setB(bdt);
        }
        Color color = new Color();
        color.setVal(colorVal);
        rPr.setColor(color);

        //背景
//        rPr.setBdr();


        HpsMeasure sz = new HpsMeasure();
        sz.setVal(new BigInteger(hpsMeasureSize));
        rPr.setSz(sz);
        rPr.setSzCs(sz);

        return rPr;
    }

    /**
     * 设置字体的样式，宋体，黑色，18号
     *
     * @param isBlod 是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author delinz
     */
    private static RPr getRPr(boolean isBlod) {
        return getRPr("宋体", "000000", "18", STHint.EAST_ASIA, isBlod);
    }

    /**
     * 设置字体的样式，黑色，18号
     *
     * @param fontFamily 字体
     * @param isBlod     是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author delinz
     */
    private static RPr getRPr(String fontFamily, boolean isBlod) {
        return getRPr(fontFamily, "000000", "18", STHint.EAST_ASIA, isBlod);
    }

    /**
     * 设置字体的样式，黑色
     *
     * @param fontFamily     字体
     * @param hpsMeasureSize 字号的大小
     * @param isBlod         是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author delinz
     */
    private static RPr getRPr(String fontFamily, String hpsMeasureSize, boolean isBlod) {
        return getRPr(fontFamily, "000000", hpsMeasureSize, STHint.EAST_ASIA, isBlod);
    }

    /**
     * 合并单元格
     * 表示合并第startRow（开始行）行中的第startCol（开始列）列到（startCol + colSpan - 1）列 </BR>
     * 表示合并第startCol（开始列）行中的第startRow（开始行）列到（startRow + rowSpan - 1）行
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param rowSpan    合并的行数，大于1才表示合并
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @param colSpan    合并的列数，大于1才表示合并
     * @author delinz
     */
    public static void setCellMerge(Tc tc, int currentRow, int startRow, int rowSpan, int currentCol, int startCol, int colSpan) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        TcPr tcpr = tc.getTcPr();
        if (tcpr == null) {
            tcpr = factory.createTcPr();
        }
        //表示合并列
        if (colSpan > 1) {
            //表示从第startRow行开始
            if (currentRow == startRow) {
                //表示从第startRow行的第startCol列开始合并，合并到第startCol + colSpan - 1列
                if (currentCol == startCol) {
                    TcPrInner.HMerge hm = factory.createTcPrInnerHMerge();
                    hm.setVal("restart");
                    tcpr.setHMerge(hm);
                    tc.setTcPr(tcpr);
                } else if (currentCol > startCol && currentCol <= (startCol + colSpan - 1)) {
                    TcPrInner.HMerge hm = factory.createTcPrInnerHMerge();
                    tcpr.setHMerge(hm);
                    tc.setTcPr(tcpr);
                }
            }
        }
        //表示合并行
        if (rowSpan > 1) {
            //表示从第startCol列开始
            if (currentCol == startCol) {
                //表示从第startCol列的第startRow行始合并，合并到第startRow + rowSpan - 1行
                if (currentRow == startRow) {
                    TcPrInner.VMerge vm = factory.createTcPrInnerVMerge();
                    vm.setVal("restart");
                    tcpr.setVMerge(vm);
                    tc.setTcPr(tcpr);
                } else if (currentRow > startRow && currentRow <= (startRow + rowSpan - 1)) {
                    TcPrInner.VMerge vm = factory.createTcPrInnerVMerge();
                    tcpr.setVMerge(vm);
                    tc.setTcPr(tcpr);
                }
            }
        }

    }

    /**
     * 合并单元格，相当于跨列的效果
     * 表示合并第startRow（开始行）行中的第startCol（开始列）列到（startCol + colSpan - 1）列 </BR>
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @param colSpan    合并的列数，大于1才表示合并
     * @author delinz
     */
    public static void setCellHMerge(Tc tc, int currentRow, int startRow, int currentCol, int startCol, int colSpan) {
        setCellMerge(tc, currentRow, startRow, 1, currentCol, startCol, colSpan);
    }

    /**
     * 合并单元格，相当于跨行的效果</BR>
     * 表示合并第startCol（开始列）行中的第startRow（开始行）列到（startRow + rowSpan - 1）行
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param rowSpan    合并的行数，大于1才表示合并
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @author delinz
     */
    public static void setCellVMerage(Tc tc, int currentRow, int startRow, int rowSpan, int currentCol, int startCol) {
        setCellMerge(tc, currentRow, startRow, rowSpan, currentCol, startCol, 1);
    }


    /**
     * 设置文档是否只读，包括内容和样式
     *
     * @param wordPackage 文档处理包对象
     * @param isReadOnly  是否只读
     * @throws Exception
     * @author delinz
     */
    public static void setReadOnly(WordprocessingMLPackage wordPackage, boolean isReadOnly) throws Exception {
        byte[] bt = "".getBytes();
        if (isReadOnly) {
            bt = "123456".getBytes();
        }
        ObjectFactory factory = Context.getWmlObjectFactory();
        //创建设置文档对象
        DocumentSettingsPart ds = wordPackage.getMainDocumentPart().getDocumentSettingsPart();
        if (ds == null) {
            ds = new DocumentSettingsPart();
        }
        CTSettings cs = ds.getJaxbElement();
        if (cs == null) {
            cs = factory.createCTSettings();
        }
        //创建文档保护对象
        CTDocProtect cp = cs.getDocumentProtection();
        if (cp == null) {
            cp = new CTDocProtect();
        }
        //设置加密方式
        cp.setCryptProviderType(STCryptProv.RSA_AES);
        cp.setCryptAlgorithmClass(STAlgClass.HASH);
        //设置任何用户
        cp.setCryptAlgorithmType(STAlgType.TYPE_ANY);
        cp.setCryptAlgorithmSid(new BigInteger("4"));
        cp.setCryptSpinCount(new BigInteger("50000"));
        //只读
        if (isReadOnly) {
            cp.setEdit(STDocProtect.READ_ONLY);
            cp.setHash(bt);
            cp.setSalt(bt);
            //设置内容不可编辑
            cp.setEnforcement(true);
            //设置格式不可编辑
            cp.setFormatting(true);
        } else {
            cp.setEdit(STDocProtect.NONE);
            cp.setHash(null);
            cp.setSalt(null);
            //设置内容不可编辑
            cp.setEnforcement(false);
            //设置格式不可编辑
            cp.setFormatting(false);
        }

        cs.setDocumentProtection(cp);
        ds.setJaxbElement(cs);
        //添加到文档主体中
        wordPackage.getMainDocumentPart().addTargetPart(ds);
    }

    /**
     * 设置文档是否只读，包括内容和样式
     *
     * @param fileName   文件
     * @param isReadOnly 是否只读
     * @return 返回值：设置成功，则返回true，否则返回false
     * @throws Exception
     * @author delinz
     */
    public static boolean setReadOnly(String fileName, boolean isReadOnly) throws Exception {
        try {
            File file = new File(fileName);
            if (!file.exists()) {
                return false;
            }
            //加载需要设置只读的文件
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(file);
            //设置只读
            setReadOnly(wordPackage, isReadOnly);
            //保存文件
            saveWordPackage(wordPackage, file);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public static Tbl createTable(WordprocessingMLPackage wordPackage, int rows, int cols) throws Exception {
        return createTable(wordPackage, rows, cols, true);
    }

    /**
     * 创建文档表格
     *
     * @param rows 行数
     * @param cols 列数
     * @return 返回值：返回表格对象
     * @author delinz
     */
    public static Tbl createTable(WordprocessingMLPackage wordPackage, int rows, int cols, boolean hasTh) throws Exception {
        ObjectFactory factory = Context.getWmlObjectFactory();
        Tbl tbl = factory.createTbl();
        // w:tblPr
//        StringBuffer tblSb = new StringBuffer();
//        tblSb.append("<w:tblPr ").append(Namespaces.W_NAMESPACE_DECLARATION).append(">");
//        tblSb.append("<w:tblStyle w:val=\"TableGrid\"/>");
//        tblSb.append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
//        //上边框双线
//        tblSb.append("<w:tblBorders><w:top w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
//        //左边无边框
//        tblSb.append("<w:left w:val=\"none\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
//        //下边框双线
//        tblSb.append("<w:bottom w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
//        //右边无边框
//        tblSb.append("<w:right w:val=\"none\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
//        tblSb.append("</w:tblBorders>");
//        tblSb.append("<w:tblLook w:val=\"04A0\"/>");
//        tblSb.append("</w:tblPr>");
//        TblPr tblPr = null;
//        try {
//            tblPr = (TblPr) XmlUtils.unmarshalString(tblSb.toString());
        //        } catch (JAXBException e) {
        //            e.printStackTrace();
        //        }
        TblPr tblPr = new TblPr();
        CTBorder border = new CTBorder();
        border.setColor("dcdcdc");
        border.setSz(new BigInteger("10"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideV(border);
        borders.setInsideH(border);
        tblPr.setTblBorders(borders);

        tbl.setTblPr(tblPr);
        if (tblPr != null) {
            Jc jc = factory.createJc();
            //单元格居中对齐
            jc.setVal(JcEnumeration.CENTER);
            tblPr.setJc(jc);
            CTTblLayoutType tbll = factory.createCTTblLayoutType();
            // 固定列宽
            tbll.setType(STTblLayoutType.FIXED);
            tblPr.setTblLayout(tbll);
        }
        // <w:tblGrid><w:gridCol w:w="4788"/>
        TblGrid tblGrid = factory.createTblGrid();
        tbl.setTblGrid(tblGrid);
        // Add required <w:gridCol w:w="4788"/>
        for (int i = 1; i <= cols; i++) {
            TblGridCol gridCol = factory.createTblGridCol();
            gridCol.setW(BigInteger.valueOf(getWritableWidth(wordPackage) / cols));
            tblGrid.getGridCol().add(gridCol);
        }
        // Now the rows
        for (int j = 1; j <= rows; j++) {
            Tr tr = factory.createTr();
            tbl.getContent().add(tr);
            TrPr trPr = tr.getTrPr();

            if (trPr == null) {
                trPr = factory.createTrPr();
            }
            //tr样式设置，此处设置高
            CTHeight ctHeight = new CTHeight();
            ctHeight.setVal(new BigInteger("29"));
            TrHeight trHeight = new TrHeight(ctHeight);
            trHeight.set(trPr);
            tr.setTrPr(trPr);

            for (int i = 1; i <= cols; i++) {
                Tc tc = factory.createTc();
                tr.getContent().add(tc);
                //tc样式设置，此处设置背景色及宽
                TcPr tcPr = factory.createTcPr();
                if (j == 1 && hasTh) {
                    CTShd shd = new CTShd();
                    shd.setFill("7ecef4");
                    tcPr.setShd(shd);
                }
                tc.setTcPr(tcPr);

                // <w:tcW w:w="4788" w:type="dxa"/>
                TblWidth cellWidth = factory.createTblWidth();
                tcPr.setTcW(cellWidth);
                cellWidth.setType("dxa");
                cellWidth.setW(BigInteger.valueOf(getWritableWidth(wordPackage) / cols));
                tc.getContent().add(factory.createP());
            }

        }
        return tbl;
    }

    /**
     * 创建情报文档表格
     *
     * @param rows 行数
     * @param cols 列数
     * @return 返回值：返回表格对象
     * @author delinz
     */
    public static Tbl createTableDT(WordprocessingMLPackage wordPackage, int rows, int cols, boolean hasTh) throws Exception {
        ObjectFactory factory = Context.getWmlObjectFactory();
        Tbl tbl = factory.createTbl();
        TblPr tblPr = new TblPr();
        CTBorder border = new CTBorder();
        //白线
        border.setColor("ffffff");
        border.setSz(new BigInteger("6"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideV(border);
        borders.setInsideH(border);
        tblPr.setTblBorders(borders);

        tbl.setTblPr(tblPr);
        if (tblPr != null) {
            Jc jc = factory.createJc();
            //单元格居中对齐
            jc.setVal(JcEnumeration.CENTER);
            tblPr.setJc(jc);
            // 固定列宽
            CTTblLayoutType tbll = factory.createCTTblLayoutType();
            tbll.setType(STTblLayoutType.FIXED);
            tblPr.setTblLayout(tbll);
        }
        // <w:tblGrid><w:gridCol w:w="4788"/>
        TblGrid tblGrid = factory.createTblGrid();
        tbl.setTblGrid(tblGrid);
        // Add required <w:gridCol w:w="4788"/>
        for (int i = 1; i <= cols; i++) {
            TblGridCol gridCol = factory.createTblGridCol();
            gridCol.setW(BigInteger.valueOf(getWritableWidth(wordPackage) / cols));
            tblGrid.getGridCol().add(gridCol);
        }
        // Now the rows
        for (int j = 1; j <= rows; j++) {
            Tr tr = factory.createTr();
            tbl.getContent().add(tr);
            TrPr trPr = tr.getTrPr();

            if (trPr == null) {
                trPr = factory.createTrPr();
            }
            //tr样式设置，此处设置高
            CTHeight ctHeight = new CTHeight();
            if (j == 1) {
                ctHeight.setVal(new BigInteger("413"));
            } else {
                ctHeight.setVal(new BigInteger("263"));
            }
            TrHeight trHeight = new TrHeight(ctHeight);
            trHeight.set(trPr);
            tr.setTrPr(trPr);
            //列
            for (int i = 1; i <= cols; i++) {
                Tc tc = factory.createTc();
                tr.getContent().add(tc);
                //tc样式设置，此处设置背景色及宽
                TcPr tcPr = factory.createTcPr();
                //第一行蓝底色
                if (j == 1 && hasTh) {
                    CTShd shd = new CTShd();
                    shd.setFill("40B7EA");
                    tcPr.setShd(shd);
                } else {
                    //其它杭灰底色
                    CTShd shd = new CTShd();
                    shd.setFill("f2f2f2");
                    tcPr.setShd(shd);
                }
                tc.setTcPr(tcPr);

                // <w:tcW w:w="4788" w:type="dxa"/>
                TblWidth cellWidth = factory.createTblWidth();
                tcPr.setTcW(cellWidth);
                cellWidth.setType("dxa");
                //风险资产列表
                if (cols == 5) {
                    switch (i) {
                        case 1:
                            cellWidth.setW(BigInteger.valueOf(685));
                            break;
                        case 2:
                            cellWidth.setW(BigInteger.valueOf(2135));
                            break;
                        case 3:
                            cellWidth.setW(BigInteger.valueOf(2694));
                            break;
                        case 4:
                            cellWidth.setW(BigInteger.valueOf(1134));
                            break;
                        case 5:
                            cellWidth.setW(BigInteger.valueOf(3118));
                            break;
                    }
                }
                if (cols == 3) {
                    switch (i) {
                        case 1:
                            cellWidth.setW(BigInteger.valueOf(694));
                            break;
                        case 2:
                            cellWidth.setW(BigInteger.valueOf(7088));
                            break;
                        case 3:
                            cellWidth.setW(BigInteger.valueOf(1984));
                            break;
                    }
                }
                if (cols == 4) {
                    switch (i) {
                        case 1:
                            cellWidth.setW(BigInteger.valueOf(585));
                            break;
                        case 2:
                            cellWidth.setW(BigInteger.valueOf(2094));
                            break;
                        case 3:
                            cellWidth.setW(BigInteger.valueOf(5244));
                            break;
                        case 4:
                            cellWidth.setW(BigInteger.valueOf(1701));
                            break;
                    }
                }

//                cellWidth.setW(BigInteger.valueOf(getWritableWidth(wordPackage) / cols));
//                cellWidth.setW();
                tc.getContent().add(factory.createP());
            }

        }
        return tbl;
    }


    /**
     * 填充表格内容
     *
     * @param wordPackage    文档处理包对象
     * @param tbl            表格对象
     * @param dataList       表格数据
     * @param titleList      表头数据
     * @param isFixedTitle   是否固定表头
     * @param tFontFamily    表头字体
     * @param tFontSize      表头字体大小
     * @param tIsBlod        表头是否加粗
     * @param tJcEnumeration 表头对齐方式
     * @param fontFamily     表格字体
     * @param fontSize       表格字号
     * @param isBlod         表格内容是否加粗
     * @param jcEnumeration  表格对齐方式
     * @author delinz
     */
    private static void fillTableData(WordprocessingMLPackage wordPackage, Tbl tbl, List<List<String>> dataList, List<String> titleList, boolean isFixedTitle, String tFontFamily, String tFontSize, String tFontColor, boolean tIsBlod, JcEnumeration tJcEnumeration, String fontFamily, String fontSize, String fontColor, boolean isBlod, JcEnumeration jcEnumeration) {
        List rowList = tbl.getContent();
        //整个表格的行数
        int rows = rowList.size();
        int tSize = titleList == null ? dataList.get(0).size() : titleList.size();
//      Object[] tobj = (Object[]) titleList.get(t);
        Tr tr0 = (Tr) XmlUtils.unwrap(rowList.get(0));
        List colList = tr0.getContent();
        if (titleList != null) {
            for (int c = 0; c < colList.size(); c++) {
                Tc tc0 = (Tc) XmlUtils.unwrap(colList.get(c));
                //填充表头数据
                fillCellData(tc0, converObjToStr(titleList.get(c), ""), tFontFamily, tFontSize, tFontColor, tIsBlod, tJcEnumeration);
            }
            if (isFixedTitle) {
                //设置固定表头
                fixedTitle(tr0);
            }
        }

        int offset = titleList == null ? 0 : 1;
        for (int i = offset; i < dataList.size() + offset; i++) {
            Tr tr = (Tr) XmlUtils.unwrap(rowList.get(i));
            List<String> objs = null;
            //如果表格内容不为空，则取出相应的数据进行填充
            if (dataList != null) {
                objs = dataList.get(i - offset);
            }
            List colsList = tr.getContent();
            for (int j = 0; j < dataList.get(i - offset).size(); j++) {
//                Tc tc = (Tc) row.getContent().get(j);
//                tc.getContent().add(getPText(dataList.get(i - 1).get(j)));
                Tc tc = (Tc) XmlUtils.unwrap(colsList.get(j));
                //填充表格数据
                if (objs != null) {
                    fillCellData(tc, converObjToStr(objs.get(j), "-"), fontFamily, fontSize, fontColor, isBlod, jcEnumeration);
                } else {
                    fillCellData(tc, "", fontFamily, fontSize, fontColor, isBlod, jcEnumeration);
                }
            }
        }

    }


    /**
     * 填充报告表格内容 -风险资产列表颜色需要动态改变
     *
     * @param wordPackage    文档处理包对象
     * @param tbl            表格对象
     * @param dataList       表格数据
     * @param titleList      表头数据
     * @param isFixedTitle   是否固定表头
     * @param tFontFamily    表头字体
     * @param tFontSize      表头字体大小
     * @param tIsBlod        表头是否加粗
     * @param tJcEnumeration 表头对齐方式
     * @param fontFamily     表格字体
     * @param fontSize       表格字号
     * @param isBlod         表格内容是否加粗
     * @param jcEnumeration  表格对齐方式
     * @author delinz
     */
    private static void fillTableDataDT(WordprocessingMLPackage wordPackage, Tbl tbl, List<List<String>> dataList, List<String> titleList, boolean isFixedTitle, String tFontFamily, String tFontSize, String tFontColor, boolean tIsBlod, JcEnumeration tJcEnumeration, String fontFamily, String fontSize, String fontColor, boolean isBlod, JcEnumeration jcEnumeration, boolean isDiffColor) {
        List rowList = tbl.getContent();
        //整个表格的行数
        int rows = rowList.size();
        int tSize = titleList == null ? dataList.get(0).size() : titleList.size();
//      Object[] tobj = (Object[]) titleList.get(t);
        Tr tr0 = (Tr) XmlUtils.unwrap(rowList.get(0));
        List colList = tr0.getContent();
        if (titleList != null) {
            for (int c = 0; c < colList.size(); c++) {
                Tc tc0 = (Tc) XmlUtils.unwrap(colList.get(c));
                //填充表头数据
                fillCellData(tc0, converObjToStr(titleList.get(c), ""), tFontFamily, "18", tFontColor, tIsBlod, tJcEnumeration);
            }
            if (isFixedTitle) {
                //设置固定表头
                fixedTitle(tr0);
            }
        }

        int offset = titleList == null ? 0 : 1;
        for (int i = offset; i < dataList.size() + offset; i++) {
            Tr tr = (Tr) XmlUtils.unwrap(rowList.get(i));
            List<String> objs = null;
            //如果表格内容不为空，则取出相应的数据进行填充
            if (dataList != null) {
                objs = dataList.get(i - offset);
            }
            List colsList = tr.getContent();
            for (int j = 0; j < dataList.get(i - offset).size(); j++) {
//                Tc tc = (Tc) row.getContent().get(j);
//                tc.getContent().add(getPText(dataList.get(i - 1).get(j)));
                Tc tc = (Tc) XmlUtils.unwrap(colsList.get(j));
                //填充表格数据
                if (objs != null) {
                    if (isDiffColor) {
                        fillCellDataDT(tc, converObjToStr(objs.get(j), "-"), fontFamily, fontSize, fontColor, isBlod, jcEnumeration, true);
                    } else {
                        fillCellData(tc, converObjToStr(objs.get(j), "-"), fontFamily, fontSize, fontColor, isBlod, jcEnumeration);
                    }
                } else {
                    fillCellData(tc, "", fontFamily, fontSize, fontColor, isBlod, jcEnumeration);
                }
            }
        }

    }

    /**
     * 填充表格内容，固定表头，表头宋体加粗，小五号，表格内容宋体，小五号，表格居中对齐</BR>
     * 其中表格数据跟表头数据结构要一致，适用于简单的n行m列的普通表格
     *
     * @param wordPackage 文档处理包对象
     * @param tbl         表格对象
     * @param dataList    表格数据
     * @param titleList   表头数据，如果不需要表头信息，则只要传入null即可
     * @author delinz
     */
    public static void fillTableDataDT(WordprocessingMLPackage wordPackage, Tbl tbl, List dataList, List titleList, boolean isDiffColor) {
        fillTableDataDT(wordPackage, tbl, dataList, titleList, true, "微软雅黑", "15", "ffffff", true, JcEnumeration.CENTER, "微软雅黑", "15", "817f80", false, JcEnumeration.CENTER, isDiffColor);
    }

    /**
     * 填充表格内容，固定表头，表头宋体加粗，小五号，表格内容宋体，小五号，表格居中对齐</BR>
     * 其中表格数据跟表头数据结构要一致，适用于简单的n行m列的普通表格
     *
     * @param wordPackage 文档处理包对象
     * @param tbl         表格对象
     * @param dataList    表格数据
     * @param titleList   表头数据，如果不需要表头信息，则只要传入null即可
     * @author delinz
     */
    public static void fillTableData(WordprocessingMLPackage wordPackage, Tbl tbl, List dataList, List titleList) {
        fillTableData(wordPackage, tbl, dataList, titleList, true, "微软雅黑", "15", "ffffff", true, JcEnumeration.CENTER, "微软雅黑", "15", "000000", false, JcEnumeration.CENTER);
    }

    /**
     * 固定表头
     *
     * @param tr 行对象
     * @author delinz
     */
    public static void fixedTitle(Tr tr) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        BooleanDefaultTrue bdt = factory.createBooleanDefaultTrue();
        //表示固定表头
        bdt.setVal(true);
        TrPr trpr = tr.getTrPr();
        if (trpr == null) {
            trpr = factory.createTrPr();
        }
        trpr.getCnfStyleOrDivIdOrGridBefore().add(factory.createCTTrPrBaseTblHeader(bdt));
        tr.setTrPr(trpr);
    }

    /**
     * 填充单元格内容
     *
     * @param tc            单元格对象
     * @param data          内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param isBlod        是否加粗
     * @param jcEnumeration 对齐方式
     * @author delinz
     */
    private static void fillCellData(Tc tc, String data, String fontFamily, String fontSize, String fontColor, boolean isBlod, JcEnumeration jcEnumeration) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = (P) XmlUtils.unwrap(tc.getContent().get(0));
        //设置表格内容的对齐方式
        setCellContentStyle(p, jcEnumeration);
        Text t = factory.createText();
        // 每个字符后面添加零宽度空格，来解决长文本pdf换行问题
        // https://stackoverflow.com/questions/27230307/export-to-pdf-table-columns-not-wrapped
        String zeroWidthSpace = Character.toString((char) 8203);
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < data.length(); i++) {
            sb.append(data.charAt(i));
            sb.append(zeroWidthSpace);
        }
        t.setValue(sb.toString());

        R run = factory.createR();
        //设置表格内容字体样式
        run.setRPr(getRPr(fontFamily, fontColor, fontSize, STHint.EAST_ASIA, isBlod));

        TcPr tcpr = tc.getTcPr();
        if (tcpr == null) {
            tcpr = factory.createTcPr();
        }
        //设置内容垂直居中
        CTVerticalJc valign = factory.createCTVerticalJc();
        valign.setVal(STVerticalJc.CENTER);
        tcpr.setVAlign(valign);
        run.getContent().add(t);
        p.getContent().add(run);
    }

    /**
     * 填充报告单元格内容，部分内容有背景色
     *
     * @param tc            单元格对象
     * @param data          内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param isBlod        是否加粗
     * @param jcEnumeration 对齐方式
     * @author delinz
     */
    private static void fillCellDataDT(Tc tc, String data, String fontFamily, String fontSize, String fontColor, boolean isBlod, JcEnumeration jcEnumeration, boolean isDiffColor) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = (P) XmlUtils.unwrap(tc.getContent().get(0));
        //设置表格内容的对齐方式
        setCellContentStyle(p, jcEnumeration);
        Text t = factory.createText();
        // 每个字符后面添加零宽度空格，来解决长文本pdf换行问题
        // https://stackoverflow.com/questions/27230307/export-to-pdf-table-columns-not-wrapped
        String zeroWidthSpace = Character.toString((char) 8203);
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < data.length(); i++) {
            sb.append(data.charAt(i));
            sb.append(zeroWidthSpace);
        }
        t.setValue(sb.toString());

        R run = factory.createR();
        //设置表格内容字体样式
        run.setRPr(getRPr(fontFamily, fontColor, fontSize, STHint.EAST_ASIA, isBlod));

        TcPr tcpr = tc.getTcPr();
        if (tcpr == null) {
            tcpr = factory.createTcPr();
        }
        //设置内容垂直居中
        CTVerticalJc valign = factory.createCTVerticalJc();
        valign.setVal(STVerticalJc.CENTER);
        tcpr.setVAlign(valign);
        //背景色
        if (isDiffColor) {
            CTShd shd = new CTShd();
            switch (data) {
                case "已失陷":
                    shd.setFill("FF1B2D");
                    run.setRPr(getRPr(fontFamily, "ffffff", fontSize, STHint.EAST_ASIA, isBlod));
                    break;
                case "高风险":
                    shd.setFill("FF621B");
                    run.setRPr(getRPr(fontFamily, "ffffff", fontSize, STHint.EAST_ASIA, isBlod));
                    break;
                case "低风险":
                    shd.setFill("FF9725");
                    run.setRPr(getRPr(fontFamily, "ffffff", fontSize, STHint.EAST_ASIA, isBlod));
                    break;
                default:
                    shd.setFill("f2f2f2");
                    break;
            }
            tcpr.setShd(shd);
        }

        run.getContent().add(t);
        p.getContent().add(run);
    }


    /**
     * 填充单元格内容，小五号，宋体，内容居中
     *
     * @param tc     单元格对象
     * @param data   数据
     * @param isBlod 是否加粗
     * @author delinz
     */
    public static void fillCellData(Tc tc, String data, boolean isBlod) {
        fillCellData(tc, data, "Microsoft YaHei", "10", "FFFFFF", isBlod, JcEnumeration.CENTER);
    }

    /**
     * 获取单元格边框样式
     *
     * @param type   单元格类型，0表示无边框，2表示双线边框，其他表示单线边框
     * @param color  边框颜色
     * @param border 边框大小
     * @param space  间距
     * @return 返回值：返回边框对象
     * @author delinz
     */
    private static CTBorder getCTBorder(int type, String color, String border, String space) {
        CTBorder ctb = new CTBorder();
        if (type == 0) {
            ctb.setVal(STBorder.NIL);
        } else {
            ctb.setColor(color);
            ctb.setSz(new BigInteger(border));
            ctb.setSpace(new BigInteger(space));
            if (type == 2) {
                ctb.setVal(STBorder.DOUBLE);
            } else {
                ctb.setVal(STBorder.SINGLE);
            }
        }
        return ctb;
    }

    /**
     * 设置单元格内容对齐方式
     *
     * @param p             内容
     * @param jcEnumeration 对齐方式
     * @author delinz
     */
    public static void setCellContentStyle(P p, JcEnumeration jcEnumeration) {
        PPr pPr = p.getPPr();
        if (pPr == null) {
            ObjectFactory factory = Context.getWmlObjectFactory();
            pPr = factory.createPPr();
        }
        //设置单元格内缩进
        PPrBase.Ind ind = factory.createPPrBaseInd();
        ind.setFirstLine(BigInteger.valueOf(0));
//        ind.setHanging(BigInteger.valueOf(0));
        ind.setLeft(BigInteger.valueOf(0));
        pPr.setInd(ind);
        Jc jc = pPr.getJc();
        if (jc == null) {
            jc = new Jc();
        }
        jc.setVal(jcEnumeration);
        pPr.setJc(jc);
        p.setPPr(pPr);
    }

    /**
     * 设置单元格内容对齐方式，居中对齐
     *
     * @param p 内容
     * @author delinz
     */
    public static void setCellContentStyle(P p) {
        setCellContentStyle(p, JcEnumeration.CENTER);
    }

    /**
     * Object数据转换为String类型
     *
     * @param obj
     * @param defaultStr 如果obj对象为空，则返回的值
     * @return
     * @author delinz
     */
    public static String converObjToStr(Object obj, String defaultStr) {
        if (obj != null && !"".equals(obj)) {
            return obj.toString();
        }
        return defaultStr;
    }

    /**
     * 提取文档中所有书签
     *
     * @param mainDocumentPart
     * @return
     */
    public static RangeFinder getBookMarks(MainDocumentPart mainDocumentPart) {
        Document wmlDoc = (Document) mainDocumentPart.getJaxbElement();
        Body body = wmlDoc.getBody();
        // 提取正文中所有段落
        List<Object> paragraphs = body.getContent();
        // 提取书签并创建书签的游标
        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);
        return rt;
    }

    /**
     * 把关键词替换成表格
     * 关键词最好是单独占用一行
     *
     * @param mlPackage
     * @param replace
     * @param table
     */
    public static void replaceTable(WordprocessingMLPackage mlPackage, String replace, Tbl table) {
        List<Object> paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
        for (Object text : paragraphs) {
            Text content = (Text) text;
            if (content.getValue().equals(replace)) {
                R r = (R) content.getParent();
                r.getContent().clear();
                P p = (P) r.getParent();
                ContentAccessor doc = (ContentAccessor) p.getParent();
                doc.getContent().set(doc.getContent().indexOf(p), table);
            }
        }
    }

    public static void setParagraphNum(P p, String ilvlStr, String id) {
        PPr ppr = getPPr(p);
        PPrBase.NumPr numPr = new PPrBase.NumPr();
        PPrBase.NumPr.Ilvl ilvl = new PPrBase.NumPr.Ilvl();
        ilvl.setVal(new BigInteger(ilvlStr));
        PPrBase.NumPr.NumId numId = new PPrBase.NumPr.NumId();
        numId.setVal(new BigInteger(id));
        numPr.setIlvl(ilvl);
        numPr.setNumId(numId);
        ppr.setNumPr(numPr);
    }

    /**
     * @Description: 设置段落缩进信息 1厘米≈567
     */
    public static void setParagraphIndInfo(P p, String firstLine,
                                           String firstLineChar, String hanging, String hangingChar,
                                           String right, String rigthChar, String left, String leftChar) {
        PPr ppr = getPPr(p);
        PPrBase.Ind ind = ppr.getInd();
        if (ind == null) {
            ind = new PPrBase.Ind();
            ppr.setInd(ind);
        }
        if (StringUtils.isNotBlank(firstLine)) {
            ind.setFirstLine(new BigInteger(firstLine));
        }
        if (StringUtils.isNotBlank(firstLineChar)) {
            ind.setFirstLineChars(new BigInteger(firstLineChar));
        }
        if (StringUtils.isNotBlank(hanging)) {
            ind.setHanging(new BigInteger(hanging));
        }
        if (StringUtils.isNotBlank(hangingChar)) {
            ind.setHangingChars(new BigInteger(hangingChar));
        }
        if (StringUtils.isNotBlank(left)) {
            ind.setLeft(new BigInteger(left));
        }
        if (StringUtils.isNotBlank(leftChar)) {
            ind.setLeftChars(new BigInteger(leftChar));
        }
        if (StringUtils.isNotBlank(right)) {
            ind.setRight(new BigInteger(right));
        }
        if (StringUtils.isNotBlank(rigthChar)) {
            ind.setRightChars(new BigInteger(rigthChar));
        }
    }

    /**
     * @param distance    :距正文距离 1厘米=567
     * @param start       :起始编号(0开始)
     * @param countBy     :行号间隔
     * @param restartType :STLineNumberRestart.CONTINUOUS(continuous连续编号)<br/>
     *                    STLineNumberRestart.NEW_PAGE(每页重新编号)<br/>
     *                    STLineNumberRestart.NEW_SECTION(每节重新编号)
     * @Description: 设置行号
     */
    public static void setDocInNumType(WordprocessingMLPackage wordPackage,
                                       String countBy, String distance, String start,
                                       STLineNumberRestart restartType) {
        SectPr sectPr = getDocSectPr(wordPackage);
        CTLineNumber lnNumType = sectPr.getLnNumType();
        if (lnNumType == null) {
            lnNumType = new CTLineNumber();
            sectPr.setLnNumType(lnNumType);
        }
        if (StringUtils.isNotBlank(countBy)) {
            lnNumType.setCountBy(new BigInteger(countBy));
        }
        if (StringUtils.isNotBlank(distance)) {
            lnNumType.setDistance(new BigInteger(distance));
        }
        if (StringUtils.isNotBlank(start)) {
            lnNumType.setStart(new BigInteger(start));
        }
        if (restartType != null) {
            lnNumType.setRestart(restartType);
        }
    }

    public static PPr getPPr(P p) {
        PPr ppr = p.getPPr();
        if (ppr == null) {
            ppr = new PPr();
            p.setPPr(ppr);
        }
        return ppr;
    }

    public static ParaRPr getParaRPr(PPr ppr) {
        ParaRPr parRpr = ppr.getRPr();
        if (parRpr == null) {
            parRpr = new ParaRPr();
            ppr.setRPr(parRpr);
        }
        return parRpr;
    }

    public static SectPr getDocSectPr(WordprocessingMLPackage wordPackage) {
        SectPr sectPr = wordPackage.getDocumentModel().getSections().get(0)
                .getSectPr();
        return sectPr;
    }

    // 设置段间距-->行距 段前段后距离
    // 段前段后可以设置行和磅 行距只有磅
    // 段前磅值和行值同时设置，只有行值起作用
    // TODO 1磅=20 1行=100 单倍行距=240 为什么是这个值不知道

    /**
     * @param jcEnumeration     对齐方式
     * @param isSpace           是否设置段前段后值
     * @param before            段前磅数
     * @param after             段后磅数
     * @param beforeLines       段前行数
     * @param afterLines        段后行数
     * @param isLine            是否设置行距
     * @param lineValue         行距值
     * @param sTLineSpacingRule 自动auto 固定exact 最小 atLeast
     */
    public static void setParagraphSpacing(P p,
                                           JcEnumeration jcEnumeration, boolean isSpace, String before,
                                           String after, String beforeLines, String afterLines,
                                           boolean isLine, String lineValue,
                                           STLineSpacingRule sTLineSpacingRule) {
        PPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = factory.createPPr();
        }
        Jc jc = pPr.getJc();
        if (jc == null) {
            jc = new Jc();
        }
        jc.setVal(jcEnumeration);
        pPr.setJc(jc);

        PPrBase.Spacing spacing = new PPrBase.Spacing();
        if (isSpace) {
            if (before != null) {
                // 段前磅数
                spacing.setBefore(new BigInteger(before));
            }
            if (after != null) {
                // 段后磅数
                spacing.setAfter(new BigInteger(after));
            }
            if (beforeLines != null) {
                // 段前行数
                spacing.setBeforeLines(new BigInteger(beforeLines));
            }
            if (afterLines != null) {
                // 段后行数
                spacing.setAfterLines(new BigInteger(afterLines));
            }
        }
        if (isLine) {
            if (lineValue != null) {
                spacing.setLine(new BigInteger(lineValue));
            }
            spacing.setLineRule(sTLineSpacingRule);
        }
        pPr.setSpacing(spacing);
        p.setPPr(pPr);
    }

    public static void mergeCellsHorizontalByGridSpan(Tbl tbl, int row, int fromCell,
                                                      int toCell) {
        if (row < 0 || fromCell < 0 || toCell < 0) {
            return;
        }
        List<Tr> trList = getTblAllTr(tbl);
        if (row > trList.size()) {
            return;
        }
        Tr tr = trList.get(row);
        List<Tc> tcList = getTrAllCell(tr);
        for (int cellIndex = Math.min(tcList.size() - 1, toCell); cellIndex >= fromCell; cellIndex--) {
            Tc tc = tcList.get(cellIndex);
            TcPr tcPr = getTcPr(tc);
            if (cellIndex == fromCell) {
                TcPrInner.GridSpan gridSpan = tcPr.getGridSpan();
                if (gridSpan == null) {
                    gridSpan = new TcPrInner.GridSpan();
                    tcPr.setGridSpan(gridSpan);
                }
                gridSpan.setVal(BigInteger.valueOf(Math.min(tcList.size() - 1,
                        toCell) - fromCell + 1));
            } else {
                tr.getContent().remove(cellIndex);
            }
        }
    }

    /**
     * @Description: 跨列合并
     */
    public static void mergeCellsHorizontal(Tbl tbl, int row, int fromCell, int toCell) {
        if (row < 0 || fromCell < 0 || toCell < 0) {
            return;
        }
        List<Tr> trList = getTblAllTr(tbl);
        if (row > trList.size()) {
            return;
        }
        Tr tr = trList.get(row);
        List<Tc> tcList = getTrAllCell(tr);
        for (int cellIndex = fromCell, len = Math
                .min(tcList.size() - 1, toCell); cellIndex <= len; cellIndex++) {
            Tc tc = tcList.get(cellIndex);
            TcPr tcPr = getTcPr(tc);
            TcPrInner.HMerge hMerge = tcPr.getHMerge();
            if (hMerge == null) {
                hMerge = new TcPrInner.HMerge();
                tcPr.setHMerge(hMerge);
            }
            if (cellIndex == fromCell) {
                hMerge.setVal("restart");
            } else {
                hMerge.setVal("continue");
            }
        }
    }

    /**
     * @Description: 跨行合并
     */
    public static void mergeCellsVertically(Tbl tbl, int col, int fromRow, int toRow) {
        if (col < 0 || fromRow < 0 || toRow < 0) {
            return;
        }
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            Tc tc = getTc(tbl, rowIndex, col);
            if (tc == null) {
                break;
            }
            TcPr tcPr = getTcPr(tc);
            TcPrInner.VMerge vMerge = tcPr.getVMerge();
            if (vMerge == null) {
                vMerge = new TcPrInner.VMerge();
                tcPr.setVMerge(vMerge);
            }
            if (rowIndex == fromRow) {
                vMerge.setVal("restart");
            } else {
                vMerge.setVal("continue");
            }
        }
    }

    /**
     * @Description:得到指定位置的表格
     */
    public static Tc getTc(Tbl tbl, int row, int cell) {
        if (row < 0 || cell < 0) {
            return null;
        }
        List<Tr> trList = getTblAllTr(tbl);
        if (row >= trList.size()) {
            return null;
        }
        List<Tc> tcList = getTrAllCell(trList.get(row));
        if (cell >= tcList.size()) {
            return null;
        }
        return tcList.get(cell);
    }

    /**
     * @Description: 获取所有的单元格
     */
    public static List<Tc> getTrAllCell(Tr tr) {
        List<Object> objList = getAllElementFromObject(tr, Tc.class);
        List<Tc> tcList = new ArrayList<Tc>();
        if (objList == null) {
            return tcList;
        }
        for (Object tcObj : objList) {
            if (tcObj instanceof Tc) {
                Tc objTc = (Tc) tcObj;
                tcList.add(objTc);
            }
        }
        return tcList;
    }

    public static TcPr getTcPr(Tc tc) {
        TcPr tcPr = tc.getTcPr();
        if (tcPr == null) {
            tcPr = new TcPr();
            tc.setTcPr(tcPr);
        }
        return tcPr;
    }

    /**
     * @Description: 得到表格所有的行
     */
    public static List<Tr> getTblAllTr(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
        List<Tr> trList = new ArrayList<Tr>();
        if (objList == null) {
            return trList;
        }
        for (Object obj : objList) {
            if (obj instanceof Tr) {
                Tr tr = (Tr) obj;
                trList.add(tr);
            }
        }
        return trList;
    }

    /**
     * 文件内容替换
     *
     * @param filePath
     * @param oldstr
     * @param newStr
     */
    public synchronized static void autoReplace(String filePath, String oldstr, String newStr) {
        File file = new File(filePath);
        Long fileLength = file.length();
        byte[] fileContext = new byte[fileLength.intValue()];
        FileInputStream in = null;
        PrintWriter out = null;
        try {
            in = new FileInputStream(filePath);
            in.read(fileContext);
            // 避免出现中文乱码
            String str = new String(fileContext, "utf-8");
            str = str.replace(oldstr, newStr);
            out = new PrintWriter(filePath);
            out.write(str);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.flush();
                out.close();
                in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public static void createHyperlink(MainDocumentPart mainPart,
                                       String url, String value, String cnFontName, String enFontName,
                                       String fontSize) throws Exception {
        ObjectFactory objectFactory = new ObjectFactory();
        P paragraph = objectFactory.createP();
        mainPart.getContent().add(paragraph);
        if (StringUtils.isBlank(enFontName)) {
            enFontName = "Times New Roman";
        }
        if (StringUtils.isBlank(cnFontName)) {
            cnFontName = "微软雅黑";
        }
        if (StringUtils.isBlank(fontSize)) {
            fontSize = "22";
        }
        org.docx4j.relationships.ObjectFactory reFactory = new org.docx4j.relationships.ObjectFactory();
        org.docx4j.relationships.Relationship rel = reFactory
                .createRelationship();
        rel.setType(Namespaces.HYPERLINK);
        rel.setTarget(url);
        rel.setTargetMode("External");
        mainPart.getRelationshipsPart().addRelationship(rel);
        StringBuffer sb = new StringBuffer();
        // addRelationship sets the rel's @Id
        sb.append("<w:hyperlink r:id=\"");
        sb.append(rel.getId());
        sb.append("\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
        sb.append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >");
        sb.append("<w:r><w:rPr><w:rStyle w:val=\"Hyperlink\" />");
        sb.append("<w:rFonts  w:ascii=\"");
        sb.append(enFontName);
        sb.append("\"  w:hAnsi=\"");
        sb.append(enFontName);
        sb.append("\"  w:eastAsia=\"");
        sb.append(cnFontName);
        sb.append("\" w:hint=\"eastAsia\"/>");
        sb.append("<w:sz w:val=\"");
        sb.append(fontSize);
        sb.append("\"/><w:szCs w:val=\"");
        sb.append(fontSize);
        sb.append("\"/></w:rPr><w:t>");
        sb.append(value);
        sb.append("</w:t></w:r></w:hyperlink>");

        P.Hyperlink link = (P.Hyperlink) XmlUtils.unmarshalString(sb.toString());
        paragraph.getContent().add(link);
    }

    public static void createHyperlink(MainDocumentPart mdp, String url, String linkText) {
        try {

            ObjectFactory objectFactory = new ObjectFactory();
            P paragraph = objectFactory.createP();
            mdp.getContent().add(paragraph);

            org.docx4j.relationships.ObjectFactory reFactory = new org.docx4j.relationships.ObjectFactory();
            org.docx4j.relationships.Relationship rel = reFactory
                    .createRelationship();
            rel.setType(Namespaces.HYPERLINK);
            rel.setTarget(url);
            rel.setTargetMode("External");
            mdp.getRelationshipsPart().addRelationship(rel);

            // addRelationship sets the rel's @Id
            String hpl = "<w:hyperlink r:id=\"" + rel.getId() + "\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >" +
                    "<w:r>" +
                    "<w:rPr>" +
                    "<w:rStyle w:val=\"a9\" />" +  // TODO: enable this style in the document!
                    "</w:rPr>" +
                    "<w:t>" + linkText + "</w:t>" +
                    "</w:r>" +
                    "</w:hyperlink>";

            paragraph.getContent().add((P.Hyperlink) XmlUtils.unmarshalString(hpl));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    public static P.Hyperlink createPureHyperlink(MainDocumentPart mdp, String url, String linkText) {
        try {
            org.docx4j.relationships.ObjectFactory reFactory = new org.docx4j.relationships.ObjectFactory();
            org.docx4j.relationships.Relationship rel = reFactory
                    .createRelationship();
            rel.setType(Namespaces.HYPERLINK);
            rel.setTarget(url);
            rel.setTargetMode("External");
            mdp.getRelationshipsPart().addRelationship(rel);
            // addRelationship sets the rel's @Id
            String hpl = "<w:hyperlink r:id=\"" + rel.getId() + "\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >" +
                    "<w:r>" +
                    "<w:rPr>" +
                    "<w:rStyle w:val=\"a9\" />" +  // TODO: enable this style in the document!
                    "</w:rPr>" +
                    "<w:t>" + linkText + "</w:t>" +
                    "</w:r>" +
                    "</w:hyperlink>";
            return (P.Hyperlink) XmlUtils.unmarshalString(hpl);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }


    public static void replaceTextByHyperlink(WordprocessingMLPackage wordMLPackage, String threatInfo) {
        List texts = getAllElementFromObject(wordMLPackage.getMainDocumentPart(), Text.class);
        for (Object text : texts) {
            Text content = (Text) text;
            if (content.getValue().contains(threatInfo)) {
                if (content.getValue().contains(THREAT_DBAPPSECURITY_URL)){
                    if (!content.getValue().contains("了解更多")) {
                        R r = (R) content.getParent();
                        r.getContent().clear();
                        P p = (P) r.getParent();
                        Object ob1 = p.getContent().get(0);
                        p.getContent().clear();
                        p.getContent().add(createPureHyperlink(wordMLPackage.getMainDocumentPart(), content.getValue(), content.getValue()));
                        p.getContent().add(ob1);
                    }
            } else {
                R r = (R) content.getParent();
                P p = (P) r.getParent();
                StringBuilder herf = new StringBuilder();
                List<Object> objList = p.getContent();
                for (int i = 0; i < objList.size(); i++) {
                    Object object = objList.get(i);
                    if (object.getClass().getName().contains("JAXBElement")) {
                        appendJAXBElementContent(object, herf);
                    } else {
                        if (object.getClass().getName().contains("org.docx4j.wml.R")) {
                            R secondeR = (R) object;
                            Object thirdObject = secondeR.getContent().get(0);
                            if (thirdObject.getClass().getName().contains("JAXBElement")) {
                                appendJAXBElementContent(thirdObject, herf);
                            }
                        }
                    }
                }

                p.getContent().clear();
                p.getContent().add(createPureHyperlink(wordMLPackage.getMainDocumentPart(), herf.toString(), herf.toString()));

            }

        }
    }

}

    private static void appendJAXBElementContent(Object object, StringBuilder herf) {
        JAXBElement jaxbelement = (JAXBElement) object;
        Object secondeObject = jaxbelement.getValue();
        if (secondeObject.getClass().getName().contains("Text")) {
            Text secondeText = (Text) secondeObject;
            herf.append(secondeText.getValue());
        }

    }

    public static void replaceContent(WordprocessingMLPackage mlPackage, List<ReplaceContent> replaceContentList) {
        try {
            List paragraphs = getAllElementFromObject(mlPackage.getMainDocumentPart(), Text.class);
            for (Object text : paragraphs) {
                Text content = (Text) text;
                ReplaceContent replaceContent = getRpContentByField(replaceContentList, content.getValue());
                if (replaceContent != null) {
                    switch (replaceContent.getType()) {
                        case RPType.TEXT:
                            if (replaceContent.getContent() != null) {
                                content.setValue(replaceContent.getContent().toString());
                                content.setSpace("preserve");
                            } else {
                                R r = (R) content.getParent();
                                P p = (P) r.getParent();//P para=(P) p.getParent();P paragraph=(P)para.getParent();ContentAccessor page=(ContentAccessor)paragraph.getParent();page.getContent().remove(paragraph);
                                ContentAccessor doc = (ContentAccessor) p.getParent();
                                doc.getContent().remove(doc.getContent().indexOf(p) - 1);
                                doc.getContent().remove(doc.getContent().indexOf(p) + 1);
                                doc.getContent().remove(p);
                            }
                            break;
                        case RPType.TABLE:
                            List contentTable = (List) replaceContent.getContent();
                            List th = (List) contentTable.get(0);
                            List data = (List) contentTable.get(1);
                            Boolean isFirstColMerge = (Boolean) contentTable.get(2);
                            Boolean isDiffColor = (Boolean) contentTable.get(3);
                            R r = (R) content.getParent();
                            r.getContent().clear();
                            P p = (P) r.getParent();
                            int rowNum = data.size() + (th == null ? 0 : 1);
                            int colsNum = th == null ? ((List) data.get(0)).size() : th.size();
                            Tbl table = DocxUtils.createTableDT(mlPackage, rowNum, colsNum, th != null);
                            DocxUtils.fillTableDataDT(mlPackage, table, data, th, isDiffColor);
                            if (isFirstColMerge) {
                                int mergeStart = 0;
                                int mergeEnd = 0;
                                String colValue = null;
                                for (Object datum : data) {
                                    String colValue1 = (String) ((List) datum).get(0);
                                    if (colValue1.equals(colValue)) {
                                        mergeEnd++;
                                    } else {
                                        if (colValue != null) {
                                            if (mergeStart < mergeEnd) {
                                                mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                                            }
                                        }
                                        colValue = colValue1;
                                        mergeEnd++;
                                        mergeStart = mergeEnd;
                                    }
                                }
                                if (mergeEnd > mergeStart) {
                                    mergeCellsVertically(table, 0, mergeStart, mergeEnd);
                                }
                            }
                            ContentAccessor doc = (ContentAccessor) p.getParent();
                            doc.getContent().set(doc.getContent().indexOf(p), table);
                            break;

                        case RPType.IMAGE:
                            break;


                        case RPType.Bar:
                            break;
                    }

                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static ReplaceContent getRpContentByField(List<ReplaceContent> replaceContentList, String value) {
        for (ReplaceContent replaceContent : replaceContentList) {
            if (replaceContent.getField().equals(value)) {
                return replaceContent;
            }
        }
        return null;
    }
}

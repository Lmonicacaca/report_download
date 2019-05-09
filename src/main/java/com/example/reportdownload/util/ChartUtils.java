package com.example.reportdownload.util;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.axis.*;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.*;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.AbstractRenderer;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.chart.renderer.category.StackedBarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.renderer.xy.AbstractXYItemRenderer;
import org.jfree.chart.renderer.xy.StandardXYBarPainter;
import org.jfree.chart.renderer.xy.XYBarRenderer;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.time.Day;
import org.jfree.data.time.Minute;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.ui.RectangleInsets;
import org.jfree.ui.TextAnchor;
import org.jfree.util.ArrayUtilities;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Jfreechart工具类
 * 解决中文乱码问题
 * 用来创建类别图表数据集、创建饼图数据集、时间序列图数据集
 * 用来对柱状图、折线图、饼图、堆积柱状图、时间序列图的样式进行渲染
 * 设置X-Y坐标轴样式
 *
 * @author delin
 */
public class ChartUtils {
    private static String NO_DATA_MSG = "数据加载失败";

    private static final Pattern urlPattern=Pattern.compile("^(?=^.{3,255}(?:$|/))([a-zA-Z0-9][-_a-zA-Z0-9]{0,62}(\\.[a-zA-Z0-9][-_a-zA-Z0-9]{0,62})+(?::\\d{1,5})?|(?:(?:\\d|[1-9]\\d|1\\d{2}|2[0-4]\\d|25[0-5])\\.){3}(?:\\d|[1-9]\\d|\\d{1,2}|1\\d{2}|2[0-4]\\d|25[0-5]):(\\d{1,5}))(?:$|/.*)");

    private static Font FONT = new Font("宋体", Font.PLAIN, 12);
    public static Color[] CHART_COLORS = {
            new Color(31, 129, 188), new Color(92, 92, 97), new Color(144, 237, 125), new Color(255, 188, 117),
            new Color(153, 158, 255), new Color(255, 117, 153), new Color(253, 236, 109), new Color(128, 133, 232),
            new Color(158, 90, 102), new Color(255, 204, 102)};// 颜色

    static {
        setChartTheme();
    }

    public ChartUtils() {
    }

    /**
     * 中文主题样式 解决乱码
     */
    public static void setChartTheme() {
        // 设置中文主题样式 解决乱码
        StandardChartTheme chartTheme = new StandardChartTheme("CN");
        // 设置标题字体
        chartTheme.setExtraLargeFont(FONT);
        // 设置图例的字体
        chartTheme.setRegularFont(FONT);
        // 设置轴向的字体
        chartTheme.setLargeFont(FONT);
        chartTheme.setSmallFont(FONT);
        chartTheme.setTitlePaint(new Color(51, 51, 51));
        chartTheme.setSubtitlePaint(new Color(85, 85, 85));

        chartTheme.setLegendBackgroundPaint(Color.WHITE);// 设置标注
        chartTheme.setLegendItemPaint(Color.BLACK);//
        chartTheme.setChartBackgroundPaint(Color.WHITE);
        // 绘制颜色绘制颜色.轮廓供应商
        // paintSequence,outlinePaintSequence,strokeSequence,outlineStrokeSequence,shapeSequence

        Paint[] OUTLINE_PAINT_SEQUENCE = new Paint[]{Color.WHITE};
        // 绘制器颜色源
        DefaultDrawingSupplier drawingSupplier = new DefaultDrawingSupplier(CHART_COLORS, CHART_COLORS, OUTLINE_PAINT_SEQUENCE,
                DefaultDrawingSupplier.DEFAULT_STROKE_SEQUENCE, DefaultDrawingSupplier.DEFAULT_OUTLINE_STROKE_SEQUENCE,
                DefaultDrawingSupplier.DEFAULT_SHAPE_SEQUENCE);
        chartTheme.setDrawingSupplier(drawingSupplier);

        chartTheme.setPlotBackgroundPaint(Color.WHITE);// 绘制区域
        chartTheme.setPlotOutlinePaint(Color.WHITE);// 绘制区域外边框
        chartTheme.setLabelLinkPaint(new Color(8, 55, 114));// 链接标签颜色
        chartTheme.setLabelLinkStyle(PieLabelLinkStyle.CUBIC_CURVE);

        chartTheme.setAxisOffset(new RectangleInsets(5, 12, 5, 12));
        chartTheme.setDomainGridlinePaint(new Color(192, 208, 224));// X坐标轴垂直网格颜色
        chartTheme.setRangeGridlinePaint(new Color(192, 192, 192));// Y坐标轴水平网格颜色

        chartTheme.setBaselinePaint(Color.WHITE);
        chartTheme.setCrosshairPaint(Color.BLUE);// 不确定含义
        chartTheme.setAxisLabelPaint(new Color(51, 51, 51));// 坐标轴标题文字颜色
        chartTheme.setTickLabelPaint(new Color(67, 67, 72));// 刻度数字
        chartTheme.setBarPainter(new StandardBarPainter());// 设置柱状图渲染
        chartTheme.setXYBarPainter(new StandardXYBarPainter());// XYBar 渲染

        chartTheme.setItemLabelPaint(Color.black);
        chartTheme.setThermometerPaint(Color.white);// 温度计

        ChartFactory.setChartTheme(chartTheme);
    }

    public static byte[] createBarChart(String title, String categoygAxisLable, String valueAxisLabel, CategoryLabelPositions CategoryLabelPositions, DefaultCategoryDataset dataset, int width, int height) {
        return createBarChart(title, categoygAxisLable, valueAxisLabel, CategoryLabelPositions, dataset, width, height, false);
    }

    public static byte[] createBarChartDT(String title, String categoygAxisLable, String valueAxisLabel, CategoryLabelPositions CategoryLabelPositions, DefaultCategoryDataset dataset, int width, int height) {
        return createBarChartDT(title, categoygAxisLable, valueAxisLabel, CategoryLabelPositions, dataset, width, height, false);
    }

    public static byte[] createBarChart(String title, String categoygAxisLable, String valueAxisLabel, CategoryLabelPositions CategoryLabelPositions, DefaultCategoryDataset dataset, int width, int height, boolean allDataIsZero) {
        JFreeChart chart = ChartFactory.createBarChart(title, categoygAxisLable, valueAxisLabel, dataset);
        ChartUtils.setAntiAlias(chart);
        ChartUtils.setBarRenderer(chart.getCategoryPlot(), false);//
        ChartUtils.setXAixs(chart.getCategoryPlot());
        ChartUtils.setYAixs(chart.getCategoryPlot(), allDataIsZero);
        CategoryPlot plotBar = chart.getCategoryPlot();
        NumberAxis na = (NumberAxis) plotBar.getRangeAxis();
        na.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

        chart.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions);

        return chart2Bytes(width, height, chart);
    }
    public static byte[] createBarChartDT(String title, String categoygAxisLable, String valueAxisLabel, CategoryLabelPositions CategoryLabelPositions, DefaultCategoryDataset dataset, int width, int height, boolean allDataIsZero) {
        JFreeChart chart = ChartFactory.createBarChart(title, categoygAxisLable, valueAxisLabel, dataset);
        ChartUtils.setAntiAlias(chart);
        ChartUtils.setBarRendererDT(chart.getCategoryPlot(), false);//
        ChartUtils.setXAixs(chart.getCategoryPlot());
        ChartUtils.setYAixs(chart.getCategoryPlot(), allDataIsZero);
        CategoryPlot plotBar = chart.getCategoryPlot();
        NumberAxis na = (NumberAxis) plotBar.getRangeAxis();
        na.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

        chart.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions);

        return chart2Bytes(width, height, chart);
    }

    public static byte[] createLineChart(String title, String categoygAxisLable, String valueAxisLabel, CategoryLabelPositions CategoryLabelPositions, DefaultCategoryDataset dataset, int width, int height) {
        JFreeChart chart = ChartFactory.createLineChart(title, categoygAxisLable, valueAxisLabel, dataset);
        ChartUtils.setAntiAlias(chart);
        ChartUtils.setLineRender(chart.getCategoryPlot(), false,true);//
        ChartUtils.setXAixs(chart.getCategoryPlot());
        ChartUtils.setYAixs(chart.getCategoryPlot());
        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

        CategoryPlot plotBar = chart.getCategoryPlot();

        LineAndShapeRenderer renderer = (LineAndShapeRenderer) plotBar.getRenderer();
        renderer.setUseFillPaint(true);//设置线条是否被显示填充颜色
        renderer.setDrawOutlines(true);//设置拐点不同用不同的形状
        renderer.setBaseShapesVisible(true);
        plotBar.setRangeGridlinesVisible(true); //是否显示格子线

        NumberAxis na = (NumberAxis) plotBar.getRangeAxis();
        na.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

        chart.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions);

        return chart2Bytes(width, height, chart);
    }

    private static byte[] chart2Bytes(int width, int height, JFreeChart chart) {
        byte[] bytes = new byte[0];
        try {
            ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
            ChartUtilities.writeChartAsJPEG(byteOutputStream, 1f, chart, width, height, null);
            bytes = byteOutputStream.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return bytes;
    }

    /**
     * 创建x轴为时间折线图
     *
     * @param title
     * @param categoygAxisLable
     * @param valueAxisLabel
     * @param dataset
     * @param width
     * @param height
     * @Title: createBarChartForDateX
     * @return: void
     */
    public static byte[] createLineChartForDateX(String title, String categoygAxisLable,
                                                String valueAxisLabel, TimeSeriesCollection dataset, int width, int height,
                                                boolean defaultYData,DateTickUnit dateTickUnit) {
        JFreeChart chart=ChartFactory.createTimeSeriesChart(title, categoygAxisLable, valueAxisLabel, dataset, true, true, true);

        XYPlot plotBar = chart.getXYPlot();

        plotBar.setNoDataMessage(NO_DATA_MSG);
        plotBar.setInsets(new RectangleInsets(10, 10, 5, 10));
        AbstractXYItemRenderer renderer = (XYLineAndShapeRenderer) plotBar.getRenderer();
        setBarColor(renderer);

        Color lineColor = new Color(220, 220, 220);
        plotBar.getDomainAxis().setAxisLinePaint(lineColor);// X坐标轴颜色
        plotBar.getDomainAxis().setTickMarkPaint(lineColor);// X坐标轴标记|竖线颜色

        NumberAxis axis = (NumberAxis) plotBar.getRangeAxis();
        axis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        axis.setAxisLinePaint(lineColor);// Y坐标轴颜色
        axis.setTickMarkPaint(lineColor);// Y坐标轴标记|竖线颜色
        // Y刻度
        axis.setAxisLineVisible(true);
        axis.setTickMarksVisible(true);
        //显示正向箭头
        axis.setPositiveArrowVisible(true);

        Font font = new Font("宋体", Font.PLAIN, 18);
        chart.getTitle().setFont(font);
        NumberAxis na = (NumberAxis) plotBar.getRangeAxis();
        ValueAxis xAxis = plotBar.getDomainAxis();
        xAxis.setUpperMargin(0.15);
        na.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        na.setLabelFont(font);
        xAxis.setLabelFont(font);

        if (defaultYData) {
            na.setLowerBound(0);
            na.setUpperBound(1);
        }

        // 这里是关键
        XYPlot xyplot = (XYPlot) chart.getPlot();
        DateAxis domainAxis = (DateAxis) xyplot.getDomainAxis(); // x轴设置
        domainAxis.setTickUnit(dateTickUnit);
        return chart2Bytes(width, height, chart);
    }

    /**
     * 设置柱状图颜色
     *
     * @param renderer
     * @Title: setBarColor
     * @Description: TODO
     * @return: void
     */
    public static void setBarColor(AbstractRenderer renderer) {
        renderer.setSeriesPaint(0, new Color(42, 183, 183));
        renderer.setSeriesPaint(1, new Color(4, 183, 26));
        renderer.setSeriesPaint(2, new Color(1, 14, 183));
        renderer.setSeriesPaint(3, new Color(183, 8, 1));
        renderer.setSeriesPaint(4, new Color(183, 178, 5));
        renderer.setSeriesPaint(5, new Color(135, 8, 183));
    }

    //报告柱状图
    public static void setBarColorDT(AbstractRenderer renderer) {
        renderer.setSeriesPaint(0, new Color(0, 187, 230));
    }

    public static byte[] createPieChart(String title, DefaultPieDataset dataset, int width, int height) {
        JFreeChart chart = ChartFactory.createPieChart(title, dataset);
        ChartUtils.setAntiAlias(chart);
        ChartUtils.setPieRender(chart.getPlot());
//        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));
//        chart.getLegend().setPosition(RectangleEdge.RIGHT);
        chart.removeLegend();
        ChartUtils.setNumberFormat(chart);

        return chart2Bytes(width, height, chart);
    }

    /**
     * 设置数据小数点位数
     */
    public static void setNumberFormat(JFreeChart chart) {
        PiePlot pieplot = (PiePlot) chart.getPlot();
        pieplot.setLabelFont(new Font("宋体", 0, 12));
        pieplot.setNoDataMessage("无数据");
        pieplot.setCircular(true);
        pieplot.setLabelGap(0.002D);
        pieplot.setLabelGenerator(new StandardPieSectionLabelGenerator("{0}:{2}",
                NumberFormat.getNumberInstance(), new DecimalFormat("0.000%")));
    }

    /**
     * 必须设置文本抗锯齿
     */
    public static void setAntiAlias(JFreeChart chart) {
        chart.setTextAntiAlias(false);

    }

    /**
     * 设置图例无边框，默认黑色边框
     */
    public static void setLegendEmptyBorder(JFreeChart chart) {
        chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

    }

    /**
     * 创建类别数据集合
     */
    public static DefaultCategoryDataset createDefaultCategoryDataset(ArrayList<Serie> series, String[] categories) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        for (Serie serie : series) {
            String name = serie.getName();
            Vector<Object> data = serie.getData();
            if (data != null && categories != null && data.size() == categories.length) {
                for (int index = 0; index < data.size(); index++) {
                    String value = data.get(index) == null ? "" : formatDoubleNumber((Double) data.get(index));
                    if (isPercent(value)) {
                        value = value.substring(0, value.length() - 1);
                    }
                    if (isNumber(value)) {
                        dataset.setValue(Double.parseDouble(value), name, categories[index]);
                    }
                }
            }

        }
        return dataset;

    }

    /**
     * 创建饼图数据集合
     */
    public static DefaultPieDataset createDefaultPieDataset(String[] categories, Object[] datas) {
        DefaultPieDataset dataset = new DefaultPieDataset();
        for (int i = 0; i < categories.length && categories != null; i++) {
            String value = datas[i].toString();
            if (isPercent(value)) {
                value = value.substring(0, value.length() - 1);
            }
            if (isNumber(value)) {
                dataset.setValue(categories[i], Double.valueOf(value));
            }
        }
        return dataset;

    }

    /**
     * 创建饼图数据集合
     */
    public static DefaultPieDataset createDefaultPieDataset(Map<String, Integer> datas) {
        DefaultPieDataset dataset = new DefaultPieDataset();
        for (Map.Entry<String, Integer> entry : datas.entrySet()) {
            dataset.setValue(entry.getKey(), entry.getValue());
        }
        return dataset;

    }

    public static DefaultPieDataset createDefaultPieDatasetForLong(Map<String, Long> datas) {
        DefaultPieDataset dataset = new DefaultPieDataset();
        for (Map.Entry<String, Long> entry : datas.entrySet()) {
            dataset.setValue(entry.getKey(), entry.getValue());
        }
        return dataset;

    }

    /**
     *  //曲线图数据集
     * @param xAxis  x轴
     * @param data  曲线对应的数据 key曲线名 value数据列表
     * @return
     */
    public static DefaultCategoryDataset createLinePicDataset(List<String> xAxis,Map<String,List<Number>> data) {

        DefaultCategoryDataset linedataset = new DefaultCategoryDataset();

        //series 曲线名
        for (String series:data.keySet()){
            List<Number> val = data.get(series);
            for (int i = 0;i<xAxis.size()&&i<val.size();i++){
                linedataset.addValue(val.get(i),series,xAxis.get(i));
            }
        }
        return linedataset;
    }

    public static TimeSeriesCollection createTimeSeriesCollection(List<Date> timeList, Map<String, List<Number>> data) {

        TimeSeriesCollection timeSeriesCollection = new TimeSeriesCollection();
        for (Map.Entry<String, List<Number>> entry : data.entrySet()) {
            String lineName = entry.getKey();
            TimeSeries timeSeries = new TimeSeries(lineName, Minute.class);
            List<Number> series = entry.getValue();
            for (int i = 0 ; i < timeList.size() && i < series.size(); i++) {
                timeSeries.add(new Minute(timeList.get(i)), series.get(i));
            }
            timeSeriesCollection.addSeries(timeSeries);
        }
        return timeSeriesCollection;
    }

    /**
     * 堆积柱状图数据集
     * @param rowKeys
     * @param columnKeys
     * @param data
     * @return
     */
    public static CategoryDataset createCategoryDataset(List<String> rowKeys, List<String> columnKeys, List<List<Number>> data) {
        if (ArrayUtilities.hasDuplicateItems(rowKeys.toArray())) {
            throw new IllegalArgumentException("Duplicate items in 'rowKeys'.");
        } else if (ArrayUtilities.hasDuplicateItems(columnKeys.toArray())) {
            throw new IllegalArgumentException("Duplicate items in 'columnKeys'.");
        } else if (rowKeys.size() != data.size()) {
            throw new IllegalArgumentException("The number of row keys does not match the number of rows in the data array.");
        } else {
            int columnCount = 0;

            for(int r = 0; r < data.size(); ++r) {
                columnCount = Math.max(columnCount, data.get(r).size());
            }

            if (columnKeys.size() != columnCount) {
                throw new IllegalArgumentException("The number of column keys does not match the number of columns in the data array.");
            } else {
                DefaultCategoryDataset result = new DefaultCategoryDataset();

                for(int r = 0; r < data.size(); ++r) {
                    String rowKey = rowKeys.get(r);

                    for(int c = 0; c < data.get(r).size(); ++c) {
                        String columnKey = columnKeys.get(c);
                        result.addValue(data.get(r).get(c), rowKey, columnKey);
                    }
                }

                return result;
            }
        }
    }

    /**
     * 折线图
     * @param linedataset  //s数据集
     * @param title         //图表名
     * @param xAxisLabel    //x轴名称
     * @param yAxisLabel    //y轴名称
     * @param width         //宽
     * @param height        //高
     * @return
     */
    public static byte[] createLineChart(DefaultCategoryDataset linedataset,String title,
                                         String xAxisLabel,String yAxisLabel,CategoryLabelPositions labelPositions,int width,int height){
        // 定义图表对象
        JFreeChart chart = ChartFactory.createLineChart(title, //折线图名称
                xAxisLabel, // 横坐标名称
                yAxisLabel, // 纵坐标名称
                linedataset, // 数据
                PlotOrientation.VERTICAL, // 水平显示图像
                true, // include legend
                true, // tooltips
                false // urls
        );
        CategoryPlot plot = chart.getCategoryPlot();
        LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();
        renderer.setSeriesPaint(0,Color.RED);
        renderer.setSeriesPaint(1,Color.BLUE);
        renderer.setSeriesPaint(2,Color.GREEN);
        renderer.setSeriesPaint(3,Color.ORANGE);
        renderer.setSeriesPaint(4,Color.BLACK);
        renderer.setSeriesPaint(5,Color.PINK);
        renderer.setSeriesPaint(7,Color.MAGENTA);
        renderer.setUseFillPaint(true);//设置线条是否被显示填充颜色
        renderer.setDrawOutlines(true);//设置拐点不同用不同的形状
        renderer.setBaseShapesVisible(true);
        plot.setRangeGridlinesVisible(true); //是否显示格子线
        //plot.setBackgroundAlpha(0.3f); //设置背景透明度
        NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
        rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        rangeAxis.setAutoRangeIncludesZero(true);
        rangeAxis.setUpperMargin(0.20);
        rangeAxis.setLabelAngle(Math.PI / 2.0);
        if (labelPositions != null){
            chart.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(labelPositions);
        }
        return chart2Bytes(width, height, chart);
    }

    /**
     *
     * @param chartTitle 标题
     * @param xName     x轴
     * @param yName     y轴
     * @param dataset   数据集
     * @param width
     * @param height
     * @return
     */
    public static byte[] createStackedBarChart(String chartTitle, String xName,
                                               String yName, CategoryDataset dataset,CategoryLabelPositions categoryLabelPositions,int width,int height){
        //JFreeChart对象
        JFreeChart chart = ChartFactory.createStackedBarChart(
                chartTitle, //图表标题
                xName, //目录轴的显示标签
                yName, //数值轴的显示标签
                dataset, //数据集
                PlotOrientation.VERTICAL, //图表方向：水平、垂直
                true, //是否显示图例
                false, //是否生成工具
                false //是否生成URL链接
        );
        //图例字体清晰
        chart.setTextAntiAlias(false);
        //图表背景色
        chart.setBackgroundPaint(Color.WHITE);
        //建立图表标题字体
        Font titleFont = new Font("宋体", Font.BOLD, 14);
        //建立图表图例字体
        Font legendFont = new Font("宋体", Font.PLAIN, 12);
        //建立x,y轴坐标的字体
        Font axisFont = new Font("SansSerif", Font.TRUETYPE_FONT, 12);
        //设置图表的标题字体
        chart.getTitle().setFont(titleFont);
        //设置图表的图例字体
        chart.getLegend().setItemFont(legendFont);

        //Plot对象是图形的绘制结构对象
        //CategoryPlot plot = chart.getCategoryPlot();

        //设置横虚线可见
        //plot.setRangeGridlinesVisible(true);
        //虚线色彩
        //plot.setRangeGridlinePaint(Color.black);
        //设置柱的透明度(如果是3D的必须设置才能达到立体效果，如果是2D的设置则使颜色变淡)
        //plot.setForegroundAlpha(0.65f);

        /** ----------  RangeAxis (范围轴，相当于 y 轴)---------- **/
        //数据轴精度
        //NumberAxis numberAxis = (NumberAxis) plot.getRangeAxis();
        //numberAxis.setLabelFont(axisFont);//轴标题字体
        //numberAxis.setTickLabelFont(axisFont);//轴数值字体
        //设置最高的一个 Item 与图片顶端的距离
        //numberAxis.setUpperMargin(0.15);
        //设置最低的一个 Item 与图片底端的距离
        //numberAxis.setLowerMargin(0.15);
        //设置最大值是1
        //numberAxis.setUpperBound(1);
        //设置数据轴坐标从0开始
        //numberAxis.setAutoRangeIncludesZero(true);
        //数据显示格式是百分比
        //DecimalFormat df = new DecimalFormat("0.00%");
        //数据轴数据标签的显示格式
        //numberAxis.setNumberFormatOverride(df);

        /** ----------  DomainAxis (区域轴，相当于 x 轴)---------- **/
        //CategoryAxis domainAxis = (CategoryAxis)plot.getDomainAxis();
        //domainAxis.setLabelFont(axisFont);//轴标题字体
        //domainAxis.setTickLabelFont(axisFont);//轴数值字体

        // x轴坐标太长，建议设置倾斜，如下两种方式选其一，两种效果相同
        // 倾斜（1）横轴上的 Lable 45度倾斜
        // domainAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
        // 倾斜（2）Lable（Math.PI 3.0）度倾斜
        // domainAxis.setCategoryLabelPositions(CategoryLabelPositions
        // .createUpRotationLabelPositions(Math.PI / 3.0));

        //横轴上的 Lable 是否完整显示
        //domainAxis.setMaximumCategoryLabelWidthRatio(1f);

        /** ----------  Renderer (图形绘制单元)---------- **/
        //Renderer 对象是图形的绘制单元
        StackedBarRenderer renderer = new StackedBarRenderer();
        //设置柱子宽度
        //renderer.setMaximumBarWidth(0.05D);
        //设置柱子高度
        //renderer.setMinimumBarLength(0.1D);
        //设置柱的边框颜色
        //renderer.setBaseOutlinePaint(Color.BLACK);
        //设置柱的边框可见
        //renderer.setDrawBarOutline(true);
        //设置柱的颜色(可设定也可默认)
        renderer.setSeriesPaint(0, new Color(0, 255, 0));
        renderer.setSeriesPaint(1, new Color(0, 0, 255));
        //设置每个平行柱的之间距离
        //renderer.setItemMargin(0.4);

        // 显示每个柱的数值，并修改该数值的字体属性
        //renderer.setIncludeBaseInRange(true);
        //renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        //renderer.setBaseItemLabelsVisible(true);

        //plot.setRenderer(renderer);
        if (categoryLabelPositions != null){
            chart.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(categoryLabelPositions);
        }
        return chart2Bytes(width, height, chart);
    }

    /**
     * 创建时间序列数据
     *
     * @param category   类别
     * @param dateValues 日期-值 数组
     * @return
     */
    public static TimeSeries createTimeseries(String category, ArrayList<Object[]> dateValues) {
        TimeSeries timeseries = new TimeSeries(category);

        if (dateValues != null) {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            SimpleDateFormat complexFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm");
            int type = 0;
            for (Object[] objects : dateValues) {
                Date date = null;
                String buf = objects[0].toString();
                if (buf.length() > 11) {
                    type = 1;
                    try {
                        date = complexFormat.parse(objects[0].toString());
                    } catch (ParseException e) {
                    }
                } else {
                    try {
                        date = dateFormat.parse(objects[0].toString());
                    } catch (ParseException e) {
                    }
                }
                String sValue = objects[1].toString();
                long dValue = 0;
                if (date != null && isNumber(sValue)) {
                    dValue = Long.parseLong(sValue);
                    if (type == 1) {
                        timeseries.add(new Minute(date), dValue);
                    } else {
                        timeseries.add(new Day(date), dValue);
                    }
                }
            }
        }

        return timeseries;
    }

    /**
     * 设置 折线图样式
     *
     * @param plot
     * @param isShowDataLabels 是否显示数据标签 默认不显示节点形状
     */
    public static void setLineRender(CategoryPlot plot, boolean isShowDataLabels) {
        setLineRender(plot, isShowDataLabels, false);
    }

    /**
     * 设置折线图样式
     *
     * @param plot
     * @param isShowDataLabels 是否显示数据标签
     */
    public static void setLineRender(CategoryPlot plot, boolean isShowDataLabels, boolean isShapesVisible) {
        plot.setNoDataMessage(NO_DATA_MSG);
        plot.setInsets(new RectangleInsets(10, 10, 0, 10), false);
        LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();

        renderer.setStroke(new BasicStroke(1.5F));
        if (isShowDataLabels) {
            renderer.setBaseItemLabelsVisible(true);
            renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator(StandardCategoryItemLabelGenerator.DEFAULT_LABEL_FORMAT_STRING,
                    NumberFormat.getInstance()));
            renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_CENTER));// weizhi
        }
        renderer.setBaseShapesVisible(isShapesVisible);// 数据点绘制形状
        setXAixs(plot);
        setYAixs(plot);

    }

    /**
     * 设置时间序列图样式
     *
     * @param plot
     * @param isShowData      是否显示数据
     * @param isShapesVisible 是否显示数据节点形状
     */
    public static void setTimeSeriesRender(Plot plot, boolean isShowData, boolean isShapesVisible) {

        XYPlot xyplot = (XYPlot) plot;
        xyplot.setNoDataMessage(NO_DATA_MSG);
        xyplot.setInsets(new RectangleInsets(10, 10, 5, 10));

        XYLineAndShapeRenderer xyRenderer = (XYLineAndShapeRenderer) xyplot.getRenderer();

        xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
        xyRenderer.setBaseShapesVisible(false);
        if (isShowData) {
            xyRenderer.setBaseItemLabelsVisible(true);
            xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
            xyRenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_CENTER));// weizhi
        }
        xyRenderer.setBaseShapesVisible(isShapesVisible);// 数据点绘制形状

        DateAxis domainAxis = (DateAxis) xyplot.getDomainAxis();
        domainAxis.setAutoTickUnitSelection(false);
        DateTickUnit dateTickUnit = dateTickUnit = new DateTickUnit(DateTickUnitType.HOUR, 1, new SimpleDateFormat("yyyy-MM")); // 第二个参数是时间轴间距
        domainAxis.setTickUnit(dateTickUnit);

        StandardXYToolTipGenerator xyTooltipGenerator = new StandardXYToolTipGenerator("{1}:{2}", new SimpleDateFormat("yyyy-MM-dd"), new DecimalFormat("0"));
        xyRenderer.setBaseToolTipGenerator(xyTooltipGenerator);

        setXY_XAixs(xyplot);
        setXY_YAixs(xyplot);

    }

    public static void setTimeSeriesRenderWithTimeType(Plot plot, boolean isShowData, boolean isShapesVisible, int dateType) {

        XYPlot xyplot = (XYPlot) plot;
        xyplot.setNoDataMessage(NO_DATA_MSG);
        xyplot.setInsets(new RectangleInsets(10, 10, 5, 10));

        XYLineAndShapeRenderer xyRenderer = (XYLineAndShapeRenderer) xyplot.getRenderer();

        xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
        xyRenderer.setBaseShapesVisible(false);
        if (isShowData) {
            xyRenderer.setBaseItemLabelsVisible(true);
            xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
            xyRenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_CENTER));// weizhi
        }
        xyRenderer.setBaseShapesVisible(isShapesVisible);// 数据点绘制形状

        DateAxis domainAxis = (DateAxis) xyplot.getDomainAxis();
        domainAxis.setVerticalTickLabels(true);
        domainAxis.setAutoTickUnitSelection(false);
        DateTickUnit dateTickUnit = null;
        if (dateType == 1) {
            dateTickUnit = new DateTickUnit(DateTickUnitType.HOUR, 1, new SimpleDateFormat("dd-HH")); // 第二个参数是时间轴间距
        } else {
            dateTickUnit = new DateTickUnit(DateTickUnitType.DAY, 1, new SimpleDateFormat("MM-dd")); // 第二个参数是时间轴间距
        }
        domainAxis.setTickUnit(dateTickUnit);

        NumberAxis numberAxis = (NumberAxis) xyplot.getRangeAxis();
        NumberFormat numberFormat = new DecimalFormat("#.###");
        numberAxis.setNumberFormatOverride(numberFormat);


        StandardXYToolTipGenerator xyTooltipGenerator = new StandardXYToolTipGenerator("{1}:{2}", new SimpleDateFormat("yyyy-MM-dd"), new DecimalFormat("0"));
        xyRenderer.setBaseToolTipGenerator(xyTooltipGenerator);

        setXY_XAixs(xyplot);
        setXY_YAixs(xyplot);

    }

    /**
     * 设置时间序列图样式 -默认不显示数据节点形状
     *
     * @param plot
     * @param isShowData 是否显示数据
     */

    public static void setTimeSeriesRender(Plot plot, boolean isShowData) {
        setTimeSeriesRender(plot, isShowData, false);
    }

    /**
     * 设置时间序列图渲染：但是存在一个问题：如果timeseries里面的日期是按照天组织， 那么柱子的宽度会非常小，和直线一样粗细
     *
     * @param plot
     * @param isShowDataLabels
     */

    public static void setTimeSeriesBarRender(Plot plot, boolean isShowDataLabels) {

        XYPlot xyplot = (XYPlot) plot;
        xyplot.setNoDataMessage(NO_DATA_MSG);

        XYBarRenderer xyRenderer = new XYBarRenderer(0.1D);
        xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());

        if (isShowDataLabels) {
            xyRenderer.setBaseItemLabelsVisible(true);
            xyRenderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
        }

        StandardXYToolTipGenerator xyTooltipGenerator = new StandardXYToolTipGenerator("{1}:{2}", new SimpleDateFormat("yyyy-MM-dd"), new DecimalFormat("0"));
        xyRenderer.setBaseToolTipGenerator(xyTooltipGenerator);
        setXY_XAixs(xyplot);
        setXY_YAixs(xyplot);

    }

    /**
     * 设置柱状图渲染
     *
     * @param plot
     * @param isShowDataLabels
     */
    public static void setBarRenderer(CategoryPlot plot, boolean isShowDataLabels) {
        plot.setNoDataMessage(NO_DATA_MSG);
        plot.setInsets(new RectangleInsets(10, 10, 5, 10));
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setMaximumBarWidth(0.085);// 设置柱子最大宽度
        setBarColor(renderer);
        if (isShowDataLabels) {
            renderer.setBaseItemLabelsVisible(true);
        }

        setXAixs(plot);
        setYAixs(plot);
    }

    /**
     * 设置报告柱状图渲染
     *
     * @param plot
     * @param isShowDataLabels
     */
    public static void setBarRendererDT(CategoryPlot plot, boolean isShowDataLabels) {
        plot.setNoDataMessage(NO_DATA_MSG);
        plot.setInsets(new RectangleInsets(10, 10, 5, 10));
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setMaximumBarWidth(0.085);// 设置柱子最大宽度
        setBarColorDT(renderer);
        if (isShowDataLabels) {
            renderer.setBaseItemLabelsVisible(true);
        }

        setXAixs(plot);
        setYAixs(plot);
    }
    /**
     * 设置堆积柱状图渲染
     *
     * @param plot
     */

    public static void setStackBarRender(CategoryPlot plot) {
        plot.setNoDataMessage(NO_DATA_MSG);
        plot.setInsets(new RectangleInsets(10, 10, 5, 10));
        StackedBarRenderer renderer = (StackedBarRenderer) plot.getRenderer();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        plot.setRenderer(renderer);
        setXAixs(plot);
        setYAixs(plot);
    }

    /**
     * 设置类别图表(CategoryPlot) X坐标轴线条颜色和样式
     *
     * @param plot
     */
    public static void setXAixs(CategoryPlot plot) {
        Color lineColor = new Color(220, 220, 220);
        plot.getDomainAxis().setAxisLinePaint(lineColor);// X坐标轴颜色
        plot.getDomainAxis().setTickMarkPaint(lineColor);// X坐标轴标记|竖线颜色
    }

    /**
     * 设置类别图表(CategoryPlot) Y坐标轴线条颜色和样式 同时防止数据无法显示
     *
     * @param plot
     */

    public static void setYAixs(CategoryPlot plot) {
        setYAixs(plot, false);
    }

    public static void setYAixs(CategoryPlot plot, boolean allDataIsZero) {
        Color lineColor = new Color(220, 220, 220);
//		ValueAxis axis = plot.getRangeAxis();
        NumberAxis axis = (NumberAxis) plot.getRangeAxis();
        axis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        axis.setAxisLinePaint(lineColor);// Y坐标轴颜色
        axis.setTickMarkPaint(lineColor);// Y坐标轴标记|竖线颜色
        // Y刻度
        axis.setAxisLineVisible(true);
        axis.setTickMarksVisible(true);
        if (allDataIsZero) {
            axis.setLowerBound(0);
            axis.setUpperBound(10);
        }
        //显示正向箭头
        axis.setPositiveArrowVisible(true);
        // Y轴网格线条
//		plot.setRangeGridlinePaint(new Color(192, 192, 192));
//		plot.setRangeGridlineStroke(new BasicStroke(1));

//		plot.getRangeAxis().setUpperMargin(0.1);// 设置顶部Y坐标轴间距,防止数据无法显示
//		plot.getRangeAxis().setLowerMargin(0.1);// 设置底部Y坐标轴间距

    }

    /**
     * 设置XY图表(XYPlot) X坐标轴线条颜色和样式
     *
     * @param plot
     */
    public static void setXY_XAixs(XYPlot plot) {
        Color lineColor = new Color(220, 220, 220);
        plot.getDomainAxis().setAxisLinePaint(lineColor);// X坐标轴颜色
        plot.getDomainAxis().setTickMarkPaint(lineColor);// X坐标轴标记|竖线颜色

    }

    /**
     * 设置XY图表(XYPlot) Y坐标轴线条颜色和样式 同时防止数据无法显示
     *
     * @param plot
     */
    public static void setXY_YAixs(XYPlot plot) {
        Color lineColor = new Color(192, 208, 224);
        ValueAxis axis = plot.getRangeAxis();
        axis.setAxisLinePaint(lineColor);// X坐标轴颜色
        axis.setTickMarkPaint(lineColor);// X坐标轴标记|竖线颜色
        // 隐藏Y刻度
        axis.setAxisLineVisible(false);
        axis.setTickMarksVisible(false);
        // Y轴网格线条
        plot.setRangeGridlinePaint(new Color(192, 192, 192));
        plot.setRangeGridlineStroke(new BasicStroke(1));
        plot.setDomainGridlinesVisible(false);

        plot.getRangeAxis().setUpperMargin(0.12);// 设置顶部Y坐标轴间距,防止数据无法显示
        plot.getRangeAxis().setLowerMargin(0.12);// 设置底部Y坐标轴间距

    }

    /**
     * 设置饼状图渲染
     */
    public static void setPieRender(Plot plot) {

        plot.setNoDataMessage(NO_DATA_MSG);
        plot.setInsets(new RectangleInsets(10, 10, 5, 10));
        PiePlot piePlot = (PiePlot) plot;
        piePlot.setInsets(new RectangleInsets(0, 0, 0, 0));
        piePlot.setCircular(true);// 圆形

        // piePlot.setSimpleLabels(true);// 简单标签
        piePlot.setLabelGap(0.01);
        piePlot.setInteriorGap(0.05D);
        piePlot.setLegendItemShape(new Rectangle(10, 10));// 图例形状
        piePlot.setIgnoreNullValues(true);
        piePlot.setLabelBackgroundPaint(null);// 去掉背景色
        piePlot.setLabelShadowPaint(null);// 去掉阴影
        piePlot.setLabelOutlinePaint(null);// 去掉边框
        piePlot.setShadowPaint(null);
        // 0:category 1:value:2 :percentage
        piePlot.setLabelGenerator(new StandardPieSectionLabelGenerator("{0}:{2}"));// 显示标签数据
    }

    /**
     * 格式化数据为柱状图数据源
     *
     * @param list
     * @return
     */
    public static DefaultCategoryDataset formatBarDcd(List<Map<String, Object>> list, String legend) {
        return formatBarDcd(list, legend, "name", "value");
    }

    public static DefaultCategoryDataset formatBarDcd(List<Map<String, Object>> list, String legend, String name, String value) {
        ArrayList<Serie> series = new ArrayList<>();
        String[] categories = new String[list.size()];
        Double[] values = new Double[list.size()];
        for (int i = 0; i < list.size(); i++) {
            Map<String, Object> map = list.get(i);
            categories[i] = (String) map.get(name);
            values[i] = castObj2Double(map.get(value));
        }
        series.add(new Serie(legend, values));
        DefaultCategoryDataset dataset = ChartUtils.createDefaultCategoryDataset(series, categories);
        return dataset;
    }
    public static DefaultCategoryDataset formatBarsDcd(Map<String, List<Map<String, Object>>> maplist, String name, String value) {
        ArrayList<Serie> series = new ArrayList<>();
        String[] categories = null;
        for (Map.Entry<String, List<Map<String, Object>>> listEntry : maplist.entrySet()) {
            String legend = listEntry.getKey();
            List<Map<String, Object>> list = listEntry.getValue();
            if (categories == null){
                categories = new String[list.size()];
            }
            Double[] values = new Double[list.size()];
            for (int i = 0; i < list.size(); i++) {
                Map<String, Object> map = list.get(i);
                categories[i] = (String) map.get(name);
                values[i] = castObj2Double(map.get(value));
            }
            series.add(new Serie(legend, values));
        }
        DefaultCategoryDataset dataset = ChartUtils.createDefaultCategoryDataset(series, categories);
        return dataset;
    }

    public static TimeSeriesCollection formatTimeLineTsc(List<Map<String, Object>> list, String legend, String name, String value) {
        TimeSeriesCollection dataset = new TimeSeriesCollection();
        TimeSeries series = new TimeSeries(legend);

        String[] categories = new String[list.size()];
        Double[] values = new Double[list.size()];
        for (int i = 0; i < list.size(); i++) {
            Map<String, Object> map = list.get(i);
            categories[i] = (String) map.get(name);
            values[i] = castObj2Double(map.get(value));
            series.add(new Minute(DateUtil.toDate(categories[i], "yyyy-MM-dd HH:mm:ss")),values[i]);
        }
        dataset.addSeries(series);
        return dataset;
    }

    /**
     * 是不是一个%形式的百分比
     *
     * @param str
     * @return
     */
    public static boolean isPercent(String str) {
        return str != null ? str.endsWith("%") && isNumber(str.substring(0, str.length() - 1)) : false;
    }

    /**
     * 是不是一个数字
     *
     * @param str
     * @return
     */
    public static boolean isNumber(String str) {
        return str != null ? str.matches("^[-+]?(([0-9]+)((([.]{0})([0-9]*))|(([.]{1})([0-9]+))))$") : false;
    }

    /**
     * 删除图片
     *
     * @param imgName
     */
    public static void deleteImg(String imgName) {
        File img = null;
        try {
            img = new File(imgName);
            boolean b = img.exists();
            if (b) {
                System.out.println(img.getPath());
                boolean delete = img.delete();
                System.out.println(delete);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static double castObj2Double(Object obj) {
        if (obj instanceof Integer) {
            return ((Integer) obj).doubleValue();
        } else if (obj instanceof Long) {
            return ((Long) obj).doubleValue();
        } else if (obj instanceof Double) {
            return (double) obj;
        } else if (obj instanceof String) {
            return Double.parseDouble((String) obj);
        }
        return 0.0d;
    }

    public static String formatDoubleNumber(double value) {
        String retValue = null;
        DecimalFormat df = new DecimalFormat();
        df.setMinimumFractionDigits(0);
        df.setMaximumFractionDigits(2);
        retValue = df.format(value);
        retValue = retValue.replaceAll(",", "");
        return retValue;
    }

    public static byte[] createTimeSeriesChart(String title, String timeAxisLabel, String valueAxisLabel, TimeSeries series, int width, int height, int dateType) {
        TimeSeriesCollection collection = new TimeSeriesCollection();
        collection.addSeries(series);
        JFreeChart chart = ChartFactory.createTimeSeriesChart(title, timeAxisLabel, valueAxisLabel, collection);
        ChartUtils.setAntiAlias(chart);
        ChartUtils.setTimeSeriesRenderWithTimeType(chart.getPlot(), false, true, dateType);
        return chart2Bytes(width, height, chart);
    }

    //URL判断
    public static boolean containsBadCharURL(String URL){
        Matcher m = urlPattern.matcher(URL);
        if (m.find()) {
            return false;
        } else {
            return true;
        }
    }

}

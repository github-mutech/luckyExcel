package com.report.common.utils.luckyexcel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

@Data
public class LuckySheet {
    public static Set<String> fieldNames = Arrays.stream(LuckySheet.class.getDeclaredFields()).map(Field::getName).collect(Collectors.toSet());
    /**
     * 工作表名称
     */
    private String name;
    /**
     * 工作表颜色
     */
    private String color;
    /**
     * 工作表索引
     */
    private Integer index;
    /**
     * 激活状态
     */
    private Integer status;
    /**
     * 工作表的顺序
     */
    private String order;
    /**
     * 是否隐藏
     */
    private Integer hide;
    /**
     * 行数
     */
    private Integer row;
    /**
     * 列数
     */
    private Integer column;
    /**
     * 配置
     */
    private Config config;
    /**
     * 初始化使用的单元格数据
     */
    private JSONArray celldata;
    /**
     * 更新和存储使用的单元格数据
     */
    private JSONArray data;
    /**
     * 左右滚动条位置
     */
    private Integer scrollLeft;
    /**
     * 上下滚动条位置
     */
    private Integer scrollTop;
    /**
     * 选中的区域
     */
    private JSONArray luckysheet_select_save;
    /**
     * 条件格式
     */
    private JSONArray luckysheet_conditionformat_save;
    /**
     * 公式链
     */
    private JSONArray calcChain;
    /**
     * 是否数据透视表
     */
    private Boolean isPivotTable;
    /**
     * 数据透视表设置
     */
    private JSONObject pivotTable;
    /**
     * 筛选范围
     */
    private JSONObject filter_select;
    /**
     * 筛选配置
     */
    private JSONObject filter;
    /**
     * 交替颜色
     */
    private JSONArray luckysheet_alternateformat_save;
    /**
     * 自定义交替颜色
     */
    private JSONArray luckysheet_alternateformat_save_modelCustom;
    /**
     * 冻结行列
     */
    private JSONObject freezen;
    /**
     * 图表配置
     */
    private JSONArray chart;
    /**
     * 所有行的位置
     */
    private JSONArray visibledatarow;
    /**
     * 所有列的位置
     */
    private JSONArray visibledatacolumn;
    /**
     * 工作表区域的宽度
     */
    private Integer ch_width;
    /**
     * 工作表区域的高度
     */
    private Integer rh_height;
    /**
     * 已加载过此sheet的标识
     */
    private String load;
    /**
     * ? 默认高
     */
    private Integer defaultRowHeight;
    /**
     * ? 默认宽
     */
    private Integer defaultColWidth;


    private String mirror;
    private String dynamicArray;
    private String images;
    /**
     * 数据验证配置
     */
    private String dataVerification;
    private String luckysheet_selection_range;
    /**
     * 缩放比例
     */
    private String zoomRatio;

    public static List<LuckySheet> parseLuckySheetList(String excelData) {
        List<LuckySheet> luckySheets = new ArrayList<>();
        List<JSONObject> jsonObjects = JSON.parseArray(excelData, JSONObject.class);
        for (JSONObject jsonObject : jsonObjects) {
            Set<String> keySet = new HashSet<>(jsonObject.keySet());
            keySet.removeAll(fieldNames);
            System.out.println("没有映射的key" + keySet);
            Set<String> uselessFieldNames = new HashSet<>(fieldNames);
            uselessFieldNames.removeAll(jsonObject.keySet());
            System.out.println("没有用上的字段" + uselessFieldNames);
            LuckySheet luckySheet = new LuckySheet();
            luckySheet.setName(jsonObject.getString("name"));
            luckySheet.setColor(jsonObject.getString("color"));
            luckySheet.setIndex(jsonObject.getInteger("index"));
            luckySheet.setStatus(jsonObject.getInteger("status"));
            luckySheet.setOrder(jsonObject.getString("order"));
            luckySheet.setHide(jsonObject.getInteger("hide"));
            luckySheet.setRow(jsonObject.getInteger("row"));
            luckySheet.setColumn(jsonObject.getInteger("column"));
            luckySheet.setConfig(jsonObject.getObject("config", Config.class));
            luckySheet.setCelldata(jsonObject.getJSONArray("celldata"));
            luckySheet.setData(jsonObject.getJSONArray("data"));
            luckySheet.setScrollLeft(jsonObject.getInteger("scrollLeft"));
            luckySheet.setScrollTop(jsonObject.getInteger("scrollTop"));
            luckySheet.setLuckysheet_select_save(jsonObject.getJSONArray("luckysheet_select_save"));
            luckySheet.setLuckysheet_conditionformat_save(jsonObject.getJSONArray("luckysheet_conditionformat_save"));
            luckySheet.setCalcChain(jsonObject.getJSONArray("jsonObject"));
            luckySheet.setIsPivotTable(jsonObject.getBoolean("isPivotTable"));
            luckySheet.setPivotTable(jsonObject.getJSONObject("pivotTable"));
            luckySheet.setFilter_select(jsonObject.getJSONObject("filter_select"));
            luckySheet.setFilter(jsonObject.getJSONObject("filter"));
            luckySheet.setLuckysheet_alternateformat_save(jsonObject.getJSONArray("luckysheet_alternateformat_save"));
            luckySheet.setLuckysheet_alternateformat_save_modelCustom(jsonObject.getJSONArray("luckysheet_alternateformat_save_modelCustom"));
            luckySheet.setFreezen(jsonObject.getJSONObject("freezen"));
            luckySheet.setChart(jsonObject.getJSONArray("chart"));
            luckySheet.setVisibledatarow(jsonObject.getJSONArray("visibledatarow"));
            luckySheet.setVisibledatacolumn(jsonObject.getJSONArray("visibledatacolumn"));
            luckySheet.setCh_width(jsonObject.getInteger("ch_width"));
            luckySheet.setRh_height(jsonObject.getInteger("rh_height"));
            luckySheet.setLoad(jsonObject.getString("load"));
            luckySheet.setDefaultRowHeight(jsonObject.getInteger("defaultRowHeight"));
            luckySheet.setDefaultColWidth(jsonObject.getInteger("defaultColWidth"));
            luckySheets.add(luckySheet);
        }
        return luckySheets;
    }
}

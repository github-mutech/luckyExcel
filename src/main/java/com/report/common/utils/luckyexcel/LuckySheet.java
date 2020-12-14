package com.report.common.utils.luckyexcel;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Data;

@Data
public class LuckySheet {
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
}

package com.report.common.utils.luckyexcel;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Data;

@Data
public class Config {
    /**
     * 合并单元格
     */
    private JSONObject merge;
    /**
     * 表格行高
     */
    private JSONObject rowlen;
    /**
     * 表格列宽
     */
    private JSONObject columnlen;
    /**
     * 隐藏行
     */
    private JSONObject rowhidden;
    /**
     * 隐藏列
     */
    private JSONObject colhidden;
    /**
     * 边框
     */
    private JSONArray borderInfo;
}
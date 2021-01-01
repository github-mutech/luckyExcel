package com.report.common.utils.luckyexcel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.thymeleaf.util.StringUtils;
import sun.misc.BASE64Decoder;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;
import java.util.regex.Pattern;

/**
 * LuckySheet导入导出工具类
 *
 * @author h
 */
public class LuckyExcelUtils {

    private static final Map<Integer, String> CODE_FONT_MAP = new HashMap<>();
    private static final Map<String, Integer> FONT_CODE_MAP = new HashMap<>();
    /**
     * 设置边框样式map
     */
    private static final Map<Integer, BorderStyle> BORD_MAP = new HashMap<>();
    private static final Pattern PATTERN = Pattern.compile("^[-+]?[\\d]*$");

    static {
        CODE_FONT_MAP.put(-1, "Arial");
        CODE_FONT_MAP.put(0, "Times New Roman");
        CODE_FONT_MAP.put(1, "Arial");
        CODE_FONT_MAP.put(2, "Tahoma");
        CODE_FONT_MAP.put(3, "Verdana");
        CODE_FONT_MAP.put(4, "微软雅黑");
        CODE_FONT_MAP.put(5, "宋体");
        CODE_FONT_MAP.put(6, "黑体");
        CODE_FONT_MAP.put(7, "楷体");
        CODE_FONT_MAP.put(8, "仿宋");
        CODE_FONT_MAP.put(9, "新宋体");
        CODE_FONT_MAP.put(10, "华文新魏");
        CODE_FONT_MAP.put(11, "华文行楷");
        CODE_FONT_MAP.put(12, "华文隶书");
        CODE_FONT_MAP.forEach((k, v) -> FONT_CODE_MAP.put(v, k));
        Arrays.stream(BorderStyle.values()).forEach(borderStyle -> BORD_MAP.put((int) borderStyle.getCode(), borderStyle));
    }

    /**
     * 功能: LuckySheet导出方法
     *
     * @param excelData 数据
     * @param response  用来获取输出流
     * @param request   针对火狐浏览器导出时文件名乱码的问题,也可以不传入此值
     */
    public static void exportExcel(String excelData, HttpServletRequest request, HttpServletResponse response) {
        // 解析对象，可以参照官方文档:https://mengshukeji.github.io/LuckysheetDocs/zh/guide/#%E6%95%B4%E4%BD%93%E7%BB%93%E6%9E%84
        List<LuckySheet> luckySheets = LuckySheet.parseLuckySheetList(excelData);
        XSSFWorkbook workbook = new XSSFWorkbook();
        for (LuckySheet luckySheet : luckySheets) {
            // 默认高
            int defaultRowHeight = luckySheet.getDefaultRowHeight() == null ? 20 : luckySheet.getDefaultRowHeight();
            // 默认宽
            int defaultColWidth = luckySheet.getDefaultColWidth() == null ? 74 : luckySheet.getDefaultColWidth();
            // 读取了模板内所有sheet内容
            XSSFSheet sheet = workbook.createSheet(luckySheet.getName());
            JSONObject columnlen = null;
            JSONObject rowlen = null;
            JSONArray borderInfo = null;
            Config config = luckySheet.getConfig();
            if (config != null) {
                columnlen = config.getColumnlen();
                rowlen = config.getRowlen();
                borderInfo = config.getBorderInfo();
            }
            // 如果这行没有了，整个公式都不会有自动计算的效果的
            sheet.setForceFormulaRecalculation(true);
            // TODO 固定行列
            setFreezePane(sheet, luckySheet.getFreezen());
            // 设置行高列宽
            setCellWH(sheet, columnlen, rowlen);
            // TODO 图片插入
            // setImages(workbook, sheet, null, columnlen, rowlen, defaultRowHeight, defaultColWidth);
            // 设置单元格值及格式
            setCellValue(workbook, sheet, luckySheet.getCelldata(), columnlen, rowlen, defaultRowHeight, defaultColWidth);
            // TODO 设置数据验证
            // settDataValidation(luckySheet.getdataVerification, sheet);
            // 设置边框
            setBorder(borderInfo, sheet);
        }
        // 如果只有一个sheet那就是get(0),有多个那就对应取下标
        try {
            String disposition = "attachment;filename=";
            if (request != null && request.getHeader("USER-AGENT") != null && StringUtils.contains(request.getHeader("USER-AGENT"), "Firefox")) {
                disposition += new String(("template.xlsx").getBytes(), "ISO8859-1");
            } else {
                disposition += URLEncoder.encode("template.xlsx", "UTF-8");
            }
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
            response.setHeader("Content-Disposition", disposition);
            // 修改模板内容导出新模板
            OutputStream out;
            out = response.getOutputStream();
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String importExecl(InputStream is) throws IOException {
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(is);
        List<LuckySheet> luckySheets = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            LuckySheet luckySheet = new LuckySheet();
            luckySheet.setName(sheet.getSheetName());
            JSONArray data = new JSONArray();
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                JSONArray luckyRow = new JSONArray();
                XSSFRow row = (XSSFRow) rowIterator.next();
                Iterator<Cell> iterator = row.cellIterator();
                while (iterator.hasNext()) {
                    XSSFCell cell = (XSSFCell) iterator.next();
                    JSONObject luckyCell = new JSONObject();
                    JSONObject ctValue = new JSONObject();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                DateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy", LocaleUtil.getUserLocale());
                                sdf.setTimeZone(LocaleUtil.getUserTimeZone());
                                sdf.format(cell.getDateCellValue());
                                ctValue.put("fa", null);
                                ctValue.put("t", "n");
                                luckyCell.put("v", cell.getDateCellValue());
                                luckyCell.put("m", sdf.format(cell.getDateCellValue()));
                                break;
                            }
                            ctValue.put("fa", null);
                            ctValue.put("t", "n");
                            luckyCell.put("v", cell.getNumericCellValue());
                            luckyCell.put("m", Double.toString(cell.getNumericCellValue()));
                            break;
                        case STRING:
                            ctValue.put("fa", null);
                            ctValue.put("t", "s");
                            luckyCell.put("v", cell.getRichStringCellValue());
                            luckyCell.put("m", cell.getRichStringCellValue().toString());
                            break;
                        case FORMULA:
                            ctValue.put("fa", null);
                            ctValue.put("t", null);
                            luckyCell.put("v", cell.getCellFormula());
                            luckyCell.put("m", cell.getCellFormula());
                            break;
                        case BLANK:
                            ctValue.put("fa", null);
                            ctValue.put("t", null);
                            luckyCell.put("v", null);
                            luckyCell.put("m", "");
                            break;
                        case BOOLEAN:
                            ctValue.put("fa", null);
                            ctValue.put("t", null);
                            luckyCell.put("v", cell.getBooleanCellValue());
                            luckyCell.put("m", cell.getBooleanCellValue() ? "TRUE" : "FALSE");
                            break;
                        case ERROR:
                            ctValue.put("fa", null);
                            ctValue.put("t", null);
                            luckyCell.put("v", cell.getErrorCellValue());
                            luckyCell.put("m", ErrorEval.getText(cell.getErrorCellValue()));
                            break;
                        default:
                            ctValue.put("fa", null);
                            ctValue.put("t", null);
                            luckyCell.put("v", cell.getCellType());
                            luckyCell.put("m", "Unknown Cell Type: " + cell.getCellType());
                            break;
                    }

                    luckyCell.put("ct", ctValue);

                    XSSFCellStyle cellStyle = cell.getCellStyle();
                    // luckyCell.put("bg", cellStyle.getFillBackgroundColorColor());// TODO
                    luckyCell.put("bl", cellStyle.getFont().getBold() ? 1 : 0);
                    luckyCell.put("it", cellStyle.getFont().getItalic() ? 1 : 0);
                    luckyCell.put("ff", FONT_CODE_MAP.get(cellStyle.getFont().getFontName()));
                    luckyCell.put("fs", cellStyle.getFont().getFontHeightInPoints());
                    luckyCell.put("fc", getRgbValue(cellStyle.getFont().getXSSFColor()));
                    if (HorizontalAlignment.CENTER.equals(cellStyle.getAlignment())) {
                        luckyCell.put("ht", 0);
                    } else if (HorizontalAlignment.LEFT.equals(cellStyle.getAlignment())) {
                        luckyCell.put("ht", 1);
                    } else if (HorizontalAlignment.RIGHT.equals(cellStyle.getAlignment())) {
                        luckyCell.put("ht", 2);
                    }
                    if (VerticalAlignment.CENTER.equals(cellStyle.getVerticalAlignment())) {
                        luckyCell.put("vt", 0);
                    } else if (VerticalAlignment.TOP.equals(cellStyle.getVerticalAlignment())) {
                        luckyCell.put("vt", 1);
                    } else if (VerticalAlignment.BOTTOM.equals(cellStyle.getVerticalAlignment())) {
                        luckyCell.put("vt", 2);
                    }
                    luckyRow.add(luckyCell);
                }
                data.add(luckyRow);
            }
            luckySheet.setData(data);

            luckySheets.add(luckySheet);
        }
        return JSON.toJSONString(luckySheets);
    }

    private static String getRgbValue(XSSFColor xssfColor) {
        byte[] rgb = xssfColor.getRGB();
        return String.format("rgb(%d, %d, %d)", rgb[0] < 0 ? 256 + rgb[0] : rgb[0],
                rgb[1] < 0 ? 256 + rgb[1] : rgb[1], rgb[2] < 0 ? 256 + rgb[2] : rgb[2]);
    }

    private static CellStyle createCellStyle(XSSFSheet sheet, XSSFWorkbook wb, JSONObject v) {
        XSSFCellStyle cellStyle = wb.createCellStyle();
        // 合并单元格
        if (v.get("mc") != null && ((JSONObject) v.get("mc")).get("rs") != null && ((JSONObject) v.get("mc")).get("cs") != null) {
            // 主单元格的行号,开始行号
            int r = Integer.parseInt(((JSONObject) v.get("mc")).get("r").toString());
            // 合并单元格占的行数,合并多少行
            int rs = Integer.parseInt(((JSONObject) v.get("mc")).get("rs").toString());
            // 主单元格的列号,开始列号
            int c = Integer.parseInt(((JSONObject) v.get("mc")).get("c").toString());
            // 合并单元格占的列数,合并多少列
            int cs = Integer.parseInt(((JSONObject) v.get("mc")).get("cs").toString());
            CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, c, c + cs - 1);
            sheet.addMergedRegion(region);
        }
        XSSFFont font = wb.createFont();
        // 字体
        if (v.get("ff") != null) {
            if (v.get("ff").toString().matches("^(-?\\d+)(\\.\\d+)?$")) {
                font.setFontName(CODE_FONT_MAP.get(v.getInteger("ff")));
            } else {
                font.setFontName(v.get("ff").toString());
            }
        }
        // 字体颜色
        if (v.get("fc") != null) {
            String fc = v.get("fc").toString();
            XSSFColor color = toColorFromString(fc);
            font.setColor(color);
        }
        // 粗体
        if (v.get("bl") != null) {
            font.setBold("1".equals(v.get("bl").toString()));
        }
        // 斜体
        if (v.get("it") != null) {
            font.setItalic("1".equals(v.get("it").toString()));
        }
        // 删除线
        if (v.get("cl") != null) {
            font.setStrikeout("1".equals(v.get("cl").toString()));
        }
        // 下滑线
        if (v.get("un") != null) {
            font.setUnderline("1".equals(v.get("un").toString()) ? FontUnderline.SINGLE : FontUnderline.NONE);
        }
        // 字体大小
        if (v.get("fs") != null) {
            font.setFontHeightInPoints(new Short(v.get("fs").toString()));
        }
        cellStyle.setFont(font);
        // 水平对齐
        if (v.get("ht") != null) {
            switch (v.getInteger("ht")) {
                case 0:
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case 1:
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case 2:
                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                default:
                    break;
            }
        }
        // 垂直对齐
        if (v.get("vt") != null) {
            switch (v.getInteger("vt")) {
                case 0:
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    break;
                case 1:
                    cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
                    break;
                case 2:
                    cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    break;
                default:
                    break;
            }
        }
        // 背景颜色
        if (v.get("bg") != null) {
            String bg = v.get("bg").toString();
            cellStyle.setFillForegroundColor(toColorFromString(bg));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        cellStyle.setWrapText(true);
        return cellStyle;
    }

    /**
     * 字符串转换成Color对象
     *
     * @param colorStr 16进制颜色字符串
     * @return Color对象
     */
    private static XSSFColor toColorFromString(String colorStr) {
        if (colorStr.contains("#")) {
            colorStr = colorStr.substring(1);
            Color color = new Color(Integer.parseInt(colorStr, 16));
            return new XSSFColor(color, new DefaultIndexedColorMap());
        } else {
            int strStartIndex = colorStr.indexOf("(");
            int strEndIndex = colorStr.indexOf(")");
            String[] strings = colorStr.substring(strStartIndex + 1, strEndIndex).split(",");
            String R = Integer.toHexString(Integer.parseInt(strings[0].replaceAll(" ", "")));
            R = R.length() < 2 ? ('0' + R) : R;
            String B = Integer.toHexString(Integer.parseInt(strings[1].replaceAll(" ", "")));
            B = B.length() < 2 ? ('0' + B) : B;
            String G = Integer.toHexString(Integer.parseInt(strings[2].replaceAll(" ", "")));
            G = G.length() < 2 ? ('0' + G) : G;
            String cStr = R + B + G;
            Color color1 = new Color(Integer.parseInt(cStr, 16));
            return new XSSFColor(color1, new DefaultIndexedColorMap());
        }
    }

    /**
     * 获取图片位置
     * dx1：起始单元格的x偏移量，如例子中的255表示直线起始位置距A1单元格左侧的距离；
     * dy1：起始单元格的y偏移量，如例子中的125表示直线起始位置距A1单元格上侧的距离；
     * dx2：终止单元格的x偏移量，如例子中的1023表示直线起始位置距C3单元格左侧的距离；
     * dy2：终止单元格的y偏移量，如例子中的150表示直线起始位置距C3单元格上侧的距离；
     * col1：起始单元格列序号，从0开始计算；竖
     * row1：起始单元格行序号，从0开始计算，如例子中col1=0,row1=0就表示起始单元格为A1；横
     * col2：终止单元格列序号，从0开始计算；
     * row2：终止单元格行序号，从0开始计算，如例子中col2=2,row2=2就表示起始单元格为C3；
     *
     * @param imageDefault     imageDefault
     * @param defaultRowHeight defaultRowHeight
     * @param defaultColWidth  defaultColWidth
     * @param columnlen        columnlenObject
     * @param rowlen           rowlen
     */
    private static Map<String, Integer> getAnchorMap(JSONObject imageDefault, int defaultRowHeight, int defaultColWidth, JSONObject columnlen, JSONObject rowlen) {
        int left = imageDefault.getInteger("left") == null ? 0 : imageDefault.getInteger("left");
        int top = imageDefault.getInteger("top") == null ? 0 : imageDefault.getInteger("top");
        int width = imageDefault.getInteger("width") == null ? 0 : imageDefault.getInteger("width");
        int height = imageDefault.getInteger("height") == null ? 0 : imageDefault.getInteger("height");
        // 算起始最大列
        int colMax1 = (int) Math.ceil((double) left / defaultColWidth);
        // 算起始最大行
        int rowMax1 = (int) Math.ceil((double) top / defaultRowHeight);
        // 算终止最大列
        int colMax2 = (int) Math.ceil((double) (left + width) / defaultColWidth);
        // 算终止最大行
        int rowMax2 = (int) Math.ceil((double) (top + height) / defaultRowHeight);
        // 宽 行
        BigDecimal dx1 = new BigDecimal(left);
        // 高 列
        BigDecimal dy1 = new BigDecimal(top);
        BigDecimal dx2 = new BigDecimal(left + width);
        BigDecimal dy2 = new BigDecimal(top + height);
        int col1 = 0;
        int row1 = 0;
        int col2 = 0;
        int row2 = 0;
        // 算起始列的序号和偏移量
        for (int index = 0; index <= colMax1; index++) {
            BigDecimal col = null;
            if (columnlen != null && columnlen.getString(Integer.toString(index)) != null) {
                // 看当前列是否重新赋值
                col = new BigDecimal(columnlen.getString(Integer.toString(index)));
            }
            // 算起始列
            if (col == null && dx1.compareTo(new BigDecimal(defaultColWidth)) < 0) {
                col1 = index;
                break;
            }
            // 算起始X偏移
            if (col == null && dx1.compareTo(new BigDecimal(defaultColWidth)) >= 0) {
                dx1 = dx1.subtract(new BigDecimal(defaultColWidth));
            }
            // 算起始列
            if (col != null && dx1.compareTo(col) < 0) {
                col1 = index;
                break;
            }
            // 算起始X偏移
            if (col != null) {
                dx1 = dx1.subtract(col);
            }
        }
        // 算起始行的序号和偏移量
        for (int index = 0; index <= rowMax1; index++) {
            BigDecimal row = null;
            if (rowlen != null && rowlen.getString(Integer.toString(index)) != null) {
                // 看当前行是否重新赋值
                row = new BigDecimal(rowlen.getString(Integer.toString(index)));
            }
            // 算起始行
            if (row == null && dy1.compareTo(new BigDecimal(defaultRowHeight)) < 0) {
                row1 = index;
                break;
            }
            // 算起始y偏移
            if (row == null && dy1.compareTo(new BigDecimal(defaultRowHeight)) >= 0) {
                dy1 = dy1.subtract(new BigDecimal(defaultRowHeight));
            }
            // 算起始行
            if (row != null && dy1.compareTo(row) < 0) {
                row1 = index;
                break;
            }
            // 算起始y偏移
            if (row != null) {
                dy1 = dy1.subtract(row);
            }
        }
        // 算最终列的序号和偏移量
        for (int index = 0; index <= colMax2; index++) {
            BigDecimal col = null;
            if (columnlen != null && columnlen.getString(Integer.toString(index)) != null) {
                // 看当前列是否重新赋值
                col = new BigDecimal(columnlen.getString(Integer.toString(index)));
            }
            // 算最终列
            if (col == null && dx2.compareTo(new BigDecimal(defaultColWidth)) < 0) {
                col2 = index;
                break;
            }
            // 算最终X偏移
            if (col == null && dx2.compareTo(new BigDecimal(defaultColWidth)) >= 0) {
                dx2 = dx2.subtract(new BigDecimal(defaultColWidth));
            }
            // 算最终列
            if (col != null && dx2.compareTo(col) < 0) {
                col2 = index;
                break;
            }
            // 算最终X偏移
            if (col != null) {
                dx2 = dx2.subtract(col);
            }
        }
        // 算最终行的序号和偏移量
        for (int index = 0; index <= rowMax2; index++) {
            // 行高
            BigDecimal row = null;
            if (rowlen != null && rowlen.getString(Integer.toString(index)) != null) {
                row = new BigDecimal(rowlen.getString(Integer.toString(index)));// 看当前行是否重新赋值
            }
            // 算最终行
            if (row == null && dy2.compareTo(new BigDecimal(defaultRowHeight)) < 0) {
                row2 = index;
                break;
            }
            // 算最终y偏移
            if (row == null && dy2.compareTo(new BigDecimal(defaultRowHeight)) >= 0) {
                dy2 = dy2.subtract(new BigDecimal(defaultRowHeight));
            }
            // 算最终行
            if (row != null && dy2.compareTo(row) < 0) {
                row2 = index;
                break;
            }
            // 算最终Y偏移
            if (row != null) {
                dy2 = dy2.subtract(row);
            }
        }
        Map<String, Integer> map = new HashMap<>();
        map.put("dx1", dx1.multiply(new BigDecimal(Units.EMU_PER_PIXEL)).setScale(0, BigDecimal.ROUND_HALF_UP).intValue());
        map.put("dy1", dy1.multiply(new BigDecimal(Units.EMU_PER_PIXEL)).setScale(0, BigDecimal.ROUND_HALF_UP).intValue());
        map.put("dx2", dx2.multiply(new BigDecimal(Units.EMU_PER_PIXEL)).setScale(0, BigDecimal.ROUND_HALF_UP).intValue());
        map.put("dy2", dy2.multiply(new BigDecimal(Units.EMU_PER_PIXEL)).setScale(0, BigDecimal.ROUND_HALF_UP).intValue());
        map.put("col1", col1);
        map.put("row1", row1);
        map.put("col2", col2);
        map.put("row2", row2);
        return map;
    }

    /**
     * 行列冻结
     *
     * @param sheet  sheet
     * @param frozen frozen
     */
    private static void setFreezePane(XSSFSheet sheet, JSONObject frozen) {
        if (frozen != null) {
            Map<String, Object> frozenMap = frozen.getInnerMap();
            // 首行
            if ("row".equals(frozenMap.get("type").toString())) {
                sheet.createFreezePane(0, 1);
            }
            // 首列
            if ("column".equals(frozenMap.get("type").toString())) {
                sheet.createFreezePane(1, 0);
            }
            // 行列
            if ("both".equals(frozenMap.get("type").toString())) {
                sheet.createFreezePane(1, 1);
            }
            // 几行
            if ("rangeRow".equals(frozenMap.get("type").toString())) {
                JSONObject value = (JSONObject) frozenMap.get("range");
                sheet.createFreezePane(0, value.getInteger("row_focus") + 1);
            }
            // 几列
            if ("rangeColumn".equals(frozenMap.get("type").toString())) {
                JSONObject value = (JSONObject) frozenMap.get("range");
                sheet.createFreezePane(value.getInteger("column_focus") + 1, 0);
            }
            // 几行列
            if ("rangeBoth".equals(frozenMap.get("type").toString())) {
                JSONObject value = (JSONObject) frozenMap.get("range");
                sheet.createFreezePane(value.getInteger("column_focus") + 1, value.getInteger("row_focus") + 1);
            }
        }
    }

    /**
     * 设置非默认宽高
     *
     * @param sheet     sheet
     * @param columnlen columnlen
     * @param rowlen    rowlen
     */
    private static void setCellWH(XSSFSheet sheet, JSONObject columnlen, JSONObject rowlen) {
        // 我们都知道excel是表格，即由一行一行组成的，那么这一行在java类中就是一个XSSFRow对象，我们通过XSSFSheet对象就可以创建XSSFRow对象
        // 如：创建表格中的第一行（我们常用来做标题的行)  XSSFRow firstRow = sheet.createRow(0); 注意下标从0开始
        // 根据luckysheet创建行列
        // 创建行和列
        if (rowlen != null) {
            Map<String, Object> rowMap = rowlen.getInnerMap();
            for (Map.Entry<String, Object> rowEntry : rowMap.entrySet()) {
                XSSFRow row = sheet.createRow(Integer.parseInt(rowEntry.getKey()));// 创建行
                BigDecimal hei = new BigDecimal(rowEntry.getValue() + "");
                // 转化excle行高参数1
                BigDecimal excleHei1 = new BigDecimal(72);
                // 转化excle行高参数2
                BigDecimal excleHei2 = new BigDecimal(96);
                row.setHeightInPoints(hei.multiply(excleHei1).divide(excleHei2).floatValue());// 行高px值
                if (columnlen != null) {
                    Map<String, Object> cloMap = columnlen.getInnerMap();
                    for (Map.Entry<String, Object> cloEntry : cloMap.entrySet()) {
                        BigDecimal wid = new BigDecimal(cloEntry.getValue().toString());
                        // 转化excle列宽参数35.7   调试后我改为33   --具体多少没有算
                        BigDecimal excleWid = new BigDecimal(33);
                        sheet.setColumnWidth(Integer.parseInt(cloEntry.getKey()), wid.multiply(excleWid).setScale(0, BigDecimal.ROUND_HALF_UP).intValue());// 列宽px值
                        row.createCell(Integer.parseInt(cloEntry.getKey()));// 创建列
                    }
                }
            }
        }
    }

    /**
     * @param wb               wb
     * @param sheet            sheet
     * @param images           所有图片
     * @param columnlenObject  columnlenObject
     * @param rowlenObject     rowlenObject
     * @param defaultRowHeight defaultRowHeight
     * @param defaultColWidth  defaultColWidth
     */
    private static void setImages(XSSFWorkbook wb, XSSFSheet sheet, JSONObject images, JSONObject columnlenObject, JSONObject rowlenObject, int defaultRowHeight, int defaultColWidth) {
        // 图片插入
        if (images != null) {
            Map<String, Object> map = images.getInnerMap();
            JSONObject finalColumnlenObject = columnlenObject;
            JSONObject finalRowlenObject = rowlenObject;
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                XSSFDrawing patriarch = sheet.createDrawingPatriarch();
                // 图片信息
                JSONObject iamgeData = (JSONObject) entry.getValue();
                // 图片的位置宽 高 距离左 距离右
                JSONObject imageDefault = ((JSONObject) iamgeData.get("default"));
                // 算坐标
                Map<String, Integer> colrowMap = getAnchorMap(imageDefault, defaultRowHeight, defaultColWidth, finalColumnlenObject, finalRowlenObject);
                XSSFClientAnchor anchor = new XSSFClientAnchor(colrowMap.get("dx1"), colrowMap.get("dy1"), colrowMap.get("dx2"), colrowMap.get("dy2"), colrowMap.get("col1"), colrowMap.get("row1"), colrowMap.get("col2"), colrowMap.get("row2"));
                // TODO
                anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE);
                BASE64Decoder decoder = new BASE64Decoder();
                byte[] decoderBytes = new byte[0];
                boolean flag = true;
                try {
                    if (iamgeData.get("src") != null) {
                        decoderBytes = decoder.decodeBuffer(iamgeData.get("src").toString().split(";base64,")[1]);
                        flag = iamgeData.get("src").toString().split(";base64,")[0].contains("png");
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
                if (flag) {
                    patriarch.createPicture(anchor, wb.addPicture(decoderBytes, HSSFWorkbook.PICTURE_TYPE_PNG));
                } else {
                    patriarch.createPicture(anchor, wb.addPicture(decoderBytes, HSSFWorkbook.PICTURE_TYPE_JPEG));
                }
            }
        }
    }

    /**
     * 设置单元格
     *
     * @param wb               wb
     * @param sheet            sheet
     * @param celldata         celldata
     * @param columnlen        columnlen
     * @param rowlen           rowlen
     * @param defaultRowHeight defaultRowHeight
     * @param defaultColWidth  defaultColWidth
     */
    private static void setCellValue(XSSFWorkbook wb, XSSFSheet sheet, JSONArray celldata, JSONObject columnlen, JSONObject rowlen, int defaultRowHeight, int defaultColWidth) {
        for (int index = 0; index < celldata.size(); index++) {
            JSONObject luckyCellData = celldata.getJSONObject(index);

            JSONObject v = luckyCellData.getJSONObject("v");
            String m = "";
            if (v.getString("m") != null && v.getString("v") != null) {
                m = v.getString("m");
            }
            Integer r = luckyCellData.getInteger("r");
            if (sheet.getRow(r) == null) {
                sheet.createRow(r);
            }
            XSSFRow row = sheet.getRow(r);
            Integer c = luckyCellData.getInteger("c");
            if (row.getCell(c) == null) {
                row.createCell(c);
            }
            XSSFCell cell = row.getCell(c);
            // 设置单元格样式
            CellStyle cellStyle = createCellStyle(sheet, wb, v);
            // 如果单元格内容是数值类型，涉及到金钱（金额、本、利），则设置cell的类型为数值型，设置data的类型为数值类型
            // 此处设置数据格式
            XSSFDataFormat dataFormat = wb.createDataFormat();
            boolean isNumber = false;
            boolean isString = false;
            boolean isDate = false;
            SimpleDateFormat sdf;
            JSONObject ct = v.getJSONObject("ct");
            if (ct != null) {
                cellStyle.setDataFormat(dataFormat.getFormat(ct.getString("fa")));
                String t = ct.getString("t");
                if ("n".equals(t)) {
                    isNumber = true;
                }
                if ("d".equals(t)) {
                    isDate = true;
                }
                if ("s".equals(t)) {
                    isString = true;
                }
            }
            if (isNumber) {
                // 设置单元格格式
                cell.setCellStyle(cellStyle);
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(m);
            } else if (isDate) {
                String fa = ((JSONObject) v.get("ct")).getString("fa");
                if (fa.contains("AM/PM")) {
                    sdf = new SimpleDateFormat(fa.replaceAll("AM/PM", "aa"), Locale.ENGLISH);
                } else {
                    sdf = new SimpleDateFormat(fa);
                }
                try {
                    Date date = sdf.parse(m);
                    cell.setCellStyle(cellStyle);
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(date);
                } catch (ParseException e) {
                    e.printStackTrace();
                }
            } else if (isString) {
                // 设置单元格格式
                cell.setCellStyle(cellStyle);
                cell.setCellType(CellType.STRING);
                cell.setCellValue(m);
            } else {
                // 设置单元格格式
                cell.setCellStyle(cellStyle);
                cell.setCellValue(m);
            }
            // 设置公式
            String f = v.getString("f");
            if (f != null) {
                if (!f.substring(1).contains("AK")) {
                    cell.setCellFormula(f.substring(1));
                }
            }
            // TODO 设置批注
            JSONObject ps = v.getJSONObject("ps");
            if (ps != null) {
                XSSFDrawing drawing = sheet.createDrawingPatriarch();
                // 后四个坐标待定
                // 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
                System.out.println("-----------------------------------------");
                System.out.println(ps);
                System.out.println(defaultRowHeight);
                System.out.println(defaultColWidth);
                System.out.println(columnlen);
                System.out.println(rowlen);
                Map<String, Integer> anchorMap = getAnchorMap(ps, defaultRowHeight, defaultColWidth, columnlen, rowlen);
                XSSFClientAnchor anchor = new XSSFClientAnchor(anchorMap.get("dx1"), anchorMap.get("dy1"), anchorMap.get("dx2"), anchorMap.get("dy2"), anchorMap.get("col1"), anchorMap.get("row1"), anchorMap.get("col2"), anchorMap.get("row2"));
                System.out.printf("%s,%s,%s,%s,%s,%s,%s,%s%n", anchorMap.get("dx1"), anchorMap.get("dy1"), anchorMap.get("dx2"), anchorMap.get("dy2"), anchorMap.get("col1"), anchorMap.get("row1"), anchorMap.get("col2"), anchorMap.get("row2"));
                XSSFComment comment = drawing.createCellComment(anchor);
                // 输入批注信息
                comment.setString(new XSSFRichTextString(ps.getString("value")));
                // 添加状态
                comment.setVisible("true".equals(ps.getString("isshow")));
                // 将批注添加到单元格对象中
                cell.setCellComment(comment);
            }

        }
    }

    /**
     * 设置边框样式
     *
     * @param borderInfo borderInfo
     * @param sheet      sheet
     */
    private static void setBorder(JSONArray borderInfo, XSSFSheet sheet) {
        if (borderInfo == null) {
            return;
        }
        // 一定要通过 cell.getCellStyle()  不然的话之前设置的样式会丢失
        // 设置边框
        for (Object o : borderInfo) {
            JSONObject borderInfoObject = (JSONObject) o;
            if ("cell".equals(borderInfoObject.get("rangeType"))) {// 单个单元格
                JSONObject borderValueObject = borderInfoObject.getJSONObject("value");

                JSONObject l = borderValueObject.getJSONObject("l");
                JSONObject r = borderValueObject.getJSONObject("r");
                JSONObject t = borderValueObject.getJSONObject("t");
                JSONObject b = borderValueObject.getJSONObject("b");

                int row = borderValueObject.getInteger("row_index");
                int col = borderValueObject.getInteger("col_index");

                XSSFCell cell = sheet.getRow(row).getCell(col);
                XSSFCellStyle xssfCellStyle = cell.getCellStyle();

                if (l != null) {
                    // 左边框
                    xssfCellStyle.setBorderLeft(BORD_MAP.get(l.getInteger("style")));
                    XSSFColor color = toColorFromString(l.getString("color"));
                    // 左边框颜色
                    xssfCellStyle.setLeftBorderColor(color);
                }
                if (r != null) {
                    // 右边框
                    xssfCellStyle.setBorderRight(BORD_MAP.get(r.getInteger("style")));
                    XSSFColor color = toColorFromString(r.getString("color"));
                    // 右边框颜色
                    xssfCellStyle.setRightBorderColor(color);
                }
                if (t != null) {
                    // 顶部边框
                    xssfCellStyle.setBorderTop(BORD_MAP.get(t.getInteger("style")));
                    XSSFColor color = toColorFromString(t.getString("color"));
                    // 顶部边框颜色
                    xssfCellStyle.setTopBorderColor(color);
                }
                if (b != null) {
                    // 底部边框
                    xssfCellStyle.setBorderBottom(BORD_MAP.get(b.getInteger("style")));
                    XSSFColor color = toColorFromString(b.getString("color"));
                    xssfCellStyle.setBottomBorderColor(color);
                }
                cell.setCellStyle(xssfCellStyle);
            }
            // 选区
            else if ("range".equals(borderInfoObject.get("rangeType"))) {
                XSSFColor color = toColorFromString(borderInfoObject.getString("color"));
                int style_ = borderInfoObject.getInteger("style");

                JSONObject rangObject = (JSONObject) ((JSONArray) (borderInfoObject.get("range"))).get(0);

                JSONArray rowList = rangObject.getJSONArray("row");
                JSONArray columnList = rangObject.getJSONArray("column");

                for (int row_ = rowList.getInteger(0); row_ < rowList.getInteger(rowList.size() - 1) + 1; row_++) {
                    for (int col_ = columnList.getInteger(0); col_ < columnList.getInteger(columnList.size() - 1) + 1; col_++) {
                        if (sheet.getRow(row_) == null) {
                            sheet.createRow(row_);
                        }
                        if (sheet.getRow(row_).getCell(col_) == null) {
                            sheet.getRow(row_).createCell(col_);
                        }
                        XSSFCell cell = sheet.getRow(row_).getCell(col_);
                        XSSFCellStyle xssfCellStyle = cell.getCellStyle();
                        // 左边框
                        xssfCellStyle.setBorderLeft(BORD_MAP.get(style_));
                        // 左边框颜色
                        xssfCellStyle.setLeftBorderColor(color);
                        // 右边框
                        xssfCellStyle.setBorderRight(BORD_MAP.get(style_));
                        // 右边框颜色
                        xssfCellStyle.setRightBorderColor(color);
                        // 顶部边框
                        xssfCellStyle.setBorderTop(BORD_MAP.get(style_));
                        // 顶部边框颜色
                        xssfCellStyle.setTopBorderColor(color);
                        // 底部边框
                        xssfCellStyle.setBorderBottom(BORD_MAP.get(style_));
                        // 底部边框颜色 }
                        xssfCellStyle.setBottomBorderColor(color);
                        cell.setCellStyle(xssfCellStyle);
                    }
                }


            }
        }
    }

    /**
     * 设置数据筛选
     *
     * @param dataVerification 数据筛选规则
     * @param sheet            sheet
     */
    private static void settDataValidation(JSONObject dataVerification, XSSFSheet sheet) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        Map<String, Integer> opTypeMap = new HashMap<>();
        opTypeMap.put("bw", DVConstraint.OperatorType.BETWEEN);// "bw"(介于)
        opTypeMap.put("nb", DVConstraint.OperatorType.NOT_BETWEEN);// "nb"(不介于)
        opTypeMap.put("eq", DVConstraint.OperatorType.EQUAL);// "eq"(等于)
        opTypeMap.put("ne", DVConstraint.OperatorType.NOT_EQUAL);// "ne"(不等于)
        opTypeMap.put("gt", DVConstraint.OperatorType.GREATER_THAN);// "gt"(大于)
        opTypeMap.put("lt", DVConstraint.OperatorType.LESS_THAN);// lt"(小于)
        opTypeMap.put("gte", DVConstraint.OperatorType.GREATER_OR_EQUAL);// "gte"(大于等于)
        opTypeMap.put("lte", DVConstraint.OperatorType.LESS_OR_EQUAL);// "lte"(小于等于)
        opTypeMap.put("number", DVConstraint.ValidationType.ANY);// 数字
        opTypeMap.put("number_integer", DVConstraint.ValidationType.INTEGER);// 整数
        opTypeMap.put("number_decimal", DVConstraint.ValidationType.DECIMAL);// 小数
        opTypeMap.put("text_length", DVConstraint.ValidationType.TEXT_LENGTH);// 文本长度
        opTypeMap.put("date", DVConstraint.ValidationType.DATE);// 日期
        if (dataVerification != null) {
            Map<String, Object> dataVe = dataVerification.getInnerMap();
            for (Map.Entry<String, Object> dataEntry : dataVe.entrySet()) {
                String[] colRow = dataEntry.getKey().split("_");
                CellRangeAddressList dstAddrList = new CellRangeAddressList(Integer.parseInt(colRow[0]), Integer.parseInt(colRow[0]), Integer.parseInt(colRow[1]), Integer.parseInt(colRow[1]));// 规则一单元格范围
                JSONObject dataVeValue = (JSONObject) dataEntry.getValue();
                DataValidation dstDataValidation = null;
                if ("dropdown".equals(dataVeValue.getString("type"))) {
                    if (dataVeValue.getString("value1").contains(",")) {
                        String[] textlist = dataVeValue.getString("value1").split(",");
                        dstDataValidation = helper.createValidation(helper.createExplicitListConstraint(textlist), dstAddrList);
                    } else {
                        dstDataValidation = helper.createValidation(helper.createFormulaListConstraint(dataVeValue.getString("value1")), dstAddrList);
                    }
                }
                if ("checkbox".equals(dataVeValue.getString("type"))) {
                    // // TODO: 2020/11/30
                }
                if ("number".equals(dataVeValue.getString("type"))) {
                    // number判断是整数还是小数
                    boolean booleanValue1;
                    boolean booleanValue2;
                    booleanValue1 = PATTERN.matcher(dataVeValue.getString("value1")).matches();
                    booleanValue2 = PATTERN.matcher(dataVeValue.getString("value2")).matches();
                    DataValidationConstraint dvc;
                    if (booleanValue1 && booleanValue2) {
                        dvc = helper.createIntegerConstraint(opTypeMap.get(dataVeValue.getString("type2")), dataVeValue.getString("value1"), dataVeValue.getString("value2"));
                    } else {
                        dvc = helper.createDecimalConstraint(opTypeMap.get(dataVeValue.getString("type2")), dataVeValue.getString("value1"), dataVeValue.getString("value2"));
                    }
                    dstDataValidation = helper.createValidation(dvc, dstAddrList);
                }
                if ("number_integer".equals(dataVeValue.getString("type"))
                        || "number_decimal".equals(dataVeValue.getString("type"))
                        || "text_length".equals(dataVeValue.getString("type"))) {
                    DataValidationConstraint dvc = helper.createNumericConstraint(opTypeMap.get(dataVeValue.getString("type")), opTypeMap.get(dataVeValue.getString("type2")), dataVeValue.getString("value1"), dataVeValue.getString("value2"));
                    dstDataValidation = helper.createValidation(dvc, dstAddrList);
                }
                if ("date".equals(dataVeValue.getString("type"))) {
                    // 日期
                    DataValidationConstraint dvc = new XSSFDataValidationConstraint(opTypeMap.get(dataVeValue.getString("type")), opTypeMap.get(dataVeValue.getString("type2")), dataVeValue.getString("value1"), dataVeValue.getString("value2"));
                    dstDataValidation = helper.createValidation(dvc, dstAddrList);
                }
                if ("text_content".equals(dataVeValue.getString("type"))) {
                    // TODO: 2020/11/30

                }
                if ("validity".equals(dataVeValue.getString("type"))) {
                    // TODO: 2020/12/1
                }
                dstDataValidation.createPromptBox("提示:", dataVeValue.getString("hintText"));
                dstDataValidation.setShowErrorBox(dataVeValue.getBoolean("prohibitInput"));
                dstDataValidation.setShowPromptBox(dataVeValue.getBoolean("hintShow"));
                sheet.addValidationData(dstDataValidation);
            }
        }
//        CellReference cr = new CellReference("A1");
    }
}

# luckysheet_demo
## 如果您觉得有帮助，请点右上角 "Star" 支持一下谢谢
>1.springboot后台导出excel 文件demo。启动后访问：http://localhost/
>
>2.参考博客：https://blog.csdn.net/zzq10066/article/details/110424977
## 导出说明
```lua
│  │  ├─createCellStyle -- 样式添加（合并单元格、字体、字体颜色、粗体、斜体、删除线、下划线、字体大小、水平对齐、垂直对齐、背景颜色）
│  │  ├─toColorFromString -- 颜色转化
│  │  ├─setFreezePane -- 固定行列（冻结行列）
│  │  ├─setCellWH -- 设置非默认宽高（宽度转化有点问题 网上列宽参数35.7   调试后我改为33   --具体多少没有算）
│  │  ├─setImages -- 插入图片
│  │  ├─setCellValue -- 设置单元格值及格式
│  │  ├─settDataValidation -- 数据验证 （数字、下拉、整数、小数、文本长度、日期等已做   复选框、文本内容、有效性代做）
│  │  ├─setBorder -- 设置边框
│  │  ├─getColRowMap -- 获取图片等具体位置
```
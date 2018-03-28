package com.deliverik.inthink.utils.excel;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel简单导出
 *
 * @author 刘旭 (LiuXu)
 * <p>
 * 2017年4月5日上午11:13:57
 */
public class SimpleExport {

    /**
     * 不创建sheet
     */
    public static final String NOT_CREATE_SHEET = null;

    /**
     * Excel版本
     *
     * @author liuxu
     * @date 17-10-13下午2:49
     */
    public enum ExcelVersionEnum {
        EXCEL2003(".xls", true), EXCEL2007(".xlsx", false);
        /**
         * 文件后缀
         */
        private String suffix;
        /**
         * 工作簿对象
         */
        private boolean isExcel2003;

        ExcelVersionEnum(String suffix, boolean isExcel2003) {
            this.suffix = suffix;
            this.isExcel2003 = true;
        }

        public String getSuffix() {
            return suffix;
        }

        /**
         * Excel版本识别
         *
         * @author liuxu
         * @date 17-10-17上午9:10
         */
        public boolean isExcel2003() {
            return isExcel2003;
        }

        /**
         * 创建工作簿
         *
         * @author liuxu
         * @date 17-10-13下午3:28
         */
        private Workbook createWorkbook() {
            if (isExcel2003) {
                return new HSSFWorkbook();
            } else {
                return new XSSFWorkbook();
            }
        }
    }

    /**
     * Excel版本
     */
    private ExcelVersionEnum excelVersionEnum;

    /**
     * 输出流
     */
    private OutputStream os;

    /**
     * 工作簿
     */
    private Workbook workbook;

    /**
     * sheet页
     */
    private Sheet sheet;

    /**
     * 主标题样式
     */
    private CellStyle mainTitleStyle;

    /**
     * 副标题样式
     */
    private CellStyle deputyTitleStyle;

    /**
     * 内容样式
     */
    private CellStyle contentStyle;

    /**
     * 列宽（默认3000）
     */
    private int columnWidth = 3000;

    /**
     * 表格列宽集合
     */
    private Map<Integer, Integer> columnsWidth;

    /**
     * 行高:POI规定如果此值为-1则自动行高
     */
    private float rowHeight = 35;

    /**
     * 表格行高集合
     */
    private List<Float> rowsHeight;

    /**
     * 占用记录表
     */
    private int[][] record;

    /**
     * 第一维：行<br>
     * 第二维：列<br>
     * 第三维：样式<br>
     * vals[?][?][0]-value<br>
     * vals[?][?][1]-X轴跨度（列）<br>
     * vals[?][?][2]-Y轴跨度（行）<br>
     * vals[?][?][3]-样式<br>
     */
    private List<Object[][]> vals;

    /**
     * 测试模式（默认关闭）
     */
    private boolean isTest = false;

    // ----------------------------------------基本设置----------------------------------------

    /**
     * 简单导出工具构造
     *
     * @param response         响应对象
     * @param fileName         文件名
     * @param excelVersionEnum 选择Excel版本
     * @author liuxu
     * @date 17-10-16下午3:20
     */
    public SimpleExport(HttpServletResponse response, String fileName, ExcelVersionEnum excelVersionEnum) throws IOException {
        //设置响应头
        setResponse(response, fileName, excelVersionEnum);
        //执行初始化
        init(response.getOutputStream(), excelVersionEnum, "sheet1");
    }

    /**
     * 简单导出工具构造
     *
     * @param response         响应对象
     * @param fileName         文件名
     * @param excelVersionEnum 选择Excel版本
     * @param sheetName        sheet页名称
     * @author liuxu
     * @date 17-10-16下午3:21
     */
    public SimpleExport(HttpServletResponse response, String fileName, ExcelVersionEnum excelVersionEnum, String sheetName) throws IOException {
        //设置响应头
        setResponse(response, fileName, excelVersionEnum);
        //执行初始化
        init(response.getOutputStream(), excelVersionEnum, sheetName);
    }

    /**
     * 简单导出工具构造
     *
     * @param file             文件对象
     * @param excelVersionEnum 选择Excel版本
     * @author liuxu
     * @date 17-10-16下午3:22
     */
    public SimpleExport(File file, ExcelVersionEnum excelVersionEnum) throws FileNotFoundException {
        init(new FileOutputStream(file), excelVersionEnum, "sheet1");
    }

    /**
     * 简单导出工具构造
     *
     * @param file             文件对象
     * @param excelVersionEnum 选择Excel版本
     * @param sheetName        sheet页名称
     * @author liuxu
     * @date 17-10-16下午3:23
     */
    public SimpleExport(File file, ExcelVersionEnum excelVersionEnum, String sheetName) throws FileNotFoundException {
        init(new FileOutputStream(file), excelVersionEnum, sheetName);
    }

    /**
     * 简单导出工具构造
     *
     * @param os               输出流
     * @param excelVersionEnum 选择Excel版本
     * @param sheetName        sheet页名称
     * @author liuxu
     * @date 17-10-16下午3:23
     */
    public SimpleExport(OutputStream os, ExcelVersionEnum excelVersionEnum, String sheetName) {
        init(os, excelVersionEnum, sheetName);
    }

    /**
     * 初始化
     *
     * @author liuxu
     * @date 17-10-16下午1:53
     */
    private void init(OutputStream os, ExcelVersionEnum excelVersionEnum, String sheetName) {
        //初始化Excel操作对象
        this.excelVersionEnum = excelVersionEnum;
        workbook = this.excelVersionEnum.createWorkbook();
        //sheet名称不存在，则不创建sheet
        sheet = sheetName != NOT_CREATE_SHEET ? createSheet(sheetName) : null;
        this.os = os;

        //创建默认样式
        mainTitleStyle = createMainTitleStyle();
        deputyTitleStyle = createDeputyTitleStyle();
        contentStyle = createContentStyle();
    }

    /**
     * 设置响应头
     *
     * @param response
     * @param fileName
     * @param excelVersionEnum
     * @author liuxu
     * @date 18-3-21上午10:34
     */
    private void setResponse(HttpServletResponse response, String fileName, ExcelVersionEnum excelVersionEnum) throws UnsupportedEncodingException {
        response.setHeader("Content-disposition",
                "attachment;filename=" + fileNameFormat(fileName + excelVersionEnum.getSuffix()));
        response.setContentType("application/msexcel");
    }

    /**
     * 创建数据容器
     *
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:16:47
     */
    private void createVals() {
        //数据容器创建
        this.vals = new ArrayList<Object[][]>();
        //列宽集合创建
        this.columnsWidth = new HashMap<Integer, Integer>();
        //行高集合创建
        this.rowsHeight = new ArrayList<Float>();
    }

    /**
     * 获取工作簿
     *
     * @author liuxu
     * @date 17-10-13下午3:21
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 创建sheet
     *
     * @param sheetName
     * @author liuxu
     * @date 17-10-13下午3:02
     */
    public Sheet createSheet(String sheetName) {
        //创建新数据容器
        createVals();
        //创建Sheet
        sheet = workbook.createSheet(sheetName);
        return sheet;
    }

    /**
     * 获取Sheet对象
     *
     * @author liuxu
     * @date 17-10-16下午2:55
     */
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * 设置Sheet对象,改变当前操作的sheet
     *
     * @param sheet
     * @author liuxu
     * @date 17-10-16下午2:56
     */
    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    // ----------------------------------------设置打印----------------------------------------

    /**
     * 设置打印参数
     *
     * @param topMargin
     * @param bottomMargin
     * @param leftMargin
     * @param rightMargin
     * @return Sheet
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日上午11:54:25
     */
    public Sheet setPrintParam(double topMargin, double bottomMargin, double leftMargin, double rightMargin) {
        // 设置打印参数
        sheet.setMargin(Sheet.TopMargin, topMargin);// 页边距（上）
        sheet.setMargin(Sheet.BottomMargin, bottomMargin);// 页边距（下）
        sheet.setMargin(Sheet.LeftMargin, leftMargin);// 页边距（左）
        sheet.setMargin(Sheet.RightMargin, rightMargin);// 页边距（右
        return sheet;
    }

    /**
     * 打印页面自适应
     *
     * @param isAutobreaks
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午1:09:57
     */
    public void setAutobreaks(boolean isAutobreaks) {
        sheet.setAutobreaks(isAutobreaks);
    }

    /**
     * 设置打印方向
     *
     * @param isHorizontal true：横向、false：纵向(默认)
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午1:15:47
     */
    public PrintSetup setLandscape(boolean isHorizontal) {
        PrintSetup ps = sheet.getPrintSetup();
        ps.setLandscape(isHorizontal);
        return ps;
    }

    /**
     * 设置纸张类型
     *
     * @param paperType 例：PrintSetup.A4_PAPERSIZE
     * @return PrintSetup
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午1:21:12
     */
    public PrintSetup setPaperType(short paperType) {
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(paperType);
        return ps;
    }

    // ----------------------------------------样式设置----------------------------------------

    /**
     * 设置列宽
     *
     * @param columnIndex
     * @param width
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午4:22:13
     */
    public void setColumnWidth(int columnIndex, int width) {
        columnsWidth.put(columnIndex, width);
    }

    /**
     * 设置所有列宽度
     *
     * @param columnWidth
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午4:10:04
     */
    public void setAllColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }

    /**
     * 执行设置列宽操作
     *
     * @param columnSum
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午4:17:00
     */
    private void executeSetAllColumnWidth(int columnSum, int colOffset) {
        for (int i = 0; i < columnSum; i++) {
            //获取高度
            Integer width = columnsWidth.get(i);
            if (width != null) {
                sheet.setColumnWidth(i + colOffset, width);
            } else {
                sheet.setColumnWidth(i + colOffset, columnWidth);
            }
        }
    }

    /**
     * 设置行高
     *
     * @param rowHeight
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午4:38:56
     */
    public void setRowHeight(float rowHeight) {
        this.rowHeight = rowHeight;
    }

    /**
     * 获取默认行高
     *
     * @author liuxu
     * @date 17-11-8下午4:21
     */
    public float getRowHeight() {
        return this.rowHeight;
    }


    /**
     * 创建字体
     *
     * @param fontName 字体
     * @param fontSize 字号
     * @param bold     字体加粗
     * @author liuxu
     * @date 17-10-13下午3:09
     */
    public Font createFont(String fontName, short fontSize, boolean bold) {
        Font font = workbook.createFont();
        font.setFontName(fontName);//字体
        font.setFontHeightInPoints(fontSize);// 字号
        font.setBold(bold);//字体加粗
        return font;
    }

    /**
     * 创建样式
     *
     * @author liuxu
     * @date 17-10-13下午3:57
     */
    public CellStyle createStyle() {
        return workbook.createCellStyle();
    }

    /**
     * 获取主标题样式
     *
     * @return CellStyle
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:45:47
     */
    public CellStyle createMainTitleStyle() {
        // 主标题样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(createFont("黑体", (short) 26, true));
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        titleStyle.setWrapText(true);//设置自动换行

        return titleStyle;
    }


    /**
     * 获取副标题样式
     *
     * @return CellStyle
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:46:07
     */
    public CellStyle createDeputyTitleStyle() {
        // 副标题样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(createFont("仿宋_GB2312", (short) 12, true));//设置字体
        setBorder(titleStyle, true, true, true, true);//设置边框线
        titleStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        titleStyle.setWrapText(true);//设置自动换行

        return titleStyle;
    }

    /**
     * 获取内容样式
     *
     * @return CellStyle
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:46:31
     */
    public CellStyle createContentStyle() {
        //内容字体
        Font contentFont = createFont("仿宋_GB2312", (short) 12, false);
        // 内容样式
        CellStyle contentStyle = workbook.createCellStyle();

        contentStyle.setFont(contentFont);//设置字体
        setBorder(contentStyle, true, true, true, true);//设置边框线
        contentStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        contentStyle.setWrapText(true);//设置自动换行

        return contentStyle;
    }

    /**
     * 设置边框线
     *
     * @param cellStyle POI样式对象
     * @param top       设置上边框
     * @param bottom    设置下边框
     * @param left      设置左边框
     * @param right     设置右边框
     * @author liuxu
     * @date 17-10-26下午5:08
     */
    public void setBorder(CellStyle cellStyle, boolean top, boolean bottom, boolean left, boolean right) {
        //设置上边框
        if (top) {
            cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        }
        //设置下边框
        if (bottom) {
            cellStyle.setBorderBottom(BorderStyle.THIN);// 下边框
        }
        //设置左边框
        if (left) {
            cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        }
        //设置右边框
        if (right) {
            cellStyle.setBorderRight(BorderStyle.THIN);// 右边框
        }
    }

    /**
     * 获取默认主标题样式
     */
    public CellStyle getMainTitleStyle() {
        return mainTitleStyle;
    }

    /**
     * 设置默认主标题样式
     */
    public void setMainTitleStyle(CellStyle mainTitleStyle) {
        this.mainTitleStyle = mainTitleStyle;
    }

    /**
     * 获取默认副标题样式
     */
    public CellStyle getDeputyTitleStyle() {
        return deputyTitleStyle;
    }

    /**
     * 设置默认副标题样式
     */
    public void setDeputyTitleStyle(CellStyle deputyTitleStyle) {
        this.deputyTitleStyle = deputyTitleStyle;
    }

    /**
     * 获取默认内容样式
     */
    public CellStyle getContentStyle() {
        return contentStyle;
    }

    /**
     * 设置默认内容样式
     */
    public void setContentStyle(CellStyle contentStyle) {
        this.contentStyle = contentStyle;
    }

    // ----------------------------------------设值函数----------------------------------------

    /**
     * 创建行
     *
     * @param cellStyle 行样式
     * @param rowHeight 行高
     * @param columns   单元格数据
     * @author liuxu
     * @date 17-10-26下午4:59
     */
    public void createRow(CellStyle cellStyle, float rowHeight, Object[]... columns) {
        //设置行高
        rowsHeight.add(rowHeight);
        //设置样式
        for (int i = 0; i < columns.length; i++) {
            //set方法设置样式优先级高于createRow方法设置样式的优先级
            if (columns[i][3] == null) {
                columns[i][3] = cellStyle;
            }
        }
        //添加到数据集
        this.vals.add(columns);
    }

    /**
     * 创建行
     *
     * @param columns
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:17:20
     */
    public void createRow(Object[]... columns) {
        createRow(contentStyle, this.rowHeight, columns);
    }

    /**
     * 创建行，并为行内每个元素设置相同样式
     *
     * @param cellStyle
     * @param columns
     * @author 刘旭 (LiuXu)
     * @CreateTime 2017年6月6日下午1:29:50
     */
    public void createRow(CellStyle cellStyle, Object[]... columns) {
        createRow(cellStyle, this.rowHeight, columns);
    }

    /**
     * 创建行
     *
     * @param rowHeight 行高
     * @param columns   单元格数据
     * @author liuxu
     * @date 17-10-26下午5:04
     */
    public void createRow(float rowHeight, Object[]... columns) {
        createRow(contentStyle, rowHeight, columns);
    }

    /**
     * 创建行
     *
     * @param cellStyle 整行cell样式
     * @param rowHeight 行参数集合
     * @param columns   行数据
     * @author liuxu
     * @date 17-10-27上午11:03
     */
    public void createRow(CellStyle cellStyle, float rowHeight, List<Object[]> columns) {
        createRow(cellStyle, rowHeight, list2Array(columns));
    }

    /**
     * 创建行
     *
     * @param cellStyle 整行cell样式
     * @param columns   行参数集合
     * @author liuxu
     * @date 17-10-13下午1:22
     */
    public void createRow(CellStyle cellStyle, List<Object[]> columns) {
        createRow(cellStyle, rowHeight, columns);
    }

    /**
     * 创建行
     *
     * @param columns 行参数集合
     * @author liuxu
     * @date 17-10-13下午1:22
     */
    public void createRow(List<Object[]> columns) {
        createRow(contentStyle, rowHeight, columns);
    }

    /**
     * 创建行
     *
     * @param rowHeight 行参数集合
     * @param columns   行数据
     * @author liuxu
     * @date 17-10-27上午11:06
     */
    public void createRow(float rowHeight, List<Object[]> columns) {
        createRow(contentStyle, rowHeight, columns);
    }

    /**
     * 创建多行，将columns按rowLen拆分为多行
     *
     * @param cellStyle 行样式
     * @param rowLen    行长度
     * @param valList   数据集合
     * @author liuxu
     * @date 17-10-27上午9:31
     */
    public void createRows(CellStyle cellStyle, int rowLen, List<Object[]> valList) {
        // 非法值检测
        if (rowLen <= 0 || valList.size() == 0) {
            return;
        }
        Object[][] rowData = null;
        for (int i = 0, j = 1; i < valList.size(); i++, j++) {
            // 创建行数据
            if (j == 1) {
                int remainder = valList.size() % rowLen;
                rowData = new Object[i + 1 < valList.size() - remainder ? rowLen : remainder][];
            }
            // 折行操作
            rowData[j - 1] = valList.get(i);
            if (j == rowLen || i == valList.size() - 1) {
                j = 0;
                createRow(cellStyle, this.rowHeight, rowData);
            }
        }
    }

    /**
     * 设置单元格内容
     *
     * @param obj       数据对象:会自动识别数据对象类型,如果此对象为字符串数组则单元格为下拉框
     * @param colSpan   列合并,最小值为1
     * @param rowSpan   行合并,最小值为1
     * @param cellStyle 单元格样式
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj, int colSpan, int rowSpan, CellStyle cellStyle) {
        return new Object[]{obj, colSpan - 1, rowSpan - 1, cellStyle};
    }

    /**
     * 设置单元格内容
     *
     * @param obj 数据对象
     *            <br>会自动识别数据对象类型
     *            <br>如果对象为字符串数组则单元格为下拉框_结构:String{{"默认值"},{"选项1","选项2"}}
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj) {
        return set(obj, 1, 1, null);
    }

    /**
     * 设置单元格内容
     *
     * @param obj       数据对象
     *                  <br>会自动识别数据对象类型
     *                  <br>如果对象为字符串数组则单元格为下拉框_结构:String{{"默认值"},{"选项1","选项2"}}
     * @param cellStyle 单元格样式
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj, CellStyle cellStyle) {
        return set(obj, 1, 1, cellStyle);
    }

    /**
     * 设置单元格内容
     *
     * @param obj     数据对象:
     *                <br>会自动识别数据对象类型
     *                <br>如果对象为字符串数组则单元格为下拉框_结构:String{{"默认值"},{"选项1","选项2"}}
     * @param colSpan 列合并,最小值为1
     * @param rowSpan 行合并,最小值为1
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj, int colSpan, int rowSpan) {
        return set(obj, colSpan, rowSpan, null);
    }

    /**
     * 设置单元格内容
     *
     * @param obj     数据对象
     *                <br>会自动识别数据对象类型
     *                <br>如果对象为字符串数组则单元格为下拉框_结构:String{{"默认值"},{"选项1","选项2"}}
     * @param colSpan 列合并,最小值为1
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj, int colSpan) {
        return set(obj, colSpan, 1, null);
    }

    /**
     * 设置单元格内容
     *
     * @param obj       数据对象
     *                  <br>会自动识别数据对象类型
     *                  <br>如果对象为字符串数组则单元格为下拉框_结构:String{{"默认值"},{"选项1","选项2"}}
     * @param colSpan   列合并,最小值为1
     * @param cellStyle 单元格样式
     * @author liuxu
     * @date 17-10-27上午8:47
     */
    public Object[] set(Object obj, int colSpan, CellStyle cellStyle) {
        return set(obj, colSpan, 1, cellStyle);
    }

    /**
     * 生成下拉选项值
     * <br>生成下拉选项数组,在set方法中使用
     *
     * @param options 下拉选项数组
     * @param defVal  默认值
     * @author liuxu
     * @date 17-11-8下午3:31
     */
    public String[][] generateOptions(String[] options, String defVal) {
        return new String[][]{{defVal}, options};
    }

    /**
     * 生成下拉选项值
     * <br>生成下拉选项数组,在set方法中使用
     *
     * @param options 下拉选项数组
     * @author liuxu
     * @date 17-11-8下午3:31
     */
    public String[][] generateOptions(String[] options) {
        return new String[][]{{""}, options};
    }

    // ----------------------------------------执行函数----------------------------------------

    /**
     * 计算表格式实际大小
     *
     * @return int[]
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:28:40
     */
    private int[] calculateTableSize() {
        //Excel表规格
        int[] tableSize = new int[2];
        // Y轴 （Y轴合并单元格数量数）
        int countY = 0;
        for (int i = 0; i < vals.size(); i++) {
            // X轴（获取最大列宽）
            int maxX = 0;
            for (int j = 0; j < vals.get(i).length; j++) {
                maxX += (Integer) vals.get(i)[j][1] + 1;
                countY += (Integer) vals.get(i)[j][1] * (Integer) vals.get(i)[j][2];
            }
            if (maxX > tableSize[0]) {
                tableSize[0] = maxX;
            }
        }
        // 除数不能为0
        int maxX = tableSize[0] == 0 ? 1 : tableSize[0];
        // Y轴
        tableSize[1] = vals.size() + (countY / maxX) + (countY % maxX > 0 ? 1 : 0);
        return tableSize;
    }

    /**
     * 设置单元格值
     *
     * @param cellObj
     * @param val
     * @author liuxu
     * @date 17-10-17下午1:47
     */
    private void setCellValues(Cell cellObj, Object val) {
        if (val != null) {
            if (val instanceof String) {
                cellObj.setCellValue((String) val);
            } else if (val instanceof Double) {
                cellObj.setCellValue((Double) val);
            } else if (val instanceof Integer) {
                cellObj.setCellValue((Integer) val);
            } else if (val instanceof Long) {
                cellObj.setCellValue((Long) val);
            } else if (val instanceof BigDecimal) {
                cellObj.setCellValue(((BigDecimal) val).doubleValue());
            } else {
                cellObj.setCellValue(String.valueOf(val));
            }
        }
    }

    /**
     * 创建下拉选项
     *
     * @param val      选项值
     * @param firstRow 起始行
     * @param lastRow  结束行
     * @param firstCol 起始列
     * @param lastCol  结束列
     * @return 返回默认值
     * @author liuxu
     * @date 17-11-7下午8:10
     */
    private Object createSelect(Object val, int firstRow, int lastRow, int firstCol, int lastCol) {

        //数据类型校验
        if (!(val instanceof String[][])) {
            return val;
        }

        //下拉列表:String{{"默认值"},{"选项1","选项2"}}
        String[][] options = (String[][]) val;

        //数据格式校验
        if (options.length != 2 || options[0] == null || options[1] == null) {
            return val;
        }

        //如果值为数组则生成下拉菜单
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);//设置合并范围
        DataValidation dataValidation;
        if (excelVersionEnum.isExcel2003()) {
            DVConstraint constraint = DVConstraint.createExplicitListConstraint(options[1]);//设置下拉内容
            dataValidation = new HSSFDataValidation(regions, constraint);//生成下拉菜单
        } else {
            DataValidationHelper dvHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(options[1]);//设置下拉内容
            dataValidation = dvHelper.createValidation(constraint, regions);//生成下拉菜单
        }
        sheet.addValidationData(dataValidation);

        return options[0][0];

    }

    /**
     * 刷新单元格使用情况,并设置单元格样式
     *
     * @param cellStyle 单元格样式
     * @param firstRow  起始行
     * @param lastRow   结束行
     * @param firstCol  起始列
     * @param lastCol   结束列
     * @author liuxu
     * @date 17-11-7下午8:26
     */
    private void refreshUsedAndSetStyle(CellStyle cellStyle, int[] startPoint, int firstRow, int lastRow, int firstCol, int lastCol) {
        // 写入占位,并设置单元格样式
        for (int m = firstRow; m <= lastRow; m++) {
            for (int n = firstCol; n <= lastCol; n++) {
                //获取cell并设置样式
                createOrGetCell(createOrGetRow(m + startPoint[1]), n + startPoint[0]).setCellStyle(cellStyle);
                //写入占位
                record[m][n] = 1;
            }
        }
    }

    /**
     * 合并计算
     *
     * @param startPoint startPoint[0]-X轴[列] <br>
     *                   startPoint[1]-Y轴[行]
     * @author liuxu
     * @date 17-10-17下午2:01
     */
    public void mergeCalculate(int[] startPoint) {
        int[] tableSize = calculateTableSize();
        //设置列宽
        executeSetAllColumnWidth(tableSize[0], startPoint[0]);
        // 占用记录表
        record = new int[tableSize[1]][tableSize[0]];
        // 当前行
        int row = 0;
        for (int i = 0; i < vals.size(); i++) {
            // 当前列
            int col = 0;
            // 创建行
            Row rowObj = createOrGetRow(startPoint[1] + row);

            // 设置行高
            rowObj.setHeightInPoints(rowsHeight.get(i));

            for (int j = 0; j < vals.get(i).length; j++) {

                // 检查占用，获取可用位置
                stop:
                for (int m = row; m < tableSize[1]; m++) {
                    for (int n = col; n < tableSize[0]; n++) {
                        if (record[m][n] == 0) {
                            row = m;
                            col = n;
                            break stop;
                        }
                    }
                }

                Object val = vals.get(i)[j][0];//待输出值

                int rowMergeIncrement = (Integer) vals.get(i)[j][2];//行合并增量
                int colMergeIncrement = (Integer) vals.get(i)[j][1];//列合并增量

                // 合并单元格计算（startPoint[0]-X轴[列] startPoint[1]-Y轴[行]）
                int firstRow = startPoint[1] + row;
                int lastRow = startPoint[1] + row + rowMergeIncrement;
                int firstCol = startPoint[0] + col;
                int lastCol = startPoint[0] + col + colMergeIncrement;

                //如果值为数组则生成下拉菜单
                val = createSelect(val, firstRow, lastRow, firstCol, lastCol);

                //普通合并
                if (rowMergeIncrement != 0 || colMergeIncrement != 0) {
                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
                }

                //设置单元格内容
                setCellValues(createOrGetCell(rowObj, firstCol), val);

                // 刷新单元格使用情况
                refreshUsedAndSetStyle((CellStyle) vals.get(i)[j][3], startPoint, row, row + rowMergeIncrement, col, col + colMergeIncrement);

                // 占用情况显示
                if (isTest) {
                    testOccupation();
                }

                //当前列位置计算
                col += (Integer) vals.get(i)[j][1] + 1;
            }
            //当前行位置计算
            ++row;
        }
    }

    /**
     * 合并计算(默认起点为0,0)
     *
     * @author liuxu
     * @date 17-10-17下午4:02
     */
    public void mergeCalculate() {
        mergeCalculate(new int[]{0, 0});
    }

    /**
     * 执行生成Excel
     *
     * @author liuxu
     * @date 17-10-17下午2:06
     */
    public void executeGenerate() throws IOException {
        workbook.write(os);
    }

    /**
     * 生成Excel
     *
     * @param startPoint startPoint[0]-X轴[列] <br>
     *                   startPoint[1]-Y轴[行]
     * @throws IOException
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午3:10:27
     */
    public void excelGenerate(int[] startPoint) throws IOException {
        mergeCalculate(startPoint);
        executeGenerate();
    }

    /**
     * 生成Excel
     * <p>
     * 将导出有效区域整体移动
     *
     * @param colSpan 列跨度
     * @param rowSpan 行跨度
     * @author liuxu
     * @date 17-11-28下午7:01
     */
    public void excelGenerate(int colSpan, int rowSpan) throws IOException {
        excelGenerate(new int[]{colSpan, rowSpan});
    }

    /**
     * 生成Excel(默认起点为0,0)
     *
     * @author liuxu
     * @date 17-10-17下午4:02
     */
    public void excelGenerate() throws IOException {
        mergeCalculate();
        executeGenerate();
    }

    /**
     * 创建or获取行
     *
     * @param rowIndex
     * @return Row
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午7:43:48
     */
    public Row createOrGetRow(int rowIndex) {
        return (sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex));
    }

    /**
     * 创建or获取单元格
     *
     * @param row
     * @param column
     * @return Cell
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午7:44:41
     */
    public Cell createOrGetCell(Row row, int column) {
        return (row.getCell(column) == null ? row.createCell(column) : row.getCell(column));
    }

    /**
     * 关闭IO
     *
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午1:55:00
     */
    public void close() {
        IOUtils.closeQuietly(os);
    }

    // ----------------------------------------测试函数----------------------------------------

    /**
     * 开启测试
     *
     * @param isTest
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午2:42:00
     */
    public void startTest(boolean isTest) {
        this.isTest = isTest;
    }

    /**
     * 打印当前单元格使用情况
     *
     * @author 刘旭 (LiuXu)
     * <p>
     * Create time: 2017年4月5日下午1:37:09
     */
    public void testOccupation() {
        System.out.println("----------------------------------------");
        for (int i = 0; i < record.length; i++) {
            StringBuilder sb = new StringBuilder(" | ");
            for (int j = 0; j < record[i].length; j++) {
                if (record[i][j] == 1) {
                    sb.append("    ■");
                } else {
                    sb.append("    □");
                }
            }
            System.out.println(sb.toString());
        }
    }

    // ----------------------------------------工具函数----------------------------------------

    /**
     * 集合转数组
     *
     * @param list
     * @author liuxu
     * @date 17-10-13下午1:16
     */
    private Object[][] list2Array(List<Object[]> list) {
        return list.toArray(new Object[list.size()][]);
    }

    /**
     * 格式化文件名
     *
     * @param fileName 待格式化的文件名
     * @author liuxu
     * @date 17-10-16下午3:15
     */
    private String fileNameFormat(String fileName) throws UnsupportedEncodingException {
        return new String(fileName.getBytes("gb2312"), "ISO8859-1");
    }

}

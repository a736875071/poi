package com.poi;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

/**
 * @author YangQing
 * @version 1.0.0
 */

@Component("excelExportUtil")
@Scope("prototype")
public class ExcelExportUtil  implements ExportUtil {
    private static HSSFWorkbook wb;
    private static CellStyle titleStyle;
    private static Font titleFont;
    private static CellStyle dateStyle;
    private static Font dateFont;
    private static CellStyle headStyle;
    private static Font headFont;
    private static CellStyle contentStyle;
    private static Font contentFont;
    private LinkedHashMap<String, List<? extends Map>> objsMap;
    private String title;
    private List<? extends String[]> headNames;
    private List<? extends String[]> fieldNames;
    private List<? extends Map> boningList;
    private HSSFRow titleRow;
    private SimpleDateFormat df;

    public ExcelExportUtil()
    {
        this.objsMap = null;

        this.title = null;

        this.headNames = null;

        this.fieldNames = null;

        this.boningList = null;

        this.titleRow = null;
        this.df = new SimpleDateFormat("yyyy-MM-dd");
    }

    public void setType(FileTypeEnum type)
    {
    }

    public void setData(List<? extends Map> data) {
        if (StringUtils.isEmpty(this.title)) {
            this.boningList = data;
        } else {
            LinkedHashMap map = new LinkedHashMap();
            map.put(this.title, data);
            this.objsMap = map;
        }
    }

    public void setData(LinkedHashMap<String, List<? extends Map>> objsMap)
    {
        this.objsMap = objsMap;
    }

    public void writeToOuputstream(OutputStream os)
    {
        try
        {
            init();
            if ((this.objsMap == null) || (this.objsMap.isEmpty())) {
                throw new RuntimeException("需要导出的数据为空");
            }
            Set entrySet = this.objsMap.entrySet();

            String[] sheetNames = new String[this.objsMap.size()];
            int sheetNameNum = 0;
            for (Iterator localIterator1 = entrySet.iterator(); localIterator1.hasNext(); ) { Map.Entry entry = (Map.Entry)localIterator1.next();

                sheetNames[sheetNameNum] = ((String)entry.getKey());
                ++sheetNameNum;
            }

            HSSFSheet[] sheets = getSheets(this.objsMap.size(), sheetNames);
            int sheetNum = 0;
            for (Iterator localIterator2 = entrySet.iterator(); localIterator2.hasNext(); ) { Map.Entry entry = (Map.Entry)localIterator2.next();

                List list = (List)entry.getValue();

                creatTableHeadRow(sheets, sheetNum);

                String[] fieldNames = (String[])this.fieldNames.get(sheetNum);

                int rowNum = 1;
                for (Iterator localIterator3 = list.iterator(); localIterator3.hasNext(); ) { Map map = (Map)localIterator3.next();
                    HSSFRow contentRow = sheets[sheetNum].createRow(rowNum);
                    contentRow.setHeight((short) 300);

                    HSSFCell[] cells = getCells(contentRow, ((String[])this.fieldNames.get(sheetNum)).length);
                    if ((fieldNames == null) || (fieldNames.length == 0)) {
                        throw new RuntimeException("请指定Field");
                    }
                    int cellNum = 0;
                    for (int i = 0; i < fieldNames.length; ++i) {
                        Object value = map.get(fieldNames[i]);

                        setValue(cells[cellNum], value, this.titleRow.getCell(cellNum), sheets[sheetNum], cellNum);
                        ++cellNum;
                    }
                    ++rowNum;
                }

                adjustColumnSize(sheets, sheetNum);
                ++sheetNum;
            }
            wb.write(os);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
        finally {
            destroy();
        }
    }

    public void export() {
        try {
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\"+title+".xls");
            writeToOuputstream(fileOut);
            fileOut.close();
        } catch (Exception e) {
            System.out.println(e.toString());
        }
//        HttpServletResponse response = response();
//        response.reset();
//        response.setContentType("application/vnd.ms-excel");
//        try {
//            response.setHeader("Content-disposition", "attachment; filename=" + new String(this.title.getBytes("GB2312"), "ISO-8859-1") + ".xls");
//            writeToOuputstream(response.getOutputStream());
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }
    }

    public void setHeaders(String[] headers) {
        List list = new ArrayList();
        list.add(headers);
        this.headNames = list;
    }

    public void setHeaders(List<? extends String[]> headers)
    {
        this.headNames = headers;
    }

    public void setFields(String[] fields)
    {
        List list = new ArrayList();
        list.add(fields);
        this.fieldNames = list;
    }

    public void setFields(List<? extends String[]> fields)
    {
        this.fieldNames = fields;
    }

    public void setTitle(String title)
    {
        if (StringUtils.isEmpty(title))
            throw new RuntimeException("模版名称不能为空");

        this.title = title;
        if ((this.boningList != null) && (!(this.boningList.isEmpty())))
        {
            LinkedHashMap map = new LinkedHashMap();
            map.put(this.title, this.boningList);
            this.objsMap = map;
        }
    }
    /**
     * 导出excel工具
     *
     * @param title   文件名
     * @param headers 表头(与列名一一对应)
     * @param fields  列名(与表头一一对应)
     * @param list    需导出数据集合
     * @return 操作信息
     */
    public String exportUtil(String title, String[] headers, String[] fields, List<?> list) {
        List<Map<String, Object>> mapList = transBean2Map(list, "yyyy-MM-dd HH:mm:ss");
        if (mapList.isEmpty()) {
            return "需导出数据为空";
        }
        if (headers.length != fields.length) {
            return "表头数与列数不一致";
        } else {
            //导出数据
            this.setData(mapList);
            this.setHeaders(headers);
            this.setFields(fields);
            this.setTitle(title);
            //导出
            this.export();
        }
        return "成功";
    }
    /**
     * list<bean>转list<map>
     *
     * @param list list<bean>
     * @return list<map>
     */
    private static List<Map<String, Object>> transBean2Map(List<?> list, String dateFormat) {
        SimpleDateFormat simpleDateFormat = null;
        if (dateFormat != null) {
            simpleDateFormat = new SimpleDateFormat(dateFormat);
        }
        List<Map<String, Object>> mapList = new ArrayList();
        for (int i = 0; i < list.size(); i++) {
            Object obj = list.get(i);
            if (obj == null) {
                return Collections.emptyList();
            }
            Map<String, Object> map = new HashMap();
            try {
                BeanInfo beanInfo = Introspector.getBeanInfo(obj.getClass());
                PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
                for (PropertyDescriptor property : propertyDescriptors) {
                    String key = property.getName();
                    // 过滤class属性
                    if (!key.equals("class")) {
                        // 得到property对应的getter方法
                        Method getter = property.getReadMethod();
                        Object value = getter.invoke(obj);
                        //处理时间类型
                        if (simpleDateFormat != null && value != null && value instanceof Date) {
                            value = simpleDateFormat.format(value);
                        }
                        map.put(key, value);
                    }
                }
            } catch (Exception e) {
                System.out.println("transBean2Map Error:"+ e.getMessage());
            }
            mapList.add(map);
        }
        return mapList;
    }
    private HttpServletResponse response()
    {
        HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();
        return response;
    }

    private void init()
    {
        wb = new HSSFWorkbook();
        titleFont = wb.createFont();
        titleStyle = wb.createCellStyle();
        dateStyle = wb.createCellStyle();
        dateFont = wb.createFont();
        headStyle = wb.createCellStyle();
        headFont = wb.createFont();
        contentStyle = wb.createCellStyle();
        contentFont = wb.createFont();

        initTitleCellStyle();

        initTitleFont();

        initDateCellStyle();

        initDateFont();

        initHeadCellStyle();

        initHeadFont();

        initContentCellStyle();

        initContentFont();
    }

    private void setValue(HSSFCell valCell, Object val, HSSFCell titleCell, HSSFSheet sheet, int i)
    {
        String ttt = titleCell.getStringCellValue();
        int titleWidth = ttt.toString().length();
        int length = 0;
        if (val != null) {
            length = val.toString().length();
            valCell.setCellValue(val.toString());
        }
        if (length < titleWidth) {
            length = titleWidth;
        }

        if (length > 50)
            length = 50;

        sheet.setColumnWidth(i, length);
    }

    private void adjustColumnSize(HSSFSheet[] sheets, int sheetNum)
    {
        int i = 0; for (int len = ((String[])this.headNames.get(sheetNum)).length + 1; i < len; ++i)
        sheets[sheetNum].autoSizeColumn(i, true);
    }

    private void createTableTitleRow(HSSFSheet[] sheets, int sheetNum)
    {
        String[] fieldNames = (String[])this.fieldNames.get(sheetNum);
        if ((fieldNames == null) || (fieldNames.length == 0)) {
            throw new RuntimeException("请指定Filed");
        }
        CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, fieldNames.length);
        sheets[sheetNum].addMergedRegion(titleRange);
        HSSFRow titleRow = sheets[sheetNum].createRow(0);
        titleRow.setHeight((short) 800);
        HSSFCell titleCell = titleRow.createCell(0);
        titleCell.setCellStyle(titleStyle);

        titleCell.setCellValue(sheets[sheetNum].getSheetName());
    }

    private void createTableDateRow(HSSFSheet[] sheets, int sheetNum)
    {
        CellRangeAddress dateRange = new CellRangeAddress(1, 1, 0, ((String[])this.fieldNames.get(sheetNum)).length);
        sheets[sheetNum].addMergedRegion(dateRange);
        HSSFRow dateRow = sheets[sheetNum].createRow(1);
        dateRow.setHeight((short) 350);
        HSSFCell dateCell = dateRow.createCell(0);
        dateCell.setCellStyle(dateStyle);
        dateCell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
    }

    private void creatTableHeadRow(HSSFSheet[] sheets, int sheetNum)
    {
        HSSFRow headRow = sheets[sheetNum].createRow(0);
        this.titleRow = headRow;
        headRow.setHeight((short) 350);

        if ((this.headNames == null) || (this.headNames.isEmpty()))
            if (!(this.fieldNames.isEmpty())) {
                this.headNames = this.fieldNames;
            } else {
                throw new RuntimeException("请设置表头");
            }

        int num = 0; for (int len = ((String[])this.headNames.get(sheetNum)).length; num < len; ++num) {
        HSSFCell headCell = headRow.createCell(num);
        headCell.setCellStyle(headStyle);
        String[] s = (String[])this.headNames.get(sheetNum);
        headCell.setCellValue(s[num]);
    }
    }

    private HSSFSheet[] getSheets(int num, String[] names)
    {
        HSSFSheet[] sheets = new HSSFSheet[num];
        for (int i = 0; i < num; ++i)
            sheets[i] = wb.createSheet(names[i]);

        return sheets;
    }

    private HSSFCell[] getCells(HSSFRow contentRow, int num)
    {
        HSSFCell[] cells = new HSSFCell[num];
        int i = 0; for (int len = cells.length; i < len; ++i) {
        cells[i] = contentRow.createCell(i);
        cells[i].setCellStyle(contentStyle);
    }

        return cells;
    }

    private void initTitleCellStyle()
    {
        titleStyle.setAlignment((short) 2);
        titleStyle.setVerticalAlignment((short) 1);
        titleStyle.setFont(titleFont);
        titleStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.index);
    }

    private void initDateCellStyle()
    {
        dateStyle.setAlignment((short) 6);
        dateStyle.setVerticalAlignment((short) 1);
        dateStyle.setFont(dateFont);
        dateStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.index);
    }

    private void initHeadCellStyle()
    {
        headStyle.setAlignment((short) 2);
        headStyle.setVerticalAlignment((short) 1);
        headStyle.setFont(headFont);
        headStyle.setFillBackgroundColor(IndexedColors.YELLOW.index);
        headStyle.setBorderTop((short) 2);
        headStyle.setBorderBottom((short) 1);
        headStyle.setBorderLeft((short) 1);
        headStyle.setBorderRight((short) 1);
        headStyle.setTopBorderColor(IndexedColors.BLUE.index);
        headStyle.setBottomBorderColor(IndexedColors.BLUE.index);
        headStyle.setLeftBorderColor(IndexedColors.BLUE.index);
        headStyle.setRightBorderColor(IndexedColors.BLUE.index);
    }

    private void initContentCellStyle()
    {
        contentStyle.setAlignment((short) 2);
        contentStyle.setVerticalAlignment((short) 1);
        contentStyle.setFont(contentFont);
        contentStyle.setBorderTop((short) 1);
        contentStyle.setBorderBottom((short) 1);
        contentStyle.setBorderLeft((short) 1);
        contentStyle.setBorderRight((short) 1);
        contentStyle.setTopBorderColor(IndexedColors.BLUE.index);
        contentStyle.setBottomBorderColor(IndexedColors.BLUE.index);
        contentStyle.setLeftBorderColor(IndexedColors.BLUE.index);
        contentStyle.setRightBorderColor(IndexedColors.BLUE.index);

        contentStyle.setWrapText(true);
    }

    private void initTitleFont()
    {
        titleFont.setFontName("华文楷体");
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBoldweight((short) 700);
        titleFont.setCharSet(1);
        titleFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    private void initDateFont()
    {
        dateFont.setFontName("隶书");
        dateFont.setFontHeightInPoints((short) 10);
        dateFont.setBoldweight((short) 700);
        dateFont.setCharSet(1);
        dateFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    private void initHeadFont()
    {
        headFont.setFontName("宋体");
        headFont.setFontHeightInPoints((short) 10);
        headFont.setBoldweight((short) 700);
        headFont.setCharSet(1);
        headFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    private void initContentFont()
    {
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 10);
        contentFont.setBoldweight((short) 400);
        contentFont.setCharSet(1);
        contentFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    private void destroy()
    {
        this.headNames = null;
        this.fieldNames = null;
        this.title = null;
        this.objsMap = null;
        this.boningList = null;
    }
}
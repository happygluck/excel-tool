package com.liutao.excel.common;

import com.liutao.excel.common.annotation.ExcelColumn;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel导入导出
 */
@Component
public class Excel {
    /**
     * Excel导入
     *
     * @param inputStream Excel文件流
     * @param classType   Class
     * @param <T>         返回数据实体
     * @return 实体数据列表
     * @throws IOException
     */
    public <T> List<T> importXlsx(InputStream inputStream, Class<T> classType) throws IOException {
        List<T> list = new ArrayList<>();
        //读取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        //读取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        Field[] fields = classType.getDeclaredFields();
        sheet.forEach(row -> {
            T t = null;
            try {
                t = classType.newInstance();
            } catch (Exception e) {
            }
            if (row.getRowNum() > 0) {
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                        int index = column.index();
                        Cell cell = row.getCell(index);
                        if (cell == null) {
                            t = null;
                            break;
                        }
                        String cellValue = cell.getStringCellValue();
                        try {
                            setFieldValue(t, field, cellValue);
                        } catch (Exception e) {
                        }
                    }
                }
                if (t != null)
                    list.add(t);
            }
        });
        //关闭工作簿
        inputStream.close();
        workbook.close();
        return list;
    }

    /**
     * Excel导出
     *
     * @param title     导出内容标题
     * @param data      导出数据
     * @param classType Class
     * @return
     * @throws IOException
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    public ByteArrayOutputStream export(String title, List data, Class classType) throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        // 声明一个工作薄(初始化1000行)
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
        workbook.setCompressTempFiles(true);
        // 标题样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBold(true);
        titleStyle.setFont(titleFont);
        // 列头样式
        CellStyle headerStyle = workbook.createCellStyle();
        //headerStyle.setFillPattern(FillPatternType.THICK_BACKWARD_DIAG);
        headerStyle.setFillBackgroundColor((short) 20);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 13);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        Font cellFont = workbook.createFont();
        cellFont.setBold(false);
        cellStyle.setFont(cellFont);
        // 生成一个(带标题)表格
        SXSSFSheet sheet = workbook.createSheet();
        Field[] fields = classType.getDeclaredFields();
        // 设置列宽
        int i = 0;
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                int width = column.width();
                if (width < 10)
                    width = 10;
                if (width > 100)
                    width = 100;
                sheet.setColumnWidth(i, width * 256);
                i++;
            }
        }
        // 遍历集合数据，产生数据行
        int rowIndex = 0;
        for (Object obj : data) {
            if (rowIndex == 65535 || rowIndex == 0) {
                if (rowIndex != 0)
                    // 如果数据超过了，则在第二页显示
                    sheet = workbook.createSheet();

                // 标题 rowIndex=0
                SXSSFRow titleRow = sheet.createRow(0);
                titleRow.createCell(0).setCellValue(title);
                titleRow.getCell(0).setCellStyle(titleStyle);
                int lastCol = i - 1;
                if (lastCol <= 0)
                    lastCol = 1;
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol));

                // 列头 rowIndex =1
                SXSSFRow headerRow = sheet.createRow(1);
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                        headerRow.createCell(column.index()).setCellValue(column.name());
                        headerRow.getCell(column.index()).setCellStyle(headerStyle);
                    }

                }
                rowIndex = 2;//数据内容从 rowIndex=2开始
            }

            SXSSFRow dataRow = sheet.createRow(rowIndex);
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                    SXSSFCell newCell = dataRow.createCell(column.index());
                    String cellValue = getFieldValue(obj, field);
                    newCell.setCellValue(cellValue);
                    newCell.setCellStyle(cellStyle);
                }
            }
            rowIndex++;
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        workbook.dispose();
        return outputStream;
    }

    // 获取实体字段值
    private String getFieldValue(Object obj, Field field) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String fieldValue = "";
        // 获取属性的名字
        String name = field.getName();
        // 将属性的首字符大写，方便构造get，set方法
        name = name.substring(0, 1).toUpperCase() + name.substring(1);
        // 获取属性的类型
        Type fieldType = field.getType();
        if (fieldType == String.class) {
            Method mGet = obj.getClass().getMethod("get" + name);
            // 调用getter方法获取属性值
            String value = (String) mGet.invoke(obj);
            if (value != null) {
                fieldValue = value;
            }
        }
        if (fieldType == BigDecimal.class) {
            Method mGet = obj.getClass().getMethod("get" + name);
            // 调用getter方法获取属性值
            BigDecimal value = (BigDecimal) mGet.invoke(obj);
            if (value != null) {
                fieldValue = value.toString();
            }
        }
        return fieldValue;
    }

    // 设置实体字段值
    private void setFieldValue(Object obj, Field field, String value) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // 获取属性的名字
        String name = field.getName();
        // 将属性的首字符大写，方便构造get，set方法
        name = name.substring(0, 1).toUpperCase() + name.substring(1);
        // 获取属性的类型
        Type fieldType = field.getType();
        if (fieldType == String.class) {
            Method mSet = obj.getClass().getMethod("set" + name, new Class[]{String.class});
            // 调用setter方法设置属性值
            mSet.invoke(obj, new Object[]{new String(value)});
        }
        if (fieldType == BigDecimal.class) {
            Method mSet = obj.getClass().getMethod("set" + name, new Class[]{BigDecimal.class});
            // 调用setter方法设置属性值
            mSet.invoke(obj, new Object[]{new BigDecimal(value)});
        }
    }
}

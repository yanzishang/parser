package com.example.excel;

import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;

public class ExcelUtils<T> {

    /**
     * 导入excel,输出bean对象
     * 
     * @param inputStream 传入的流
     * @param map 格式： {"测试":"test"},其中的test为T的类成员
     * @param clazz 类
     * @param type 文件格式
     * @return 数据集
     */
    public List<T> excelReader(InputStream inputStream, Map<String, String> map, Class<T> clazz, String type) {
        List<List<String>> excelValue = readExcel(inputStream, type);
        List<T> tList = Lists.newArrayList();
        try {
            for (int i = 1; i < excelValue.size(); i++) {
                T t = clazz.newInstance();
                for (int num = 0; num < excelValue.get(0).size(); num++) {
                    for (String key : map.keySet()) {
                        if (excelValue.get(0).get(num).equals(key)) {
                            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(map.get(key), clazz);
                            Method method = propertyDescriptor.getWriteMethod();
                            method.invoke(t, excelValue.get(i).get(num));
                        }
                    }
                }
                tList.add(t);
            }
        } catch (Exception e) {
            throw new RuntimeException("解析excel失败", e);
        }
        return tList;
    }

    /**
     * 读取excel
     * 
     * @param inputStream 输入流
     * @param type 文件类型
     * @return excel数值
     */
    private List<List<String>> readExcel(InputStream inputStream, String type) {
        List<List<String>> excelValue = Lists.newArrayList();
        try {
            Workbook wb;
            if ("xls".equals(type)) {
                wb = new HSSFWorkbook(inputStream);
            } else if ("xlsx".equals(type)) {
                wb = new XSSFWorkbook(inputStream);
            } else {
                throw new RuntimeException("文件格式不正确");
            }
            Sheet sheet = wb.getSheet(wb.getSheetName(0));
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                // row为null,直接跳到下次循环
                if (row == null) {
                    continue;
                }
                List<String> values = Lists.newArrayList();
                for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    // 所有cell都按照string来处理
                    cell.setCellType(CellType.STRING);
                    String cellValue = cell.getStringCellValue().trim();
                    if (StringUtils.isNotEmpty(cellValue)) {
                        values.add(cellValue);
                    }
                }
                if (values.size() != 0) {
                    excelValue.add(values);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("解析excel失败", e);
        }
        return excelValue;
    }

    /**
     * 导出excel
     *
     * @param headerList 有序列
     * @param headers 包含列,Map<String, String>类型,格式： {"测试":"test"},其中的test为T的类成员
     * @param dataList 要输出的数据集
     * @return workbook对象
     */
    public Workbook excelWriter(String[] headerList, Map<String, String> headers, List<T> dataList) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        // 设置cell的样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontName("微软雅黑");
        titleFont.setColor(Font.COLOR_NORMAL);
        titleStyle.setFont(titleFont);
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.cloneStyleFrom(titleStyle);
        Font bodyFont = workbook.createFont();
        bodyFont.setFontName("宋体");
        bodyFont.setColor(Font.COLOR_NORMAL);
        bodyStyle.setFont(bodyFont);
        // 过滤有序列和包含列的相同部分
        List<String> orderHeaders = Lists.newArrayList();
        for (String header : headerList) {
            for (String key : headers.keySet()) {
                if (header.equals(key)) {
                    orderHeaders.add(key);
                }
            }
        }
        // 第一行
        Row firstRow = sheet.createRow(0);
        // 第一行数据
        for (int i = 0; i < orderHeaders.size(); i++) {
            String header = orderHeaders.get(i);
            Cell cell = firstRow.createCell(i);
            cell.setCellStyle(titleStyle);
            cell.setCellValue(header);
        }

        int rowIndex = 1;
        for (T t : dataList) {
            Row row = sheet.createRow(rowIndex++);
            for (int i = 0; i < orderHeaders.size(); i++) {
                Cell cell = row.createCell(i);
                for (String key : headers.keySet()) {
                    if (firstRow.getCell(i).getStringCellValue().equals(key)) {
                        try {
                            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(headers.get(key),
                                    t.getClass());
                            Method method = propertyDescriptor.getReadMethod();
                            Object cellObj = method.invoke(t);
                            cell.setCellStyle(bodyStyle);
                            setCellValue(cell, cellObj);
                        } catch (Exception e) {
                            throw new RuntimeException("输出excel失败", e);
                        }

                    }
                }
            }
        }
        return workbook;
    }

    /**
     * 设置列值
     *
     * @param cell 传入的列对象
     * @param cellObj T的字段值
     */
    private void setCellValue(Cell cell, Object cellObj) {
        String textValue;
        if (cellObj instanceof Date) {
            Date date = (Date) cellObj;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            textValue = sdf.format(date);
        } else {
            // 其它数据类型都当作字符串处理
            textValue = cellObj.toString();
        }
        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
        Matcher matcher = p.matcher(cellObj.toString());
        if (matcher.matches()) {
            // 是数字当作double处理
            cell.setCellValue(Double.parseDouble(textValue));
        } else {
            cell.setCellValue(textValue);
        }
    }
}

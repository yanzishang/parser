package com.example;

import java.io.*;
import java.util.List;
import java.util.Map;

import com.example.excel.ExcelUtils;
import org.junit.Assert;
import org.junit.Test;

import com.example.model.User;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class ExcelUtilsTest {

    @Test
    public void readExcelTest() throws Exception {
        ExcelUtils<User> excelUtils = new ExcelUtils<User>();
        File file = new File("src/main/resources/excel/in.xls");
        String type = file.getName().split("\\.")[1];
        InputStream inputStream = new FileInputStream(file);
        Map<String, String> headers = Maps.newHashMap();
        headers.put("姓名", "name");
        headers.put("年龄", "age");
        headers.put("性别", "sex");
        List<User> userList = excelUtils.excelReader(inputStream, headers, User.class, type);
        inputStream.close();
        Assert.assertTrue("userList.size = 3", userList.size() == 3);
    }

    @Test
    public void writeExcelTest() throws Exception {
        ExcelUtils<User> excelUtils = new ExcelUtils<User>();
        File file = new File("src/main/resources/excel/out.xls");
        Map<String, String> headers = Maps.newHashMap();
        headers.put("姓名", "name");
        headers.put("年龄", "age");
        headers.put("性别", "sex");
        String[] headerList = new String[] { "姓名", "性别", "年龄" };
        List<User> userList = Lists.newArrayList();
        for (int i = 0; i < 3; i++) {
            User user = new User();
            user.setName("name" + i);
            user.setSex("男" + i);
            user.setAge(i + "");
            userList.add(user);
        }
        OutputStream outputStream = new FileOutputStream(file);
        excelUtils.excelWriter(headerList, headers, userList).write(outputStream);
        outputStream.flush();
        outputStream.close();
    }
}

package com.example.csv;

import java.io.File;
import java.io.FileReader;
import java.io.Reader;
import java.util.List;
import java.util.Scanner;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.StringUtils;

import com.example.model.User;
import com.google.common.collect.Lists;

public class CsvUtils {

    private static final String DEFAULT_SEPARATOR = ",";

    public static List<String> csvReader(File file) throws Exception {
        return csvReader(file, DEFAULT_SEPARATOR);
    }

    public static List<String> csvReader(File file, String separators) throws Exception {
        Scanner scanner = new Scanner(file);
        scanner.useDelimiter(separators);
        List<String> result = Lists.newArrayList();

        while (scanner.hasNext()) {
            String cvsLine = scanner.next();
            if (StringUtils.isEmpty(cvsLine)) {
                continue;
            }
            result.add(cvsLine);
        }

        return result;
    }

    public static List<User> read(String fileName) throws Exception {
        Reader in = new FileReader(fileName);
        List<User> users = Lists.newArrayList();
        Iterable<CSVRecord> records = CSVFormat.EXCEL.withFirstRecordAsHeader().parse(in);
        for (CSVRecord record : records) {
            User user = new User();
            user.setName(record.get("姓名"));
            user.setSex(record.get("性别"));
            user.setAge(record.get("年龄"));
            users.add(user);
        }
        return users;
    }
}

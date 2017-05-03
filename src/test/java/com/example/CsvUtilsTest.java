package com.example;

import java.io.File;
import java.util.List;

import org.junit.Test;

import com.example.csv.CsvUtils;
import com.example.model.User;

/**
 * Created by zack on 17-5-3.
 */
public class CsvUtilsTest {

    @Test
    public void readCsvTest() throws Exception {
        File file = new File("src/main/resources/csv/in.csv");
        List<String> result = CsvUtils.csvReader(file);
        for (String r : result) {
            System.out.println(r);
        }
    }

    @Test
    public void read() throws Exception {
        File file = new File("src/main/resources/csv/in.csv");
        List<User> result = CsvUtils.read("src/main/resources/csv/in.csv");
        for (User r : result) {
            System.out.println(r.getName());
        }
    }

}

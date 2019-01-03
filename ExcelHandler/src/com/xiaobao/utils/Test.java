package com.xiaobao.utils;

import java.io.File;
import java.io.*;
import java.util.*;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;

import com.alibaba.fastjson.*;


public class Test {

    public static void main(String[] args) throws IOException {
        String jsonFile = "/Users/wuyong/Documents/GitHub/ExcelHanlder/ExcelHandler/course.json";


        FileInputStream fileInputStream = new FileInputStream(jsonFile);
        InputStreamReader reader = new InputStreamReader(fileInputStream);
        BufferedReader bufferedReader = new BufferedReader(reader);

        StringBuilder stringBuilder = new StringBuilder();
        String line;

        while ((line = bufferedReader.readLine()) != null) {
            stringBuilder.append(line);
        }

//        System.out.println(stringBuilder);
        JSONArray jsonArray = JSON.parseArray(String.valueOf(stringBuilder));

        System.out.println(jsonArray);

    }

}

package com.xiaobao.utils;

import java.io.*;
import java.util.*;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;

import com.alibaba.fastjson.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelTxt {

    public List<String> getCourseInfo(String filePath) {

        List<String> list = new ArrayList<>();

        String lineTxt;

        try {

            FileInputStream fileInputStream = new FileInputStream(filePath);
            InputStreamReader reader = new InputStreamReader(fileInputStream);
            BufferedReader bufferedReader = new BufferedReader(reader);


            while ((lineTxt = bufferedReader.readLine()) != null) {

                list.add(lineTxt);

            }
        } catch (IOException e) {
            e.printStackTrace();
        }


        return list;

    }

    public String count(String data) {

        String count;
        String[] list = data.split(":");

        count = list[list.length - 1];

        return count;

    }

    public String[] courseList(String data) {


        String[] list = data.split(":");

        String courses = list[list.length - 1];


        return courses.split(",");

    }

    public int randomInt(int max) {

        Random random = new Random();

        return random.nextInt(max);

    }


    public void writeExcel(String filePath, int col, int row, String value) {
        try {
            File file = new File(filePath);
            if (file.exists()) {

                Workbook workbook = Workbook.getWorkbook(file);

                WritableWorkbook wb = Workbook.createWorkbook(file, workbook);

                WritableSheet sheet = wb.getSheet(0);

                Label label = new Label(col, row, value);

                sheet.addCell(label);

                wb.write();
                wb.close();

            }
        } catch (IOException | BiffException | WriteException e) {
            e.printStackTrace();
        }
    }

    public static void writeXlsx(String filePath) throws IOException {

        FileInputStream fs = new FileInputStream(filePath);
        POIFSFileSystem pfs = new POIFSFileSystem(fs);
        XSSFWorkbook wb = new XSSFWorkbook(String.valueOf(pfs));

        XSSFSheet sheet = wb.getSheetAt(0);

        FileOutputStream out = new FileOutputStream(filePath);

    }

    public static void main(String[] args) throws JSONException, IOException {

        String excelPath = "/Users/wuyong/SparksInterface/Validator/ImportSchedule/test.xlsx";
        String jsonFile = "/Users/wuyong/Documents/GitHub/ExcelHanlder/ExcelHandler/course.json";


        ExcelTxt excelTxt = new ExcelTxt();

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

//        System.out.println(jsonArray);
        for (Object aJsonArray : jsonArray) {

            JSONObject cellData = (JSONObject) aJsonArray;
            int cellRow = cellData.getIntValue("row");
            int cellCol = cellData.getIntValue("col");
            String value = cellData.getString("course");

            excelTxt.writeExcel(excelPath, cellCol, cellRow, value);

//            System.out.println(String.format("Cell (%s, %s) write done.", cellRow, cellCol));

        }
        System.out.println("All done.");
    }

}
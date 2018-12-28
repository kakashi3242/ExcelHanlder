package com.xiaobao.utils;

import java.io.*;
import java.util.*;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;

import org.json.*;

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

    public static void main(String[] args) throws JSONException {

        String excelPath = "/Users/wuyong/SparksInterface/Validator/ImportSchedule/VerticalVersion.xls";
        String timetableInfo = "/Users/wuyong/SparksInterface/Validator/ImportSchedule/course.txt";


        ExcelTxt excelTxt = new ExcelTxt();

        List<String> data = excelTxt.getCourseInfo(timetableInfo);

        int row = Integer.parseInt(excelTxt.count(data.get(0)));

        int col = Integer.parseInt(excelTxt.count(data.get(1)));

//        System.out.println(row + "-" + col);

        String[] course = excelTxt.courseList(data.get(2));


//        JSONArray timetableList = new JSONArray();
        JSONObject timetables = new JSONObject();

        List list = new ArrayList();

        for (int r = 3; r < row; r++) {
            for (int c = 1; c < col; c++) {

                JSONObject timetable = new JSONObject();

                timetable.put("row", r);
                timetable.put("col", c);
                timetable.put("course", course[excelTxt.randomInt(course.length)]);

                list.add(timetable);

            }
        }
        timetables.put("timetable", list);

//        System.out.println(timetables);
        JSONArray jsonArray = (JSONArray) timetables.get("timetable");


//        System.out.println(jsonArray);
        for (int i = 0; i < jsonArray.length(); i++) {

            JSONObject cellData = (JSONObject) jsonArray.get(i);
            int cellRow = cellData.getInt("row");
            int cellCol = cellData.getInt("col");
            String value = cellData.getString("course");

            excelTxt.writeExcel(excelPath, cellCol, cellRow, value);

//            System.out.println(String.format("Cell (%s, %s) write done.", cellRow, cellCol));

        }
        System.out.println("All done.");
    }

}
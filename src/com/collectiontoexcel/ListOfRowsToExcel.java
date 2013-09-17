package com.collectiontoexcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ListOfRowsToExcel {

    public void generateExcel(List<Object[]> listOfRows) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet");
        int rownum = 0;
        for (Object[] objArr : listOfRows) {
            Row row = sheet.createRow(rownum++);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof Date) {
                    cell.setCellValue(obj.toString()); //use SimpleDateFormat instead
                } else if (obj instanceof Boolean) {
                    cell.setCellValue((Boolean) obj);
                } else if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File("report.xls"));
            /*
             * In webapp case, pass ServletOutputStream to workbook.write()
             * response.setContentType("application/vnd.ms-excel; charset=UTF-8");
             * response.setCharacterEncoding("UTF-8");
             * 
             * response.addHeader("content-disposition", "attachment; filename=" + filename);
             *    // in case of downloading as attachment
             * response.addHeader("content-disposition", "inline; filename=" + filename);
             *   // in case of displaying inline    
             * ServletOutputStream out = response.getOutputStream();
             */
            workbook.write(out); 
            out.close();
            System.out.println("Excel written successfully. Find report.xls in your project root folder.");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Calendar cal = Calendar.getInstance();
        List<Object[]> listOfRows = new ArrayList<Object[]>();
        listOfRows.add(new Object[] { "Emp No.", "Name", "Salary", "DOB" });
        listOfRows.add(new Object[] { 1d, "こんにちは世界", 1500000d, cal.getTime() });
        listOfRows.add(new Object[] { 2d, "नमस्ते विश्व", 800000d, cal.getTime() });
        listOfRows.add(new Object[] { 3d, "హలో వరల్డ్", 700000d, cal.getTime() });
        listOfRows.add(new Object[] { 4d, "привет мир", 700000d, cal.getTime() });
        listOfRows.add(new Object[] { 5d, "Hello World", 700000d, cal.getTime() });
        new ListOfRowsToExcel().generateExcel(listOfRows);
    }
}

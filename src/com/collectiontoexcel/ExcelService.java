package com.collectiontoexcel;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {
	
	public void download(HttpServletResponse response, String fileName, Object searchList, Object searchObjType) throws IOException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		/*
		 * Use java reflection and call member functions
		 */
	    final BeanWrapper wrapper = new BeanWrapperImpl(searchObjType);
		List<List<Object>> listOfRows = new ArrayList<List<Object>>();
		List<Object> listOfColNames = new ArrayList<Object>();
		for (final PropertyDescriptor descriptor : wrapper.getPropertyDescriptors()) {
			listOfColNames.add(descriptor.getName());
		}
		listOfRows.add(listOfColNames);
		for(Object searchObj : (List<Object>)searchList) {
			List<Object> listOfCols = new ArrayList<Object>();
			for (final PropertyDescriptor descriptor : wrapper.getPropertyDescriptors()) {
				listOfCols.add(descriptor.getReadMethod().invoke(searchObj));
			}
			listOfRows.add(listOfCols);
		}
		download(response, fileName,listOfRows);
	}
	
	public void download(HttpServletResponse response, String fileName, List<List<Object>> listOfRows) throws IOException {

		response.setContentType("application/vnd.ms-excel;charset=UTF-8");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment; filename=report.xls");
		response.setHeader("Pragma", "public");
        response.setHeader("Cache-Control", "no-store");
        response.addHeader("Cache-Control", "max-age=0");
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");
		 		 
		int rownum = 0;
		for (List<Object> objArr : listOfRows) {
		    Row row = sheet.createRow(rownum++);
		    int cellnum = 0;
		    for (Object obj : objArr) {
		        Cell cell = row.createCell(cellnum++);
		        if(obj instanceof Date)
		            cell.setCellValue((Date)obj);
		        else if(obj instanceof Boolean)
		            cell.setCellValue((Boolean)obj);
		        else if(obj instanceof String)
		            cell.setCellValue((String)obj);
		        else if(obj instanceof Double)
		            cell.setCellValue((Double)obj);
		    }
		}
		
		ServletOutputStream outStream = response.getOutputStream();
		workbook.write(outStream);
		outStream.flush();
		outStream.close();
	
	}
	
}

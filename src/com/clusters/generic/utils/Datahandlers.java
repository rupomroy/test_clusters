package com.clusters.generic.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Datahandlers {
	public static String getDataFromExcel(String Filename, String sheetName, int row_index, int  cell_index) {
		String  data=null;
		try
		{
			File f = new File("./test-data" +Filename+".xlsx");
			FileInputStream fis = new FileInputStream(f);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet s = wb.getSheet(sheetName);
			Row r = s.getRow(row_index);
			Cell c = r.getCell(cell_index);
			data=c.toString();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return data;
	}
	public static void storedatatoexcel(String filename, String sheetname, int row, int cell, String data )
	{
		
		try
		{
			File f=new File("./test-data"+filename+ ".xlsx");
			FileInputStream fis=new FileInputStream(f);
			Workbook wb=WorkbookFactory.create(fis);
			Sheet ss= wb.getSheet(sheetname);
			  Row rr = ss.createRow(row);
			Cell c1  =rr.createCell(cell);
		          c1.setCellValue(data);
		         FileOutputStream fos=new FileOutputStream(f);
		         wb.write(fos);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	

	}
}

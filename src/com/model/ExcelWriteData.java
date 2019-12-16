package com.model;

import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteData {
	public void setcelldata(String filenm,String sheetnm,int rowno,int colnm,String dataval) throws IOException
	{
		FileInputStream fis1=new FileInputStream(filenm);
		XSSFWorkbook wb1=new XSSFWorkbook(fis1);
		XSSFSheet sheet1=wb1.getSheet(sheetnm);
		XSSFRow row1=sheet1.createRow(rowno);
		XSSFCell cell1=row1.createCell(colnm);
		cell1.setCellValue(dataval);
		
		FileOutputStream fo=new FileOutputStream(filenm);
		wb1.write(fo);
	}

	
}

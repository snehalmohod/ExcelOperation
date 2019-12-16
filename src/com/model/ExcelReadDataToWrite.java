package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadDataToWrite {
	
	public void excelRead(String filename,String Sheetname) throws IOException
	{    int rowno1=0;
	     int colno1=0;
	     System.out.println("ExcelRead Operation");
		//int arrayExcel[][]=null;
		//code to address for file
		FileInputStream fis=new FileInputStream(filename);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet(Sheetname);
		//fetch data form index 2,2
		XSSFRow row=sheet.getRow(2);
		XSSFCell cell=row.getCell(2);
		String cellval=cell.getStringCellValue();
		System.out.println(cellval);
		//find rowcount
		int rows=sheet.getLastRowNum();
		System.out.println("No of rows :"+rows);
		int rowcount=rows+1;
		System.out.println("Total no of rows "+rowcount);
		//find columncount
		int columncount=sheet.getRow(rows).getLastCellNum();
		System.out.println("No of columns are "+columncount);
		//fetch the data from excel sheet
		int arrayExcel[][]=new int [rowcount][columncount];
		for(int i=0;i<rowcount;i++)
		{
			for(int j=0;j<columncount;j++)
			{
				DataFormatter dataformat=new DataFormatter();
			System.out.println(dataformat.formatCellValue(sheet.getRow(i).getCell(j)));
			ExcelWriteData writedata=new ExcelWriteData();
			writedata.setcelldata("C:\\Users\\Nitin\\workspace\\ExcelOperation\\WriteStudentDetails.xlsx", "Sheet1", rowno1++, colno1, dataformat.formatCellValue(sheet.getRow(i).getCell(j)));
			
			}
		}
		
	}

	public static void main(String[] args) throws IOException {
		ExcelReadDataToWrite data1=new ExcelReadDataToWrite();
		data1.excelRead("C:\\Users\\Nitin\\workspace\\ExcelOperation\\StudentDetails.xlsx", "Sheet1");

	}

}

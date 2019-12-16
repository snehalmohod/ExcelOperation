package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadXSSFDemo {
	
	public void readExcel(String filename,String sheetname) throws IOException
	{ 	
				
	int arrayExcel [][]=null;
	//code for any file address which we used
	FileInputStream fis=new FileInputStream(filename);
	XSSFWorkbook wb=new XSSFWorkbook(fis);
	XSSFSheet sheet=wb.getSheet(sheetname);
	//fetch data from from(2,2)
	//XSSFCell row=sheet.getRow(2).getCell(2);
	XSSFRow row=sheet.getRow(2);
	XSSFCell cell=row.getCell(2);
	     String val=cell.getStringCellValue();
	     System.out.println("The data at 2,2 index is :"+val);
	     //find rowcount
	    int rows=sheet.getLastRowNum();
	    System.out.println("No of rows are :"+rows);
	    int rowcount=rows+1;
	    System.out.println("Total rowcount is :"+rowcount);
	    //find columncount
	    int columncount=sheet.getRow(rows).getLastCellNum();
	    System.out.println("No of columns are :"+columncount);
	    //fetch data from excel
	    arrayExcel=new int [rowcount][columncount];
	    for(int i=0;i<rowcount;i++)
	    {
	    	for(int j=0;j<columncount;j++)
	    	{
	    		
	    		System.out.println(sheet.getRow(i).getCell(j));
	    	}
	    }
	
	}

	public static void main(String[] args) throws IOException {
		ExcelReadXSSFDemo data=new ExcelReadXSSFDemo();
		data.readExcel("C:\\Users\\Nitin\\workspace\\ExcelOperation\\StudentDetails.xlsx", "sheet1");

	}

}

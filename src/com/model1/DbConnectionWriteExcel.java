package com.model1;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import com.model.ExcelWriteData;

public class DbConnectionWriteExcel {
	
	public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException {
		int rowno=0;
		int colno=0;
      Class.forName("com.mysql.jdbc.Driver");
    Connection con=DriverManager.getConnection("jdbc:mysql://localhost:3306/test", "root", "root");
	System.out.println("1");  
	Statement smt=con.createStatement();
	ResultSet res=smt.executeQuery("SELECT * FROM empdetails1");
	while(res.next())
	{
		String empnm=res.getString("EmpName");
		System.out.println(empnm);
		ExcelWriteData dbdata=new ExcelWriteData();
		dbdata.setcelldata("C:\\Users\\Nitin\\workspace\\ExcelOperation\\WritedataSql.xlsx", "Sheet1", rowno++, colno, empnm);
	
		
	}
	
	}

}

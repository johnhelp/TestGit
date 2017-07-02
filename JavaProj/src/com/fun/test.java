package com.fun;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class test {

	private static long rowCount = 0;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		System.out.println("First java code1");
		FileInputStream fis = new FileInputStream("D:\\tobeDelete\\Testdata.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		rowCount  = sheet.getLastRowNum();
		System.out.println(rowCount);
		//Iterator<itr> 
		
		XSSFCell value = sheet.getRow(0).getCell(0);		
		System.out.println(value);
	}

}

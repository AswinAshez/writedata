package org.writedata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Writedata {
	public static void main(String[] args)throws IOException {
		
		File file=new File("C:\\Users\\Ashwin\\eclipse-workspace\\Writedata\\Excel\\Book1.xlsx");
		FileInputStream fileInputStream=new FileInputStream(file);
		Workbook workbook=new XSSFWorkbook(fileInputStream);
		Sheet sheet=workbook.getSheet("Sheet1");
		Row row=sheet.createRow(8);
		Cell cell=row.createCell(2);
		cell.setCellValue("Delna");
		FileOutputStream fileOutputStream=new FileOutputStream(file);
		workbook.write(fileOutputStream);
	}}
	


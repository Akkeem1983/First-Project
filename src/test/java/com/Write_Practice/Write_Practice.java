package com.Write_Practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Practice {

	public static void main(String[] args) throws IOException {

		File ak = new File("C:\\Users\\admin\\eclipse-workspace\\DataDriven\\Excel\\Data Read.xlsx");
		FileInputStream ak1 = new FileInputStream(ak);
		Workbook ak2 = new XSSFWorkbook(ak1);
		ak2.createSheet("Akkeem1").createRow(0).createCell(0).setCellValue("Student Name");
		ak2.getSheet("Akkeem1").getRow(0).createCell(1).setCellValue("RollNumber");
		ak2.getSheet("Akkeem1").getRow(0).createCell(2).setCellValue("Course Name");
		ak2.getSheet("Akkeem1").getRow(0).createCell(3).setCellValue("Course Completed Status");

		ak2.getSheet("Akkeem1").createRow(1).createCell(0).setCellValue("Saran");
		ak2.getSheet("Akkeem1").getRow(1).createCell(1).setCellValue(12346);
		ak2.getSheet("Akkeem1").getRow(1).createCell(2).setCellValue("B.A Tamil");
		ak2.getSheet("Akkeem1").getRow(1).createCell(3).setCellValue("Completed");

		ak2.getSheet("Akkeem1").createRow(2).createCell(0).setCellValue("Naren");
		ak2.getSheet("Akkeem1").getRow(2).createCell(1).setCellValue(576576);
		ak2.getSheet("Akkeem1").getRow(2).createCell(2).setCellValue("M.A Tamil");
		ak2.getSheet("Akkeem1").getRow(2).createCell(3).setCellValue("In Complete");

		ak2.getSheet("Akkeem1").createRow(3).createCell(0).setCellValue("Kavin");
		ak2.getSheet("Akkeem1").getRow(3).createCell(1).setCellValue(678687);
		ak2.getSheet("Akkeem1").getRow(3).createCell(2).setCellValue("M.A Tamil");
		ak2.getSheet("Akkeem1").getRow(3).createCell(3).setCellValue("Completed");

		FileOutputStream ak3 = new FileOutputStream(ak);
		ak2.write(ak3);

	}

}

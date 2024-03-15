package com.read_Practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Practice {
	
	
	public static void allData_Read() throws IOException {
		
		File ak = new File("C:\\Users\\admin\\eclipse-workspace\\DataDriven\\Excel\\Data Read.xlsx");
		FileInputStream ak1 = new FileInputStream(ak);
		Workbook ak2 = new XSSFWorkbook(ak1);
		Sheet sheetAt = ak2.getSheetAt(0);
		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
		for(int i=0; i<physicalNumberOfRows;i++) {
			Row row = sheetAt.getRow(i);
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			for(int j=0; j<physicalNumberOfCells;j++) {
			Cell cell = row.getCell(j);	
			CellType cellType = cell.getCellType();
			if(cellType.equals(CellType.STRING)) {
				System.out.println(cell.getStringCellValue());
			}
			else if(cellType.equals(cellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int a =(int)numericCellValue;
			System.out.println(a);
			}
			}
		}
		
		}
	
	
public static void ColumnData_Read() throws IOException {
		
		File ak = new File("C:\\Users\\admin\\eclipse-workspace\\DataDriven\\Excel\\Data Read.xlsx");
		FileInputStream ak1 = new FileInputStream(ak);
		Workbook ak2 = new XSSFWorkbook(ak1);
		Sheet sheetAt = ak2.getSheetAt(0);
		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
		for(int i=0; i<physicalNumberOfRows;i++) {
			Row row = sheetAt.getRow(i);
			Cell cell = row.getCell(0);	
			CellType cellType = cell.getCellType();
			if(cellType.equals(CellType.STRING)) {
				System.out.println(cell.getStringCellValue());
			}
			else if(cellType.equals(cellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int a =(int)numericCellValue;
			System.out.println(a);
			}
			}
		}
			
public static void rowData_Read() throws IOException {
		
		File ak = new File("C:\\Users\\admin\\eclipse-workspace\\DataDriven\\Excel\\Data Read.xlsx");
		FileInputStream ak1 = new FileInputStream(ak);
		Workbook ak2 = new XSSFWorkbook(ak1);
		Sheet sheetAt = ak2.getSheetAt(0);
			Row row = sheetAt.getRow(2);
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			for(int j=0; j<physicalNumberOfCells;j++) {
			Cell cell = row.getCell(j);	
			CellType cellType = cell.getCellType();
			if(cellType.equals(CellType.STRING)) {
				System.out.println(cell.getStringCellValue());
			}
			else if(cellType.equals(cellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				int a =(int)numericCellValue;
			System.out.println(a);
			}
			}
		}

public static void Particular_Data_Read() throws IOException {
	
	File ak = new File("C:\\Users\\admin\\eclipse-workspace\\DataDriven\\Excel\\Data Read.xlsx");
	FileInputStream ak1 = new FileInputStream(ak);
	Workbook ak2 = new XSSFWorkbook(ak1);
	Sheet sheetAt = ak2.getSheetAt(0);
		Row row = sheetAt.getRow(2);
		Cell cell = row.getCell(2);	
		CellType cellType = cell.getCellType();
		if(cellType.equals(CellType.STRING)) {
			System.out.println(cell.getStringCellValue());
		}
		else if(cellType.equals(cellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int a =(int)numericCellValue;
		System.out.println(a);
		}
		}
	
		
		
public static void main(String[] args) throws IOException {
	System.out.println("==========Read All Data from Excel=============");
	Read_Practice.allData_Read();
	System.out.println("==========Read All Data from Row=============");
	Read_Practice.rowData_Read();
	System.out.println("==========Read  Data from Particular Cell=============");
	Read_Practice.Particular_Data_Read();
	System.out.println("==========Read All Data from Particular Column=============");
	Read_Practice.ColumnData_Read();
}
}




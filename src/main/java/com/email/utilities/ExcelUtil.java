package com.email.utilities;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.email.config.ObjectRespo;

public class ExcelUtil extends ObjectRespo{

	//Get Excel File
	public static void getExcel(String filePath) throws IOException {
		try {
			dataFile = new FileInputStream(filePath);
			wbFile = new XSSFWorkbook(dataFile);
		} catch (Exception e) {
			System.out.println("File not found");
			e.printStackTrace();
		}
	}

	//Get Excel Sheet
	public static void getSheet(int sheetno) throws IOException {
		try {
			shFile = wbFile.getSheetAt(sheetno);
		} catch (Exception e) {
			System.out.println("Sheet not found");
			e.printStackTrace();
		}
	}

	//Row Count
	public static int getRowCount() {
		int rowCount = 0; 
		try {
			rowCount = shFile.getLastRowNum()+1;
		}catch(Exception e) {
			System.out.println("Did not get Rows");
			e.printStackTrace();
		}
		return rowCount;
	}

	//Column Count
	public static int getColumnCount(){
		int colCount=0;
		try {
			colCount = shFile.getRow(0).getLastCellNum();
		}catch(Exception e) {
			System.out.println("Did not get Columns");
			e.printStackTrace();
		}
		return colCount;
	}

	//Get String Value
	public static String getStringValue(int rowNum, int colNum) {
		String cellData = null;
		try {
			cellData = shFile.getRow(rowNum).getCell(colNum).getStringCellValue();
			cellData = cellData.trim();
		}catch(Exception e) {
			System.out.println("Data not found");
			e.printStackTrace();
		}
		return cellData;
	}

	//DataProvider from Excel
	public static Object[][] getData() throws IOException{

		int rowCount = ExcelUtil.getRowCount();
		int colCount = ExcelUtil.getColumnCount();

		Object[][] data = new Object[rowCount-1][colCount];

		for(int i=1; i<rowCount; i++)
		{
			for(int j=0; j<colCount; j++)
			{
				//change values to string
				data[i-1][j] = ExcelUtil.setCellDataToString(i, j);
			}
		}
		return data;
	}


	//Change cell values to String
	public static String setCellDataToString(int rowNum, int colNum) {
		XSSFCell cell = null;
		String cellData = null;
		try {
			cell = shFile.getRow(rowNum).getCell(colNum);
			cell.setCellType(CellType.STRING);
			cellData = cell.getStringCellValue().trim();
			cellData = cellData.trim();
		}catch(Exception e) {
			System.out.println("Data not found");
			e.printStackTrace();
		}
		return cellData;
	}
}
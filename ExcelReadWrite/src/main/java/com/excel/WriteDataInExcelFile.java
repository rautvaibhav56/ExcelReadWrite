package com.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataInExcelFile {

	public static void main(String[] args) throws IOException {

		// Step 1:Create Workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Step 2:Create Sheet
		XSSFSheet sheetTestData = workbook.createSheet("sheetTestData");

		// Step 3:Create Row
		Row row_0 = sheetTestData.createRow(0);

		// Step 4:Create Cell
		Cell cell_A = row_0.createCell(0);
		Cell cell_B = row_0.createCell(1);

		// Step 5:Add data in excel using FIS object
		cell_A.setCellValue("Vaibhav");
		cell_B.setCellValue("Raut");

		// Step 6:Add data in excel using FIS object
		File file = new File(System.getProperty("user.dir") + "/src/main/resources/testData/ExcelFiles/Excel.xlsx");
		FileOutputStream fo = new FileOutputStream(file);
		workbook.write(fo);
		
		// Step 7:Close
		fo.close();

		System.out.println("TestData Added sucessfully....!");
	}

}

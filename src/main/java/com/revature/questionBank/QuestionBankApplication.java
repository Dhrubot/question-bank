package com.revature.questionBank;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;

@SpringBootApplication
public class QuestionBankApplication {

	public static void main(String[] args) {

		SpringApplication.run(QuestionBankApplication.class, args);
		excelReader();

	}

	public static void excelReader() {
		String fileLocation = "/Users/dhrubo/Desktop/Citi Bank - Interview Questions.xlsx";
		try (
				FileInputStream file = new FileInputStream(fileLocation);
				Workbook wb = new XSSFWorkbook(file);
		){
//			for (int i = 0; i < wb.getNumberOfSheets(); i ++) {
//
//			}
			Sheet sheet = wb.getSheetAt(0);
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum();

			for (int i = rowStart; i < rowEnd; i ++) {
				Row row = sheet.getRow(i);
				if (row != null) {
					Cell cell = row.getCell(0);
					System.out.println(cell.getStringCellValue());
				} else {
					System.out.println("null");
				}
			}
			System.out.println("===============================");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}

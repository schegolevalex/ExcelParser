package com.schegolevalex.excel.ExcelParser;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class ExcelParserApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ExcelParserApplication.class, args);

		String inputFile = "C:\\Users\\Алексей\\Desktop\\Трубы ГОСТ 10704.xlsx";
		String outputFile = "C:\\Users\\Алексей\\Desktop\\Переделали Трубы ГОСТ 10704.xlsx";


		//читаем
		XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(inputFile));
		XSSFSheet myExcelSheet = myExcelBook.getSheet("Лист1");
		List<PipeTemp> pipes = new ArrayList<PipeTemp>();

		for (int rowIndex = 1; rowIndex <= 74; rowIndex++) {
			for (int cellIndex = 1; cellIndex <= 45; cellIndex++) {
				XSSFRow row = myExcelSheet.getRow(rowIndex);
				if (row.getCell(cellIndex).getNumericCellValue() != 0) {

					Double outerDiameter = row.getCell(0).getNumericCellValue();
					Double wallThickness = myExcelSheet.getRow(0).getCell(cellIndex).getNumericCellValue();
					Double mass = row.getCell(cellIndex).getNumericCellValue();

					PipeTemp pipe = new PipeTemp(outerDiameter, wallThickness, mass);
					pipes.add(pipe);
					System.out.println(pipes);
				}
			}
		}
		myExcelBook.close();


		//пишем
		XSSFWorkbook outputBook = new XSSFWorkbook();
		XSSFSheet sheet = outputBook.createSheet("Лист1");

		for (PipeTemp pipe : pipes) {
            XSSFRow outputRow = sheet.createRow(pipes.indexOf(pipe));

            XSSFCell outerDiameterCell = outputRow.createCell(0);
            outerDiameterCell.setCellValue(pipe.getOuterDiameter());

			XSSFCell wallThicknessCell = outputRow.createCell(1);
			wallThicknessCell.setCellValue(pipe.getWallThickness());

			XSSFCell massCell = outputRow.createCell(2);
			massCell.setCellValue(pipe.getMass());
		}

		outputBook.write(new FileOutputStream(outputFile));
		outputBook.close();

	}

}

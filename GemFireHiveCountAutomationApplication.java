package com.example.demo;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;

@EnableScheduling
@SpringBootApplication
public class GemFireHiveCountAutomationApplication {
			static Logger logger = Logger.getLogger(GemFireHiveCountAutomationApplication.class.getName());
	private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");

    
	public static void main(String[] args) {
		SpringApplication.run(GemFireHiveCountAutomationApplication.class, args);
		
	}
	static String staticPath = "E:\\GemfireCount\\";
	static boolean finalPopupValue = false;
	
    @Scheduled(cron = "0 27 15 * * ?", zone="IST")  //- Fires at 3:27 PM every day: // Follows 24hr Format
	public static void logic() {
		
		logger.setLevel(Level.INFO);
        logger.info("<<<<<<<<<<<<<<<<<--------------------->>>>>>>>>>>>>>>>>>>>");
		
		File root = new File(staticPath);
		if (root.exists()) {
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy");
			LocalDateTime now = LocalDateTime.now(); 
			String outputPath = staticPath + "\\output_"+dtf.format(now) + "\\";
			File input = new File(staticPath + "\\" + dtf.format(now));
			if (input.exists()) {
				String[] list = input.list();
				if (list.length > 0) {
					for (String list1 : list) {
						File soureceFile = new File(input + "\\" + list1);
						String extension = FilenameUtils.getExtension(soureceFile.toString());
						if (extension.equalsIgnoreCase("xlsx")) {
							if (soureceFile.exists()) {
								//File output = new File(outputPath + dtf.format(now));
								  File output = new File(outputPath);
								if (!output.exists()) {
									output.mkdirs();
								}
								String out = output + "\\" + list1;
									finalPopupValue = mainLogic(soureceFile.toString(), out,finalPopupValue);
							} else {
								finalPopupValue = false;
								logger.info("In Todays folder XLSX not exist.");
							}
						}
					}
				} else {
					finalPopupValue = false;
					logger.info("In Todays folder File not exist.");
				}
			} else {
				finalPopupValue = false;
				logger.info("Todays folder not exist in E Drive - GemfireCount");
			}
		} else {
			finalPopupValue = false;
			logger.info("Create a folder in E Driver name as GemfireCount");
		}
		if (finalPopupValue) {
			System.out.println("Verified the column values.....Success");
			logger.info("Xls File executed at -   " + dateTimeFormatter.format(LocalDateTime.now()));
		}
	}
	static boolean mainLogic(String filePath, String outputPath, Boolean finalPopupValue) {
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(new FileInputStream(filePath));
			Sheet sheet = wb.getSheetAt(0);
			Row row1 = sheet.getRow(0);
			CellStyle styleForStatus = wb.createCellStyle();
			Font fontForStatus = wb.createFont();
			fontForStatus.setColor(IndexedColors.BLACK.getIndex());
			Cell cellForStatus = row1.createCell(5);
			cellForStatus.setCellValue("Status");
			fontForStatus.setBold(true);
			styleForStatus.setFont(fontForStatus);
			cellForStatus.setCellStyle(styleForStatus);
			for (int j = 1; j < sheet.getLastRowNum() + 1; j++) {
				Row row = sheet.getRow(j);
				Cell cell = row.getCell(1);
				cell.setCellType(CellType.STRING);
				Cell cell1 = row.getCell(2);
				cell1.setCellType(CellType.STRING);
				Cell cell2 = row.getCell(3);
				cell2.setCellType(CellType.STRING);
				Cell cell3 = row.createCell(5);
				if (cell.getStringCellValue().equalsIgnoreCase(cell1.getStringCellValue())
						&& cell1.getStringCellValue().equalsIgnoreCase(cell2.getStringCellValue())
						&& cell.getStringCellValue().equalsIgnoreCase(cell2.getStringCellValue())) {
					cell3.setCellValue("True");
					CellStyle styleForMatching = wb.createCellStyle();
					Font fontForMatching = wb.createFont();
					fontForMatching.setColor(IndexedColors.GREEN.getIndex());
					styleForMatching.setFont(fontForMatching);
					cell3.setCellStyle(styleForMatching);
					fontForMatching.setBold(true);

				} else {
					CellStyle styleForMatching = wb.createCellStyle();
					cell3.setCellValue("False");
					Font fontMatching = wb.createFont();
					fontMatching.setColor(IndexedColors.RED.getIndex());
					fontMatching.setBold(true);
					styleForMatching.setFont(fontMatching);
					cell3.setCellStyle(styleForMatching);
					fontMatching.setBold(true);
				}
			}
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		try {
			OutputStream fileOut = new FileOutputStream(outputPath);
			wb.write(fileOut);
			fileOut.close();
			finalPopupValue = true;
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			finalPopupValue = false;
			logger.info("Excel Sheet Opened, Please close gemfirecount_result Excel file");
		} catch (IOException e) {
			e.printStackTrace();
		}
		return finalPopupValue;
	}

}

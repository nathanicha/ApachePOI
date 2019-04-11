package com.gec.readexcelfile;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Date;


public class App {
	private static final String FILE_NAME = "./data/sample_data.xls";

	public static void main(String[] args) throws FileNotFoundException, IOException, ParseException {

			Workbook workbook = null;
			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			if (FILE_NAME.endsWith("xlsx")) {
				workbook = new XSSFWorkbook(excelFile);
			} else if (FILE_NAME.endsWith("xls")) {
				workbook = new HSSFWorkbook(excelFile);
			}
			int numberOfSheets = workbook.getNumberOfSheets();

			for (int i = 0; i < numberOfSheets; i++) {
				Sheet aSheet = workbook.getSheetAt(i);
				System.out.println(aSheet.getSheetName());
				Sheet datatypeSheet = workbook.getSheetAt(i);
				Iterator<Row> iterator = datatypeSheet.iterator();

				while (iterator.hasNext()) {

					Row currentRow = iterator.next();
					Iterator<Cell> cellIterator = currentRow.iterator();

					while (cellIterator.hasNext()) {

						Cell currentCell = cellIterator.next();
				
						switch (currentCell.getCellTypeEnum()) {
							case BOOLEAN:
								System.out.print(currentCell.getBooleanCellValue() + " -- ");
								break;
							case STRING:
								String str = currentCell.getStringCellValue();
								
								if(isDate(str)) {
									System.out.print(convertStringToDate(str) + " -- ");
								} else {
									System.out.print(str + " -- ");
								}
								
								break;
							case NUMERIC:
								if(DateUtil.isCellDateFormatted(currentCell)) {
									System.out.print(convertDateFormat(currentCell.getDateCellValue()) + " -- ");
								}else {
									System.out.print(currentCell.getNumericCellValue() + " -- ");	
								}
								break;
							case FORMULA:
								System.out.print(currentCell.getCellFormula() + " -- ");
								break;
							default:
								System.out.print("");
						}
					}
					System.out.println();
				}
				System.out.println();
			}
		

	}
	
	public static boolean isDate(String date) {
		/* #Case Date is String
		check string format: dd/mm/yyyy
		allow leading zeros to be omitted */
		String regex = "^[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}$";
		
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(date);
		
		return matcher.matches();
	}
	
	public static String convertStringToDate(String strDateOld) throws ParseException {
		/* #Case Date is String
		Set string to date format that we wanted */
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");  
		//convert string to SimpleDateFormat (Thu Dec 31 00:00:00 IST 1998)
		Date date = formatter.parse(strDateOld); 
		//convert SimpleDateFormat to string date format that we want (dd/mm/yyyy)
		String strDateNew = formatter.format(date);
		
		return strDateNew;
	}
	
	public static String convertDateFormat(Date dateOld) {
		/* #Case Numeric is date
		 * we don't care about date format input. It will convert to date format (dd/MM/yyyy) */
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy"); 
		String dateNew = formatter.format(dateOld); 
		
		return dateNew;
	}
}

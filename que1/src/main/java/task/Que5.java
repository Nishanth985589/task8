package task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;


public class Que5 {
	


	public static void main(String[] args) throws IOException {

		File excelfile = new File("C:\\Users\\Nishanth\\test\\tour.xlsx");
		FileInputStream fil = new FileInputStream(excelfile);
		XSSFWorkbook workbook = new XSSFWorkbook(fil);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rowcount = sheet.getPhysicalNumberOfRows();
		
		for (int i =0; i< rowcount; i++) {
			XSSFRow row = sheet.getRow(i);
			
		int cellcount = row.getPhysicalNumberOfCells();
		for (int j=0; j<cellcount; j++) {
			XSSFCell cell = row.getCell(j);
		    String cellvalue = getCellvalue(cell);
		    System.out.print("||"+ cellvalue);
		
		}
		
		System.out.println();
		}
		workbook.close();
		fil.close();
	}
		
		public static String getCellvalue(XSSFCell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case STRING:
			return cell.getStringCellValue();
        default:
        	return cell.getStringCellValue();
        	
		}}	}
		
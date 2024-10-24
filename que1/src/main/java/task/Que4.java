package task;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Que4 {

	public static void main(String[] args) throws IOException {
		
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Plan");
		
		sheet.createRow(0);
		sheet.getRow(0).createCell(0).setCellValue("Kodaikanal");
		sheet.getRow(0).createCell(1).setCellValue("Vagamon");
		sheet.getRow(0).createCell(2).setCellValue("Valparai");
		
		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue("Nirmal");
		sheet.getRow(1).createCell(1).setCellValue("Ajith");
		sheet.getRow(1).createCell(2).setCellValue("Jothi");
		
		File file = new File("C:\\Users\\Nishanth\\eclipse-workspace\\que1\\Excelfile\\tour.xls");
		workbook.write(file);
		workbook.close();
	   
		

	}

}

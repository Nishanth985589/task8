package task;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Que123 {

	public static void main(String[] args) throws IOException {
		
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sheet 1");
	    sheet.createRow(0);
	    sheet.getRow(0).createCell(0).setCellValue("Name");
	    sheet.getRow(0).createCell(1).setCellValue("Age");
	    sheet.getRow(0).createCell(2).setCellValue("Email");
	    
	    sheet.createRow(1);
	    sheet.getRow(1).createCell(0).setCellValue("John Doe");
	    sheet.getRow(1).createCell(1).setCellValue("30");
	    sheet.getRow(1).createCell(2).setCellValue("john@test");
	    
	    sheet.createRow(2);
	    sheet.getRow(2).createCell(0).setCellValue("Jane Doe");
	    sheet.getRow(2).createCell(1).setCellValue("28");
	    sheet.getRow(2).createCell(2).setCellValue("jane@test");
	    
	    sheet.createRow(3);
	    sheet.getRow(3).createCell(0).setCellValue("Bob Smith");
	    sheet.getRow(3).createCell(1).setCellValue("35");
	    sheet.getRow(3).createCell(2).setCellValue("jacky@example.com");
	    
	    sheet.createRow(4);
	    sheet.getRow(4).createCell(0).setCellValue("Swapnil");
	    sheet.getRow(4).createCell(1).setCellValue("37");
	    sheet.getRow(4).createCell(2).setCellValue("swapnil@example.com");
	    
	    File file = new File("C:\\Users\\Nishanth\\eclipse-workspace\\que1\\Excelfile\\test.xls");
	    workbook.write(file);    
	    workbook.close();
	
	}
	

}

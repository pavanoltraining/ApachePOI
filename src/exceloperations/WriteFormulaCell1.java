package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell1 {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Numbers");
		
		XSSFRow row=sheet.createRow(0);
			
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		
		row.createCell(3).setCellFormula("A1*B1*C1");
		
		
		FileOutputStream fos=new FileOutputStream(".\\datafiles\\calc.xlsx");
		
		workbook.write(fos);
		fos.close();
		
		System.out.println("calc.xlsx created with formula cell...");

	}

}

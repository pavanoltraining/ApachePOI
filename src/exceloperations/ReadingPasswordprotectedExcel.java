package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingPasswordprotectedExcel {

	public static void main(String[] args) throws IOException {
	
		FileInputStream fis=new FileInputStream(".\\datafiles\\customers.xlsx");
		String password="test123";
		
		//XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFWorkbook workbook=(XSSFWorkbook)WorkbookFactory.create(fis,password);
		XSSFSheet sheet=workbook.getSheetAt(0);
		 
		//read data from sheet using for loop
		/*int rows=sheet.getLastRowNum();
		System.out.println(rows);   //5  started from 0
		int cols=sheet.getRow(0).getLastCellNum();
		System.out.println(cols);  //3   started from 1
		
		for(int r=0;r<=rows;r++)
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell=row.getCell(c);
				switch(cell.getCellType())
				{
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case BOOLEAN: System.out.print(cell.getNumericCellValue());break;
				case FORMULA: System.out.print(cell.getNumericCellValue());break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}*/
		
		
		//read data from sheet using iterator
		Iterator<Row>  iterator=sheet.iterator();
		
		while(iterator.hasNext())
			{
					
			Row nextrow=iterator.next();
			
			Iterator<Cell> celliterator=nextrow.cellIterator();
			
				while(celliterator.hasNext())
				{
					Cell cell=celliterator.next();
					
					switch(cell.getCellType())
					{
					case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
					case STRING: System.out.print(cell.getStringCellValue()); break;
					case BOOLEAN: System.out.print(cell.getNumericCellValue());break;
					case FORMULA: System.out.print(cell.getNumericCellValue());break;
					}
					System.out.print(" | ");
				}
				System.out.println();
			
			}
				
		
		workbook.close();
		fis.close();
	}

}








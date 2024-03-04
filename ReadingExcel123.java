package task17;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;
public class ReadingExcel123 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		String excelFilePath =".\\ExcelReading\\ExcelRead.xlsx" ;
		//String excelFilePath="C:\\Users\\lenovo\\eclipse-workspace\\SeleniumTraining\\ExcelReading\\ExcelRead.xlsx";
		
		FileInputStream inputstream =new FileInputStream(excelFilePath);
        
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    
	    
	    //using For loop
	    
	    int rows = sheet.getLastRowNum();
	    int column = sheet.getRow(1).getLastCellNum();
	    
	    for(int r=0;r<=rows;r++)        //represent row
	    {
	    	XSSFRow row = sheet.getRow(r);
	    	
	    	for(int c=0;c<column;c++)                           //represent column
	    	{
	    		XSSFCell cell= row.getCell(c);
	    		
	    		switch(cell.getCellType())
	    		{
	    		case STRING: System.out.print(cell.getStringCellValue());
	    		break;
	    		case NUMERIC:System.out.print(cell.getNumericCellValue());
	    		break;
	    		case BOOLEAN:System.out.print(cell.getBooleanCellValue());
	    		break;
	    		
	    		}
	    		
	    		    System.out.print(" | ");
	    	}
	    	    System.out.println();
	    }
	    
		
	}

}




OUTPUT:
      
Name | Age | Email | 
John Doe | 30.0 | john@test.com | 
Jane Doe  | 28.0 | jane@test.com | 
Bob Smith | 35.0 | bob@example.com | 
Swapnil | 37.0 | swapnil@example.com | 





   



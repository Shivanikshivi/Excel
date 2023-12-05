package com.ApachePOI;
import java.io.File; 
import java.io.FileInputStream; 
import java.io.IOException; 
import org.apache.poi.hssf.usermodel.HSSFSheet; 
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.FormulaEvaluator; 
import org.apache.poi.ss.usermodel.Row; 
public class ReadExcel {

	public static void main(String[] args) {
		
		try { 
			  
            
            FileInputStream file = new FileInputStream( 
                new File("C:\\Users\\user\\eclipse-workspace\\ExcelFiles.xlsx")); 
  
          
            XSSFWorkbook workbook = new XSSFWorkbook(file); 
  
           
            XSSFSheet sheet = workbook.getSheetAt(0); 
  
            
            Iterator<Row> rowIterator = sheet.iterator(); 
  
            
            while (rowIterator.hasNext()) { 
  
                Row row = rowIterator.next(); 
  
               
                Iterator<Cell> cellIterator 
                    = row.cellIterator(); 
  
                while (cellIterator.hasNext()) { 
  
                    Cell cell = cellIterator.next(); 
  
                    switch (cell.getCellType()) { 
  
                    
                    case Cell.CELL_TYPE_NUMERIC: 
                        System.out.print( 
                            cell.getNumericCellValue() 
                            + "t"); 
                        break; 
  
                    case Cell.CELL_TYPE_STRING: 
                        System.out.print( 
                            cell.getStringCellValue() 
                            + "t"); 
                        break; 
                    } 
                } 
  
                System.out.println(""); 
            } 
  
            file.close(); 
        } 
  
    
        catch (Exception e) { 
        	e.printStackTrace(); 
        } 
    } 

	}



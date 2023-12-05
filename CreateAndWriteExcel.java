package com.ApachePOI;
import java.io.File; 
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.xssf.usermodel.XSSFRow; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
import java.io.FileOutputStream; 
import java.util.Map; 
import java.util.Set; 
import java.util.TreeMap; 
public class CreateAndWriteExcel {

	
	public static void main(String[] args) throws Exception 
    { 
        
        XSSFWorkbook workbook = new XSSFWorkbook(); 
  
        
        XSSFSheet spreadsheet 
            = workbook.createSheet(" Sheet1 "); 
  
        
        XSSFRow row; 
  
        Map<String, Object[]> studentData = new TreeMap<String, Object[]>(); 
  
        studentData.put(  "1", new Object[] { "Name", "Age", "Email" }); 
  
        studentData.put("2", new Object[] { "John Doe", "30", 
                                            "john@test.com" }); 
  
        studentData.put( "3", new Object[] { "Jane Doe", "28", "jane@test.com" }); 
  
        studentData.put("4", new Object[] { "Bob Smith", "25", 
                                            "jacky@example.com" }); 
  
        studentData.put("5", new Object[] { "Swapnil", "37", 
                                            "joe@example.com" });  
        Set<String> keyid = studentData.keySet(); 
  
        int rowid = 0;  
        for (String key : keyid) { 
  
            row = spreadsheet.createRow(rowid++); 
            Object[] objectArr = studentData.get(key); 
            int cellid = 0; 
  
            for (Object obj : objectArr) { 
                Cell cell = row.createCell(cellid++); 
                cell.setCellValue((String)obj); 
            } 
        } 
  
       
        FileOutputStream out = new FileOutputStream( 
            new File("C:\\Users\\user\\eclipse-workspace\\ExcelFiles.xlsx")); 
  
        workbook.write(out); 
        out.close(); 
    } 
}

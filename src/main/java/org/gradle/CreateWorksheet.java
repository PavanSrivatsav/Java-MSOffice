package org.gradle;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Date;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorksheet {
	public static void main(String[] args)throws Exception {
		
//		String sheetName;
//		String fileName;
//		
//		CreateWorksheet(){
//			
//		}
	//Create Blank workbook
    XSSFWorkbook workbook = new XSSFWorkbook(); 
    
//    XSSFFont font = workbook.createFont();
//    font.setFontHeightInPoints((short) 30);
//    font.setFontName("IMPACT");
//    font.setItalic(true);
//    font.setColor(HSSFColor.HSSFColorPredefined.BLUE);
//    
//    //Set font into style
//    XSSFCellStyle style = workbook.createCellStyle();
//    style.setFont(font);
    
  //Create a blank spreadsheet
    XSSFSheet spreadsheet = workbook.createSheet("Sample Sheet 1");
    
    XSSFRow row = spreadsheet.createRow((short) 2);
    row.createCell(0).setCellValue("Type of Cell");
    row.createCell(1).setCellValue("cell value");
    
    row = spreadsheet.createRow((short) 3);
    row.createCell(0).setCellValue("set cell type BLANK");
    row.createCell(1);
    
    row = spreadsheet.createRow((short) 4);
    row.createCell(0).setCellValue("set cell type BOOLEAN");
    row.createCell(1).setCellValue(true);
    
    row = spreadsheet.createRow((short) 6);
    row.createCell(0).setCellValue("set cell type date");
    row.createCell(1).setCellValue(new Date(0, 0, 0));
    
    row = spreadsheet.createRow((short) 7);
    row.createCell(0).setCellValue("set cell type numeric");
    row.createCell(1).setCellValue(20 );
    
    row = spreadsheet.createRow((short) 8);
    row.createCell(0).setCellValue("set cell type string");
    row.createCell(1).setCellValue("A String");
    
    XSSFSheet spreadsheet1 = workbook.createSheet("Sample Sheet 2");;
    
    XSSFRow row1 = spreadsheet1.createRow((short) 2);
    row1.createCell(0).setCellValue("Type of Cell");
    row1.createCell(1).setCellValue("cell value");
    
    

    //Create file system using specific name
    FileOutputStream out = new FileOutputStream(new File("Sample File.xlsx"));

    //write operation workbook using file out object 
    workbook.write(out);
    out.close();
    System.out.println("createworkbook.xlsx written successfully");

}
}

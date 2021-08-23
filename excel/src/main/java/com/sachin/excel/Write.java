package com.sachin.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write {

	public static void main(String[] args) throws IOException {
		  XSSFWorkbook workbook = new XSSFWorkbook();
	      
	      
	      XSSFSheet spreadsheet = workbook.createSheet( " Score Sheet ");

	 
	      XSSFRow row;

	      
	      Map < String, Object[] > empinfo = new TreeMap < String, Object[] >();
	    //  empinfo.put( "6", new Object[] { "Id", "Name", "Runs","Balls","Boundaries" });
	      
	      empinfo.put( "6", new Object[] {  "6","RohitSharma" });
	      
	      empinfo.put( "7", new Object[] {  "7","ViratKohli"  });
	 
	      Set < String > keyid = empinfo.keySet();
	      int rowid = 0;
	      
	      for (String key : keyid) {
	         row = spreadsheet.createRow(rowid++);
	         Object [] objectArr = empinfo.get(key);
	         int cellid = 0;
	         
	         for (Object obj : objectArr){
	            Cell cell = row.createCell(cellid++);
	            cell.setCellValue((String)obj);
	         }
	      }
	     
	      FileOutputStream out = new FileOutputStream(
	         new File("C:\\Users\\Hp\\Desktop\\javA\\java\\stud.xls"));
	      
	      workbook.write(out);
	      out.close();
	      System.out.println("Writesheet.xlsx written successfully");

	}

}

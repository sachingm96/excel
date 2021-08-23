package com.sachin;

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

		XSSFSheet spreadsheet = workbook.createSheet(" Score Sheet ");

		XSSFRow row;

		Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
		empinfo.put("1", new Object[] { "Id", "Name", "mobileNO" });

		empinfo.put("2", new Object[] { "1", "Sanjay", "9090909090" });

		empinfo.put("3", new Object[] { "2", "Suresh", "9988878878" });

		empinfo.put("4", new Object[] { "3", "Ramesh", "8989099908" });

		empinfo.put("5", new Object[] { "4", "Sangesh", "909088689" });

		empinfo.put("6", new Object[] { "5", "Sunil", "8989890889" });

		Set<String> keyid = empinfo.keySet();
		int rowid = 0;

		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = empinfo.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}

		FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Hp\\Desktop\\javA\\java\\newdata.xlsx"));

		workbook.write(out);
		out.close();
		System.out.println("successfull");
	}

}

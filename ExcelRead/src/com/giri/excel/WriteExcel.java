package com.giri.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("palace sheet");

		// This data needs to be written (Object[])
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		// data.put("7", new Object[] {"ID", "PALACE",
		// "OWNED","year","State","City","Country"});
		data.put("7", new Object[] { 7, "Mysore Palce", "Vadeyar", "1997", "KARNATAK", "MYSORE", "INDIA" });
		data.put("8", new Object[] { 8, "Palce", "Vadeyar", "1997", "KARNATAK", "MYSORE", "INDIA" });

		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 7;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 7;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}
		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\teju\\Documents\\palaceSheet.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("palace sheet written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}

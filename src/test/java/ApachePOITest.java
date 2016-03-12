
// package com.sqa.em;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.testng.annotations.Test;

public class ApachePOITest {

	@Test(enabled = true)
	public void f() {
		try {
			// Get File based on class loader
			ClassLoader cl = this.getClass().getClassLoader();
			// Get InputStream via Class Loader
			InputStream fileA = cl.getResourceAsStream("poi-example.xls");
			// Get InputStream by creating a new FileInputStream and passing a
			// path of a resource
			InputStream fileB = new FileInputStream(cl.getResource("poi-example.xls").getPath());
			// Get the file using basic File and relative path to directory
			InputStream fileC = new FileInputStream(new File("src/main/resources/poi-example.xls"));
			// Get the workbook instance for XLS file
			HSSFWorkbook workbook = new HSSFWorkbook(fileA);
			// Get first sheet from the workbook
			HSSFSheet sheet = workbook.getSheetAt(0);
			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						// System.out.println("Calling a boolean value!!!!");
						System.out.print(cell.getBooleanCellValue() + "\t\t\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					}
				}
				System.out.println("");
			}
			fileA.close();
			FileOutputStream out = new FileOutputStream(new File("Output-File.xls"));
			workbook.write(out);
			out.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

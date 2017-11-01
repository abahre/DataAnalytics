package main.java.com.excelManipulate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.commons.codec.language.bm.Rule.RPattern;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

//import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Manipulate_Excel_Columns {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		//FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_section_heating_energy.xlsx"));
		//FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\Updated_input_section_heating_energy_110.xlsx"));
		
		// FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_air_temperature.xlsx"));
		 //FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\Updated_input_wall_temperature_110.xlsx"));
		 FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\Updated_output_section_temperature_110.xlsx"));
		// FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\test.xlsx"));
		
		 // Get the workbook instance for XLS file
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		// Get first sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Get iterator to all the rows in current sheet
		Iterator<Row> rowIterator = sheet.iterator();

		// Create New Sheet
		XSSFWorkbook newWorkBook = new XSSFWorkbook();
		
		XSSFSheet newSheet = newWorkBook.createSheet("Updated_output_section_temperature_110_4_Col");
		int rownum = 0;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			ArrayList<Double> myArray = new ArrayList<Double>();
			int col = 0;
			Iterator<Cell> cellIterator = row.cellIterator();
			Row newRow = newSheet.createRow(rownum++);
			int cellCount = 0;
			Cell timeStamp = newRow.createCell(cellCount);

			Cell timeStampCol = cellIterator.next();
			timeStamp.setCellValue(timeStampCol.getNumericCellValue());

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				// System.out.println(cellCount);
				myArray.add(cell.getNumericCellValue());
				col++;
				if (col == 4) { //Change this value to 4 if want to merge 4 columns
					cellCount++; 
					Cell newCell = newRow.createCell(cellCount);
					newCell.setCellValue((myArray.get(0) + myArray.get(1) + myArray.get(2) + myArray.get(3))/ col); //// + myArray.get(2) + myArray.get(3)
					myArray.clear();
					col = 0;
				}
			}

		}
		try {
			FileOutputStream out = new FileOutputStream(
					new File(
							"C:\\ovgu\\proj data\\Updated_output_section_temperature_4_110.xlsx"));
			newWorkBook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
			System.out.println("Final set of rows : " + rownum);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
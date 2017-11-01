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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Expand_Material_Temp {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

FileInputStream material_temp_file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_material_temperature.xlsx"));
		
		//FileInputStream material_temp_file = new FileInputStream(new File("C:\\ovgu\\proj data\\Book1.xlsx"));
		XSSFWorkbook mat_workbook = new XSSFWorkbook(material_temp_file);

		// Get first sheet from the workbook
		XSSFSheet mat_sheet = mat_workbook.getSheetAt(0);
		Iterator<Row> mat_rowIterator = mat_sheet.iterator();
		XSSFWorkbook newWorkBook = new XSSFWorkbook();
			
			XSSFSheet newSheet = newWorkBook.createSheet("output_section_temperature_110");

			int mat_rownum = 0;
			int oldVal = 0;
			int cellNum=0;
			ArrayList<Double> mat_Array = new ArrayList<Double>();
			while (mat_rowIterator.hasNext()) {
				Row row = mat_rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				
				Cell timeStampCol = cellIterator.next();

				
				for(int i=oldVal; i<(timeStampCol.getNumericCellValue());i++){
					Iterator<Cell> cellIterNew = row.cellIterator();
					Row newRow = newSheet.createRow(oldVal++);
					while(cellIterNew.hasNext()){
						Cell cellToCopy = cellIterNew.next();
						Cell newCell = newRow.createCell(cellNum++);  
						newCell.setCellValue(cellToCopy.getNumericCellValue());

							}
					cellNum=0;				}
					
					


			}
					try {
			FileOutputStream out = new FileOutputStream(
					new File(
							"C:\\ovgu\\proj data\\temp.xlsx"));
			newWorkBook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
			System.out.println("Final set of rows : " + oldVal);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
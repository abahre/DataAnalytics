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

public class Excel_Compress_110 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		FileInputStream material_temp_file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_material_temperature.xlsx"));
		
		//FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_section_heating_energy.xlsx"));
		// FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_air_temperature.xlsx"));
		// FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_wall_temperature.xlsx"));
		 FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\output_section_temperature.xlsx"));
		// FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\test.xlsx"));
		
	 // Get the workbook instance for XLS file
		XSSFWorkbook mat_workbook = new XSSFWorkbook(material_temp_file);

		// Get first sheet from the workbook
		XSSFSheet mat_sheet = mat_workbook.getSheetAt(0);

		// Get iterator to all the rows in current sheet
		Iterator<Row> mat_rowIterator = mat_sheet.iterator();

		// Create New Sheet
		XSSFWorkbook newWorkBook = new XSSFWorkbook();
		
		XSSFSheet newSheet = newWorkBook.createSheet("output_section_temperature_110");
		int mat_rownum = 0;
		ArrayList<Double> mat_Array = new ArrayList<Double>();
		while (mat_rowIterator.hasNext()) {
			Row row = mat_rowIterator.next();

			
			int col = 0;
			Iterator<Cell> cellIterator = row.cellIterator();
			Row newRow = newSheet.createRow(mat_rownum++);
			int cellCount = 0;
			Cell timeStamp = newRow.createCell(cellCount);

			Cell timeStampCol = cellIterator.next();
			timeStamp.setCellValue(timeStampCol.getNumericCellValue());

				// System.out.println(cellCount);
				mat_Array.add(timeStampCol.getNumericCellValue());
//			}

		}
		 // Get the workbook instance for XLS file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheet.iterator();
			
		int rowCount = 0;
		int rownum =0;
		int cellnum =0;
		int cellSize = 0;
		int num =0;
		Double finalValue=0.0;
		int row_num_mat_input =0;
		int row_to_add=1;
		int mat_arr_size =mat_Array.size();
HashMap<Integer, ArrayList<Double>> myHashmap = new HashMap<Integer, ArrayList<Double>>(); 
		
		while (rowIterator.hasNext() && (row_num_mat_input <mat_arr_size)) {
		      Row row = rowIterator.next();
		      ArrayList<Double> myArray = new ArrayList<Double>();
		     
		      Iterator <Cell> cellIterator = row.cellIterator();
		      cellIterator.next();
		      while (cellIterator.hasNext()) {
		        Cell cell = cellIterator.next();
		      
		        myArray.add(cell.getNumericCellValue());
		       
		      }
		      cellSize = myArray.size();
		      rowCount++;
		      myHashmap.put(rowCount, myArray);
		     // System.out.println(myHashmap);
num++;

		      if(num == Math.ceil(mat_Array.get(row_num_mat_input))){
		    	  row_num_mat_input++;
		    	Row newRow = newSheet.createRow(rownum++);
		    	Cell timeStamp = newRow.createCell(cellnum++);
		    	timeStamp.setCellValue(Math.ceil(mat_Array.get(row_num_mat_input-1)));
		    	
		    	  for(int i=0;i<cellSize;i++){
		    		  Cell newCell = newRow.createCell(cellnum++);  
		    		   for(int j=row_to_add;j<num+1;j++){
		    		  ArrayList<Double> x = myHashmap.get(j);
		    
		    			finalValue=finalValue + x.get(i);
		    		   }
		    		   int total_rows = num-row_to_add+1;
		    		   finalValue =finalValue/(total_rows);
		    		   newCell.setCellValue(finalValue);
		    		   finalValue=0.0;
		    	  }
		
		    	  row_to_add =num+1;
		    	  cellnum=0;
		      }
		      
		}
		try {
			FileOutputStream out = new FileOutputStream(
					new File(
							"C:\\ovgu\\proj data\\Updated_output_section_temperature_110.xlsx"));
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
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

public class Manipulate_Excel_Rows {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_section_heating_energy.xlsx"));
		//FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_air_temperature.xlsx"));
	//	FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\input_material_temperature.xlsx"));
		//FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\output_section_temperature.xlsx"));
	//	FileInputStream file = new FileInputStream(new File("C:\\ovgu\\proj data\\test.xlsx"));
		//Get the workbook instance for XLS file 
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		//Get first sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);

		//Get iterator to all the rows in current sheet
		Iterator<Row> rowIterator = sheet.iterator();
	
		//Create New Sheet
		XSSFWorkbook newWorkBook = new XSSFWorkbook();
		//XSSFSheet newSheet = newWorkBook.createSheet("Updated_Heating_Energy");
		XSSFSheet newSheet = newWorkBook.createSheet("Updated_Heating_Energy");
		int rowCount = 0;
		int rownum =0;
		int cellnum =0;
		int cellSize = 0;
		int num =0;
		Double finalValue=0.0;
		//To store the values of each row from the excel
		
		HashMap<Integer, ArrayList<Double>> myHashmap = new HashMap<Integer, ArrayList<Double>>(); 
		
		while (rowIterator.hasNext()) {
		      Row row = rowIterator.next();
		      ArrayList<Double> myArray = new ArrayList<Double>();
		     
		      Iterator <Cell> cellIterator = row.cellIterator();
		      while (cellIterator.hasNext()) {
		        Cell cell = cellIterator.next();
		      
		        myArray.add(cell.getNumericCellValue());
		       
		      }
		      cellSize = myArray.size();
		      rowCount++;
		      myHashmap.put(rowCount, myArray);
		     // System.out.println(myHashmap);
num++;
		      if(num ==2){
		    	Row newRow = newSheet.createRow(rownum++);
		    	
		    	  for(int i=0;i<cellSize;i++){
		    		  Cell newCell = newRow.createCell(cellnum++);  
		    		   for(int j=1;j<num+1;j++){
		    		  ArrayList<Double> x = myHashmap.get(j);
		    
		    			finalValue=finalValue + x.get(i);
		    		   }
		    		   newCell.setCellValue(finalValue/num);
		    		   finalValue=0.0;
		    	  }
		    	  myHashmap.clear(); 
		    	  rowCount =0;
		    	  num=0;
		    	  cellnum=0;
		      }
		      
		}
		try {
			FileOutputStream out = 
					new FileOutputStream(new File("C:\\ovgu\\proj data\\updated_output_section_temperature_2.xlsx"));
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
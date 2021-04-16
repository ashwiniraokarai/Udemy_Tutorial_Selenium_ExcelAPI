package packageDemoApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class C02_Read_Different_DATA_TYPES {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/* Ensure you have your EXCEL spreadsheet and populated with TEST DATA
		 * Choosing XSSF API over HSSF API because the MS OFFICE/EXCEL version/ format is XLSX
		 * Hierarchy: Workbook->Worksheet("sheet")->Row->Column("Cell")
		 * RoadMap:
		 * 	-Begin with a Workbook-Object (constructor that takes file input stream object or file as input)
		 * 	-Get the Sheet Object from the Workbook Object
		 * 	-Row (Run the row-iterator on the Sheet Object)
		 * 	-Column (Referred to as the "Cell" on this API. Run the cell-iterator on the Row-Object)
		 * On this demo, we wont be looping yet but will focus on handling READING INPUTS of various DATA TYPES
		 */
		
		FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
		
		//Reaching a SPECIFIC CELL
		XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = exclWrkBk.getSheet("sheet1"); // or exclWrkBk.getSheetAt(0);
		
		/*Trying to ACCESS ROW 4. It has non-String values on both cells.
		 * COLUMN1 has a NUMERIC TYPE VALUE
		 * COLUMN2 has a DATE TYPE VALUE
		 */
		
		//READING a CELL VALUE
		/*You SHOULD be looping through the "cells" one by one on each row (which itself should be an outer a loop) and grab the values in a similar pattern
		But, at the moment we are FOCUSING on HANDLING READ requests on CELLS that may BELONG to ONE of the MANY DATA TYPES
				E.g.: Accessing the getRow(4).getCell(0) would require you to use .getBooleanCellValue or you end up with exception:
				   "Cannot get a STRING value from a BOOLEAN cell"
		*/
		
		//System.out.println(sheet.getRow(3).getCell(0).getCellTypeEnum()); //Temporary check point: Returns "NUMERIC"
		
		/*Logical structure:
		 * If sheet.getRow(3).getCell(0).getCellTypeEnum() = "STRING" then
		 * Else If = "NUMERIC" then
		 * Else If = "BOOLEAN" then
		 * ELse
		 *  also consider these Else IF's 
		 *  BLANK
		 *  FORMULA
		 *  ERROR 
		 */
		
		int rowIndex = 4;
		int columnIndex = 0;
		
		
		CellType cellType = sheet.getRow(rowIndex).getCell(columnIndex).getCellTypeEnum();
		System.out.println("Cell Type is: "+cellType);
		
		if (cellType.equals(CellType.NUMERIC)) {
			//Call the method to read numeric data
			double result = sheet.getRow(rowIndex).getCell(columnIndex).getNumericCellValue();
			System.out.println("Numeric value read:"+result); //Temporary check point
			//Even though the number "1" appears to be an integer on EXCEL, getNumericCellValue ALWAYS returns the value as a double type.
			
		} /*else if (cellType.equals(CellType.BOOLEAN)) {
			boolean result = sheet.getRow(rowIndex).getCell(columnIndex).getBooleanCellValue();
			System.out.println("Boolean value read:"+result); //Temporary check point 
			
		}*/ else if (cellType.equals(CellType.STRING)) {
			String result = sheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue();	
			System.out.println("String value read:"+result); //Temporary check point 
			
		}else if(cellType.equals(CellType.BLANK)) {
			String result = sheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue();	//getStringCellValue returns an empty string for blank cells
			System.out.println("Blank value read"+result); //Temporary check point
			
		}else if(cellType.equals(CellType.FORMULA)) {
				
				CellType formulaType = sheet.getRow(rowIndex).getCell(columnIndex).getCachedFormulaResultTypeEnum();
				System.out.println(formulaType);//Temporary check point
				
				if (formulaType==CellType.STRING) {
					String result = sheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue();
					System.out.println("String value read:"+result); //Temporary check point
				}
				//Include else if code sections to handle each formula type here
				
				else {
					
					System.out.println("New formula type to be handled: "+cellType);
				}
				
		}
		//Include else if code sections to handle each type here
		else {
			System.out.println("New type to be handled: "+cellType);
			String result = sheet.getRow(rowIndex).getCell(columnIndex).toString();
			/*YOU CAN SIMPLY REPLACE the ENTIRE IF-ELSE LOOP with the .toString() METHOD and NOT WORRY about CHECKING the DATA-TYPE and MATCHING
			 * with the CORRECT READ METHOD to call
			 */
			System.out.println("String converted value is: "+result);
		}
			
			
	}

}

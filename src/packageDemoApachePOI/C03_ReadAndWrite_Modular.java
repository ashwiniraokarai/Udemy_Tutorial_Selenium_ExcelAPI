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

public class C03_ReadAndWrite_Modular {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/*Enclosing the READ and WRITE FUNCTIONALITY in REUSABLE getter and setter-like METHODS
		 * RoadMap:
		 * Mark the code that could be potential "METHOD" candidates and wrap them in methods
		 * -fnGetCellValue
		 * -fnSetCellValue (should return a value, eventually)
		 * Set the ACCESS MODIFIER, RETURN TYPE, etc
		 * Call the METHODS to TEST if the CALL itself is working fine (BASIC TEST) BEFORE MAKING ANY MASSIVE ALTERATIONS
		 * Identify and declare VARIABLES that need to be OUTSIDE the METHODS in order to facilitate METHOD CALLS
		 * Add ARGUMENTS to the METHODS as NECESSARY to suit the METHOD CALL NEEDS
		 * Receive RETURNED VALUES (if applicable)
		 */
		
		//Method calls
		C03_ReadAndWrite_Modular rdWrt = new C03_ReadAndWrite_Modular();
		//rdWrt.fnGetCellValue(4,0);
		//rdWrt.fnGetCellValue(6,0);
		//rdWrt.fnGetCellValue(4,2);
		rdWrt.fnSetCellValue(5,0);
		rdWrt.fnGetCellValue(5,0);
		
	}
		public void fnGetCellValue(int rowIndex,int columnIndex) throws IOException{
			//Code to READ data from EXCEL
			FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
			
			//Reaching a SPECIFIC CELL
			XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream);
			XSSFSheet sheet = exclWrkBk.getSheet("sheet1"); // or exclWrkBk.getSheetAt(0);
			
			//Ensuring the ARGUMENTS received are within range
			int lastRowNum = sheet.getLastRowNum();
			
			if (rowIndex>lastRowNum || rowIndex<0) {
				exclWrkBk.close();	//Close the workbook
				System.out.println("Received row index out of range("+rowIndex+"). Exiting method fnGetCellValue");
				return; //exit the method
			}else {
				short lastCellNum = sheet.getRow(lastRowNum).getLastCellNum();
				lastCellNum-=1;
				if(columnIndex>lastCellNum || lastCellNum<0 ) { 
					System.out.println("Received column index out of range("+columnIndex+"). Exiting method fnGetCellValue");
					exclWrkBk.close(); //Close the workbook
					return;
				}
			}
			
			CellType cellType = sheet.getRow(rowIndex).getCell(columnIndex).getCellTypeEnum();
			System.out.println("Cell Type is: "+cellType);
			exclWrkBk.close(); //Close the workbook
			
			if (cellType.equals(CellType.NUMERIC)) {
				//Call the method to read numeric data
				double result = sheet.getRow(rowIndex).getCell(columnIndex).getNumericCellValue();
				System.out.println("Numeric value read:"+result); //Temporary check point
				//Even though the number "1" appears to be an integer on EXCEL, getNumericCellValue ALWAYS returns the value as a double type.
				
			} else if (cellType.equals(CellType.BOOLEAN)) {
				boolean result = sheet.getRow(rowIndex).getCell(columnIndex).getBooleanCellValue();
				System.out.println("Boolean value read:"+result); //Temporary check point 
				
			} else if (cellType.equals(CellType.STRING)) {
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
		
		public void fnSetCellValue(int rowIndex,int columnIndex) throws IOException{
			//Code to WRITE data to EXCEL
			
			FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
			/*You COULD choose to pass the file path as an argument instead. It can be reused on the OutputStream.
			 * Also including a similar argument on fnGetCellValue would be a good idea
			 */
			
			XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream); 
			/*Keep in mind to instantiate with the constructor that takes filestream input. Eclipse will not provide "suggestions" because 
			constructor XSSFWorkbook() with no arguments is valid too!*/
			
			XSSFSheet sheet = exclWrkBk.getSheet("Sheet1");
			//You COULD choose to pass the sheet as an argument instead and iterate through the workbook using the sheetIterator

			//OVER-WRITING VALUE on a SPECIFIC CELL (non-blank cell) on a ROW IN “RANGE” (existing/ occupied cell)
			
			//Verify the ROW and CELL you are writing to are in the occupied row range and cell range
			int lastRowNum = sheet.getLastRowNum();
			
			if (rowIndex>lastRowNum || rowIndex<0) {
				exclWrkBk.close(); //Close the workbook
				System.out.println("Received row index out of range("+rowIndex+"). Exiting method fnSetCellValue");
				return; //exit the method
			}else {
				short lastCellNum = sheet.getRow(lastRowNum).getLastCellNum();
				lastCellNum-=1;
				if(columnIndex>lastCellNum || lastCellNum<0 ) { 
					System.out.println("Received column index out of range("+columnIndex+"). Exiting method fnSetCellValue");
					exclWrkBk.close(); //Close the workbook
					return;
				}
			}
			
			sheet.getRow(rowIndex).getCell(columnIndex).setCellValue("02/08/2018");
			
			FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
			exclWrkBk.write(fileOutputStream);
			exclWrkBk.close(); //Close the workbook
			
		}
}	
		
		
				




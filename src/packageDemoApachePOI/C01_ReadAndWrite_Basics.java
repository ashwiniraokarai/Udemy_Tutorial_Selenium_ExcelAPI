package packageDemoApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class C01_ReadAndWrite_Basics {

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
		 * This DEMO covers the basic read and write with SPECIFIC INDEX VALUES
		 * The actual iteration is built as an IMPROVISATION of this demo in the following exercise
		 */
		
		/* This constructor BELOW works great when creating a new EXCEL workbook but not when you want to read an existing EXCEL workbook (file)
		XSSFWorkbook exclWrkBk = new XSSFWorkbook();
		Let's choose a more suitable one that take a File as input
		*/
		
		/*Create an input stream for your EXCEL file and pass it to the XSSFWB constructor. 
		 * Note that input stream is not recommended by Apache due to the higher
		 * memory consumption
		 */
		
		FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
		
		//Reaching a SPECIFIC CELL
		XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = exclWrkBk.getSheet("sheet1"); // or exclWrkBk.getSheetAt(0);
		
		XSSFRow row = sheet.getRow(0); //Returns XSSFRow representing the rownumber or null if its not defined on the sheet
		XSSFCell cell = row.getCell(0); //Returns XSSFCell at the given (0 based) index
		
		//READING a CELL VALUE
		System.out.println(cell.getStringCellValue()); //Returns the FIRST VALUE under EXCEL's inbuilt COLUMN A
		/*You can loop through the "cells" one by one on each row and grab the values in a similar pattern
		However, getStringCellValue MAY NOT work at all Time. Choose your METHOD WISELY by doing a getCellTypeEnum first
		E.g.: Accessing the getRow(4).getCell(0) would require your to use .getBooleanCellValue or you end up with exception:
		   "Cannot get a STRING value from a BOOLEAN cell"
		*/
		
		//READING a CELL VALUE from the LAST ROW 
		int lastRowNum = sheet.getLastRowNum(); //0 based result. So, it outputs the ROW INDEX. Good.
		System.out.println(lastRowNum);
		System.out.println(sheet.getRow(lastRowNum).getCell(0).getBooleanCellValue());
		
		//WRITING VALUE to a SPECIFIC CELL on the row AFTER THE CURRENT LAST ROW
		lastRowNum+=1;
		
		/*The row you want to write to is NOT WITHIN the ACCESS RANGE (anything outside the EXISTING POPULATED CELLS i.e, BLANKS ones are considered out of RANGE). 
		 * You will NEED to CREATE such ENTITIES (Sheet/Row/Cell) as necessary BEFORE SETTING a VALUE)
		 */
		sheet.createRow(lastRowNum).createCell(0);
		
		sheet.getRow(lastRowNum).getCell(0).setCellValue("test");
		
		/*When you want to WRITE to EXCEL,you need to call the write() method on your WORKBOOK OBJECT. This guy .write() takes a file OutputStream
		 * as an argument.So, besides opening up the FILE INPUT STREAM you also NEED to open up the FILE OUTPUT STREAM!
		 */
		FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
		//exclWrkBk.write(fileOutputStream);
		//exclWrkBk.close();
		
		/*Commenting the 
		 * exclWrkBk.close()
		 * exclWrkBk.write(fileOutputStream) in the previous section as we are trying to do ANOTHER WRITE
		 * Keeping the opened fileOutputStream as-is. We are going to need it anyway!
		 */
		
		//OVER-WRITING VALUE on a SPECIFIC CELL (non-blank cell) on a ROW IN “RANGE” (existing/ occupied cell)
		/*Over-write the VALUE on the LAST CELL on row 4
		 * When over-writing, you know the cell already has content and therefore a DATA "TYPE" is already associated with the cell
		 * All setCellValue methods IMPLICITLY CONVERT the CELL to the INPUT TYPE!
		 * You DON'T NEED to KNOW the CELL TYPE when WRITING.
		 * We will STILL DO IT just this ONCE to see the magic work
		 */
		short lastCell = sheet.getRow(3).getLastCellNum(); 
		/*getLastCellNum() returns last cell INDEX PLUS ONE (NON-ZERO base. This was done in the interest of SIMPLIFIED INDEX BASED LOOPING(<UpperLimit)
		*/
		System.out.println(lastCell); 
		lastCell -= 1; //Since you are NOT LOOPING, you will remain out of RANGE and invite an exception if you miss this important step
		
		System.out.println(sheet.getRow(3).getCell(lastCell).getCellTypeEnum());
		
		//Trying to overwrite a NUMERIC TYPE cell with a non-NUMERIC VALUE. 
		sheet.getRow(3).getCell(lastCell).setCellValue("02/07/2018");
		exclWrkBk.write(fileOutputStream);
		exclWrkBk.close();
		
	}

	
}

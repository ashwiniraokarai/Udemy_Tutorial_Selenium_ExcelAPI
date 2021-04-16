package packageDemoApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class C05_Read_Iteration {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		/*
		 * Enclosing the READ and WRITE FUNCTIONALITY in REUSABLE getter and setter-like
		 * METHODS RoadMap: Mark the code that could be potential "METHOD" candidates
		 * and wrap them in methods -fnGetCellValue -fnSetCellValue (should return a
		 * value, eventually). Set the ACCESS MODIFIER, RETURN TYPE, etc. Call the
		 * METHODS to TEST if the CALL itself is working fine (BASIC TEST) BEFORE MAKING
		 * ANY MASSIVE ALTERATIONS. Identify and declare VARIABLES that need to be
		 * OUTSIDE the METHODS in order to facilitate METHOD CALLS Add ARGUMENTS to the
		 * METHODS as NECESSARY to suit the METHOD CALL NEEDS Receive RETURNED VALUES
		 * (if applicable)
		 */

		// Method calls
		C05_Read_Iteration rdWrt = new C05_Read_Iteration();
		// READS:
		// rdWrt.fnGetCellValue(4, 0);
		// rdWrt.fnGetCellValue(6,0);
		// rdWrt.fnGetCellValue(4,2);
		// rdWrt.fnGetCellValue(1,1);
		// rdWrt.fnGetCellValue(5,1);
		// rdWrt.fnGetCellValue(5,0);

		// WRITES:
		//rdWrt.fnSetCellValue(5, 0);
		rdWrt.fnSetCellValue(6, 0);

	}

	public void fnGetCellValue(int rowIndex, int columnIndex) throws IOException {
		// Code to READ data from EXCEL
		FileInputStream fileInputStream = new FileInputStream(
				"C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");

		// Reaching a SPECIFIC CELL
		XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream);

		XSSFSheet sheet = exclWrkBk.getSheet("sheet1"); // or exclWrkBk.getSheetAt(0);

		// You need 2 loops - outer loops to iterate over rows - inner loop to iterate
		// over cells/ columns
		// Now for the Outer For Loop (after building and testing Inner Loop)
		int j;
		int lastRowIndex = sheet.getLastRowNum(); // Returns last row index (0 based)
		int flagOuterLoop = 0;
		outerLoop: for (j = 0; j <= lastRowIndex; j++) { // <= because getLastRowNum returned the last INDEX, not
															// actually the
															// NUMBER
			if (j == rowIndex) {
				flagOuterLoop = 1;
				break outerLoop;
			} // Else---keep looping (Flag continued to be “Not found”)
		} // End outerLoop

		/*
		 * Check Flag If Flag is “found”------ do whatever it is you wanted to do
		 * //ACTUAL OPERATION HERE
		 */
		if (flagOuterLoop == 1) {
			// Begin with the Inner For loop to iterate over columns (cells) and then build
			// your way up
			XSSFRow rowHandle = sheet.getRow(rowIndex);
			int i = 0;
			int lastCellNum = rowHandle.getLastCellNum(); // Returns last cell index PLUS ONE
			int flagInnerLoop = 0;
			innerLoop: for (i = 0; i < lastCellNum; i++) {

				if (i == columnIndex) {
					flagInnerLoop = 1; // Mark of search success
					break innerLoop; // Stop looping because you have found your cell
				} // else keep looping. flag continues to be 0
			} // End innerLoop

			if (flagInnerLoop == 1) {
				// ACTUAL OPERATIONS HERE
				XSSFCell cellHandle = rowHandle.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);

				CellType cellType = cellHandle.getCellTypeEnum();
				System.out.println("Cell Type is: " + cellType);

				if (cellType.equals(CellType.BLANK)) {
					String result = "BLANK";
					System.out.println("Value read: " + result); // Temporary check point
				} else {

					String result = cellHandle.toString();
					System.out.println("String converted value is: " + result);
				}
			} else { // Else-report cell out of bounds
				System.out.println("Cell(" + columnIndex + ") out of range");
			}
		} else {
			// Else-report row out of bounds
			System.out.println("Received row index (" + rowIndex + ") out of limit");
		}

		exclWrkBk.close(); // Close the workbook
	}

	public void fnSetCellValue(int rowIndex, int columnIndex) throws IOException {
		// Code to WRITE data to EXCEL
		String cellValue = "02/09/2018";

		FileInputStream fileInputStream = new FileInputStream(
				"C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
		/*
		 * You COULD choose to pass the file path as an argument instead. It can be
		 * reused on the OutputStream. Also including a similar argument on
		 * fnGetCellValue would be a good idea
		 */

		XSSFWorkbook exclWrkBk = new XSSFWorkbook(fileInputStream);
		/*
		 * Keep in mind to instantiate with the constructor that takes filestream input.
		 * Eclipse will not provide "suggestions" because constructor XSSFWorkbook()
		 * with no arguments is valid too!
		 */

		XSSFSheet sheet = exclWrkBk.getSheet("Sheet1");
		// You COULD choose to pass the sheet as an argument instead and iterate through
		// the workbook using the sheetIterator

		// OVER-WRITING VALUE on a SPECIFIC CELL (non-blank cell) on a ROW IN “RANGE”
		// (existing/ occupied cell)

		// Verify the ROW and CELL you are writing to are in the occupied row range and
		// cell range
		int j;
		int lastRowNum = sheet.getLastRowNum();
		int flagOuterLoop = 0;
		outerLoop: for (j = 0; j <= lastRowNum; j++) {
			if (j == rowIndex) {
				flagOuterLoop = 1;
				break outerLoop;
			} // else keep looping through the rows until last row
		} // end outerLoop

		if (flagOuterLoop == 1) {
			// Iterate over cells (columns) of a row
			int i;
			short lastCellNum = sheet.getRow(lastRowNum).getLastCellNum();
			int flagInnerLoop = 0;
			innerLoop: for (i = 0; i < lastCellNum; i++) {
				if (i == columnIndex) {
					flagInnerLoop = 1;
					break innerLoop;
				} // else keep looping through cells until last cell in range

			} // end innerLoop

			if (flagInnerLoop == 1) {
				sheet.getRow(rowIndex).getCell(columnIndex).setCellValue(cellValue);

				FileOutputStream fileOutputStream = new FileOutputStream(
						"C:\\Users\\Ash\\Documents\\Trainings\\Personal Notes\\Selenium\\Udemy Course Based\\Supplements\\TestData.xlsx");
				exclWrkBk.write(fileOutputStream);

			} else {
				System.out.println("Cell(" + columnIndex + ") out of range");
			}

		} else {
			System.out.println("Row(" + rowIndex + ") out of range");
		}

		exclWrkBk.close(); // Close the workbook

	}
}

package SeleniumDataDriveTest;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class DDT {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream file = new FileInputStream("D:\\TestData.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");  //proving sheet name
		//XSSFSheet sheet = workbook.getSheetAt(0);  //proving sheet name

		int rowcount = sheet.getLastRowNum(); //returns the row count

		int colcount = sheet.getRow(0).getLastCellNum(); //returns column/cell count

		for(int i=0; i< rowcount; i++) {
			XSSFRow currentrow = sheet.getRow(i);  //focussed on current row
			for(int j=0; j<colcount; j++) {
				String value = currentrow.getCell(j).toString(); //read the value from a cell
				System.out.println(" " + value);
			}
			System.out.println();
		}
	}

}

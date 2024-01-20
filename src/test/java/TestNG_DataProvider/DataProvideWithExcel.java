package TestNG_DataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.*;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataProvideWithExcel {

	WebDriver driver;

	@BeforeMethod
	public void setup() {
		//launch chrome browser
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();

		//Open URL
		driver.get("http://www.google.com");

		//maximize browser
		driver.manage().window().maximize();
	}

	@Test(dataProvider = "searchDataProvider")
	public void searchKeyword(String keyword) {
		WebElement searchbox = driver.findElement(By.name("q"));
		searchbox.sendKeys(keyword);
		searchbox.sendKeys(Keys.ENTER);
	}

	@DataProvider(name="searchDataProvider")
	public Object[][] searchDataProviderMethod(){
		String fileName = "D:\\SearchData.xlsx";
		Object[][] searchData = getExcelData(fileName,"Sheet1");
//		Object[][] searchData = new Object[2][1];   //2rows & 1 Col
//		searchData[0][0] ="Taj Mahal"; // row=1 , col=1
//		searchData[1][0] ="India Gate"; // row=2 , col=1
		return searchData;
	}

	public String[][] getExcelData(String fileName, String sheetName){
		//declare array
		String[][] data = null;
		//Open file read open
		try {
			FileInputStream inputStream = new FileInputStream(fileName);
			//create XSSFWorkBook Class object for excel file manipulation
			XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
			XSSFSheet excelSheet = workBook.getSheet(sheetName);
			//get total no. of rows
			int ttlRows = excelSheet.getLastRowNum()+1;
			//get total no. of cells
			int ttlCells = excelSheet.getRow(0).getLastCellNum();

			//Initialize array
			data = new String[ttlRows-1][ttlCells];
			
			for(int currentRow = 1; currentRow < ttlRows;currentRow++) { // loop for Row
				for(int currentCell = 0; currentCell<ttlCells; currentCell++) {

					//System.out.println();
					data[currentRow-1][currentCell] = excelSheet.getRow(currentRow).getCell(currentCell).getStringCellValue();
				}

			}

			workBook.close();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return data;
	}

	@AfterMethod
	public void tearDown() {
		driver.quit();
	}
}

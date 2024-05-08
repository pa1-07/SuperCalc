package superCal;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Random;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class SuperCalculatorExcel {

	WebDriver driver;
	Random random;
	double expectedResult;
	String filepath = ".\\datafiles\\SuperCalcSheet.xlsx";

	@BeforeMethod
	public void setUp() {
		// Initialize WebDriver and open the browser
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://juliemr.github.io/protractor-demo/");
		random = new Random();
	}

	@Test
	public void performCalculation() throws IOException, InterruptedException {
		
		FileInputStream file = new FileInputStream(filepath);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i);
			int firstValue = (int) row.getCell(0).getNumericCellValue();
			int secondValue = (int) row.getCell(1).getNumericCellValue();
			// int expectedResult = firstValue + secondValue;
			// Locate elements and perform calculation
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement element = wait
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h3[text()='Super Calculator']")));
			System.out.println("Element is visible");
			WebElement firstInput = driver.findElement(By.xpath("//input[@ng-model='first']"));
			WebElement secondInput = driver.findElement(By.xpath("//input[@ng-model='second']"));
			Select operatorDrpDwn = new Select(driver.findElement(By.tagName("select")));
			WebElement addOperator = driver.findElement(By.xpath("//option[@value='ADDITION']"));
			WebElement goBtn = driver.findElement(By.xpath("//button[text()='Go!']"));
			List<WebElement> resultElement = driver.findElements(By.xpath("//td[@class='ng-binding'][2]"));

			firstInput.clear();
			firstInput.sendKeys(Integer.toString(firstValue));
			int randomIndex = random.nextInt(operatorDrpDwn.getOptions().size());
			operatorDrpDwn.selectByIndex(randomIndex);
			String selectedOption = operatorDrpDwn.getFirstSelectedOption().getText();

			switch (selectedOption) {
			case "+":
				expectedResult = firstValue + secondValue;
				break;
			case "-":
				expectedResult = firstValue - secondValue;
				break;
			case "*":
				expectedResult = firstValue * secondValue;
				break;
			case "/":
				expectedResult = firstValue / secondValue;
				break;
			case "%":
				expectedResult = firstValue % secondValue;
				break;
			default:
				throw new IllegalArgumentException("Invalid operator: " + selectedOption);
			}
			secondInput.clear();
			secondInput.sendKeys(String.valueOf(secondValue));
			goBtn.click();
			Thread.sleep(3000);
			WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
			WebElement result = wait1
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[@class='ng-binding'][2]")));

			String actualResultText = result.getText();
			double actualResult;
			try {
				actualResult = Double.parseDouble(actualResultText);
				// Check if the parsed value is a whole number (integer)
				if (actualResult == Math.floor(actualResult)) {
					// If it's a whole number, convert it to an integer
					actualResult = (int) actualResult;
					
				}
			} catch (NumberFormatException e) {
				System.out.println("Unable to parse actual result: " + actualResultText);
				actualResult = Double.NaN; // Set to NaN (Not a Number)
			}
			
			row.createCell(2).setCellValue(selectedOption);
			row.createCell(3).setCellValue(actualResult);
		}
		
		FileOutputStream outFile = new FileOutputStream(filepath);
		workbook.write(outFile);
		
		outFile.close();
		workbook.close();
		file.close();
	}

	@AfterMethod
	public void tearDown() {
		// Close the browser
		driver.quit();
	}

	public int getRandomNumber() {
		return random.nextInt(100) + 1;

	}

}

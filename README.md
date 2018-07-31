/**
 * 
 */
package com.fox.selenium.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 *
 */
public class Foxtesting {
	WebDriver driver = null;
	String baseUrl = null;
	WebDriverWait wait = null;

	@Before
	public void init() {
		System.setProperty("webdriver.firefox.marionette", "D:\\geckodriver.exe");
		driver = new FirefoxDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		baseUrl = "https://www.fox.com";
		wait = new WebDriverWait(driver, 5);
	}

	@Test()
	public void accountCreation() {
		driver.get(baseUrl);
		driver.manage().window().maximize();
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"path-1\"]")));
		driver.findElement(By.xpath("//*[@id=\"path-1\"]")).click();
		wait.until(
				ExpectedConditions.elementToBeClickable(By.xpath("//div[1]/div/div[2]/div[2]/div/div[4]/button[1]")));
		driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div/div[4]/button[1]")).click();

		WebElement firstName = driver
				.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[4]/div[1]/input"));
		WebElement secondName = driver
				.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[4]/div[2]/input"));
		WebElement email = driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[6]/input"));
		WebElement pasword = driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[8]/div/input"));
		WebElement gender = driver
				.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[10]/div[1]/div/div/div/a"));
		WebElement dob = driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[10]/div[2]/input"));

		firstName.sendKeys("Pratik");
		secondName.sendKeys("Marwaha");
		email.sendKeys("pmarwa779@gmail.com");
		pasword.sendKeys("pmarwa779");
		dob.sendKeys("06/04/1990");
		Select gen = new Select(gender);
		gen.selectByIndex(1);

	}

	@Test
	public void loginAndExcelGeneration() {

		driver.get(baseUrl);
		driver.manage().window().maximize();
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"path-1\"]")));
		driver.findElement(By.xpath("//*[@id=\"path-1\"]")).click();
		wait.until(
				ExpectedConditions.elementToBeClickable(By.xpath("//div[1]/div/div[2]/div[2]/div/div[4]/button[2]")));
		driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div/div[4]/button[2]")).click();

		WebElement userName = driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[4]/input"));
		WebElement password = driver
				.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[4]/div[2]/input"));
		WebElement submit = driver.findElement(By.xpath("//div[1]/div/div[2]/div[2]/div[1]/div[2]/div[13]/button"));

		userName.sendKeys("pmarwa779@gmail.com");
		password.sendKeys("pmarwa779");
		submit.click();

		JavascriptExecutor js = (JavascriptExecutor) driver;

		js.executeScript("window.scrollBy(0,1000)");

		WebElement first = driver
				.findElement(By.xpath("//div/div[3]/div/div[1]/div[1]/div[2]/div[1]/a[1]/div/span/span"));
		WebElement second = driver
				.findElement(By.xpath("//div/div[3]/div/div[1]/div[1]/div[2]/div[1]/a[2]/div/span/span"));
		WebElement third = driver
				.findElement(By.xpath("//div/div[3]/div/div[1]/div[1]/div[2]/div[1]/a[3]/div/span/span"));
		WebElement fourth = driver
				.findElement(By.xpath("//div/div[3]/div/div[1]/div[1]/div[2]/div[1]/a[4]/div/span/span"));

		
		saveDataToExcel(first.getText(),second.getText(),third.getText(),fourth.getText());

	}

	private void saveDataToExcel(String data1, String data2, String data3, String data4) {
		try (FileOutputStream outputStream = new FileOutputStream("fox.xlsx")) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("FOX-Data");
		Object[][] bookData = { { data1 }, { data2 },
				{ data3 }, { data4 }, };

		int rowCount = 0;

		for (Object[] aBook : bookData) {
			Row row = sheet.createRow(++rowCount);

			int columnCount = 0;

			for (Object field : aBook) {
				Cell cell = row.createCell(++columnCount);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}

		}	
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@After
	public void destroy() {
		driver.close();
	}

}

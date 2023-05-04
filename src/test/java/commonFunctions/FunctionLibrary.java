package commonFunctions;

import java.io.FileInputStream;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Properties;

import javax.swing.Action;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.reporters.jq.Main;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

import freemarker.core.Environment;

public class FunctionLibrary {
	public static WebDriver driver;
	public static Properties conpro;
	public static String Expected ="";
	public static String Actual = "";
	
	//method for launch browser

	public static WebDriver startBrowser() throws Throwable{
		conpro =new Properties();
		conpro.load(new FileInputStream("./PropertyFile/Environment.properties"));
		if (conpro.getProperty("Browser").equalsIgnoreCase("chrome"))
		{
			driver =new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().deleteAllCookies();
		}
		else if (conpro.getProperty("Browser").equalsIgnoreCase("firefox"))
		{
			driver =new FirefoxDriver();
			driver.manage().deleteAllCookies();
		}
		else
		{
			System.out.println("Browser value is not matching");
		}
		return driver;		

		//method for launch URL
	}

	public static void openUrl(WebDriver driver) 
	{
		driver.get(conpro.getProperty("Url"));

	}
	public static void waitforElement(WebDriver driver,String LocatorType,String LocatorValue,String waitTime) {

		WebDriverWait mywait =new WebDriverWait(driver, Integer.parseInt(waitTime));
		if(LocatorType.equalsIgnoreCase("name"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.name(LocatorValue)));

		}
		else if(LocatorType.equalsIgnoreCase("xpath"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(LocatorValue)));
		}
		else if(LocatorType.equalsIgnoreCase("id"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.id(LocatorValue)));
		}

	}
	//method for textboxes
	public static void typeAction(WebDriver driver,String LocatorType,String LocatorValue,String TestData)
	{
		if (LocatorType.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(LocatorValue)).clear();
			driver.findElement(By.id(LocatorValue)).sendKeys(TestData);

		}
		else if(LocatorType.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(LocatorValue)).clear();
			driver.findElement(By.xpath(LocatorValue)).sendKeys(TestData);
		}
		else if(LocatorType.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(LocatorValue)).clear();
			driver.findElement(By.name(LocatorValue)).sendKeys(TestData);
		}
	}
	//method for buttons,radio,checkbox,links and images

	public static void clickAction(WebDriver driver,String LocatorType,String LocatorValue)
	{
		if (LocatorType.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(LocatorValue)).click();
		}
		else if  (LocatorType.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(LocatorValue)).click();
		}
		else if (LocatorType.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(LocatorValue)).sendKeys(Keys.ENTER);
		}
	}
	// method for validating title

	public static void validateTitle(WebDriver driver,String Expected_Title)
	{
		String Actual_Title =driver.getTitle();
		try {
			Assert.assertEquals(Expected_Title,Actual_Title,"Title is not matching");
		}catch (Throwable t)
		{
			System.out.println(t.getMessage());

		}
	}
	//method for closing browser

	public static void closeBrowser(WebDriver driver)
	{
		driver.quit();
	}
	//method for mouse click
	/*public static void mouseClick (WebDriver driver) throws Throwable {
		Actions ac = new Actions(driver);
		ac.moveToElement(driver.findElement(By.xpath("//a[starts-with(text(),'Stock Items ')]"))).build().perform();
		//Thread.sleep(3000);
		ac.moveToElement(driver.findElement(By.xpath("(//a[contains(text(),'Stock Categories')])[2]"))).click().perform();


	}
	//Validate stock table
	public static void categoryTable(WebDriver driver,String ExpectedData) throws Throwable {
		//if search textbox already displayed no need to click search panel
		if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
			driver.findElement(By.xpath(conpro.getProperty("search-panel "))).click();
		driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(ExpectedData);
		Thread.sleep(3000);
		driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
		Thread.sleep(3000);
	     String ActualData = driver.findElement(By.xpath("//table[@id='tbl_a_stock_categorieslist']/tbody/tr[1]/td[4]/div/span/span")).getText();
	     System.out.println(ExpectedData +"   " + ActualData);
	     Assert .assertEquals(ExpectedData, ActualData, "Category name not matching");
	}*/

	//method for mouse click

	public static void mouseClick(WebDriver driver) throws Throwable

	{

	Actions ac = new Actions(driver);

	ac.moveToElement(driver.findElement(By.xpath("//a[starts-with(text(),'Stock Items ')]"))).perform();

	Thread.sleep(3000);

	ac.moveToElement(driver.findElement(By.xpath("(//a[contains(text(),'Stock Categories')])[2]"))).click().perform();

	}

	//method stock table

	public static void categoryTable(WebDriver driver,String ExpectedData) throws Throwable

	{

	//if search textbox already displayed no need to click search panel

	if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())

	driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();

	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(ExpectedData);

	Thread.sleep(3000);

	driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();

	Thread.sleep(3000);

	String ActualData =driver.findElement(By.xpath("//table[@id='tbl_a_stock_categorieslist']/tbody/tr[1]/td[4]/div/span/span")).getText();

	System.out.println(ExpectedData+"   "+ActualData);

	Assert.assertEquals(ExpectedData, ActualData, "Category Name Not Matching");

	}
	//method for Capture Supplier number
	public static void captureSnumber (WebDriver driver,String Locator_Type,String Locator_Value) throws Throwable
	{
		Expected =driver.findElement(By.name(Locator_Value)).getAttribute("Value");
				
	}
	//method for supplier table
	public static void supplierTable(WebDriver driver) throws Throwable {
		if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())

			driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();

			driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Expected);

			Thread.sleep(3000);

			driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();

			Thread.sleep(3000);

			Actual =driver.findElement(By.xpath("//table[@id ='tbl_a_supplierslist']/tbody/tr[1]/td[6]/div/span/span ")).getText();
			System.out.println(Expected+"   "+Actual);

			Assert.assertEquals(Expected, Actual, "Supplier Number Not Matching");
	}
	//method to capture customer number
	public static void captureCnumber (WebDriver driver,String Locator_Type,String Locator_Value)
	{
		Expected =driver.findElement(By.name(Locator_Value)).getAttribute("Value");
	}
	//method for validate customer Table
	public static void customerTable(WebDriver driver) throws Throwable {
		if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())

			driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();

			driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Expected);
			Thread.sleep(3000);
			driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
			Thread.sleep(3000);
			Actual =driver.findElement(By.xpath("//table[@id ='tbl_a_customerslist']/tbody/tr[1]/td[5]/div/span/span ")).getText();
			System.out.println(Expected+"   "+Actual);
			Assert.assertEquals(Expected, Actual, "Customer Number Not Matching");
	}
	//method for date generate
	public static String generateData() {
		java.util.Date date =new java.util.Date();
		DateFormat df =new SimpleDateFormat("YYYY/MM/DD HH/MM");
		return df.format(date);
	}
}

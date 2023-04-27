package commonFunctions;

import java.util.Properties;

import org.openqa.selenium.WebDriver;

public class FunctionLibrary {
	public static WebDriver driver;
	public static Properties conpro;
	public static String Expected =" ";
	public static String Actual = " ";
	
	//method for launch browser
	
	public static WebDriver startBrowser() throws Throwable{
		conpro =new Properties();
		return driver;
		
		
		
	}
	
	

}

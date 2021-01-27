package keyworddrivenframework;

import java.io.File;
import java.io.IOException;




import java.util.concurrent.TimeUnit;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataEngine {
	
	public WebDriver driver;
	
	@BeforeTest
	public void setUp() {
		
		//System.setProperty("webdriver.gecko.driver", "C:\\geckodriver-win32\\geckodriver.exe");
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\rangarajan_r\\Desktop\\chromedriver.exe");
		//System.setProperty("webdriver.ie.driver", "D:\\IE DRIVER 3.4.0\\IEDriverServer.exe");
		//driver = new FirefoxDriver();
		//driver = new InternetExplorerDriver();
		//ChromeOptions o = new ChromeOptions();
		//o.addArguments("disable-extensions");
		//o.addArguments("--start-maximized");
		 //driver = new ChromeDriver(o);
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		
	}

	
	public By GetElementLocator(String locatorType, String locatorValue) {
		
		switch (locatorType.toUpperCase())
		{
		 
		case "CLASSNAME" :
			return By.className(locatorValue);
			
		case "CSSSELECTOR" :
			return By.cssSelector(locatorValue);
			
		case "ID" :
			return By.id(locatorValue);
			
		case "PARTIALLINKTEXT" :
			return By.partialLinkText(locatorValue);
			
		case "NAME" :
			return By.name(locatorValue);
			
		case "XPATH" :
			return By.xpath(locatorValue);
			
		case "TAGNAME" :
			return By.tagName(locatorValue);
			
		default :
			return By.id(locatorValue);
		}
		
	}
	
	@Test(dataProvider = "testdata")
	public void performAction(String keyword, String locatorType, String locatorValue, String parameter) {
		
		switch(keyword.toUpperCase()) 
		{
		
		case "CLICK" :
			
			driver.findElement(GetElementLocator(locatorType, locatorValue)).click();
			break;
			
        case "SENDKEYS" :
			
			driver.findElement(GetElementLocator(locatorType, locatorValue)).sendKeys(parameter);
			//driver.manage().timeouts().implicitlyWait(5000,TimeUnit.MILLISECONDS);
			break;
			
        case "SELECT" :
			
			Select sel = new Select(driver.findElement(GetElementLocator(locatorType, locatorValue)));
			
			sel.selectByVisibleText(parameter);
			break;
			
        case "NAVIGATE" :
			
			driver.get(parameter);
			
			
			break;
			
        case "WAITFORPAGETOLOAD" :
			
        	driver.manage().timeouts().implicitlyWait(10000,TimeUnit.MILLISECONDS);
			
			break;
			
			default :
				
				System.out.println("Invalid Keyword");
				
				break;
			
						
		}
		
	}
	
	@DataProvider(name = "testdata")
	public Object [] [] dataProvider () throws BiffException, IOException {
		
		File file = new File ("D:\\input.xls");
		
		Workbook w = Workbook.getWorkbook(file);
		
		Sheet s = w.getSheet("Sheet1");
		
		int rows = s.getRows() - 1;
		
		int columns = s.getColumns();
		
		String inputData [] [] = new String [rows] [columns];
		
		
		for (int i=0,x=1; i<rows; i++,x++){
			
		    for (int j=0,y=0; j<columns; j++,y++){
		    
		    	Cell c = s.getCell(y, x);
		    
		    	inputData [i][j] = c.getContents();
		    
		    	//System.out.println(inputData[i][j]);
		}
		}
		return inputData;
		
	}
	
	@AfterClass
	public void tearDown() {
		driver.close();
	}
}

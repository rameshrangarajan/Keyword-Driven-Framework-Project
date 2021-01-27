package keyworddrivenframework;


import static org.testng.Assert.assertEquals;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;




import java.net.URISyntaxException;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;





//import jxl.Cell;
//import jxl.Sheet;
//import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
//import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.LocalFileDetector;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class NlproTest {
	
	public WebDriver driver;
	public ReadProperties read;
	public Properties prop;
	public Workbook workbook;
	//public WritableWorkbook workbookCopy;
	public ITestResult result;
	int rownumber [];
	int t = 0;
	public Xls_Reader xls;
	public Xls_Reader xls1;
	String testcases [];
	int rows;
	//int m;
	
	/*@BeforeClass
	public void writetoExcel() throws BiffException, IOException {
	File file = new File ("D:\\NLPROINPUT_UIMAPPING.xls");
	
	workbook = Workbook.getWorkbook(file);
	workbookCopy= Workbook.createWorkbook(new File("D:\\NLPROINPUT_UIMAPPING_COPY.xls"), workbook);
	
	}*/
	
	@BeforeTest
	public void setUp() throws Exception {
		
		read = new ReadProperties();
	    prop = read.getProperties();
	    xls = new Xls_Reader(prop.getProperty("filePath"));
	    String sheetname = prop.getProperty("sheetName");
	    //rownumber = xls.getCellRowNum(sheetname, "TCID", "measureTime");
	    //xls1 = new Xls_Reader(prop.getProperty("filePath1"));
	    //String sheetname1 = prop.getProperty("sheetName");
		//System.setProperty("webdriver.gecko.driver", "C:\\geckodriver-win32\\geckodriver.exe");
	    //DesiredCapabilities capabilities = DesiredCapabilities.firefox();
	    //capabilities.setCapability("recreateChromeDriverSessions", true);
	    //capabilities.setCapability(CapabilityType.BROWSER_NAME, "firefox");
	    //capabilities.setCapability("marionette", true);
	    //capabilities.setCapability("firefox_binary","C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe");
	    
	    //Gecko driver to be used -:
	    //System.setProperty("webdriver.gecko.driver", "D:\\Geckodriverv0.16.1win64\\geckodriver.exe");
	    //driver = new FirefoxDriver();
	    
	    //Chrome driver to be used -:
		//System.setProperty("webdriver.chrome.driver", "C:\\Users\\rangarajan_r\\Desktop\\chromedriver.exe");
	    //driver = new ChromeDriver();
			    
	    //IE driver to be used -:
		//System.setProperty("webdriver.ie.driver", "D:\\IE DRIVER 3.4.0\\IEDriverServer.exe");
		//DesiredCapabilities ieCaps = DesiredCapabilities.internetExplorer();
		//ieCaps.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, "http://www.bing.com/");
		//driver = new InternetExplorerDriver(ieCaps);
		//driver = new InternetExplorerDriver();
	    
	    
		//ChromeOptions o = new ChromeOptions();
		//o.addArguments("disable-extensions");
		//o.addArguments("--start-maximized");
		//driver = new ChromeDriver(o);
		
		//driver.manage().window().maximize();
		
	}

	
	public By GetElementLocator(String objectValue) throws Exception {
		
		//Properties prop = read.getProperties();
		
		String string = prop.getProperty(objectValue);
		String [] values = string.split(",");
		
		String locatorType = values[0];
		String locatorValue = values[1];
		
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
			
		case "LINKTEXT" :
			return By.linkText(locatorValue);
			
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
	public void performAction(String keyword, String objectValue, String parameter) throws Exception {
		
		String data = prop.getProperty(parameter);
		
		switch(keyword.toUpperCase()) 
		{
		
		case "HIGHLIGHTSPECIFICTEXT" : 
			
			//new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue))).doubleClick().build().perform();
			WebElement text = driver.findElement(GetElementLocator(objectValue));
			Actions actions = new Actions(driver);
		    actions.moveToElement(text, 10, 5).clickAndHold().moveByOffset(30, 0).release().perform();
		    //JavascriptExecutor js = (JavascriptExecutor) driver;

		    //js.executeScript("arguments[0].setAttribute('style', 'background: blue;');", text);
		    //js.executeScript("arguments[0].style.backgroundColor = 'red';", text);
		    //js.executeScript("arguments[0].select();", text);
		    //js.executeScript("document.getElementById('" + data + "').select();");
		    
		    break;
		    
		case "DOUBLECLICK" :
			
			new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue)), 35, 10).doubleClick().build().perform();
			break;
	       
		case "CLICK" :
			
			//WebDriverWait wait2 = new WebDriverWait(driver, 10);
        	//wait2.until(ExpectedConditions.visibilityOfElementLocated(GetElementLocator(objectValue)));
			new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue))).click().build().perform();
			//driver.findElement(GetElementLocator(objectValue)).click();
			break;
			
        
			
        case "SENDKEYS" :
			
			driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
			//driver.manage().timeouts().implicitlyWait(5000,TimeUnit.MILLISECONDS);
			break;
			
        case "SELECT" :
			
			Select sel = new Select(driver.findElement(GetElementLocator(objectValue)));
			
			sel.selectByVisibleText(data);
			break;
			
        case "NAVIGATE" :
        	String browserName = prop.getProperty("chromebrowser");
        	//String browserName = prop.getProperty("firefoxbrowser");
        	//String browserName = prop.getProperty("iebrowser");
        	driver = BrowserFactory.startBrowser(browserName, data);
			
			//driver.navigate().to(data);
			
			
			break;
			
              	
			
        case "VERIFYTEXT" :
 			
        	String actualText = driver.findElement(GetElementLocator(objectValue)).getText();
        	
        	Assert.assertEquals(actualText, data);
        	
        case "VERIFYPOPUPISVISIBLE" :
        	
        	Assert.assertTrue(driver.findElement(GetElementLocator(objectValue)).isDisplayed(), "Twitter Popup is not visible");
			
			break;
			
        case "WAIT" :
			
        	//driver.manage().timeouts().implicitlyWait(10000,TimeUnit.MILLISECONDS);
        	
        	Thread.sleep(5000);
			
			break;
			
        case "IMPLICITLYWAIT" :
        	
        	driver.manage().timeouts().implicitlyWait(10000,TimeUnit.MILLISECONDS);
        	
        	break;
        	
        case "SWITCHTO" :
        	Set<String> set = driver.getWindowHandles();
        	Iterator<String> iterate = set.iterator();
        	String str1 = iterate.next();
        	//String str2 = iterate.next();
        	//driver.switchTo().frame(driver.findElement(GetElementLocator(locatorType, locatorValue)));
        	driver.switchTo().window(str1);
        	break;
        	
        case "CLOSEBROWSER" :
        	
        	driver.close();
        	break;
        	
        case "SWITCHTOPARENTFRAME" :
        	
        	driver.switchTo().parentFrame();
        	break;
        	
        case "SWITCHTOALERT" :
        	
        	//String str2 = iterate.next();
        	//driver.switchTo().frame(driver.findElement(GetElementLocator(locatorType, locatorValue)));
        	driver.switchTo().alert();
        	break;
        	
        case "HIGHLIGHTTEXT" :
        	
        	driver.findElement(GetElementLocator(objectValue)).sendKeys(Keys.chord(Keys.CONTROL, "a"));
        	//new Actions(driver).keyDown(Keys.CONTROL).sendKeys(String.valueOf('\u0061')).perform();
        	
            //JavascriptExecutor js = (JavascriptExecutor) driver;

            //js.executeScript("arguments[0].setAttribute('style', 'background: blue;');", text);
           //String val = "atsign";

          //JavascriptExecutor js = (JavascriptExecutor) driver;  
          //js.executeScript("document.getElementById('" + locatorValue + "').value.replace(" + parameter + ", '<span class= '" + val + "'>'" + parameter + "'</span>)');");



        	//Actions action = new Actions(driver);
        	//action.doubleClick(driver.findElement(GetElementLocator(locatorType, locatorValue)));
        	
        	break;
        	
                	
        case "RIGHTCLICK" :
        	
        	//WebDriverWait wait = new WebDriverWait(driver, 10);
            //wait.until(ExpectedConditions.presenceOfElementLocated(GetElementLocator(objectValue)));
        	//new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue))).contextClick(driver.findElement(GetElementLocator(objectValue))).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.ENTER).build().perform();
        	new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue)), 35, 10).contextClick().build().perform();
        	//new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue))).contextClick().build().perform();
        	//new Actions(driver).contextClick(driver.findElement(GetElementLocator(objectValue))).build().perform();
        	/*Robot robot = new Robot();
        	robot.mouseMove(50, 250);
            robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
            robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);*/
        	break;
        	
        case "HOVER" :
        	
        	WebDriverWait wt = new WebDriverWait(driver, 10);
        	wt.until(ExpectedConditions.visibilityOfElementLocated(GetElementLocator(objectValue)));
        	new Actions(driver).moveToElement(driver.findElement(GetElementLocator(objectValue))).build().perform();
        	break;
        	
        case "CLEAR" :
        	driver.findElement(GetElementLocator(objectValue)).clear();
        	
        	break;
        	
        case "CALCULATETIME" :
        	
        	long start = System.currentTimeMillis();

        	//driver.get("Some url");
        	WebDriverWait wait = new WebDriverWait(driver, 900);
            wait.until(ExpectedConditions.presenceOfElementLocated(GetElementLocator(objectValue)));
        	//String temp = driver.findElement(GetElementLocator(objectValue)).getText();
            //assertEquals(temp, data);
        	//WebElement ele = driver.findElement(By.id("ID of some element on the page which will load"));
        	long finish = System.currentTimeMillis();
        	long totalTime = finish - start; 
        	System.out.println("Total Time for page load - "+totalTime); 
        	break;
        	
        case "UPLOADTEXTFILE" :
        	
        	//File file = null;
        	//JavascriptExecutor js = (JavascriptExecutor) driver;
        	//js.executeScript("document.getElementById('loadTextFile').style.visibility = 'visible';");
        	//js.executeScript("document.getElementById('loadTextFile').value ='" + data + "';");
        	//driver.findElement(By.id("loadTextFile")).sendKeys(data);
        	//js.executeScript("document.getElementById('loadTextFile').style.display = 'none';");
        	//String jsScript = "var input = document.getElementById('loadTextFile');" + "input.value='" + data + "';";
        	//JavascriptExecutor executor = (JavascriptExecutor)driver;
        	//executor.executeScript(jsScript);
        	
        	//driver.findElement(By.id("submit")).click();
        	///file = new File(data);
        	//driver.findElement(GetElementLocator(objectValue)).click();
        	//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        	//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        	//WebElement element = driver.switchTo().activeElement();
        	//Actions builder = new Actions(driver);
                 //element.sendKeys(data);
        	 //Action myAction = builder.click(driver.findElement(GetElementLocator(objectValue))).sendKeys(data).release().build();

        	    //myAction.perform();
        	    //driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        	 //Alert alert = null;
        	//try{
        		/*driver.findElement(GetElementLocator(objectValue)).click();
        	
        			WebDriverWait wait = new WebDriverWait(driver, 10);
            	    wait.until(ExpectedConditions.alertIsPresent());
            	
            	    Alert alert = driver.switchTo().alert();
            		alert.sendKeys(data);*/
        		  //alert.accept();
        	//System.getProperty();
        	//String strPath = System.getProperty("user.dir")+data;
        	//String strPath = "";
    		//By by;
    		//APP_LOGS.debug("uploading Document...");
    		//try{
    			//String strPath = data;
    			//File file = new File(data);
    			//System.out.println("path:"+strPath);
        	//String strPath = System.getProperty("user.dir") +data;
			//System.out.println("path:"+strPath);
			//driver.findElement(GetElementLocator(objectValue)).sendKeys(strPath);
    			//by = object_type_identifier(OR.getProperty(object));// element
    			//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);	
    			//Thread.sleep(3);
    			//return Constants.KEYWORD_PASS;	
    		//}catch(Exception e){
    			//System.out.println(" - Getting error while document uploading" +e.getMessage());
    			//return Constants.KEYWORD_FAIL +" - Getting error while document uploading";	
    		//}
        		  //File file = new File(strPath);
        		  //driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        			//e.printStackTrace();
        		//}
        	
        	
        	
        	/*CODE TO HANDLE UPLOAD FILE FUNCTIONALITY IN IE 
        	WebDriverWait wait = new WebDriverWait(driver, 20);
        	wait.until(ExpectedConditions.alertIsPresent());
        	Alert alert = driver.switchTo().alert();
        	alert.sendKeys(data);
        	Robot r = new Robot();
        	r.keyPress(KeyEvent.VK_ENTER);
    	    r.keyRelease(KeyEvent.VK_ENTER);*/
        	
        	
        	
        	//wait.until(ExpectedConditions.visibilityOfElementLocated(GetElementLocator(objectValue)));
        	//driver.findElement(GetElementLocator(objectValue)).click();
        	/*CODE TO HANDLE UPLOAD FILE FUNCTIONALITY IN CHROME*/
        	driver.findElement(GetElementLocator(objectValue)).sendKeys(data); 
        	    //WebDriverWait wait2 = new WebDriverWait(driver, 10);
        	    //wait2.until(ExpectedConditions.alertIsPresent());
        	    //new Actions(driver).sendKeys(data);
        	    
        	    //Thread.sleep(5000);
        	    //driver.switchTo().activeElement().sendKeys(data);
        	    // switch to the file upload window
        	  //Alert alert = driver.switchTo().alert();
              //new Actions(driver).sendKeys(data);
        	    // enter the filename
        	   //alert.sendKeys(data);
        	
        	   //new Actions(driver).keyDown(Keys.ENTER).sendKeys(Keys.ENTER).build().perform();
        	   //new Actions(driver).keyUp(Keys.ENTER).sendKeys(Keys.ENTER).build().perform();
        	

        	    // hit enter
        	    //Robot r = new Robot();
        	    /*r.keyPress(KeyEvent.VK_SHIFT);
        	    r.keyPress(KeyEvent.VK_D);
        	    r.keyRelease(KeyEvent.VK_D);
        	    r.keyRelease(KeyEvent.VK_SHIFT);
        	    r.keyPress(KeyEvent.VK_A);
        	    r.keyRelease(KeyEvent.VK_A);
        	    r.keyPress(KeyEvent.VK_T);
        	    r.keyRelease(KeyEvent.VK_T);
        	    r.keyPress(KeyEvent.VK_A);
        	    r.keyRelease(KeyEvent.VK_A);
        	    r.keyPress(KeyEvent.VK_PERIOD);
        	    r.keyRelease(KeyEvent.VK_PERIOD);
        	    r.keyPress(KeyEvent.VK_T);
        	    r.keyRelease(KeyEvent.VK_T);
        	    r.keyPress(KeyEvent.VK_X);
        	    r.keyRelease(KeyEvent.VK_X);
        	    r.keyPress(KeyEvent.VK_T);
        	    r.keyRelease(KeyEvent.VK_T);*/
        	    //r.keyPress(KeyEvent.VK_ENTER);
        	    //r.keyRelease(KeyEvent.VK_ENTER);

        	    // switch back
        	    //driver.switchTo().activeElement();

        	    
        	//driver.switchTo().window("File Upload");
        	//driver.findElement(By.id("loadTextFile")).sendKeys(data);
        	
        	//Robot r = new Robot();
        	//r.keyPress(KeyEvent.VK_ENTER);
        	//r.keyRelease(KeyEvent.VK_ENTER);
        	
        	
        	//String str2 = iterate.next();
        	//driver.switchTo().frame(driver.findElement(GetElementLocator(locatorType, locatorValue)));
        	
        	//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        	//Thread.sleep(3000);
        	//driver.switchTo().activeElement().sendKeys(data);
        	//driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        	//Alert alert = driver.switchTo().alert();

        	// enter the filename
        	//alert.sendKeys(data);
        	
        	

        	// switch back
        	//driver.switchTo().activeElement();
        	break;
        	
        case "UPLOADHTMLFILE" :
        	
        	driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
        	
        	break;
        	
        case  "VERIFYINDICATION" :
        	
        	Assert.assertTrue(driver.findElement(GetElementLocator(objectValue)).getText().contains(data));
        	
        	break;
        	
         case  "VERIFYACCORDIONNOTPRESENT" :
        	
        	 try {
        	Assert.assertTrue(!driver.findElement(GetElementLocator(objectValue)).isDisplayed());
        	 }
        	catch ( NoSuchElementException e) {System.out.println("Element not found"); }
        	 
        	break;
        	
            case "CANCELUPLOADFILE" :
            	
            	//driver.findElement(GetElementLocator(objectValue)).click();
            	//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
            	//driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
            	//WebElement element = driver.switchTo().activeElement();
            	//Actions builder = new Actions(driver);
                     //element.sendKeys(data);
            	 //Action myAction = builder.click(driver.findElement(GetElementLocator(objectValue))).sendKeys(data).release().build();
                    //Thread.sleep(5000);
            	    //myAction.perform();
            	    //driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
            	    
            	    //WebDriverWait wait1 = new WebDriverWait(driver, 10);
            	    //wait1.until(ExpectedConditions.alertIsPresent());

            	    // switch to the file upload window
            	    //Alert alert1 =  driver.switchTo().alert();

            	    // enter the filename
            	    //alert1.sendKeys(data);

            	    // hit enter
            	    Robot r1 = new Robot();
            	    r1.keyPress(KeyEvent.VK_TAB);
            	    r1.keyRelease(KeyEvent.VK_TAB);
            	    Thread.sleep(1000);
            	    r1.keyPress(KeyEvent.VK_TAB);
            	    r1.keyRelease(KeyEvent.VK_TAB);
            	    Thread.sleep(1000);
            	    r1.keyPress(KeyEvent.VK_TAB);
            	    r1.keyRelease(KeyEvent.VK_TAB);
            	    Thread.sleep(1000);
            	    r1.keyPress(KeyEvent.VK_ENTER);
            	    r1.keyRelease(KeyEvent.VK_ENTER);
            	
        	
        	break;
        	
            case "VERIFYERRORFORFILE" :
            	
            	Assert.assertTrue(driver.findElement(GetElementLocator(objectValue)).getText().contains(data));
        	
            case "PRESSENTERKEY" :
            	
            	Robot ro = new Robot();
            	ro.keyPress(KeyEvent.VK_ENTER);
        	    ro.keyRelease(KeyEvent.VK_ENTER);
        	    break;
        	    
            case  "PRESSESCAPEKEY" :
            	
            	Robot ro1 = new Robot();
            	ro1.keyPress(KeyEvent.VK_ESCAPE);
        	    ro1.keyRelease(KeyEvent.VK_ESCAPE);
        	    break;
            	
            case "CONTAINSTEXT" :
            	
            	String actualText1 = driver.findElement(GetElementLocator(objectValue)).getText();
            	Assert.assertTrue(actualText1.contains(data));
        	    break;
        	    
            case "CHECKIFELEMENTISVISIBLE" :
            	
            	WebDriverWait wa = new WebDriverWait(driver, 20);
            	wa.until(ExpectedConditions.visibilityOfElementLocated(GetElementLocator(objectValue)));
            	break;
            	
            case "VERIFYVALUEATTRIBUTE" :
            	
            	String actualText2 = driver.findElement(GetElementLocator(objectValue)).getAttribute("value");
            	Assert.assertTrue(actualText2.equals(data));
            	break;
            	
			default :
				
				System.out.println("Invalid Keyword");
				
				break;
			
						
		}
		
		
	}
	
	@DataProvider(name = "testdata")
	public Object [] [] dataProvider () throws BiffException, IOException {
		
		
		//String test = "measureTime";
		int x = 0;
		int y;
		int testrows = 0;
		
		
		String sheetname1 = prop.getProperty("sheetName1");
		int rows1 = xls.getRowCount(sheetname1);
		
		testcases  = new String [rows1];
		
		for(int r1 = 2, a = 0; r1 <= rows1; r1++, a++) {
			String runmode = xls.getCellData(sheetname1, "Runmode", r1);
			if(runmode.equals("Y")) {
				
				testcases [a] = xls.getCellData(sheetname1, "TCID", r1);
				System.out.println(testcases [a]);
			}
		}
		// rownumber = xls.getCellRowNum(sheetname, "TCID", cellValue)
				
		String sheetname = prop.getProperty("sheetName");
		rows = xls.getRowCount(sheetname);
		
		for(int r = 2; r <= rows; r++) {
			
			String tcid = xls.getCellData(sheetname, "TCID", r);
			if(tcid==null) {
				continue;
			}
			
		for(int b = 0; b < testcases.length; b++) {
			
		String tcid1 = testcases[b];
		
			
			
			if(tcid.equals(tcid1)) {
				testrows++;
			}
		}
		}
		
		rownumber = new int [testrows];
		
		int columns = xls.getColumnCount(sheetname) - 6;
		String inputData [] [] = new String [testrows] [columns];
		for(int rNum= 2; rNum <= rows; rNum++) {
			
			String tcid = xls.getCellData(sheetname, "TCID", rNum);
			
			if(tcid==null) {
				continue;
			}
		
			for( int g = 0; g < testcases.length; g++) {
				
			String test = testcases[g];
		
			
			
			if(tcid.equals(test)) {
				System.out.println(tcid + " - " + test);
				//h = xls.getCellRowNum(sheetname, "TCID", tcid);
				rownumber [x] = rNum;
				//rowN [x] = h;
				//m = rNum;
				//for(int cNum=2, x = 0, y = 0; cNum <= columns; cNum++, y++) {
					//x=0;
					y=0;
					inputData [x] [y] = xls.getCellData(sheetname, "KEYWORD", rNum);
					y++;
					inputData [x] [y] = xls.getCellData(sheetname, "OBJECT", rNum);
					y++;
					inputData [x] [y] = xls.getCellData(sheetname, "PARAMETER", rNum);
					
					x++;
					
					
					//h++;
					
			}
			
					
				//}
				
			}
		}
		return inputData;
		
		
		}
	
	
	
	
	/*@DataProvider(name = "testdata")
	public Object [] [] dataProvider () throws BiffException, IOException {
		
        String path = prop.getProperty("filePath");
		
		String sheetname = prop.getProperty("sheetName");
		//File file = new File ("D:\\NLPROINPUT_UIMAPPING.xls");
		
		FileInputStream inputStream = new FileInputStream(path);
		//Workbook w = Workbook.getWorkbook(file);
		
		workbook = new HSSFWorkbook(inputStream);
		
		Sheet s = workbook.getSheet(sheetname);
		
		int rowCount = s.getLastRowNum()-s.getFirstRowNum();
		
		
		
		int columnCount = s.getRow(0).getLastCellNum() - 4;
		
		System.out.println(rowCount + " " + columnCount);
		Sheet s = w.getSheet("Sheet1");
		
		int rows = s.getRows() - 1;
		
		int columns = s.getColumns() - 1;
		
		String inputData [] [] = new String [6] [columnCount];
		
		for(int i=389,x=0; i<395; i++,x++){
			
		
			Row row = s.getRow(i);
			
		    for (int j=1,y=0; j<4; j++,y++){
		    	
		    	//System.out.println(row.getCell(j).getStringCellValue());
		    	Cell c = row.getCell(j);
		    	 if (c == null) {
		    	    // This cell is empty
		    		 inputData [x] [y] = "";
		    		 System.out.println("blank");
		    	 }
		    	if(row.getCell(j).getStringCellValue().equals("")) {
		    		inputData [i] [j] = "";
		    	}
		    	else {
		    	//System.out.println(row.getCell(j).getStringCellValue());
		    	
		    	
		    		System.out.println(row.getCell(j).getStringCellValue());
		    	    inputData [x] [y] = row.getCell(j).getStringCellValue();
		    	
		    	}
		    
		    	//inputData [i][j] = c.getContents();
		    
		    	//System.out.println(inputData[i][j]);
		}
		}
		for (int i=0,x=1; i<rows; i++,x++){
			
		    for (int j=0,y=0; j<columns; j++,y++){
		    
		    	Cell c = s.getCell(y, x);
		    
		    	inputData [i][j] = c.getContents();
		    
		    	//System.out.println(inputData[i][j]);
		}
		}
		//workbook.close();
		return inputData;
		
	}*/
	
	@SuppressWarnings("deprecation")
	@AfterMethod
	public void writeResult(ITestResult result) throws BiffException, IOException, WriteException
	{
		int status = result.getStatus();
		
		//ITestResult result = Reporter.getCurrentTestResult();
		
        String path = prop.getProperty("filePath");
		
		String sheetname = prop.getProperty("sheetName");
		
		//String chrome = prop.getProperty("chromecolumninexcel");
		
		//String iexplorer = prop.getProperty("iecolumninexcel");
		
		//String firefox = prop.getProperty("firefoxcolumninexcel");
		
		String chromecolumn = prop.getProperty("chromecolumnnameinexcel");
		
		String iecolumn = prop.getProperty("iecolumnnameinexcel");
		
		String firefoxcolumn = prop.getProperty("firefoxcolumnnameinexcel");
		
		//int ch = Integer.parseInt(chrome);
		
		//int ie = Integer.parseInt(iexplorer);
		
		//int ff = Integer.parseInt(firefox);
		//XSSFFont font =  xls.workbook.createFont();
		//font.setBold(true);
		//CellStyle style = xls.workbook.createCellStyle();
		/*for(int rNum= 2; rNum <= rows; rNum++) {
			
			String tcid = xls.getCellData(sheetname, "TCID", rNum);
		
			for( int g = 0; g < testcases.length; g++) {
				
			String test = testcases[g];
		
			
			
			if(tcid.equals(test)) {
			System.out.println("Writing to excel " + tcid + " - " + test);
		
			rownumber = xls.getCellRowNum(sheetname, "TCID", tcid);*/
		//rownumber = xls.getCellRowNum(sheetname, "TCID", tcid);
		int rownum = rownumber [t];
			System.out.println(rownum);
			
		try
	    {
			
			
	        if(status == ITestResult.SUCCESS)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.GREEN);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "PASS", cellFormat);
	        	//s.addCell(label);
	        	
	            //cell.setCellValue("PASS");
	        	//style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	           //font.setFontName("Calibri");
	            //font.setFontHeightInPoints((short) 11);
	   		 //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 //style.setFont(font);
	   		 //cell.setCellStyle(headerStyle);
	        	xls.setCellData(sheetname, chromecolumn, rownum, "PASS");
                //xls.workbook.createFont();
	            //cell.setCellValue("PASS");
	            //font.setColor(IndexedColors.GREEN.getIndex());
	            //font.setFontName("Calibri");
	            //font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 //headerStyle.setFont(font);
	   		 //cell.setCellStyle(headerStyle);

	            //Do your excel writing stuff here
	        }
	        else if(status == ITestResult.FAILURE)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.RED);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "FAIL", cellFormat);
	        	//s.addCell(label);
	        	
	        	//Cell cell = row.createCell(3);
	        	//style.setFillForegroundColor(IndexedColors.RED.getIndex());
	            //font.setFontName("Calibri");
	            //font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 //headerStyle.setFont(font);
	   		 //xls.cell.setCellStyle(headerStyle);
	        	xls.setCellData(sheetname, chromecolumn, rownum, "FAIL");
	            ///cell.setCellValue("FAIL");
	            //font.setColor(IndexedColors.RED.getIndex());
	            //font.setFontName("Calibri");
	            //font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 //headerStyle.setFont(font);
	   		 //cell.setCellStyle(headerStyle);
	            //takeScreenshot(dateTimeStamp,driver,methodName);
	            //System.out.println("Log Message:: @AfterMethod: Method-"+methodName+"- has Failed");
	            //Do your excel writing stuff here

	        }
	        else if(status == ITestResult.SKIP)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLUE);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "SKIP", cellFormat);
	        	//s.addCell(label);
	        	//Cell cell = row.createCell(3);

	            //cell.setCellValue("SKIP");
	        	//style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
	            //font.setFontName("Calibri");
	            //font.setFontHeightInPoints((short) 11);
	            xls.setCellData(sheetname, chromecolumn, rownum, "SKIP");
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 //headerStyle.setFont(font);
	   		 //cell.setCellStyle(headerStyle);
	            //System.out.println("Log Message::@AfterMethod: Method-"+methodName+"- has Skipped");

	        }
				
	        t++;
	    
			
	    }
	    catch(Exception e)
	    {
	        System.out.println("\nLog Message::@AfterMethod: Exception caught");
	        e.printStackTrace();
	    }
		
		
	/*@AfterMethod
	public void writeResult(ITestResult result) throws BiffException, IOException, WriteException
	{
		//System.out.println("Inside writeResult method");
		int status = result.getStatus();
		
		//ITestResult result = Reporter.getCurrentTestResult();
		
        String path = prop.getProperty("filePath");
		
		String sheetname = prop.getProperty("sheetName");
		
		String chrome = prop.getProperty("chromecolumninexcel");
		
		String iexplorer = prop.getProperty("iecolumninexcel");
		
		String firefox = prop.getProperty("firefoxcolumninexcel");
		
		int ch = Integer.parseInt(chrome);
		
		int ie = Integer.parseInt(iexplorer);
		
		int ff = Integer.parseInt(firefox);
		
        //File file = new File ("D:\\NLPROINPUT_UIMAPPING.xls");
		
		FileInputStream inputStream = new FileInputStream(path);
		//Workbook w = Workbook.getWorkbook(file);
		
		workbook = new XSSFWorkbook(inputStream);
		
		Sheet s = workbook.getSheet(sheetname);
		
        
		//WritableSheet  s = workbookCopy.getSheet("Sheet1");
		//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11);
	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    
	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		
		//int rowCount = s.getLastRowNum()-s.getFirstRowNum();

	    //Get the first row from the sheet
		CellStyle headerStyle = workbook.createCellStyle();
		 Font font = workbook.createFont();
		 font.setBold(true);
		 //headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		 //font.setColor(IndexedColors.RED.getIndex());
		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		 //headerStyle.setFont(font);
		 //cell.setCellStyle(headerStyle);
		//Font font = workbook.createFont();
	    //font.setBold(true);
	    //font.setColor(arg0);
	    //style.setFont(font);
		 
	    Row row = s.getRow(rownumber);
	    
	    Cell cell = row.createCell(ch);

		//int col = s.getColumns() - 1;
		//int rows = s.getRows();
		
		//int columns = s.getColumns();
		
	    try
	    {
	        if(status == ITestResult.SUCCESS)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.GREEN);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "PASS", cellFormat);
	        	//s.addCell(label);
	        	

	            cell.setCellValue("PASS");
	            font.setColor(IndexedColors.GREEN.getIndex());
	            font.setFontName("Calibri");
	            font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 headerStyle.setFont(font);
	   		 cell.setCellStyle(headerStyle);

	            //Do your excel writing stuff here
	        }
	        else if(status == ITestResult.FAILURE)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.RED);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "FAIL", cellFormat);
	        	//s.addCell(label);
	        	
	        	//Cell cell = row.createCell(3);

	            cell.setCellValue("FAIL");
	            font.setColor(IndexedColors.RED.getIndex());
	            font.setFontName("Calibri");
	            font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 headerStyle.setFont(font);
	   		 cell.setCellStyle(headerStyle);
	            //takeScreenshot(dateTimeStamp,driver,methodName);
	            //System.out.println("Log Message:: @AfterMethod: Method-"+methodName+"- has Failed");
	            //Do your excel writing stuff here

	        }
	        else if(status == ITestResult.SKIP)
	        {
	        	//WritableFont cellFont = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLUE);
	    	    //cellFont.setBoldStyle(WritableFont.BOLD);
	    	    
	    	    //WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	        	//Label label = new Label(col, rownumber, "SKIP", cellFormat);
	        	//s.addCell(label);
	        	//Cell cell = row.createCell(3);

	            cell.setCellValue("SKIP");
	            font.setColor(IndexedColors.BLUE.getIndex());
	            font.setFontName("Calibri");
	            font.setFontHeightInPoints((short) 11);
	   		 //headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	   		 headerStyle.setFont(font);
	   		 cell.setCellStyle(headerStyle);
	            //System.out.println("Log Message::@AfterMethod: Method-"+methodName+"- has Skipped");

	        }
	    }
	    catch(Exception e)
	    {
	        System.out.println("\nLog Message::@AfterMethod: Exception caught");
	        e.printStackTrace();
	    }
	    FileOutputStream outputStream = new FileOutputStream(path);

	    //write data in the excel file
	    inputStream.close();
	    workbook.write(outputStream);
	    outputStream.close();

	    rownumber++;

	}*/
	
	/*@AfterClass
	public void closeFile()
    {
        try {
            // Closing the writable work book
        	//workbookCopy.write();
        	//workbookCopy.close();

            // Closing the original work book
        	//workbook.write(arg0);
        	workbook.close();
        } catch (Exception e)

        {
            e.printStackTrace();
        }
    }*/
	/*@AfterClass
	public void tearDown() {
		driver.close();
	}*/
}
}

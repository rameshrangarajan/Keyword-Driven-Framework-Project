package keyworddrivenframework;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class TestServiceNow {
	
	
	public WebDriver driver;
	public ReadProperties read;
	public Properties prop;
	public Workbook workbook;
	
	public ITestResult result;
	int rownumber = 1;
	
	@BeforeTest
	public void setUp() throws Exception {
		
		read = new ReadProperties();
	    prop = read.getProperties();
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\rangarajan_r\\Desktop\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		
		
		
	}
	
      public By GetElementLocator(String objectValue) throws Exception {
		
		
		
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
		
		case "CLICK" :
			
			driver.findElement(GetElementLocator(objectValue)).click();
			break;
			
        case "SENDKEYS" :
			
			driver.findElement(GetElementLocator(objectValue)).sendKeys(data);
			
			break;
			
        
        case "NAVIGATE" :
			
			driver.get(data);
			
			
			break;
			
        case "VERIFYTEXT" :
 			
        	String actualText = driver.findElement(GetElementLocator(objectValue)).getAttribute("value");
        	
        	Assert.assertEquals(actualText, data);
        	
			
			
			break;
			
        case "WAIT" :
        	
        	Thread.sleep(8000);
			
			break;
			
        case "IMPLICITLYWAIT" :
        	
        	driver.manage().timeouts().implicitlyWait(20000,TimeUnit.MILLISECONDS);
        	
        	break;
        	
        
			default :
				
				System.out.println("Invalid Keyword");
				
				break;
			
						
		}
		
		
	}
	
	@DataProvider(name = "testdata")
	public Object [] [] dataProvider () throws BiffException, IOException {
		
        String path = prop.getProperty("filePath");
		
		String sheetname = prop.getProperty("sheetName");
		
		File file =	new File(path);
		
		
		
		
		FileInputStream inputStream = new FileInputStream(file);
		
		
		workbook = new HSSFWorkbook(inputStream);
		
		
		Sheet s = workbook.getSheet(sheetname);
		
		int rowCount = s.getLastRowNum()-s.getFirstRowNum();
		
		
		
		int columnCount = s.getRow(0).getLastCellNum() - 2;
		
		//System.out.println(rowCount + " " + columnCount);
		
		
		String inputData [] [] = new String [rowCount] [columnCount];
		
		for(int i=1,x=0; i<rowCount+1; i++,x++){
			
		
			Row row = s.getRow(i);
			
		    for (int j=1,y=0; j<4; j++,y++){
		    	
		    	
		    	Cell c = row.getCell(j);
		    	 if (c == null) {
		    	    // This cell is empty
		    		 inputData [x] [y] = "";
		    		 //System.out.println("blank");
		    	 }
		    	
		    	else {
		    	
		    	
		    	
		    		//System.out.println(row.getCell(j).getStringCellValue());
		    	    inputData [x] [y] = row.getCell(j).getStringCellValue();
		    	
		    	}
		    
		    	
		}
		}
		
		//workbook.close();
		return inputData;
		
	}
	
	@AfterMethod
	public void writeResult(ITestResult result) throws BiffException, IOException, WriteException
	{
		//System.out.println("Inside writeResult method");
		int status = result.getStatus();
		
		String path = prop.getProperty("filePath");
		
		String sheetname = prop.getProperty("sheetName");
		
        File file = new File (path);
		
		FileInputStream inputStream = new FileInputStream(file);
		
		
		workbook = new HSSFWorkbook(inputStream);
		
		Sheet s = workbook.getSheet(sheetname);
		
        
		
		CellStyle headerStyle = workbook.createCellStyle();
		 Font font = workbook.createFont();
		 font.setBold(true);
		 

	    Row row = s.getRow(rownumber);
	    
	    Cell cell = row.createCell(4);

		
		
	    try
	    {
	        if(status == ITestResult.SUCCESS)
	        {
	        	
	            cell.setCellValue("PASS");
	            font.setColor(IndexedColors.GREEN.getIndex());
	            font.setFontName("Calibri");
	            font.setFontHeightInPoints((short) 11);
	   		 
	   		    headerStyle.setFont(font);
	   		    cell.setCellStyle(headerStyle);

	            //Do your excel writing stuff here
	        }
	        else if(status == ITestResult.FAILURE)
	        {
	        	

	            cell.setCellValue("FAIL");
	            font.setColor(IndexedColors.RED.getIndex());
	   		 
	   		    headerStyle.setFont(font);
	   		    cell.setCellStyle(headerStyle);
	            
	            //Do your excel writing stuff here

	        }
	        else if(status == ITestResult.SKIP)
	        {
	        	

	            cell.setCellValue("SKIP");
	            font.setColor(IndexedColors.BLUE.getIndex());
	   		 
	   		    headerStyle.setFont(font);
	   		    cell.setCellStyle(headerStyle);
	            

	        }
	    }
	    catch(Exception e)
	    {
	        System.out.println("\nLog Message::@AfterMethod: Exception caught");
	        e.printStackTrace();
	    }
	    FileOutputStream outputStream = new FileOutputStream(file);

	    //write data in the excel file
	    inputStream.close();
	    workbook.write(outputStream);
	    outputStream.close();

	    rownumber++;

	}
	
	@AfterClass
	public void closeFile()
    {
        try {
            
        	workbook.close();
        } catch (Exception e)

        {
            e.printStackTrace();
        }
    }
}

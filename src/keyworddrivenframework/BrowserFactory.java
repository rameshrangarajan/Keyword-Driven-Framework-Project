package keyworddrivenframework;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

public class BrowserFactory {

	static WebDriver driver;
	
	public static WebDriver startBrowser (String browserName, String url) {
		
		if(browserName.equalsIgnoreCase("firefox")) {
			
			System.setProperty("webdriver.gecko.driver", "D:\\Geckodriverv0.16.1win64\\geckodriver.exe");
		    driver = new FirefoxDriver();
		}
		
		else if (browserName.equalsIgnoreCase("chrome")) {
			
			//System.setProperty("webdriver.chrome.driver", "C:\\Users\\rangarajan_r\\Desktop\\chromedriver.exe");
			System.setProperty("webdriver.chrome.driver", "D:\\Chromedriver_2.33\\chromedriver_2_33.exe");
		    driver = new ChromeDriver();
		}
		
		else if (browserName.equalsIgnoreCase("IE")) {
			
			System.setProperty("webdriver.ie.driver", "D:\\IE DRIVER 3.4.0\\IEDriverServer.exe");
			DesiredCapabilities ieCaps = DesiredCapabilities.internetExplorer();
			ieCaps.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, "http://www.bing.com/");
			driver = new InternetExplorerDriver(ieCaps);
			driver = new InternetExplorerDriver();
		}
		
		driver.manage().window().maximize();
		driver.get(url);
		return driver;
		
	}
}

package keyworddrivenframework;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ReadProperties {

	public Properties p = new Properties();
	
	 
	public Properties getProperties() throws Exception{
		InputStream input = new FileInputStream(new File(System.getProperty("user.dir")+"\\data.properties"));
		p.load(input);
		return p;
	}
	
	
	
	

}

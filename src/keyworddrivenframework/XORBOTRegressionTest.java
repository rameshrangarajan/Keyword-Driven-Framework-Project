package keyworddrivenframework;

import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertTrue;

import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

public class XORBOTRegressionTest {

	
	@Test
	public void regressionTestForXorbot() {
		
		try {
			File file1 = new File("C:\\Users\\rangarajan_r\\Documents\\Regression_test_suite\\Regression_test_suite\\xorbotoutputold2.txt");
			File file2 = new File("C:\\Users\\rangarajan_r\\Documents\\Regression_test_suite\\Regression_test_suite\\xorbotoutput.txt");
			FileReader fileReader1 = new FileReader(file1);
			FileReader fileReader2 = new FileReader(file2);
			BufferedReader bufferedReader1 = new BufferedReader(fileReader1);
			BufferedReader bufferedReader2 = new BufferedReader(fileReader2);
			
			String line1, line2;
			int linecount = 1;
			while ((line1 = bufferedReader1.readLine()) != null) {
				
				line2 = bufferedReader2.readLine();
				
				if(!line2.equals(line1)) {
					
					System.out.println("Wrong response at line -: " + linecount);
					
				}
				
				linecount++;
			}
			fileReader1.close();
			fileReader2.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
}

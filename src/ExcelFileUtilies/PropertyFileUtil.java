package ExcelFileUtilies;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class PropertyFileUtil {

	public static String getValueForkey(String Key) throws IOException 
	{
		FileInputStream fis = new FileInputStream("./PropertiesFile/Environment.properties");
		
		Properties configProperties = new Properties();
		configProperties.load(fis);
		return configProperties.getProperty(Key);
		
		
	}
	
	
	
	
	
	
}

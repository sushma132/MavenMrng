package Utilities;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertyFileUtil {
	public static String getValueForkey(String key)throws Throwable
	{
		Properties configprop=new Properties();
		FileInputStream fis= new FileInputStream("D:\\mrng930batch\\ERP_Maven\\PropertyFile\\Environment.properties");
		configprop.load(fis);
		return configprop.getProperty(key);		
	}

}

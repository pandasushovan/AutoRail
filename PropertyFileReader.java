package automation.library;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;


public class PropertyFileReader {

	public static String AppConfigRead(String key) throws IOException {

		File fs = new File("./ConfigFiles/AppConfig.properties");
		FileReader fr = new FileReader(fs);
		Properties prop = new Properties();
		prop.load(fr);
		return prop.getProperty(key);

	}
	public static String locatorReader(String key) throws IOException {

		File fs = new File("./ConfigFiles/locator.properties");
		FileReader fr = new FileReader(fs);
		Properties prop = new Properties();
		prop.load(fr);
		return prop.getProperty(key);

	}

}


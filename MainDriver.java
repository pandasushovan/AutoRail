package automation.library;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;

import org.openqa.selenium.WebElement;

import automation.library.PropertyFileReader;

public class MainDriver  {
	
	public static void main(String[] args) throws IOException {
		
	// Chrome Driver Instance 
		ChromeDriver driver;
	
	// Open Browser and URL 
		System.setProperty("webdriver.chrome.driver", "./driver/chromedriver-new.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(PropertyFileReader.AppConfigRead("URL"));
	
	// log-in page: fill up user id, password and click sign-in
		driver.findElementById(PropertyFileReader.locatorReader("login_username_id"))
			.sendKeys(PropertyFileReader.AppConfigRead("Username"));
		driver.findElementById(PropertyFileReader.locatorReader("login_password_id"))
			.sendKeys(PropertyFileReader.AppConfigRead("Password"));
		// Wait for 20 second to enter captcha
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	//	driver.findElementByXPath(PropertyFileReader.locatorReader("login_button_xpath")).click();
		
	//	driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);

	//click the edit button on measurement table
		
		driver.findElementByXPath(PropertyFileReader.locatorReader("measurement_edit_button_xpath")).click();

		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);

	//Processing on Account Measurement entry
		
		// click at edit button every row

		//No.of Columns
        //List <WebElement> col = driver.findElements(By.xpath(".//*[@id=\"leftcontainer\"]/table/thead/tr/th"));
        //System.out.println("No of cols are : " +col.size()); 
        //No.of rows 
        List <WebElement> rowsMeasurementTable = driver.findElementsByXPath(PropertyFileReader.locatorReader("account_measurement_row_firstdata_xpath")); 
        System.out.println("No of rows in Measurement table are : " + rowsMeasurementTable.size());
		
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		
		
		//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
		FileInputStream fis = new FileInputStream(PropertyFileReader.AppConfigRead("excel_file_path_name"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		//call the getSheet() method of Workbook and pass the Sheet Name here.
		//In this case I have given the sheet name as “TestData”
		                //or if you use the method getSheetAt(), you can pass sheet number starting from 0. Index starts with 0.
		XSSFSheet sheet = workbook.getSheet("TestData");
		
		int excelRowCounter = 2;
		Row headerRow = sheet.createRow(excelRowCounter);
		Cell header1 = headerRow.createCell(1);
		header1.setCellValue("Serial No.");
		
		Cell header2 = headerRow.createCell(2);
		header2.setCellValue("Particulars");
		
		Cell header3 = headerRow.createCell(3);
		header3.setCellValue("N1");
		
		Cell header4 = headerRow.createCell(4);
		header4.setCellValue("N2");
		
		Cell header5 = headerRow.createCell(5);
		header5.setCellValue("N3");
		
		Cell header6 = headerRow.createCell(6);
		header6.setCellValue("Plus / Minus");
		
		Cell header7 = headerRow.createCell(7);
		header7.setCellValue("Coefficient / K");
		
		Cell header8 = headerRow.createCell(8);
		header8.setCellValue("Parameters - L");
		
		Cell header9 = headerRow.createCell(9);
		header9.setCellValue("Parameters - B");
		
		Cell header10 = headerRow.createCell(10);
		header10.setCellValue("Parameters - H");
		
		Cell header11 = headerRow.createCell(11);
		header11.setCellValue("Contents");
			
//		for (int i =1;i<rowsMeasurementTable.size();i++)
		for (int i =1; i<5; i++)

        {    
			
			System.out.println("Serial Number : " + i);
			
			

			if (i == 1) {
				
				driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				
				// click the edit row option of Measurement table
				driver.findElementByXPath("//*[@id='mbdata']/tr[" + i + "]/td[14]/button").click();

	            driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	            
	            // to fetch the web element of the modal container
	    		WebElement modalContainer = driver.findElementById("modalContainer");
	    		
            //********** reading detail line item from measurement particular table ***********//
	    		
	    		List <WebElement> rowsParticularsTable = driver.findElementsByXPath
	    				(PropertyFileReader.locatorReader("account_particular_row_firstdata_xpath")); 
	            System.out.println("No of rows in Particular table are : " + rowsParticularsTable.size());
	    		
	    		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	    		
	    			
	    		for (int j =3;j < (rowsParticularsTable.size()+3);j++) {
//	    		for (int j =3; j<5; j++) {
	    			
	    			excelRowCounter = excelRowCounter +1;
	    			
	    			String column1 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[1]/input").getAttribute("value");
	    			String column2 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[2]/input").getAttribute("value");
	    			String column3 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[3]/input").getAttribute("value");
	    			String column4 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[4]/input").getAttribute("value");
	    			String column5 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[5]/input").getAttribute("value");
	    			String column6 ="";
	    			if (driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[6]/nobr[1]/input").isSelected())
	    			{
	    			    column6 = "Plus";
	    			}
	    			if (driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[6]/nobr[2]/input").isSelected())
	    			{
	    				column6 = "Minus";
	    			}
	    			String column7 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[7]/input").getAttribute("value");
	    			String column8 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[8]/input").getAttribute("value");
	    			String column9 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[9]/input").getAttribute("value");
	    			String column10 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[10]/input").getAttribute("value");
	    			String column11 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[11]/input").getAttribute("value");
	    			System.out.println(" " + column1 +" "+ column2 +" "+ column3 +" "+ column4 +" "+ column5 +" "+ column6 +" "+ column7 +" "+ column8 +" "+ column9 +" "+ column10 +" "+ column11);
	    		// build cells of an excel row
	    			//XSSFSheet sheet = workbook.getSheetAt(0);
	    			//Now create a row number and a cell where we want to enter a value.
	    			//Here im about to write my test data in the cell B2. It reads Column B as 1 and Row 2 as 1. Column and Row values start from 0.
	    			//The below line of code will search for row number 2 and column number 2 (i.e., B) and will create a space.
	                //The createCell() method is present inside Row class.
	                Row row = sheet.createRow(excelRowCounter);
	    			Cell cell1 = row.createCell(1);
	    			//Now we need to find out the type of the value we want to enter.
	                //If it is a string, we need to set the cell type as string
	                //if it is numeric, we need to set the cell type as number
	    			cell1.setCellValue(String.valueOf(i));
	    			
	    			Cell cell2 = row.createCell(2);
	    			cell2.setCellValue(column2);
	    			
	    			Cell cell3 = row.createCell(3);
	    			cell3.setCellValue(column3);
	    			
	    			Cell cell4 = row.createCell(4);
	    			cell4.setCellValue(column4);
	    			
	    			Cell cell5 = row.createCell(5);
	    			cell5.setCellValue(column5);
	    			
	    			Cell cell6 = row.createCell(6);
	    			cell6.setCellValue(column6);
	    			
	    			Cell cell7 = row.createCell(7);
	    			cell7.setCellValue(column7);
	    			
	    			Cell cell8 = row.createCell(8);
	    			cell8.setCellValue(column8);
	    			
	    			Cell cell9 = row.createCell(9);
	    			cell9.setCellValue(column9);
	    			
	    			Cell cell10 = row.createCell(10);
	    			cell10.setCellValue(column10);
	    			
	    			Cell cell11 = row.createCell(11);
	    			cell11.setCellValue(column11);
	    			
	    		}
	    		// code to click on close modal button
	    		WebElement modalCloseButton = modalContainer.findElement
	    				(By.xpath(PropertyFileReader.locatorReader("account_particulars_close_xpath")));
	    		
	    		modalCloseButton.click();  
				
			} else{
				driver.navigate().refresh();
				driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				
				//click the edit button on measurement table
				
				driver.findElementByXPath(PropertyFileReader.locatorReader("measurement_edit_button_xpath")).click();

				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				
				// again navigate to current iteration of row
				
				driver.findElementByXPath("//*[@id='mbdata']/tr[" + i + "]/td[14]/button").click();
	
	            driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	            
	            // to fetch the web element of the modal container
	    		WebElement modalContainer = driver.findElementById("modalContainer");
	
	    		//********** reading detail line item from measurement particular table ***********//
	    		
	    		List <WebElement> rowsParticularsTable = driver.findElementsByXPath
	    				(PropertyFileReader.locatorReader("account_particular_row_firstdata_xpath")); 
	            System.out.println("No of rows in Particular table are : " + rowsParticularsTable.size());
	    		
	    		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	    		
	    			
	    		for (int j =3;j < (rowsParticularsTable.size()+3);j++) {
//	    		for (int j =3; j<5; j++) {
	    			
	    			excelRowCounter = excelRowCounter + 1;
	    			
	    			String column1 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[1]/input").getAttribute("value");
	    			String column2 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[2]/input").getAttribute("value");
	    			String column3 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[3]/input").getAttribute("value");
	    			String column4 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[4]/input").getAttribute("value");
	    			String column5 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[5]/input").getAttribute("value");
	    			String column6 ="";
	    			if (driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[6]/nobr[1]/input").isSelected())
	    			{
	    			    column6 = "Plus";
	    			}
	    			if (driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[6]/nobr[2]/input").isSelected())
	    			{
	    				column6 = "Minus";
	    			}
	    			String column7 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[7]/input").getAttribute("value");
	    			String column8 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[8]/input").getAttribute("value");
	    			String column9 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[9]/input").getAttribute("value");
	    			String column10 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[10]/input").getAttribute("value");
	    			String column11 = driver.findElementByXPath("//*[@id='datatbody']/tr[" + j + "]/td[11]/input").getAttribute("value");
	    			System.out.println(" " + column1 +" "+ column2 +" "+ column3 +" "+ column4 +" "+ column5 +" "+ column6 +" "+ column7 +" "+ column8 +" "+ column9 +" "+ column10 +" "+ column11);
	    		
	    			Row rownext = sheet.createRow(excelRowCounter);
	    			Cell cellnext1 = rownext.createCell(1);
	    			//Now we need to find out the type of the value we want to enter.
	                //If it is a string, we need to set the cell type as string
	                //if it is numeric, we need to set the cell type as number
	    			cellnext1.setCellValue(String.valueOf(i));
	    			
	    			Cell cellnext2 = rownext.createCell(2);
	    			cellnext2.setCellValue(column2);
	    			
	    			Cell cellnext3 = rownext.createCell(3);
	    			cellnext3.setCellValue(column3);
	    			
	    			Cell cellnext4 = rownext.createCell(4);
	    			cellnext4.setCellValue(column4);
	    			
	    			Cell cellnext5 = rownext.createCell(5);
	    			cellnext5.setCellValue(column5);
	    			
	    			Cell cellnext6 = rownext.createCell(6);
	    			cellnext6.setCellValue(column6);
	    			
	    			Cell cellnext7 = rownext.createCell(7);
	    			cellnext7.setCellValue(column7);
	    			
	    			Cell cellnext8 = rownext.createCell(8);
	    			cellnext8.setCellValue(column8);
	    			
	    			Cell cellnext9 = rownext.createCell(9);
	    			cellnext9.setCellValue(column9);
	    			
	    			Cell cellnext10 = rownext.createCell(10);
	    			cellnext10.setCellValue(column10);
	    			
	    			Cell cellnext11 = rownext.createCell(11);
	    			cellnext11.setCellValue(column11);
	    		
	    		}

	    		
	    		// code to click on close modal button
	    		WebElement modalCloseButton = modalContainer.findElement
	    				(By.xpath(PropertyFileReader.locatorReader("account_particulars_close_xpath")));
	    		
	    		modalCloseButton.click();  
			}
 


	// Quitting Driver Instance 
	//	driver.quit();
	
        }
	// File stream writing to excel file	
		FileOutputStream fos = new FileOutputStream(PropertyFileReader.AppConfigRead("excel_file_path_name"));
		workbook.write(fos);
		fos.close();
		System.out.println("END OF WRITING DATA IN EXCEL");
		
	}  
	
}

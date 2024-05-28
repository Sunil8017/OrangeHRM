
// call excel and function class method in driver script .
package driverFactory;

import java.io.File;


import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver  driver ; 
	String inputpath="./FileInput/DataEngine.xlsx";
	String outputpath="./FileOutput/HybridResults.xlsx";
	ExtentReports report;
	ExtentTest logger;
	String TCSheet ="MasterTestCases";
	
	//@Test
	public void startTest() throws Throwable {
		String Module_status =" ";
		String Module_new = " ";
		//create object for excel file util class
		
		ExcelFileUtil x1 = new ExcelFileUtil(inputpath);
		//iterate all test case in TCSheet
		for(int i=1;i<=x1.rowCount(TCSheet);i++)
		{
			if(x1.getCellData(TCSheet, i, 2).equalsIgnoreCase("Y"))
			{
				//store each sheet into one variable
				String TCModule =x1.getCellData(TCSheet, i, 1);
				//define path of html
				report =new ExtentReports("./target/ExtentReports/"+TCModule+FunctionLibrary.generateDate()+".html");
				logger = report.startTest("TCModule");
				
				//iterate all row in every sheet 
				for(int j=1;j<=x1.rowCount(TCModule);j++) {
					String Description = x1.getCellData(TCModule, j, 0);
					String ObjectType = x1.getCellData(TCModule, j, 1);
					String LType = x1.getCellData(TCModule, j, 2);
					String LValue = x1.getCellData(TCModule, j, 3);
					String TestData = x1.getCellData(TCModule, j, 4);
					try {
						if(ObjectType.equalsIgnoreCase("startBrowser")) {
							driver = FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("openUrl")) {
							 FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("waitForElement")) {
							 FunctionLibrary.waitForElement(LType, LValue, TestData);
							 logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("typeAction")) {
							 FunctionLibrary.typeAction(LType, LValue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("clickAction")) {
							 FunctionLibrary.clickAction(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("validateTiltle")) {
							 FunctionLibrary.validateTiltle(TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(LType, LValue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("captuserStock"))
						{
							FunctionLibrary.captuserStock(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("captureSup"))
						{
							FunctionLibrary.captureSup(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("captureCus"))
						{
							FunctionLibrary.captureCus(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(ObjectType.equalsIgnoreCase("customerTable"))
						{
							FunctionLibrary.supplierTable();
							logger.log(LogStatus.INFO, Description);
						}
						
						
						if(ObjectType.equalsIgnoreCase("closeBrowser")) {
							 FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						//write as pass into status cell in TCModule Sheet
						x1.setCellData(TCModule, j, 5, " Pass",outputpath );
						Module_status ="True";
					}catch(Throwable t) {
						System.out.println(t.getMessage());
						//write as fail into status cell in TCModule Sheet
						x1.setCellData(TCModule, j, 5, " Fail",outputpath );
						Module_status ="False";
						//take screen shot
						File screen = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						FileUtils.copyFile(screen, new File("./target/Screenshot/"+Description+FunctionLibrary.generateDate()+"invalidusername&password.png"));
					}
					if(Module_status.equalsIgnoreCase("True")) {
						x1.setCellData(TCSheet, i, 3, "Pass", outputpath);
					}
					if(Module_status.equalsIgnoreCase("False")) {
						x1.setCellData(TCSheet, i, 3, "Fail", outputpath);
					}
					report.endTest(logger);
					report.flush();
				}
				
			}else {
				//write as blocked for testcases flag to N
				
				x1.setCellData(TCSheet, i, 3, "Blocked", outputpath);
			}
		}
		
		
		
	}

	
}

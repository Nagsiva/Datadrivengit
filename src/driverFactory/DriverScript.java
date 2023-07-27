package driverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.Reporter;
import org.testng.annotations.Test;

import commonFunctions.FunctionLibrary;
import config.AppUtil;
import utilities.ExcelFileUtil;

public class DriverScript extends AppUtil {
String inputpath = "D:\\Programs\\Desktop\\DataDriven_Framework\\TestInput\\TestData.xlsx";
String outputpath = "D:\\Programs\\Desktop\\DataDriven_Framework\\TestOutput\\DataDrivenResults.xlsx";
@Test
public void startTest()throws Throwable
{
	//create object for ExcelFileutil class
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	//count no of rows in Login sheet
	int rc = xl.rowcount("Login");
	Reporter.log("No of rows in a sheet ::"+rc,true);
	for(int i=1;i<=rc;i++) 
	{
		String user = xl.getcelldata("Login", i, 0);
		String pass = xl.getcelldata("Login", i, 1);
		//call login methods
		boolean res = FunctionLibrary.verifyLogin(user, pass);
		if (res) {
			//if res is true write Login success in results cell
			xl.setcelldata("Login", i, 2, "Login success", outputpath);
			xl.setcelldata("Login", i, 3, "Pass", outputpath);
		}
		else {
			//Take screenshot for fail test
			File screen = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(screen, new File("./ScreenShot/Iterations/"+i+"_"+	"Loginpage.png"));
			xl.setcelldata("Login", i, 2, "Login Fail", outputpath);
			xl.setcelldata("Login", i, 3, "Fail", outputpath);
		}
	}
}
}

package driverFactory;

import org.openqa.selenium.support.PageFactory;
import org.testng.Reporter;
import org.testng.annotations.Test;

import commonFunctions.AddEmpPage;
import config.AppUtil1;
import utilities.ExcelFileUtil;

public class DriverScriptPOM  extends AppUtil1
{
	String inputpath = "D:\\Programs\\Desktop\\DataDriven_Framework\\TestInput\\Employee.xlsx";
	String outputpath = "D:\\Programs\\Desktop\\DataDriven_Framework\\TestOutput\\POMResilts.xlsx";
	@Test
	public void startTest()throws Throwable
	{
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);
		int rc = xl.rowcount("EmpData");
		Reporter.log("No of rows are::"+rc,true);
		for(int i=1;i<=rc;i++)
		{
			String fname = xl.getcelldata("EmpData", i, 0);
			String mname = xl.getcelldata("EmpData", i, 1);
			String lname = xl.getcelldata("EmpData", i, 2);
			AddEmpPage Emp = PageFactory.initElements(driver, AddEmpPage.class);
			boolean res = Emp.verifyEmp(fname, mname, lname);
			if (res) {
				xl.setcelldata("EmpData", i, 3, "Pass", outputpath);
			}
			else {
				xl.setcelldata("EmpData",i, 3, "Fail", outputpath);
			}
		}
	}

}

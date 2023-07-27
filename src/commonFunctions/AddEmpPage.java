package commonFunctions;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.testng.Reporter;

public class AddEmpPage {
	WebDriver driver;
	public AddEmpPage(WebDriver driver)
	{
		this.driver = driver;
	}
	//Define Repository
	@FindBy(xpath = "//b[normalize-space()='PIM']")
	WebElement clickPim;
	@FindBy(name ="btnAdd")
	WebElement clickAddBtn;
	@FindBy(name = "firstName")
	WebElement EnterFName;
	@FindBy(name = "middleName")
	WebElement EnterMName;
	@FindBy(name = "lastName")
	WebElement EnterLName;
	@FindBy(name = "employeeId")
	WebElement BeforeEid;
	@FindBy(id = "btnSave")
	WebElement ClickSaveBtn;
	@FindBy(name = "personal[txtEmployeeId]")
	WebElement AfterEid;
	public boolean verifyEmp(String FirstName,String MiddleName,String LastName)
	{
		Actions ac = new Actions(driver);
		ac.moveToElement(this.clickPim).click().perform();
		ac.moveToElement(this.clickAddBtn).click().perform();
		this.EnterFName.sendKeys(FirstName);
		this.EnterMName.sendKeys(MiddleName);
		this.EnterLName.sendKeys(LastName);
		String ExpectedEID = this.BeforeEid.getAttribute("value");
		this.ClickSaveBtn.click();
		String ActualEID = this.AfterEid.getAttribute("value");
		if(ExpectedEID.equals(ActualEID))
		{
			Reporter.log("Add Emp Success::"+ExpectedEID+"     "+ActualEID,true);
			return true;
		}
		else
		{
			Reporter.log("Add Emp Fail::"+ExpectedEID+"     "+ActualEID,true);
			return false;
		}
	}
}

package DriverFactory;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import CommonFunLibrary.ERP_Functions;
import Utilities.ExcelFileUtil;


public class DataDrivenFactory {
	WebDriver driver;
	//Creating reference object for Function library
	ERP_Functions erp=new ERP_Functions();
	
@BeforeTest
	public void launchapp()
	{
		String app=erp.launchurl("http://webapp.qedge.com");
		System.out.println(app);
		//Calling Login
		String login=erp.verifyLogin("admin", "master");
		System.out.println(login);
	}
@org.testng.annotations.Test
	public void suppliercreation() throws Throwable {
	//To Call excel Util methods
	ExcelFileUtil xl=new ExcelFileUtil();
		//Count no. of rows in a sheet
	int rc=xl.rowCount("Suppliers");
	//Count no. of columns in a sheet
	int cc=xl.colCount("Suppliers");
	//Reporter -- To print the report message in to html reports
	Reporter.log("no of rows are::"+rc+" "+"no of column are::"+cc,true);
	for(int i=1; i<=rc; i++)
	{
		String sname=xl.getData("Suppliers", i, 0);
		String address=xl.getData("Suppliers", i, 1);
		String city=xl.getData("Suppliers", i, 2);
		String country=xl.getData("Suppliers", i, 3);
		String cperson=xl.getData("Suppliers", i, 4);
		String pnumber=xl.getData("Suppliers", i, 5);
		String mail=xl.getData("Suppliers", i, 6);
		String mnumber=xl.getData("Suppliers", i, 7);
		String note=xl.getData("Suppliers", i, 8);
		String result=erp.verifySupplier(sname, address, city, country, cperson, pnumber, mail, mnumber, note);
		xl.setCellData("Suppliers", i, 9, result);
		
	}
}
@AfterTest
	public void logout() throws Throwable
	{
	String logoutapp=erp.verifyLogout();
	System.out.println(logoutapp);
	}
}


	

package com.market.ashent.PFGOnline;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.util.framework.CommonPFG;
import com.util.framework.ExcelFunctions;
import com.util.framework.RandomAction;
import com.util.framework.SendMailSSL;

import junit.framework.Assert;

public class TestPFG extends CommonPFG {

	// public testPFG(WebDriver driver) {
	// super(driver);
	// }
	public static int rowIndex;
	public static String projectPath = System.getProperty("user.dir");
	public static String inputFile = "C:\\Users\\Edge\\Desktop\\ExportEngineInput.xlsx";
	// "C:\\Users\\my\\Downloads\\ExportEngineInput.xlsx";
	// projectPath + "\\config\\ExportEngineInput.xlsx";
	public static SimpleDateFormat sdf = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
	public static String reportFile = "C:\\Users\\Edge\\Desktop\\Reports\\PFG_OG_report\\ExportSummary_PFG_"
			+ new Date().toString().replace(":", "").replace(" ", "") + ".xlsx";
	// "C:\\Users\\my\\Downloads\\Reports\\USF_OG_report\\ExportSummary_USF_"+ new
	// Date().toString().replace(":", "").replace(" ", "") + ".xlsx";
	// for Edge -
	// "C:\Users\Edge\Desktop\\Reports\\SyscoOG_report\\ExportSummary_Sysco_" +
	// PageAction.getDate().toString().replace(" ", "_");
	// + new Date().toString().replace(":", "").replace(" ", "") + ".xlsx";
	// projectPath+ "\\Output_Summary\\ExportSummary_Sysco_" + new
	// Date().toString().replace(":", "").replace(" ", "")+".xlsx";
	public static int acno;
	public static XSSFWorkbook exportworkbook;
	public static XSSFSheet inputsheet;
	public static int AcColStatus, AcColdetail;
	public static FileOutputStream out;
	public static int totalNoOfRows;
	public static String folderDate;
	public static String currList = "";
	public static String emailMessageExport = "";
	public static String path = System.getProperty("user.home") + "\\Downloads\\chromedriver_win32\\chromedriver.exe";
	// System.getProperty("user.home")+"\\Downloads\\chromedriver_win32_new\\chromedriver.exe";
	public static String project = "PFG";
	private final static Logger logger = Logger.getLogger(TestPFG.class);
	public static CommonPFG testPFG = new TestPFG();

	@BeforeTest
	public static void beforeData() throws Exception {
		System.out.println("before data.");
	}

	@AfterTest
	public static void closeResources() throws SQLException, IOException {
		System.out.println("Closing the resources!");

		if (out != null) {
			System.out.println("Closing file output stream object!");
			out.close();
		}
		if (driver != null) {
			System.out.println("Closing the browser!");
			// TestCases.driver.close();
			driver.quit();
		}

		if (exportworkbook != null) {
			exportworkbook.close();
		}
	}

	@BeforeMethod
	public static void setUp() throws IOException {
		// to get the browser on which the UI test has to be performed.
		System.out.println("***********StartTest*********");
		RandomAction.deleteFiles("C:\\Users\\Edge\\Downloads");
		driver = RandomAction.openBrowser("Chrome", path);
		System.out.println("Invoked browser .. ");
	}

	@AfterMethod
	public static void writeExcel() throws IOException {
		System.out.println("Running Excel write method!");
		out = new FileOutputStream(new File(reportFile));
		exportworkbook.write(out);
		acno++;
		try {
			driver.close();
		} catch (Exception e) {
			System.out.println("already closed");
		}
	}

	@DataProvider(name = "testData")
	public static Object[][] testData() throws IOException {
		exportworkbook = ExcelFunctions.openFile(inputFile);
		System.out.println("Test data read.");
		inputsheet = exportworkbook.getSheet(project);
		AcColStatus = ExcelFunctions.getColumnNumber("Export Status", inputsheet);
		AcColdetail = ExcelFunctions.getColumnNumber("Detailed Export Status", inputsheet);
		System.out.println("Inside Dataprovider. Creating the Object Array to store test data inputs.");
		Object[][] td = null;
		try {
			// Get TestCase sheet data
			int totalNoOfCols = inputsheet.getRow(inputsheet.getFirstRowNum()).getPhysicalNumberOfCells();
			totalNoOfRows = inputsheet.getLastRowNum();
			System.out.println(totalNoOfRows + " Accounts and Columns are: " + totalNoOfCols);
			td = new String[totalNoOfRows][totalNoOfCols];
			for (int i = 1; i <= totalNoOfRows; i++) {
				for (int j = 0; j < totalNoOfCols; j++) {
					td[i - 1][j] = ExcelFunctions.getCellValue(inputsheet, i, j);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("Test Cases captured in the Object Array. Exiting dataprovider.");
		return td;
	}

	@Test(dataProvider = "testData")
	public static void Export_Mail_OG(String active, String accountID, String purveyor, String restaurant_name,
			String username, String password, String listname, String exportstatus, String detailedstatus) {
		Boolean result;
		System.out.println("Inside OG Export : Started exporting OG for different accounts");
		XSSFCell cell1, cell2;
		TestPFG.rowIndex = Math.floorMod(TestPFG.acno, TestPFG.totalNoOfRows) + 1;

		System.out.println("Test Case test #" + TestPFG.rowIndex);

		cell1 = TestPFG.exportworkbook.getSheet(project).getRow(TestPFG.rowIndex).createCell(TestPFG.AcColStatus);
		cell1.setCellValue("");

		cell2 = TestPFG.exportworkbook.getSheet(project).getRow(TestPFG.rowIndex).createCell(TestPFG.AcColdetail);
		cell2.setCellValue("");

		exportstatus = cell1.getStringCellValue();
		detailedstatus = cell2.getStringCellValue();

		try {
			if (active.equalsIgnoreCase("Yes")) {
				// if list is not empty
				System.out.println(restaurant_name + " for purveryor " + purveyor + " is Active !!");
				if (listname != null) {
					result = testPFG.startPFG(listname.trim(), username.trim(), password.trim());
					if (result.equals(true)) {
						try {
							Thread.sleep(5000);
							SendMailSSL.sendMailActionXls(purveyor.trim(), restaurant_name.trim());
							emailMessageExport = "Pass";
							exportstatus = "Pass";
							detailedstatus = "OG exported succesfully";
						} catch (Exception e) {
							exportstatus = "Failed";
							detailedstatus = "Error : OG downloaded but failed during mail trigger";
						}
					} else {
						emailMessageExport = "Failed";
						exportstatus = "Failed";
						detailedstatus = "Some technical issue ocurred during export";
					}
				} else { // default OG
					// result = startSysco(driver, username.trim(), password.trim());
					exportstatus = "Failed";
					detailedstatus = "Error : Please provide valid List name";
				}
			} else {
				System.out.println(restaurant_name + " for purveryor " + purveyor + " is not Active !!");
				exportstatus = "Not Active";
			}
			cell1.setCellValue(exportstatus);
			cell2.setCellValue(detailedstatus);

			System.out.println("Exiting test method");

		} catch (Exception e) {
			e.printStackTrace();
			exportstatus = "Failed";
			detailedstatus = "Some technical issue ocurred during export";
			cell1.setCellValue(exportstatus);
			cell2.setCellValue(detailedstatus);
			System.out.println("Technical issue occured during export for restaurant - " + restaurant_name);
			Assert.assertTrue(false);
		}
		System.out.println(emailMessageExport.trim());
	}

	////////////////////////////////////////////////
	@AfterClass
	public static void sendMail() {
		try {
			String emailMsg = "Daily " + project + " OG Export Status: " + RandomAction.getDate();

			SendMailSSL.sendReport(emailMsg, reportFile);
			System.out.println("Email Sent with Attachment");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	////////////////////////////////////////////////////

}

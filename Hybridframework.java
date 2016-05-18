import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class Hybridframework {
	// Define global variables
	String vXLPath, vXLTC, vXLTS, vXLTD, vXLEM;
	String vXLTSResPath, vXLTCResPath, vXLTDResPath;
	int xTCRows, xTCCols, xTSRows, xTSCols, xTDRows, xTDCols, xEMCols, xEMRows;
	String[][] xTCData, xTSData, xTDData, xEMData;
	WebDriver Driver;
	String vKW, vXP, vData;
	String vResult, vError, vTCResult;

	@Test
	public void driverTest() throws Exception {

		// Define the webdriver
		Driver = new FirefoxDriver();
		Driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		Driver.get("http://test.atomic77.in/TWAv1/1/");

		
		vXLPath = "C:\\Users\\Renjusha\\workspace\\Renjusha_Hybrid_Admin_Trainer.xlsx";
		vXLTSResPath = "C:\\TestResults\\TernWave_TSRes_";
		vXLTCResPath = "C:\\TestResults\\TernWave_TCRes_";
		vXLTDResPath = "C:\\Users\\Renjusha\\workspace\\Renjusha_Hybrid_Admin_Trainer\\TestSteps_Res.xls";
		vXLTDResPath = "C:\\Users\\Renjusha\\workspace\\Renjusha_Hybrid_Admin_Trainer\\ElementMap_Res.xls";
		vXLTC = "TestCase";
		vXLTS = "TestSteps";
		vXLTD = "TestData";
		vXLEM = "ElementMap";

		// Read this excel
		xTCData = readXL(vXLPath, vXLTC); // returns TestCases Data
		xTSData = readXL(vXLPath, vXLTS); // returns TestSteps Data
		xTDData = readXL(vXLPath, vXLTD); // returns TestData Data
		xEMData = readXL(vXLPath, vXLEM); // returns ElementMap Data

		// Get the number of rows and columns in both the sheets
		xTCRows = xTCData.length;
		xTCCols = xTCData[0].length;

		xTSRows = xTSData.length;
		xTSCols = xTSData[0].length;

		xTDRows = xTDData.length;
		xTDCols = xTDData[0].length;

		xEMRows = xEMData.length;
		xEMCols = xEMData[0].length;

		// Run the KDF main code for different sets of Test Data
		long startTime, endTime, totalTime;
		for (int k = 1; k < 2 ; k++) {
			if (xTDData[k][1].equals("Y")) { // TestData to be executed??
				startTime = System.currentTimeMillis();
				System.out.println("startTime "+startTime);
				// The main code of execution is over here
				for (int i = 1; i < xTCRows; i++) { // Go through each row in TC
					if (xTCData[i][2].equals("Y")) { // Verify if TC is ready
														// for run
						System.out.println("Test Case ID to RUN: " + xTCData[i][0]);
						vTCResult = "Pass";
						for (int j = 1; j < xTSRows; j++) { // Go through each
															// row in TS
							if (xTCData[i][0].equals(xTSData[j][0])) { // Do the
																		// Test
																		// Case
																		// ID's
																		// match
								vKW = xTSData[j][3];
								//vXP = xTSData[j][4];
								System.out.println("Value of k = ="+i);
								vXP = getElementIdentifier(xTSData[j][4], i);
								vData = getTestData(xTSData[j][5], i);
								System.out.println("vXP is " + vXP);
								System.out.println("vData is " + vData);
								vResult = "Pass";
								vError = "No Error";

								try {
									executeKW(vKW, vXP, vData);
								} catch (Exception e) {
									vResult = "Fail";
									vError = "Error: " + e;
								}
								xTSData[j][6] = vResult;
								xTSData[j][7] = vError;
								// Set the Test Case to Fail as soon as even 1
								// step fails.
								if (vResult.equals("Fail")) {
									vTCResult = "Fail";
								}
								xTCData[i][3] = vTCResult;
							}
						}

					} else {
						System.out.println("Test Case Not ready for execution.");
					}

				}
				writeXL(vXLTSResPath + xTDData[k][0] + ".xls", "TestSteps", xTSData);
				writeXL(vXLTCResPath + xTDData[k][0] + ".xls", "TestCases", xTCData);
				endTime = System.currentTimeMillis();
				totalTime = (endTime - startTime) / 1000;
				System.out.println("To execute " + xTDData[k][0] + " total time taken is " + totalTime + " seconds.");
			}
		}
		Driver.quit();
	}

	// Custom Keyword Functions

	public String getTestData(String fData, int fRowNumber) {
		// Purpose : Gets the actual value of the test data variable
		// Inputs : Data variable name and the row number
		// Output : Data value
   System.out.println(fData);
		for (int a = 0; a < xTDCols; a++) {
			if (fData.equals(xTDData[0][a])) {
				return xTDData[fRowNumber][a];
			}
		}
		return fData;
		
	}

	public String getElementIdentifier(String fXP, int fRowNumber) {
		// Purpose : Gets the actual xpath of the test step
		// Inputs : Element name and the row number
		// Output : xpath value

		for (int a = 1; a < xEMRows; a++) {
			if (fXP.equals(xEMData[a][2])) {
				return xEMData[a][3];
			}
		}
		return fXP;
	}

	public void executeKW(String vKW, String vXP, String vData) throws Exception {
		// Purpose : Calls the corresponding function to execute the Test Case
		// (Keyword)
		// Inputs : KW, xP, Data
		// Output : No output
		// Chose the Keyword and call the corresponding function
		switch (vKW) {
		case "getURL":
			System.out.println("Running: " + vKW);
			getURL(vData);
			break;
		case "clickLink":
			System.out.println("Running: " + vKW);
			clickLink(vXP);
			break;
		case "typeText":
			System.out.println("Running: " + vKW);
			typeText(vXP, vData);
			break;
		case "clickElement":
			System.out.println("Running: " + vKW);
			clickElement(vXP);
			break;
		case "verifyText":

			System.out.println("Running: " + vKW);
			vResult = verifyText(vXP, vData);
			if (vResult.equals("Fail")) {
				vError = "Verification Failed";
			}

			System.out.println("Result is " + vResult);
			break;
		case "hitEnter":
			System.out.println("Running: " + vKW);
			hitEnter(vXP);
			break;
		case "isPresent":

			System.out.println("Running: " + vKW);
			vResult = isPresent(vXP);
			if (vResult.equals("Fail")) {
				vError = "Verification Failed";
			}

			System.out.println("Result is " + vResult);
			break;
		case "selectDropdown":

			System.out.println("Running: " + vKW);
			selectDropdown(vXP, vData);
			break;
		case "uploadImage":

			System.out.println("Running: " + vKW);
			uploadImage(vXP, vData);
			break;
			

		default:
			System.out.println("ALERT : Keyword is missing. " + vKW);
			break;
		}

	}

	public void typeText(String fXP, String fText) {
		// Purpose : It takes a webdriver and enters a text into it.
		// Inputs : Where to type (Xpath), What to type (text)
		// Output : No output
		Driver.findElement(By.xpath(fXP)).clear();
		Driver.findElement(By.xpath(fXP)).sendKeys(fText);

	}
	
	public void selectDropdown(String fXP, String fValue) {
		// Purpose : selects the dropdown value for specified xpath
		// Inputs : which (Xpath), What to value
		// Output : No output
		Select select = new Select(Driver.findElement(By.xpath(fXP)));
		//select.deselectAll();
		select.selectByVisibleText(fValue);

	}
	
	public void uploadImage(String fXP, String fImagePath) {
		// Purpose : selects the dropdown value for specified xpath
		// Inputs : which (Xpath), What to value
		// Output : No output
		Driver.findElement(By.xpath(fXP)).sendKeys(fImagePath);

	}

	public void clickLink(String fLinkText) {
		// Purpose : It clicks on a link
		// Inputs : Text of the link to click on.
		// Output : No output
		Driver.findElement(By.linkText(fLinkText)).click();
	}

	public void getURL(String fURL) {
		// Purpose : Navigate to a URL in our Webdriver
		// Inputs : URL
		// Output : No output
		Driver.get(fURL);
	}
	
	

	public void clickElement(String fXP) {
		// Purpose : It clicks any element
		// Inputs : xPath of the element to click on.
		// Output : No output
		Driver.findElement(By.xpath(fXP)).click();
	}

	public String isPresent(String fxp) {
		// Purpose : Check whether an element is present on page
		// Inputs : xPath of the element to check the presence.
		// Output : String
		// Driver.findElement(By.xpath(fxp)).isDisplayed() ;
		if (Driver.findElement(By.xpath(fxp)) != null) {
			System.out.println("Element is Present");
			return "Pass";
		} else {
			System.out.println("Element is Absent");
			return "Fail";
		}
	}

	public void hitEnter(String fXP) {
		// Purpose : Hits an enter over an element
		// Inputs : xPath of the element to hit enter on
		// Output : No output
		Driver.findElement(By.xpath(fXP)).sendKeys(Keys.ENTER);
	}

	public String verifyText(String fXP, String fText) {
		// Purpose : Verifies if a specific text is present on that element
		// Inputs : xPath of the element and the text to verify
		// Output : Pass or a Fail
		String fTemp;

		fTemp = Driver.findElement(By.xpath(fXP)).getText();
		if (fTemp.equals(fText)) {
			return "Pass";
		} else {
			return "Fail";
		}
	}
	
	public String verify(String fXP, String value) {
		// Purpose : Verifies if a specific text is present on that element
		// Inputs : xPath of the element and the text to verify
		// Output : Pass or a Fail
		String fTemp;

		fTemp = Driver.findElement(By.xpath(fXP)).getText();
		if (fTemp.equals(value)) {
			return "Pass";
		} else {
			return "Fail";
		}
	}
	

	// Read and Write from an excel
	// Method to write into an XL
	public void writeXL(String sPath, String iSheet, String[][] xData) throws Exception {

		File outFile = new File(sPath);
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet osheet = wb.createSheet(iSheet);
		int xR_TS = xData.length;
		int xC_TS = xData[0].length;
		for (int myrow = 0; myrow < xR_TS; myrow++) {
			XSSFRow row = osheet.createRow(myrow);
			for (int mycol = 0; mycol < xC_TS; mycol++) {
				XSSFCell cell = row.createCell(mycol);
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(xData[myrow][mycol]);
			}
			FileOutputStream fOut = new FileOutputStream(outFile);
			wb.write(fOut);
			fOut.flush();
			fOut.close();
		}
	}

	// Method to read XL
	public String[][] readXL(String sPath, String iSheet) throws Exception {
		// Purpose : Read data from an excel sheet
		// I/P : Path and Sheet name.
		// O/P : 2D Array containing the xl sheet data

		String[][] xData;
		int xRows, xCols;

		File myxl = new File(sPath);
		FileInputStream myStream = new FileInputStream(myxl);
		XSSFWorkbook myWB = new XSSFWorkbook(myStream);
		XSSFSheet mySheet = myWB.getSheet(iSheet);
		System.out.println("Sheet == "+ mySheet);
		xRows = mySheet.getLastRowNum() + 1;
		xCols = mySheet.getRow(0).getLastCellNum();
		System.out.println("Rows for this sheet == "+xRows);
		System.out.println("Columns for this sheet == "+xCols);
		xData = new String[xRows][xCols];
		for (int i = 0; i < xRows; i++) {
			XSSFRow row = mySheet.getRow(i);
			for (int j = 0; j < xCols; j++) {
				XSSFCell cell = row.getCell(j);
				String value = "-";
				if (cell != null) {
					value = cellToString(cell);
				}
				xData[i][j] = value;
				// System.out.println(value);
				// System.out.print("--");
			}
		}
		return xData;
	}

	// Change cell type
	public static String cellToString(XSSFCell cell) {
		// This function will convert an object of type excel cell to a string
		// value
		int type = cell.getCellType();
		Object result;
		switch (type) {
		case XSSFCell.CELL_TYPE_NUMERIC: // 0
			result = cell.getNumericCellValue();
			break;
		case XSSFCell.CELL_TYPE_STRING: // 1
			result = cell.getStringCellValue();
			break;
		case XSSFCell.CELL_TYPE_FORMULA: // 2
			throw new RuntimeException("We can't evaluate formulas in Java");
		case XSSFCell.CELL_TYPE_BLANK: // 3
			result = "%";
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN: // 4
			result = cell.getBooleanCellValue();
			break;
		case XSSFCell.CELL_TYPE_ERROR: // 5
			throw new RuntimeException("This cell has an error");
		default:
			throw new RuntimeException("We don't support this cell type: " + type);
		}
		return result.toString();
	}

}

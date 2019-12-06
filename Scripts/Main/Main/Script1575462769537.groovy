import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable

import java.awt.image.IndexColorModel
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import javax.swing.colorchooser.CenterLayout

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.lang.String

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileNotFoundException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//setup input excel sheet
def info = ['UserName' : '',
			'Password' : '']


def data = TestDataFactory.findTestData("LoginData")

FileInputStream fls = new FileInputStream('C:\\TestInput\\UserLoginFile.xlsx')

XSSFWorkbook myWorkbook = new XSSFWorkbook(fls)
XSSFSheet sheetin = myWorkbook.getSheet("LoginData")
 
rowCnt = 0
rowCnt = sheetin.getLastRowNum()


//Set Date and Time variables
String todaysDate = CustomKeywords.'packCommonFunctions.dateFunctions.getDateToday'()
String todayTimeStamp = CustomKeywords.'packCommonFunctions.dateFunctions.getNewtimestamd'()

//Setup the Excel Sheet
row = 0
//rowCount = 3

FileInputStream file = new FileInputStream (new File("C:\\TestOutput\\SecurityValidation.xlsx"))
XSSFWorkbook workbook = new XSSFWorkbook(file);
XSSFSheet sheet = workbook.getSheetAt(0);
XSSFRow rowi = sheet.createRow(row)
//XSSFCell celli = rowi.createCell(0)
//XSSFColor MyBlue = new XSSFColor(Color.BLUE)

//XSSFCellStyle style = workbook.createCellStyle()
//style = setborder(style)

CellStyle backgroundStyle1 = workbook.createCellStyle()
backgroundStyle1.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex())
backgroundStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND)
backgroundStyle1.setAlignment(HorizontalAlignment.CENTER)
//backgroundStyle1.setBorderBottom(CellStyle.THIN)

Font font = workbook.createFont()
font.setFontHeightInPoints((short) 12)
font.setFontName("Arial")
//font.setBoldweight(Font.BOLDWEIGHT_BOLD)
font.setColor(IndexedColors.WHITE.getIndex())
backgroundStyle1.setFont(font)

CellStyle backgroundStyle2 = workbook.createCellStyle()
Font font1 = workbook.createFont()
font1.setColor(IndexedColors.DARK_RED.getIndex())
//font1.setBoldweight(Font.BOLDWEIGHT_BOLD)
font1.setItalic(true)
backgroundStyle2.setFont(font1)

CellStyle backgroundStyle3 = workbook.createCellStyle()
backgroundStyle3.setAlignment(HorizontalAlignment.CENTER)
Font font2 = workbook.createFont()
font2.setColor(IndexedColors.GREEN.getIndex())
//font2.setBoldweight(Font.BOLDWEIGHT_BOLD)
font2.setItalic(false)
backgroundStyle3.setFont(font2)

/*XSSFRow rowHeader = sheet.createRow(row)
 sheet.addMergedRegion(new CellRangeAddress(1,1,1,3))
 XSSFCell cellHeader = rowHeader.createCell(0)
 cellHeader.setCellValue("Smoke Test Records")
 cellHeader.setCellStyle(backgroundStyle1)
 
 
 row = row + 1*/
 XSSFCell cell = rowi.createCell(1)
 cell.setCellValue("8412 Slate")
 cell.setCellStyle(backgroundStyle1)
 XSSFCell cella = rowi.createCell(2)
 cella.setCellValue("Applicant Forms")
 cella.setCellStyle(backgroundStyle1)
 XSSFCell cellb = rowi.createCell(3)
 cellb.setCellValue("Applicants")
 cellb.setCellStyle(backgroundStyle1)
 XSSFCell cellc = rowi.createCell(4)
 cellc.setCellValue("Assets")
 cellc.setCellStyle(backgroundStyle1)
 XSSFCell celld = rowi.createCell(5)
 celld.setCellValue("Billets")
 celld.setCellStyle(backgroundStyle1)
 XSSFCell celle = rowi.createCell(6)
 celle.setCellValue("Career Recruiter Package")
 celle.setCellStyle(backgroundStyle1)
 XSSFCell cellf = rowi.createCell(7)
 cellf.setCellValue("Colleges")
 cellf.setCellStyle(backgroundStyle1)
 XSSFCell cellg = rowi.createCell(8)
 cellg.setCellValue("Couse Catalog")
 cellg.setCellStyle(backgroundStyle1)
 XSSFCell cellh = rowi.createCell(9)
 cellh.setCellValue("Help Instructions")
 cellh.setCellStyle(backgroundStyle1)
 XSSFCell celli = rowi.createCell(10)
 celli.setCellValue("High/Comm Collages")
 celli.setCellStyle(backgroundStyle1)
 XSSFCell cellj = rowi.createCell(11)
 cellj.setCellValue("MEPS")
 cellj.setCellStyle(backgroundStyle1)
 XSSFCell cellk = rowi.createCell(12)
 cellk.setCellValue("Personnel")
 cellk.setCellStyle(backgroundStyle1)
 XSSFCell celll = rowi.createCell(13)
 celll.setCellValue("Scheduled Couses")
 celll.setCellStyle(backgroundStyle1)
 XSSFCell cellm = rowi.createCell(14)
 cellm.setCellValue("Surveys")
 cellm.setCellStyle(backgroundStyle1)
 XSSFCell celln = rowi.createCell(15)
 celln.setCellValue("System Health Monitor")
 celln.setCellStyle(backgroundStyle1)
 XSSFCell cello = rowi.createCell(16)
 cello.setCellValue("System Logs")
 cello.setCellStyle(backgroundStyle1)
 XSSFCell cellp = rowi.createCell(17)
 cellp.setCellValue("Units")
 cellp.setCellStyle(backgroundStyle1)
 XSSFCell cellq = rowi.createCell(18)
 cellq.setCellValue("Users")
 cellq.setCellStyle(backgroundStyle1)
 XSSFCell cellr = rowi.createCell(19)
 cellr.setCellValue("Zip Codes")
 cellr.setCellStyle(backgroundStyle1)

 
 
 while (rowCnt > row) {
		 
	 row = row + 1
	 println(row)
	 info.userName = data.getValue("UserName", row)
	 info.passWord = data.getValue("Password", row)
	 
	 XSSFRow rowa = sheet.createRow(row)
	 XSSFCell cell0 = rowa.createCell(0)
	 cell0.setCellValue(info.userName)
	 println(info.userName)
	
	
	
	WebUI.openBrowser('')
	
	//CustomKeywords.'packCommonFunctions.loginAppian.GetLogin'()
	
	WebUI.maximizeWindow()
	
	WebUI.navigateToUrl('https://mcriss-cds.appiancloud.com/suite/')
	WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Username_un'))
	WebUI.setText(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Username_un'), info.userName)
	WebUI.setText(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Password_pw'), info.passWord)
	
	WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Forgot your password_btn primary'))
	
	WebUI.click(findTestObject('Object Repository/Page_Records/div_Records'))
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/p_8412 Slates'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("8412 Slate exsists")
		XSSFCell cell1 = rowa.createCell(1)
		cell1.setCellValue("X")
		cell1.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("8412 Slate does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Applicant Forms'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Applicant Forms exsists")
		XSSFCell cell2 = rowa.createCell(2)
		cell2.setCellValue("X")
		cell2.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Applicant Forms does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Applicants'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Applicant exsists")
		XSSFCell cell3 = rowa.createCell(3)
		cell3.setCellValue("X")
		cell3.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Applicant does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Assets'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Assets exsists")
		XSSFCell cell4 = rowa.createCell(4)
		cell4.setCellValue("X")
		cell4.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Assets does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Billets'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Billets exsists")
		XSSFCell cell5 = rowa.createCell(5)
		cell5.setCellValue("X")
		cell5.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Billets does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Career Recruiter Packages'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("CRP exsists")
		XSSFCell cell6 = rowa.createCell(6)
		cell6.setCellValue("X")
		cell6.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("CRP does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Colleges'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Collages exsists")
		XSSFCell cell7 = rowa.createCell(7)
		cell7.setCellValue("X")
		cell7.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Collages does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Course Catalog'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Course Catalog exsists")
		XSSFCell cell8 = rowa.createCell(8)
		cell8.setCellValue("X")
		cell8.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Course Catalog does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Help Instructions'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Help Instructions exsists")
		XSSFCell cell9 = rowa.createCell(9)
		cell9.setCellValue("X")
		cell9.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Help Instruction does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_High SchoolCommunity Colleges'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("High Schools and Community Collages exsists")
		XSSFCell cell10 = rowa.createCell(10)
		cell10.setCellValue("X")
		cell10.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("High Schools and Community Collages does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_MEPS'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("MEPS exsists")
		XSSFCell cell11 = rowa.createCell(11)
		cell11.setCellValue("X")
		cell11.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("MEPS does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Personnel'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Personnel exsists")
		XSSFCell cell12 = rowa.createCell(12)
		cell12.setCellValue("X")
		cell12.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Personnel does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Scheduled Courses'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Scheduled Couses exsists")
		XSSFCell cell13 = rowa.createCell(13)
		cell13.setCellValue("X")
		cell13.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Scheduled Couses does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Surveys'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Surveys exsists")
		XSSFCell cell14 = rowa.createCell(14)
		cell14.setCellValue("X")
		cell14.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Survey does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_System Health Monitor'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("System Health Monitor exsists")
		XSSFCell cell15 = rowa.createCell(15)
		cell15.setCellValue("X")
		cell15.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("System Health Monitor does not Exsist")
	}
	
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_System logs'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("System Logs exsists")
		XSSFCell cell16 = rowa.createCell(16)
		cell16.setCellValue("X")
		cell16.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("System Logs does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Units'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Units exsists")
		XSSFCell cell17 = rowa.createCell(17)
		cell17.setCellValue("X")
		cell17.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Units does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/p_Users'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Users exsists")
		XSSFCell cell18 = rowa.createCell(18)
		cell18.setCellValue("X")
		cell18.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Users does not Exsist")
	}
	
	if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Records/span_Zip Codes'), 5, FailureHandling.OPTIONAL)) {
		System.out.println("Zip Codes exsists")
		XSSFCell cell19 = rowa.createCell(19)
		cell19.setCellValue("X")
		cell19.setCellStyle(backgroundStyle3)
	} else
	{	System.out.println("Users does not Exsist")
	}
	
	
	
	WebUI.closeBrowser()

}

//Close file and write results
file.close();
FileOutputStream outFile =new FileOutputStream(new File("C:\\TestOutput\\SecurityValidation_" + todayTimeStamp + ".xlsx"));
workbook.write(outFile);
outFile.close();



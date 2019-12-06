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
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import java.awt.image.IndexColorModel as IndexColorModel
import java.io.FileInputStream as FileInputStream
import java.io.FileNotFoundException as FileNotFoundException
import java.io.IOException as IOException
import java.util.Date as Date
import javax.swing.colorchooser.CenterLayout as CenterLayout
import org.apache.poi.hssf.usermodel.HSSFWorkbook as HSSFWorkbook
import org.apache.poi.sl.usermodel.Sheet as Sheet
import org.apache.poi.ss.usermodel.Cell as Cell
import org.apache.poi.ss.usermodel.Row as Row
import org.apache.poi.ss.usermodel.Workbook as Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory as WorkbookFactory
import org.apache.poi.hssf.util.HSSFColor as HSSFColor
import org.apache.poi.ss.usermodel.CellStyle as CellStyle
import org.apache.poi.ss.usermodel.FillPatternType as FillPatternType
import org.apache.poi.ss.usermodel.Font as Font
import org.apache.poi.ss.usermodel.HorizontalAlignment as HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors as IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell as XSSFCell
import org.apache.poi.xssf.usermodel.XSSFFont as XSSFFont
import org.apache.poi.xssf.usermodel.XSSFColor as XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRow as XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.lang.String as String
import java.io.FileOutputStream as FileOutputStream
import org.openqa.selenium.Keys as Keys

//setup input excel sheet
def info = [('UserName') : '', 
			('Password') : '']
def data = TestDataFactory.findTestData('LoginDataPersonnel')

FileInputStream fls = new FileInputStream('C:\\TestInput\\UserLoginFilePersonnel.xlsx')
XSSFWorkbook myWorkbook = new XSSFWorkbook(fls)
XSSFSheet sheetin = myWorkbook.getSheet('LoginData')
 
rowCnt = 0
rowCnt = sheetin.getLastRowNum()

//Set Date and Time variables
String todaysDate = CustomKeywords.'packCommonFunctions.dateFunctions.getDateToday'()
String todayTimeStamp = CustomKeywords.'packCommonFunctions.dateFunctions.getNewtimestamd'()

//Setup the Excel Sheet
row = 0
rowCount = 1

FileInputStream file = new FileInputStream(new File('C:\\TestOutput\\SecurityValidation.xlsx'))
XSSFWorkbook workbook = new XSSFWorkbook(file)
XSSFSheet sheet = workbook.getSheetAt(0)
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
font.setFontHeightInPoints(((12) as short))
font.setFontName('Arial')
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
cell.setCellValue('Add Award')
cell.setCellStyle(backgroundStyle1)
XSSFCell cella = rowi.createCell(2)
cella.setCellValue('Create Counseling')
cella.setCellStyle(backgroundStyle1)
XSSFCell cellb = rowi.createCell(3)
cellb.setCellValue('Delete Documents')
cellb.setCellStyle(backgroundStyle1)
XSSFCell cellc = rowi.createCell(4)
cellc.setCellValue('Edit Personnel')
cellc.setCellStyle(backgroundStyle1)
XSSFCell celld = rowi.createCell(5)
celld.setCellValue('Manage Production History')
celld.setCellStyle(backgroundStyle1)
XSSFCell celle = rowi.createCell(6)
celle.setCellValue('Set Production Status')
celle.setCellStyle(backgroundStyle1)
XSSFCell cellf = rowi.createCell(7)
cellf.setCellValue('Upload Personnel Certificate')
cellf.setCellStyle(backgroundStyle1)
XSSFCell cellg = rowi.createCell(8)
cellg.setCellValue('Upload Documents')
cellg.setCellStyle(backgroundStyle1)
XSSFCell cellh = rowi.createCell(9)
cellh.setCellValue('Sub Standard Performance Remarks')
cellh.setCellStyle(backgroundStyle1)

while (rowCount > row) {
	row = (row + 1)

	println(row)
	info.userName = data.getValue('UserName', row)
	info.passWord = data.getValue('Password', row)
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

	WebUI.click(findTestObject('Object Repository/Page_Records/span_Personnel (1)'))

	WebUI.setText(findTestObject('Object Repository/Page_Personnel/input_Search Personnel_269e1bf9d8766fe8325c_b9dd7f'), 'Recruiter 8412')
	WebUI.click(findTestObject('Object Repository/Page_Personnel/button_Search'))
	WebUI.click(findTestObject('Object Repository/Page_Personnel/a_Recruiter 8412'))
	WebUI.click(findTestObject('Object Repository/Page_Personnel  8412 Recruiter  RECRUITER  DAYTON/div_Related Actions'))

//WebUI.closeBrowser()

}
//file.close()
//FileOutputStream outFile = new FileOutputStream(new File(('C:\\TestOutput\\SecurityValidation_' + todayTimeStamp) + '.xlsx'))
//workbook.write(outFile)
//outFile.close()





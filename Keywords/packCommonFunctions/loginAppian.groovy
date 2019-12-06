package packCommonFunctions

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

import internal.GlobalVariable

public class loginAppian{

	@Keyword
	public void GetLogin() {
		//WebUI.openBrowser('')

		WebUI.maximizeWindow()

		WebUI.navigateToUrl('https://mcriss-cds.appiancloud.com/suite/')
		WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Username_un'))
		WebUI.setText(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Username_un'), 'dboswell')
		WebUI.setEncryptedText(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Password_pw'), 'ymCD6nePYUOBz+d/adMFDw==')

		WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Training)/input_Forgot your password_btn primary'))
	}

	@Keyword
	public void GetLoginTest() {

		WebUI.maximizeWindow()

		WebUI.navigateToUrl('https://mcriss-cdstest.appiancloud.com/suite/')
		//WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Test)/input_Remember me on this comp'))
		WebUI.setText(findTestObject('Object Repository/Page_MCRC Modernization (Test)/input_Remember me on this comp'), 'dboswell')
		WebUI.setEncryptedText(findTestObject('Object Repository/Page_MCRC Modernization (Test)/input_Remember me on this comp_13'),
				'ymCD6nePYUOBz+d/adMFDw==')

		WebUI.click(findTestObject('Object Repository/Page_MCRC Modernization (Test)/input_Forgot your password_btn'))
	}
}
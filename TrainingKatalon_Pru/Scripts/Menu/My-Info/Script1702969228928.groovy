import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import org.apache.poi.sl.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import com.google.common.collect.FilteredEntryMultimap.Keys
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.keyword.excel.ExcelKeywords



TestData data = findTestData('Data Files/TestDataMyInfo')

WebUI.click(findTestObject('MenuInfo/Page_OrangeHRM/Link/MyInfo'))

for (int baris = 1; baris <= data.getRowNumbers(); baris++) {
	
	String nama = WebUI.getText('Object Repository/MenuInfo/Page_OrangeHRM/NameRandomize')
	String dataSave = 'Excel/TestData.xlsx'
	
	Workbook excel = ExcelKeywords.getWorkBook(data)
	Sheet sheet1 = excel.getSheet('Sheet1')
	for(int i = baris; i <= baris; i++) {
		for(int j = 4; j <= 3; j++) {
			ExcelKeywords.setValueToCellByIndex(sheet1, i, j, nama)
			ExcelKeywords.saveWorkbook(dataSave, excel)
		}
	}
	
    WebUI.sendKeys(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_FirstName'), 
        Keys.chord(Keys.CONTROL, 'a', Keys.BACK_SPACE))

    WebUI.setText(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_FirstName'), 
        data.getValue('First Name', baris))

    WebUI.sendKeys(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_MiddleName'), 
        Keys.chord(Keys.CONTROL, 'a', Keys.BACK_SPACE))

    WebUI.setText(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_MiddleName'), 
        data.getValue('Middle Name', baris))

    WebUI.sendKeys(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_LastName'), 
        Keys.chord(Keys.CONTROL, 'a', Keys.BACK_SPACE))

    WebUI.setText(findTestObject('MenuInfo/Page_OrangeHRM/inputTextField/Page_OrangeHRM/Login-Menu/input_LastName'), 
        data.getValue('Last Name', baris))
}
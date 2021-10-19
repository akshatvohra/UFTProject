
'---------------------------------------------------Start-----------------------------------------\

Call ImportExcelSheet("C:\Users\vakshat\Desktop\TestCombiData.xls", "Test")

Call InitializeBrowserAndURL()

Call LoginToApp()

Datatable.SetCurrentRow Environment.Value("TestIteration")
myInputData = Datatable("FirstName",1)&"|"&Datatable("MiddleName",1)&"|"&Datatable("LastName",1)&"|"&Datatable("Gender",1)&"|"&Datatable("Date",1)&"|"&Datatable("Month",1)&"|"&Datatable("Year",1)&"|"&Datatable("Country",1)&"|"&Datatable("PostalCode",1)&"|"&Datatable("PhoneNumber",1)&"|"&Datatable("PersonName",1)

Call RegisterPatient(myInputData)

Call ConfirmPatientDetails()


'-----------------------------------------------------------------------------FUNCTIONS--------------------------------------------------------------------------------------------------------\

'Import Test Combi Data:-
Function ImportExcelSheet(ExcelSheetPath,SheetNameToImport)
	Datatable.ImportSheet ExcelSheetPath, SheetNameToImport, 1
End Function

'Initiate Browser and Navigate to URL:-
Function InitializeBrowserAndURL()
	SystemUtil.Run "chrome.exe", Parameter("URL")
	Environment("BrowserStatus") = Browser("Login").Exist
	AIUtil.SetContext Browser("Login")
End Function

'Login to Application:-
Function LoginToApp()
	If Environment("BrowserStatus") Then
		Reporter.ReportEvent micPass, "Browser Initiated Successfully", "Browser initiated."
		AIUtil("text_box", "Username").Type Parameter("UserName")
		AIUtil("text_box", "Password").Type Parameter("Password")
		AIUtil.FindTextBlock("Inpatient Ward").Click
		AIUtil("button", "Log In").Click
		wait 10
		Environment("URL_Page_Status") = Browser("Login").Page("Login").Exist
	Else
		Reporter.ReportEvent micFail, "Browser Not Initiated Successfully", "Browser not initiated. Issue Spotted!!!"
	End If
End Function

'Register A Patient:-
Function RegisterPatient(PatientDataString)
	AIUtil.FindTextBlock("Register a patient").Click
	'Split up the string value for the data of diff section:-
	myPatientData = Split(PatientDataString,"|")
	
	AIUtil("text_box", "Given (required)").Type myPatientData(0)
	AIUtil("text_box", "Middle").Type myPatientData(1)
	AIUtil("text_box", "Family Name (required)").Type myPatientData(2)
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 5
	
	AIUtil("combobox", "What's the patient's gender? (required)").Select myPatientData(3)
	
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 5
	
	
	AIUtil("text_box", "Day").Type myPatientData(4)
	
	AIUtil("combobox", "Month").Select myPatientData(5)
	
	AIUtil("text_box", "Year").Type myPatientData(6)
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 5
	
	'AIUtil("text_box", "@Birthdate").Type "Rajasthan"
	Browser("Login").Page("Login").WebEdit("xpath:=//input[@id='address1']").Set "Rajasthan"
	
	'AIUtil("text_box", "Address").Type "Rajasthan"
	Browser("Login").Page("Login").WebEdit("xpath:=//input[@id='address2']").Set "Rajasthan"
	
	AIUtil("text_box", "City/Village").Type "Rajasthan"
	
	AIUtil("text_box", "").Type "Rajasthan"
	
	AIUtil("text_box", "Country").Type myPatientData(7)
	
	AIUtil("text_box", "Postal Code").Type myPatientData(8)
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 5
	
	
	AIUtil("text_box", "What's the patient phone number?").Type myPatientData(9)
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 5
	
	AIUtil("combobox", "Who is the patient related to?").Select "Sibling"
	
	AIUtil("text_box", "Person Name").Type myPatientData(10)
	
	AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
	
	wait 10
	
	AIUtil("button", "Confirm").Click
	wait 10
End Function

Function ConfirmPatientDetails()
	Browser("Login").Page("Login").WebElement("xpath:=//ul[@id='breadcrumbs']//i[@class='icon-chevron-right link']/parent::li").WaitProperty "visible",True,5000



	myText = Browser("Login").Page("Login").WebElement("xpath:=//ul[@id='breadcrumbs']//i[@class='icon-chevron-right link']/parent::li").GetROProperty("innertext")
	
	myID = Browser("Login").Page("Login").WebElement("xpath:=//div[@class='float-sm-right']/span").GetROProperty("innertext")
	
	If StrComp(trim(myText),myName&" "&myMiddleName&" "&myLastName) = 0 Then
		Reporter.ReportEvent micPass, "Patient Name Confirmed", "Data added."
	End If
	
	
	Browser("Login").Page("Login").WebElement("HomeIcon_WebElement").Click
	
	
	
	
	AIUtil.FindTextBlock("Find Patient").Click
	
	wait 10
	
	AIUtil("text_box", "Find Patient Record").Type myID
	
	wait 10
	
	set myTableObject = Browser("Login").Page("Login").WebTable("Identifier")
	
	tableColumns = myTableObject.GetROProperty("cols")
	
	myData=""
	
	For i = 1 To tableColumns
		myData = myData&"|"&myTableObject.GetCellData(2,i)
	Next
	
	Reporter.ReportEvent micDone, "Patient Record: "&Replace(myData,"|","",1,1), "Getting the Data of Patient recently added."
	
	AIUtil("down_triangle", micNoText, micFromRight, 1).Click
	
	Browser("Login").Close
End Function

'Datatable.ImportSheet "C:\Users\vakshat\Desktop\TestCombiData.xls", "Test", 1
'
'SystemUtil.Run "chrome.exe", "https://demo.openmrs.org/openmrs/login.htm"
'AIUtil.SetContext Browser("Login")
'
'AIUtil("text_box", "Username").Type "Admin"
'AIUtil("text_box", "Password").Type "Admin123"
'AIUtil.FindTextBlock("Inpatient Ward").Click
'AIUtil("button", "Log In").Click
'
'AIUtil.FindTextBlock("Register a patient").Click
'
'myName = RandomString(5)
'myMiddleName = RandomString(5)
'myLastName = RandomString(5)
'AIUtil("text_box", "Given (required)").Type myName
'AIUtil("text_box", "Middle").Type myMiddleName
'AIUtil("text_box", "Family Name (required)").Type myLastName
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
'
'AIUtil("combobox", "What's the patient's gender? (required)").Select "Male"
'
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
'
'
'AIUtil("text_box", "Day").Type 15
'
'AIUtil("combobox", "Month").Select "May"
'
'AIUtil("text_box", "Year").Type "1970"
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
''AIUtil("text_box", "@Birthdate").Type "Rajasthan"
'Browser("Login").Page("Login").WebEdit("xpath:=//input[@id='address1']").Set "Rajasthan"
'
''AIUtil("text_box", "Address").Type "Rajasthan"
'Browser("Login").Page("Login").WebEdit("xpath:=//input[@id='address2']").Set "Rajasthan"
'
'AIUtil("text_box", "City/Village").Type "Rajasthan"
'
'AIUtil("text_box", "").Type "Rajasthan"
'
'AIUtil("text_box", "Country").Type "India"
'
'AIUtil("text_box", "Postal Code").Type "110084"
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
'
'AIUtil("text_box", "What's the patient phone number?").Type "9765412345"
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
'AIUtil("combobox", "Who is the patient related to?").Select "Sibling"
'
'AIUtil("text_box", "Person Name").Type "TIRAN"
'
'AIUtil("right_triangle", micNoText, micFromBottom, 1).Click
'
'
'
'AIUtil("button", "Confirm").Click
'
'
'
'Browser("Login").Page("Login").WebElement("xpath:=//ul[@id='breadcrumbs']//i[@class='icon-chevron-right link']/parent::li").WaitProperty "visible",True,5000
'
'
'
'myText = Browser("Login").Page("Login").WebElement("xpath:=//ul[@id='breadcrumbs']//i[@class='icon-chevron-right link']/parent::li").GetROProperty("innertext")
'
'myID = Browser("Login").Page("Login").WebElement("xpath:=//div[@class='float-sm-right']/span").GetROProperty("innertext")
'
'If StrComp(trim(myText),myName&" "&myMiddleName&" "&myLastName) = 0 Then
'	Reporter.ReportEvent micPass, "Patient Name Confirmed", "Data added."
'End If
'
'
'Browser("Login").Page("Login").WebElement("HomeIcon_WebElement").Click
'
'
'
'
'AIUtil.FindTextBlock("Find Patient").Click
'
'wait 3
'
'AIUtil("text_box", "Find Patient Record").Type myID
'
'wait 3
'
'set myTableObject = Browser("Login").Page("Login").WebTable("Identifier")
'
'tableColumns = myTableObject.GetROProperty("cols")
'
'myData=""
'
'For i = 1 To tableColumns
'	myData = myData&"|"&myTableObject.GetCellData(2,i)
'Next
'
'Reporter.ReportEvent micDone, "Patient Record: "&Replace(myData,"|","",1,1), "Getting the Data of Patient recently added."
'
'AIUtil("down_triangle", micNoText, micFromRight, 1).Click
'
'Browser("Login").Close
'
'
'Function RandomString(strLen)
'
'    Dim str
'    Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
'    For i = 1 to strLen
'        str = str & Mid( LETTERS, RandomNumber( 1, Len( LETTERS ) ), 1 )
'    Next
'    RandomString = str
'
'End Function
'
'

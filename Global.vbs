Const QC_PATH_PREFIX = "[QualityCenter] "

'@Description Determine if object matches specification
'@Documentation Determine if <theObj> matches <theDesc>
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] theObj A reference to the object to evaluate
'@InParameter [in] theDesc A reference to the object description
'@ReturnValue 'True' if the specied object matches the description; otherwise 'False'
Public Function isDescribedObject(ByRef theObj, ByRef theDesc)
	Dim propIndex
	Dim thisProp, propValue, propVType, propRegEx
	Dim objValue, objVType
	
	Dim regEx
	Dim isObj

	' Allocate RegExp object
	Set regEx = New RegExp

	' Iterate over description properties
	For propIndex = 0 to (theDesc.Count - 1)
		' Get description property object
		Set thisProp = theDesc(propIndex)

		' Get description value
		propValue = thisProp.Value
		' Get description value type
		propVType = VarType(propValue)
		' If description value is a string
		If (propVType = vbString) Then
			' Get description regular expression spec
			propRegEx = thisProp.RegularExpression
		Else ' otherwise (non-string)
			' Not regular expression
			propRegEx = False
		End If

		' Get value of test obj value
		objValue = theObj.GetROProperty(thisProp.Name)
		' Get test obj value type
		objVType = VarType(objValue)

		' If value types differ
		If (propVType <> objVType) Then
			' Not regular expression
			propRegEx = False

			' If description value is empty or uninitialized
			If ((propVType = vbEmpty) Or (propVType = vbNull))Then
				' No match
				isObj = False
			' Otherwise, if test obj value is empty or uninitialized
			ElseIf ((objVType = vbEmpty) Or (objVType = vbNull)) Then
				' No match
				isObj = False
			' Otherwise, if property value is a string
			ElseIf (propVType = vbString) Then
				' Convert property value to test obj value type
				propValue = convertString(propValue, objVType)
				' If conversion failed, no match
				isObj = Not IsNull(propValue)
			' Otherwise, if test obj value is a string
			ElseIf (objVType = vbString) Then
				' Convert test obj value to description value type
				objValue = convertString(objValue, propVType)
				' If conversion failed, no match
				isObj = Not IsNull(objValue)
			' Otherwise, if description is a boolean
			ElseIf (propVType = vbBoolean) Then
				' Convert test obj value to boolean
				objValue = CBool(objValue)
			' Otherwise, if test obj value is a boolean
			ElseIf (objVType = vbBoolean) Then
				' Convert description value to boolean
				propValue = CBool(propValue)
			' Otherwise (none of the above)
			Else
				' No match
				isObj = False
			End If
		' Otherwise (value types match)
		Else
			' Request evaluation
			isObj = True
		End If

		' If evaluation requested
		If (isObj) Then
			' If regular expression
			If (propRegEx) Then
				' Set description pattern
				regEx.Pattern = propValue
				' Evaluate test obj value
				isObj = regEx.Test(objValue)
			' Otherwise (not regex)
			Else
				' Evaluate test obj value
				isObj = (propValue = objValue)
			End If
		End If

		' If mismatched, stop checking
		If Not (isObj) Then Exit For
	Next

	' Set function result
	isDescribedObject = isObj
	
	' Release RegExp
	Set regEx = Nothing
	
End Function


'@Description Convert the specified string to the indicated type
'@Documentation Convert <theStr> to type <toType>
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] theStr The string to be converted
'@InParameter [in] toType The desired result type
'@ReturnValue On success, the specified string converted to the indicated type; otherwise 'Null'
Public Function convertString(ByVal theStr, ByVal toType)

   On Error Resume Next

	' Initialize result
	convertString = Null

	Select Case toType
		Case vbInteger
			convertString = CInt(theStr)
		Case vbLong
			convertString = CLng(theStr)
		Case vbSingle
			convertString = CSng(theStr)
		Case vbDouble
			convertString = CDbl(theStr)
		Case vbCurrency
			convertString = CCur(theStr)
		Case vbDate
			convertString = CDate(theStr)
		Case vbByte
			convertString = CByte(theStr)
		Case vbBoolean
			convertString = CBool(theStr)
	End Select
   
End Function


'@Description Return upper bound of specified array
'@Documentation Return upper bound of <theArray>
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] theArray The array whose upper bound is requested
'@ReturnValue The upper bound of the array, or -1 if the array is empty
Public Function safeUBound(ByRef theArray)

	On Error Resume Next
	theUBound = -1
	theUBound = UBound(theArray)
	safeUBound = theUBound
	
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnSetGlobalTimeouts( )
' ---------------------------------------------------------------------------------------------------
Public Sub cmnSetGlobalTimeouts(iMSec)
   Setting("WebTimeout") = iMSec
   Setting("DefaultTimeout") = iMSec   
End Sub


' ---------------------------------------------------------------------------------------------------
'                                         cmnSetObjExistsTimeout ( )
' ---------------------------------------------------------------------------------------------------
Public Sub cmnSetObjExistsTimeout(iMSec)
   Setting("WebTimeout") = iMSec
   Setting("DefaultTimeout") = iMSec   
End Sub


' --------------------------------------------------------------------------------------
'                      cmnGetGlobalTimeOuts_QC
' --------------------------------------------------------------------------------------
'                    Gets the timeout values from QC, 
'             located in the details section of the test sets

'  Dependancy: 
'		DataTable.vbs / dtDataSheetPeramExist function
' --------------------------------------------------------------------------------------
Public Function cmnGetGlobalTimeOuts_QC()

	' Declarations
	Dim defaultTimeOut
	Dim specialTimeOut
	Dim testSet
	Dim slowTimeOut
	Dim fastTimeOut

	' Initialize default timeout to medium wait
	defaultTimeOut = Environment("Wait_MEDIUM")
	' Initialize special timeout to very short wait
	specialTimeOut = Environment("Wait_VERYSHORT")

	' Get current test set
	Set testSet = QCUtil.CurrentTestSet
	' If a test set is selected
	If Not (testSet Is Nothing) Then
		' Get value of slow timeout field
		slowTimeOut = testSet.Field(SLOW_TIMEOUT)
		' If slow timeout is numeric, use it as the default timeout
		If (IsNumeric(slowTimeOut)) Then defaultTimeOut = slowTimeOut

		' Get value of fast timeout field
		fastTimeOut = testSet.Field(FAST_TIMEOUT)
		' If fast timeout is numeric, use it as the special timeout
		If (IsNumeric(fastTimeOut)) Then specialTimeOut = fastTimeOut
	End If

   ' Set default timeout parameter (add if absent)
	MySetParameter dtGlobalSheet, "DefaultTimeOut", 1, defaultTimeOut, flagAddParam

   ' Set special timeout parameter (add if absent)
	MySetParameter dtGlobalSheet, "SpecialTimeOut", 1, specialTimeOut, flagAddParam

	' Store timeout values in the environment
	Environment("envSlowTimeOut") =  defaultTimeOut
	Environment("envFastTimeOut") =  specialTimeOut

	' Send timout numbers to the test results report
	Reporter.ReportEvent micInfo, "Long Timeout: " & defaultTimeOut, "< Default Timeout >" & vbcr & defaultTimeOut
	Reporter.ReportEvent micInfo, "Short Timeout: " & specialTimeOut, "< Special Timeout >" & vbcr & specialTimeOut

End Function


' -------------------------------------------------------------------------------------
'                                    CvtTimeOrDate
' -------------------------------------------------------------------------------------
' This function formats a time/date into a string that can be used for file names and  registering an online user
Public Function CvtTimeOrDate(t)

	Dim dt
	Dim z
	
	For z = 1 to Len(t)
		If Mid(t,z,1) = "/" Then
			dt = dt + "-"
		ElseIf	Mid(t,z,1) = ":" Then
			dt = dt + "-"	
		ElseIf	Mid(t,z,1) = " " Then
			dt = dt + "-"	 
		Else
			dt = dt + Mid(t,z,1)
		End If
	Next
	
	CvtTimeOrDate = dt
	
End Function


' ---------------------------------------------------------------------------------------------------
'                                       cmnActionDesc
' ---------------------------------------------------------------------------------------------------
' Posts to the test results the name of the action and a brief description 
'of what the action is doing passed thru a param
' ---------------------------------------------------------------------------------------------------
Public Function cmnActionDesc(sDesc)
	sTitle = Environment("ActionName")
	Reporter.ReportEvent micInfo, sTitle, sDesc
End Function


' ---------------------------------------------------------------------------------------------------
'                                       cmnTestDesc
' ---------------------------------------------------------------------------------------------------
' Posts to the test results the name of the action and a brief description 
'of what the action is doing passed thru a param
' ---------------------------------------------------------------------------------------------------
Public Function cmnTestDesc()
	sTestName = Environment("TestName")
	sLocalHostName = Environment("LocalHostName")	
	
	sOperatingSys = Environment("OS")	
	sOSVersion = Environment("OSVersion")	
	sProductName = Environment("ProductName")	
	sProductVer = Environment("ProductVer")	

	sTestDir = Environment("TestDir")	
	sUserName = Environment("UserName")	
   
	sTitle = "Test Description"
	sDesc = "Test Name: " +sTestName +vbcr+vbcr+ "Local Host: " +sLocalHostName +vbcr+ "Operating System: " +sOperatingSys+ " v." +sOSVersion +vbcr+vbcr+ "Tester Name: " +sUserName +vbcr+ "Test Tool: " +sProductName+ " v." +sProductVer +vbcr+ "Test Src Folder: " +sTestDir
	
	Reporter.ReportEvent micInfo, sTitle, sDesc
End Function
  

' ---------------------------------------------------------------------------------------------------
'                       cmnReusableAction_ImportDataSheet
' ---------------------------------------------------------------------------------------------------

Public Function cmnReusableAction_ImportDataSheet()

	'Get Excel file path
	sSourceDataSheet = Environment("DataSheet_ML")

	'Get Destination Sheet
	sDestSheet = Environment("ActionName")
	
	'Get Test Set ID
	sID = Environment("env_TestSet_ID")	

	'Get Source Sheet	
	aSheet = Split(sDestSheet, " [")
	sSrcSheet = aSheet(0)
	sSrcSheet = sSrcSheet + sID 

	'Import sheet
	DataTable.ImportSheet sSourceDataSheet, sSrcSheet, sDestSheet
	
End Function


' ---------------------------------------------------------------------------------------------------
'                         Non-Reusable    cmnAction Inport DataSheet
' ---------------------------------------------------------------------------------------------------
Public Function cmnAction_ImportDataSheet()

   '@@@ add file path

	'Get Excel file path
	sSourceDataSheet = Environment("DataSheet_ML")

	'Get Destination Sheet
	sDestSheet = Environment("ActionName")

	'Get Test Set ID
	sID = Environment("env_TestSet_ID")		

	'Get Source Sheet
	sSrcSheet = sDestSheet +"." +sID

	'Import sheet
	DataTable.ImportSheet sSourceDataSheet, sSrcSheet, sDestSheet
	Wait 1

End Function


' ---------------------------------------------------------------------------------------------------
'                                                     cmnTCMapper
' ---------------------------------------------------------------------------------------------------
'Desc - 
	'Sends TC tracability information to the reporter tool

' Args -  NA

'Req - 
	'This function is designed only for actions that represent test case shells
	'and have their information set up in the same row in the GlobalSheet.
		
' ---------------------------------------------------------------------------------------------------
Public Function cmnTCMapper()

   'Get Action name - should only be for actions representing test cases
	sFeature = Environment("ActionName")
	aFeature = Split(sFeature, " ")
	sFeature = aFeature(0)

	'Loop until you find a match on the global sheet
	Do
		iRow = iRow +1
		DataTable.GlobalSheet.SetCurrentRow iRow

		'Once match is found, gather other details
		If sFeature = DataTable("Feature", dtLocalSheet) Then
			sTC = DataTable("TC", dtLocalSheet)
			sTCName = DataTable("TC_Name", dtLocalSheet)
			sTestDoc = DataTable("TestDoc", dtLocalSheet)
			Exit Do
		End If
	Loop Until "" = sFeature

   'Report TC tracking data to the reporter tool
	Reporter.ReportEvent micInfo, "TC Tractability", "< Feature >" +vbcr+ sFeature +vbcr+vbcr+ "< TC# >" +vbcr+ sTC +vbcr+vbcr+ "< Test Case Name >" +vbcr+ sTCName +vbcr+vbcr+ "< Test Case Details File >" +vbcr+ sTestDoc
	
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnAwaitObjectExist ( )
' ---------------------------------------------------------------------------------------------------

'Desc - 
	''Pass object and it will wait until object exists or timeout is reached

'Returns - 
	'Ture (if object exists) or False (if no object exists w/i timeout)

' Args - 

	'oObject -  Object reference
	'iTimeout -  Intuitive name for the object

'Req - NA
		
' ---------------------------------------------------------------------------------------------------
Public Function cmnAwaitObjectExist(oObject, iTimeout)
	Do 
		Wait 1
		
		iTimer = iTimer+1
		If iTimer > iTimeout Then
			cmnWaitObjectExist = "False"
			Exit Do
		End If
		
	Loop Until "True" = oObject.Exist

	If "True" = oObject.Exist Then
		cmnWaitObjectExist = "True"
	End If
	
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnForceClick ( )
' ---------------------------------------------------------------------------------------------------
Public Function cmnForceClick(oObject)

	x = oObject.GetROProperty("abs_x")
	y = oObject.GetROProperty("abs_y")
	Set DeviceReplay = CreateObject("Mercury.DeviceReplay") 

	On Error Resume Next
	
	wait 1
	DeviceReplay.MouseMove x +15, y+15
	wait 1
	DeviceReplay.MouseClick x +15, y+15, 0
	Set DeviceReplay = Nothing 
	
	On error goto 0
	
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnLoadDataSheet ( )
' ---------------------------------------------------------------------------------------------------
Public Function cmnLoadDataSheet (sUniqueID)
	'This code imports the Global Variables into the Global DataSheet .
		'This is the only place we use a hard-coded path.
	sSourceDataSheet = Environmental("DataSheet_ML")
	
	'Trim off the reusable action library reference from the name
	aActionName = Split(Environment("ActionName"), " ")
	sActionName = aActionName(0) +"."+ sUniqueID
	
	'Import the sheet
	Call DataTable.ImportSheet (sSourceDataSheet, sActionName, sActionName)
	wait 1
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnSetCheckBox ( )
' ---------------------------------------------------------------------------------------------------
' Desc: Sets a chkbox to "Off" unless the value is set in the sSetChk Param
' ---------------------------------------------------------------------------------------------------
Public Function cmnSetCheckBox(oObject, bSetChk, sDefaultSetting)
	Set oObject = oObject
	If "" <> bSetChk Then
		oObject.Set bSetChk
	Else
		oObject.Set sDefaultSetting
	End If
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnIterateLocalSheet ( )
' ---------------------------------------------------------------------------------------------------
' Desc: Sets athe LocalSheet to whatever row specified.
'Arg: Takes an integer.
' Tip: use environment variable to pass from outer action.
' ---------------------------------------------------------------------------------------------------
Public Function cmnIterateLocalSheet(iSetRow)
	If "" <> iSetRow Then
		DataTable.LocalSheet.SetCurrentRow CInt(iSetRow)
	Else
		DataTable.LocalSheet.SetCurrentRow 1
	End If
End Function


' ---------------------------------------------------------------------------------------------------
'                                         cmnSetEnvironment ( )
' ---------------------------------------------------------------------------------------------------
' Desc: Sets the Environment Input on or off
' Select to get Environment ID & Tests to run 
' from user input or automatically from datasheet
' ---------------------------------------------------------------------------------------------------
Public Function cmnSetEnvironment ()
	sManual = "MANUAL"
	If sManual = DataTable("Environments", dtLocalSheet) Then
		' ------------------------------------------------------------------------------------
		'      GET ENVIRONMENT ID FROM  USER INPUT
		' ------------------------------------------------------------------------------------
		
		'Get the Test Environment
			'Function sets Test Environment Env Var (rtTestEnv)
		Call cmnSetEnvID()
		
	Else
		' ------------------------------------------------------------------------------------
		'                GET TESTS  RUN FROM DATASHEET
		' ------------------------------------------------------------------------------------
		
		Environment("rtEnvID") = DataTable("Environments", dtLocalSheet) 
	End If
End Function


' ------------------------------------------------------------------------------------
'                        cmnDTReplaceInvalidChrs ()
' Replaces invalid chrs with underscore for inserting 
' and matching against the datasheet column names
' ------------------------------------------------------------------------------------
Public Function cmnDTReplaceInvalidChrs(sText)
	sText = Replace (sText," ","_", 1, -1)  	
	sText = Replace (sText,"/","_", 1, -1)
	sText = Replace (sText,"\","_", 1, -1)
	cmnDTReplace = sText
End Function


' ---------------------------------------------------------------------------------------------------
'                                       cmnGetTableContent ()
' ---------------------------------------------------------------------------------------------------
'DESC - 

	'Draws data out of a webtable and places it into a QTP DataSheet.
	'Tests can then pull data from the datasheet at will.
	'Also is a debugging visual ade. 		

'RETURNS -  

	'The New DataSheet Name
	'QTP DataSheet populated by WebTable

' ARGS -  

	'oWebTable - an object reference to a webtable
	'iHeaderRow - the row the column names start on

		'Tip: Use the checkpoint tool to look into the table object to see where the hader row is.

'EXAMPLE - 

'	Set oWebTable = Browser("Market Leader").Page("LIST BUILDER").Frame("QUICK Tab").WebTable("Prospects Returned")
'	Call cmnGetTableContent(oWebTable, 2)

'REQ -  Common.vbs
	
' ---------------------------------------------------------------------------------------------------
Public Function cmnGetTableContent(oWebTable, iHeaderRow)

   ' ------------------------------------------------------------------------
	'                CREATE NEW DATASHEET
   ' ------------------------------------------------------------------------   

	'ReSet  Table object
	Set oWebTable = oWebTable
	
	'Get WebTable name property
	sTableID = oWebTable.GetROProperty("html id")
	sTblDesc = "TableData" 

	'Use name property to identify the table displayed in added datasheet
	Environment("rtWebTableCounter") = Environment("rtWebTableCounter") +1
	sTCCount = CStr(Environment("rtWebTableCounter"))

	sSheetName = "TC" +sTCCount+ "." +sTableID+ "." + sTblDesc
	DataTable.AddSheet(sSheetName)

	'Return new datasheet name
	cmnGetTableContent = sSheetName

   ' ------------------------------------------------------------------------
	'           POP DATASHEET WITH TBL DATA
   ' ------------------------------------------------------------------------     
   
	iRowQty = oWebTable.RowCount  '-4 ' -4 because there is one leading row and three following rows in this table.
	IColumnQty = oWebTable.ColumnCount(iHeaderRow)

	sProspectsTxt = oWebTable.GetROProperty("innertext")
	aProspectsTxt = Split(sProspectsTxt, " ")
	iProspectsQty = aProspectsTxt(0)
'
'	'If the number of Prospects dows not match the number 
'	'of rows displayed on the page then report a failure
'	If iRowQty <> CInt(iProspectsQty) Then
'		Reporter.ReportEvent micFail,"LB.Prospects Table", "Expected Rows: " +CStr(iRowQty) _
'		 +vbcr+ "This value s based on the 'LB.Prospects Returrned' table displayed on the page's row count" +vbcr+vbcr+ _
'		  "Actual rows: " +CStr(iProspectsQty) +vbcr+ "This value is based on the number of Prospects returned listed above said table"
'	End If
	
	For iCol = 1 To IColumnQty
		'Use iHeaderRow to indicate where the column names are
		sColName = Trim(oWebTable.GetCellData (iHeaderRow, iCol))

		If "" <> sColName Then
			'Remove invalid chrs from name
			sColName = LCase(cmnDTReplace(sColName))
			'Get column name and populate datatable Parameter.Name
			DataTable.GetSheet(sSheetName).AddParameter sColName, ""
	
			'Get row values and populate column
			dtRow = 0
			For iRow = iHeaderRow+1 To iRowQty 
				' FYI - There is one leading row and three following rows in this table.
				dtRow = dtRow +1
				DataTable.GetSheet(sSheetName).SetCurrentRow dtRow
				
				sCellValue = Trim(oWebTable.GetCellData(iRow, iCol))
				'If the cell value is invalid exit the for loop
				If "ERROR: The specified cell does not exist." = sCellValue Then
					Exit For
				Else
					DataTable(sColName, sSheetName) = sCellValue
				End If
	
				'Reset current row to top for next column values
				DataTable.GetSheet(sSheetName).SetCurrentRow iHeaderRow
			Next
		End If
	Next 
	
End Function


'********************************************************************************
'	General and Standard Windows functions, plus registration information required by all add-in function library files
'       -------------------------
'
'   Available Functions:
'	* OpenApp - Opens a specified application (common file)
'	* AddToTestResults - Adds a Report.Event step to the Test Results (common file)
'	* VerifyProperty - Verifies the value of a specified property (for all Standard Windows test objects)
'	* OutputProperty - Returns the value of the specified property (for all Standard Windows test objects)
'	* VerifyEnable - Verifies whether a specified object is enabled (for all Standard Windows test objects)
'	* VerifyValue - Verifies the value of a specified object (for Static, WinCalendar, WinCheckBox, WinComboBox, WinEdit, WinEditor, WinList, WinListView, WinRadioButton, WinSpin, WinTab, WinTreeView, ListView20WndClass, ListViewWndClass)
'	* GetValue - Returns the object value (for Static, WinCalendar, WinCheckBox, WinComboBox, WinEdit, WinEditor, WinList, WinListView, WinRadioButton, WinSpin, WinTab, WinTreeView, ListView20WndClass, ListViewWndClass)
'
'   Version: QTP8.2 November 2004
'
'   ** Do not modify this file. It may be automatically updated by a later version, and then you will lose your changes.
'********************************************************************************


' Function OpenApp
' ------------------
' Open a specified application
' Parameter: application - the application full name (including location)
'@Description Opens an application
'@Documentation Open the  <application> application.
Function cmnOpenApp (application)
	systemUtil.Run application
End Function


' AddToTestResults
' --------------
' Add a Report.Event step to the Test Results
'Parameters:
'	status - Step status (micPass, micFail, micDone or micWarning)
'       StepName - Name of the intended step in the report (object name)
'       details - Description of the report event
'
'@Description Reports an event to the Test Results
'@Documentation Report an event to the Test Results.
Public Function cmnAddToTestResults (status, StepName, details)
	Reporter.ReportEvent status, StepName, details
End Function


' Function OutputProperty
' ------------------------
' Return the value of the specified property
' Parameters:
'    	paramName - the parameter name (to return the value)
' Returns - The property value
'
'@Description Returns the value of the specified property
'@Documentation Return the <Test object name> <test object type> <PropertyName> value.
Function cmnOutputProperty (obj, PropertyName)
	Dim property_value
	property_value = obj.GetROProperty(PropertyName)
	cmnOutputProperty = property_value
End Function


' ******** GetValue Functions - Start ***********


' Function GetValueProperty
' --------------------------
' Return the object 'Value' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetValueProperty (obj)
	cmnGetValueProperty = obj.GetROProperty("value")
End Function


' Function GetDateProperty
' --------------------------
' Return the object 'Date' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetDateProperty (obj)
	cmnGetDateProperty = obj.GetROProperty("date")
End Function


' Function GetTextProperty
' --------------------------
' Return the object 'Text' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetTextProperty (obj)
	cmnGetTextProperty = obj.GetROProperty("text")
End Function


' Function GetCheckedProperty
' --------------------------
' Return the object 'checked' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetCheckedProperty (obj)
	cmnGetCheckedProperty = obj.GetROProperty("checked")
End Function


' Function GetSelectionProperty
' --------------------------
' Return the object 'selection' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetSelectionProperty (obj)
	cmnGetSelectionProperty = obj.GetROProperty("selection")
End Function


' Function GetPositionProperty
' --------------------------
' Return the object 'position' property
'
'@Description Returns the Object value
'@Documentation Return the <Test object name> <test object type> value.
Public Function cmnGetPositionProperty (obj)
	cmnGetPositionProperty = obj.GetROProperty("position")
End Function

' ******** GetValue functions - End ***********


'@Description Convert the specified date to its epoch equivalent
'@Documentation Convert <myDate> to its epoch equivalent
'@Author sbabcoc
'@Date 07-APR-2011
'@InParameter [in] myDate, date, The date to be converted
'@ReturnValue The epoch equivalent of the specified date
Function date2epoch(myDate)
	date2epoch = DateDiff("s", "01/01/1970 00:00:00", myDate) * 1000
End Function


'@Description Convert the specified epoch to its date equivalent
'@Documentation Convert <myEpoch> to its date equivalent
'@Author sbabcoc
'@Date 07-APR-2011
'@InParameter [in] myEpoch, long, The epoch to be converted
'@ReturnValue The date equivalent of the specified epoch
Function epoch2date(myEpoch)
	epoch2date = DateAdd("s", CDbl(myEpoch) / 1000, "01/01/1970 00:00:00")
End Function


'@Description Retrieve the name of the active user
'@Documentation Retrieve the name of the active user
'@Author sbabcoc
'@Date 29-JUN-2011
'@ReturnValue The name of the active user
'@Notes If connected to Quality Center, the user name is obtained from there. Otherwise, the user name is obtained from the environment.
Function getUserName()
	Dim userName
	Dim theOffset

	' If connected to Quality Center
	If (QCUtil.IsConnected) Then
		' Get Quality Center user name
		userName = QCUtil.QCConnection.UserName
	' Otherwise (not connected)
	Else
		' Get Windows login user name
		userName = Environment("UserName")
		' Determine if this is an 'admin' user
		theOffset = InStr(1, userName, ".adm")
		' If user is 'admin'
		If (theOffset > 0) Then
			' Trim off 'admin' suffix from user name
			userName = Left(userName, theOffset)
		End If
	End If

	getUserName = userName

End Function

'@Description Download specified attachment(s) from current test case
'@Documentation Download attachment(s) <attachName> from current test case
Function downloadAttachments(attachName)
	Dim currentTest
	Dim attachFact
	Dim attachFilter
	Dim attachList
	Dim attachCount
	Dim attachIndex
	Dim pathList
	Dim tempPath
	Dim thisAttach
	Dim attachStore

	' Init result
	pathList = Null
	
	' If connected to Quality Center
	If (QCUtil.IsConnected) Then
		' Get Quality Center current test
		Set currentTest = QCUtil.CurrentTest
		' Get current test attachments factory
		Set attachFact = currentTest.Attachments
		' Get filter to find attachments
		Set attachFilter = attachFact.Filter

		' Ensure name begins with "*"
		If (Mid(attachName, 1, 1) <> "*") Then attachName = "*" & attachName
		' Set attachment filter criteria
		attachFilter.Filter("CR_REFERENCE") = attachName
		' Get filtered list of attachments
		Set attachList = attachFact.NewList(attachFilter.Text)
		' Get count of attachments
		attachCount = attachList.Count
		' If attachments were found
		If (attachCount > 0) Then
			' Allocate path list
			pathList = Array()
			' Set list dimensions
			ReDim pathList(attachCount - 1)

			' Initialize index
			attachIndex = 0
			' Get path of temp folder
			tempPath = getTempFolderPath()
			
			' Iterate over attachments
			For Each thisAttach In attachList
				' Get extended storage for attachment
				Set attachStore = thisAttach.AttachmentStorage
				' Set client-side storage path
				attachStore.ClientPath = tempPath
				' Download attachment
				thisAttach.Load True, ""
				' Add download path to list
				pathList(attachIndex) = thisAttach.FileName
				' Increment index
				attachPath = attachPath + 1
			Next
		End If

		downloadAttachments = pathList
	End If

	' Release objects
	Set currentTest = Nothing
	Set attachFact = Nothing
	Set attachFilter = Nothing
	Set attachList = Nothing
	Set thisAttach = Nothing
	Set attachStore = Nothing
End Function

Function getTempFolderPath()
	Dim fileSysObj
	Dim tempFolder

	' Create file system object
	Set fileSysObj = CreateObject("Scripting.FileSystemObject")
	' Get object for temp folder
	Set tempFolder = fileSysObj.GetSpecialFolder(2)
	' Get path of temp folder
	getTempFolderPath = tempFolder.ShortPath
	
	' Release objects
	Set fileSysObj = Nothing
	Set tempFolder = Nothing
End Function

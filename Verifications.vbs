Option Explicit

' chkVerifyParity constants
Const EXP_POS = True
Const EXP_NEG = False
Const CMP_VALUE = 0
Const CMP_STRING = 1
Const CMP_EQUAL = 2
Const CMP_LESS = 4
Const CMP_LES_EQ = 6
Const CMP_MORE = 8
Const CMP_MOR_EQ = 10
Const CMP_NOT_EQ = 12
Const CMP_ASSERT = 14
Const CMP_CASING = 16
Const CMP_HEAD = 32
Const CMP_TAIL = 64
Const CMP_BODY = 96

Const STR_EQUAL = 3 ' CMP_STRING Or CMP_EQUAL
Const STR_LESS = 5 ' CMP_STRING Or CMP_LESS
Const STR_LES_EQ = 7 ' CMP_STRING Or CMP_LES_EQ
Const STR_MORE = 9 ' CMP_STRING Or CMP_MORE
Const STR_MOR_EQ = 11 ' CMP_STRING Or CMP_MOR_EQ
Const STR_NOT_EQ = 13 ' CMP_STRING Or CMP_NOT_EQ
Const STR_HEAD = 33 ' CMP_STRING Or CMP_HEAD
Const STR_TAIL = 65 ' CMP_STRING Or CMP_TAIL
Const STR_BODY = 97 ' CMP_STRING Or CMP_BODY

Const CAS_EQUAL = 19 ' CMP_STRING Or CMP_CASING Or CMP_EQUAL
Const CAS_LESS = 21 ' CMP_STRING Or CMP_CASING Or CMP_LESS
Const CAS_LES_EQ = 23 ' CMP_STRING Or CMP_CASING Or CMP_LES_EQ
Const CAS_MORE = 25 ' CMP_STRING Or CMP_CASING Or CMP_MORE
Const CAS_MOR_EQ = 27 ' CMP_STRING Or CMP_CASING Or CMP_MOR_EQ
Const CAS_NOT_EQ = 29 ' CMP_STRING Or CMP_CASING Or CMP_NOT_EQ
Const CAS_HEAD = 49 ' CMP_STRING Or CMP_CASING Or CMP_HEAD
Const CAS_TAIL = 81 ' CMP_STRING Or CMP_CASING Or CMP_TAIL
Const CAS_BODY = 113 ' CMP_STRING Or CMP_CASING Or CMP_BODY

' chkVerifyExistence constants
Const EXPECT_ABSENT = False
Const EXPECT_EXISTS = True

' =======================================================
'                     Common Verification Functions Library
' =======================================================

'Herein contains all common reusable verifications functions
	'All actions sould accociate to this resource file.
' =======================================================

'@Description Verify that the specified sample value meets the indicated criteria
'@Documentation Verify that the specified sample value meets the indicated criteria
'@Author sbabcoc
'@Date 02-JUN-2011
'@InParameter [in] sample Sample value to be evaluated
'@InParameter [in] master Master value used for evaluation, or description of assertion for CMP_ASSERT
'@InParameter [in] compare Comparision used to evaluate sample value:
'		
'		BOOLEAN ASSERTION
'			' EXAMPLE: Verify that this shopping cart item has a product name
'			isCorrect = chkVerifyParityEx(this_item.Item(key_ProdName).Exists, "specified", CMP_ASSERT, EXP_POS, "verifyShoppingCart 04", "Item name")
'		-------------------------------------
'		CMP_ASSERT: <sample> is Boolean ('True' or 'False'), described by the string in <master>
'		
'		COMPARISONS (SAMPLE VS. MASTER)
'			MODES: CMP = binary; STR = string; CAS = case-sensitive
'			' EXAMPLE: Verify that the expected quantity was added to the shopping cart
'			isCorrect = chkVerifyParityEx(this_item.Item(key_ProdQty), requestedQty, CMP_EQUAL, EXP_POS, "verifyShoppingCart 05", "Item quantity")
'			' EXAMPLE: Verify that the product page link ends as expected
'			isCorrect = chkVerifyParityEx(this_item.Item(key_ProdLink), expectTail, STR_TAIL, EXP_POS, "verifyShoppingCart 06". "Item product link")
'		--------------------------------------
'		<mode>_EQUAL: <sample> is equal to <master>
'		<mode>_LESS: <sample> is less than <master>
'		<mode>_LES_EQ: <sample> is less or equal to <master>
'		<mode>_MORE: <sample> is more than <master>
'		<mode>_MOR_EQ: <sample> is more ot equal to <master>
'		<mode>_NOT_EQ: <sample> is not equal to <master>
'		<mode>_HEAD: <sample> begins with <master>
'		<mode>_TAIL: <sample> ends with <master>
'		<mode>_BODY: <sample> contains <master>
'		
'@InParameter [in] expect Expectation (positive/negative), which determines PASS/FAIL status
'		EXP_POS: Positive expectation - Test passes if specified condition is true
'		EXP_NEG: Negative expectation - Test fails if specified condition is true
'@InParameter [in] stepStr Step name assigned to the evaluation in the log
'@InParameter [in] descStr Description of the specified sample value
'@ReturnValue 'True' if  verification passes; otherwise 'False'
Public Function chkVerifyParityEx(ByVal sample, ByVal master, ByVal compare, ByVal expect, ByVal stepStr, ByVal descStr)

	' Declarations
	Dim strPass, strFail
	Dim actual
	Dim status
	Dim detail
	Dim cmpMod, cmpStr, cmpMsg
	Dim lenMaster, strOffset
	Dim isAssert

	cmpMod = ""
	isAssert = False

	' If string compare specified
	If (compare And CMP_STRING) Then
		' Trim whitespace
		master = Trim(master)
		sample = Trim(sample)

		' If case matching specified
		If (compare And CMP_CASING) Then
			cmpMod = " (case matched)"
		' Otherwise (ignoring case)
		Else
			' Upcase strings
			master = UCase(master)
			sample = UCase(sample)
			cmpMod = " (case ignored)"
		End If
	End If

	' If containment evaluation specified
	If (compare And CMP_BODY) Then
		' Get length of master
		lenMaster = Len(master)

		' Differentiate specified containment
		Select Case (compare And CMP_BODY)

			Case CMP_HEAD
				' Retain left end of sample
				sample = Left(sample, lenMaster)
				cmpStr = "begin with "

			Case CMP_TAIL
				' Retain right end of sample
				sample = Right(sample, lenMaster)
				cmpStr = "end with "

			Case CMP_BODY
				cmpStr = "contain "

		End Select

		' Get offset of master in sample
		strOffset = InStr(1, sample, master)
		' Offset valid if found
		actual = (strOffset > 0)
		
		' If match found
		If (actual) Then
			cmpMsg = " does " & cmpStr
		' Otherwise (no match)
		Else
			cmpMsg = " does not " & cmpStr
		End If
	' Otherwise (equality evaluation)
	Else
		' Differentiate specified comparison
		Select Case (compare And CMP_ASSERT)

			Case CMP_EQUAL
				actual = (sample = master)
				cmpStr = "equal to "

			Case CMP_LESS
				actual = (sample < master)
				cmpStr = "less than "

			Case CMP_LES_EQ
				actual = (sample <= master)
				cmpStr = "less or equal to "

			Case CMP_MORE
				actual = (sample > master)
				cmpStr = "more than "

			Case CMP_MOR_EQ
				actual = (sample >= master)
				cmpStr = "more or equal to "

			Case CMP_NOT_EQ
				actual = (sample <> master)
				cmpStr = "different than "

			Case CMP_ASSERT
				isAssert = True
				actual = sample
				cmpStr = master
				
		End Select

		' If test positive
		If (actual) Then
			cmpMsg = " is " & cmpStr
		' Otherwise (test negative)
		Else
			cmpMsg = " is not " & cmpStr
		End If
	End If

	If (actual = expect) Then
		status = micPass
		detail = Not expect
		stepStr = stepStr & " Passed"
	Else
		status = micFail
		detail = expect
		stepStr = stepStr & " Failed"
	End If

	If (isAssert) Then
		descStr = descStr & cmpMsg
	Else
		descStr = descStr & cmpMsg & "master value" & cmpMod & vbCR & "EXPECT: " & cmpStr & master

		If (detail) Then
			descStr = descStr & vbCR & "ACTUAL: " & sample
		End If
	End If

	' Report results
	Reporter.ReportEvent status, stepStr, descStr

	chkVerifyParityEx = (status = micPass)

End Function


'@Description Verify that the specified sample value meets the indicated criteria
'@Documentation Verify that the specified sample value meets the indicated criteria
'@Author sbabcoc
'@Date 02-JUN-2011
'@InParameter [in] sample Sample value to be evaluated
'@InParameter [in] master Master value used for evaluation, or description of assertion for CMP_ASSERT
'@InParameter [in] compare Comparision used to evaluate sample value:
'		
'		BOOLEAN ASSERTION
'			' EXAMPLE: Verify that this shopping cart item has a product name
'			isCorrect = chkVerifyParity(this_item.Item(key_ProdName).Exists, "specified", CMP_ASSERT, "verifyShoppingCart 04", "Item name")
'		-------------------------------------
'		CMP_ASSERT: <sample> is Boolean ('True' or 'False'), described by the string in <master>
'		
'		COMPARISONS (SAMPLE VS. MASTER)
'			MODES: CMP = binary; STR = string; CAS = case-sensitive
'			' EXAMPLE: Verify that the expected quantity was added to the shopping cart
'			isCorrect = chkVerifyParity(this_item.Item(key_ProdQty), requestedQty, CMP_EQUAL, "verifyShoppingCart 05", "Item quantity")
'			' EXAMPLE: Verify that the product page link ends as expected
'			isCorrect = chkVerifyParity(this_item.Item(key_ProdLink), expectTail, STR_TAIL, "verifyShoppingCart 06". "Item product link")
'		--------------------------------------
'		<mode>_EQUAL: <sample> is equal to <master>
'		<mode>_LESS: <sample> is less than <master>
'		<mode>_LES_EQ: <sample> is less or equal to <master>
'		<mode>_MORE: <sample> is more than <master>
'		<mode>_MOR_EQ: <sample> is more ot equal to <master>
'		<mode>_NOT_EQ: <sample> is not equal to <master>
'		<mode>_HEAD: <sample> begins with <master>
'		<mode>_TAIL: <sample> ends with <master>
'		<mode>_BODY: <sample> contains <master>
'		
'@InParameter [in] stepStr Step name assigned to the evaluation in the log
'@InParameter [in] descStr Description of the specified sample value
'@ReturnValue 'True' if  verification passes; otherwise 'False'
Public Function chkVerifyParity(ByVal sample, ByVal master, ByVal compare, ByVal stepStr, ByVal descStr)

	chkVerifyParity = chkVerifyParityEx(sample, master, compare, EXP_POS, stepStr, descStr)

End Function


' -------------------------------------------------------------------------------------
'                                    chkVerifySimpleText
' -------------------------------------------------------------------------------------
'Used to send a simple verification statement to the QTP test report
'Pass expected, actual and test verification title into method. This
'performs a InStr text comparrisom, where the Expected text value 
'is expected to be found within the Actual text.

'Agrs: 
	'Exp_Text - expected value to comapare against
	'Act_Text - actual value found
	'sTitle - short descriptive title for the verification

	'Code Sample:
'		sTitle = "Verify 'Help' link text"
'		Exp_Text = "Live Help"
'		Act_Text = Browser("Site").Page("Page").Link("Live Help").GetRoProperty("innertext")
'		
'		chkVerifySimpleText(Exp_Text, Act_Text, sTitle)
Public Function chkVerifySimpleText(Exp_Text, Act_Text, sTitle)

	chkVerifySimpleText = chkVerifyParity(Act_Text, Exp_Text, STR_EQUAL, sTitle, "Specified string")

End Function


' -------------------------------------------------------------------------------------
'                                    chkVerifyExistence
' -------------------------------------------------------------------------------------
'@Description Verify the presence or absence of the specified object
'@Documentation Verify the presence or absence of the specified object
'@Author sbabcoc
'@Date 02-JUN-2011
'@InParameter [in] obj Object whose presence/absence is being evaluated
'@InParameter [in] desc Description of the specified object
'@InParameter [in] expect Existential expectation:
'		EXPECT_EXISTS: Expect object to be present
'		EXPECT_ABSENT: Expect object to be absent
'@InParameter [in] step Step name assigned to the evaluation in the log
'@ReturnValue 'True' if  verification passes; otherwise 'False'
Public Function chkVerifyExistence(obj, desc, expect, [step])

	' Declarations
	Dim typeStr
	Dim passStr
	Dim failStr
	Dim stepStr
	Dim descStr
	Dim status

	' Get object type
	typeStr = obj.GetTOProperty("micclass")

	' If expect to exist
	If (expect) Then
		passStr = "present"
		failStr = "absent"
	Else
		passStr = "absent"
		failStr = "present"
	End If

	' If existence meets expectations
	' NOTE: Type toggle works around bug
	If (CBool(CInt(obj.Exist(0))) = expect) Then
		status = micPass
		stepStr = [step] & " Passed"
		descStr = typeStr & " [" & desc & "] is " & passStr
	Else
		status = micFail
		stepStr = [step] & " Failed"
		descStr = typeStr & " [" & desc & "] is " & failStr
	End If

	' Report results
	Reporter.ReportEvent status, stepStr, descStr

	' RESULT: 'True' if OK; otherwise 'False'
	chkVerifyExistence = (status = micPass)
	
End Function


' -------------------------------------------------------------------------------------
'                                    chkVerifyInStrText
' -------------------------------------------------------------------------------------
'Used to send a InStr verification statement to the QTP test report
'Pass expected, actual and test verification title into method. This
'performs an InStr text comparison, where the expected text value 
'is expected to show within the larger actual text value.

'Agrs: 
	'Exp_Text - expected value to find within the Act_Text 
	'Act_Text - actual value found (expect to find Exp_Text in this string)
	'sTitle - short descriptive title for the verification

	'Code Sample:
'		sTitle = "Verify 'Help' link text"
'		Exp_Text = "Live Help"
'		Act_Text = Browser("Site").Page("Page").Link("Live Help").GetRoProperty("innertext")
'		
'		chkVerifyInStrText(Exp_Text, Act_Text, sTitle)

Public Function chkVerifyInStrText(Exp_Text, Act_Text, sTitle)

	chkVerifyInStrText = chkVerifyParity(Act_Text, Exp_Text, STR_BODY, sTitle, "Specified string")

End Function

' -------------------------------------------------------------------------------------
'                                    cmnVerifyObjectExist
' -------------------------------------------------------------------------------------
Public Function chkVerifyObjectExist(oObject)

	Set oObject = oObject
   sObjectName = oObject.GetROProperty("innertext")

	If oObject.Exist Then
		bResult = micPass
		sResult = "PASS"
		sStatus = "does"
	Else
		bResult = micFail	
		sResult = "FAIL"		
		sStatus = "does NOT"		
	End If

	sTitle = "Verify " +sObjectName+ " object exists"
	sDesc = sResult+ ": The " +sObjectName+ " object " +sStatus+ " exist"

	Call Reporter.ReportEvent (bResult, sTitle, sDesc)
	cmnVerifyObjectExist = bResult

End Function

' -------------------------------------------------------------------------------------------------
'                                          chkTBL_SingleRowValue
' -------------------------------------------------------------------------------------------------

'Focus: WebTable

'Desc: Verifys expected value occures once in specified column

'Args -

'	sDataSheetName -  Datasheet populated by the cmnGetTableContent function (Str)
'	sCol -  Column name (Str)
'	sValue - Cell value (Str)

'Req(s) - 

	'Common.vbs
	'cmnGetTableContent function 

' -------------------------------------------------------------------------------------------------
Public Function chkTBL_SingleRowValue(sDataSheetName, sCol, sValue)

	'Get Table row count
	iRowCount = DataTable.GetSheet(sDataSheetName).GetRowCount

	'Loop thru table rows seek expected value...
	For iRow = 1 To iRowCount

		'Iterate thru rows in specified column seek expected value
		DataTable.GetSheet(sDataSheetName).SetCurrentRow iRow
		sCelValue = DataTable(sCol, sDataSheetName)
		
		'If expected value found...
		If Trim(LCase(sCelValue)) = Trim(LCase(sValue)) Then
			bStatus = micPass
			sRowFound = CStr(iRow)
			Exit For
		'If expected value not found...
		ElseIf iRow = iRowCount Then
			bStatus = micFail
			Exit For
		End If
		
		'Remove leading camma if value isn't empty
		If bStatus = micFail Then
			sRowFound = "zero"
		End If

	Next

	'Report pass/fail result to test
	sTitle = "Verify Table Value - Single Row" 
	sDesc = "Verify " +sValue+ " appears in ONE ROW of the " +sCol+ " Column" +vbcr+vbcr+ "Value was found in the following row '" +sRowFound+ "'" +vbcr+vbcr+ "Expected: " + sValue +vbcr+ "Actual: " +sCelValue
	Call Reporter.ReportEvent (bStatus, sTitle, sDesc)	

	chkTBL_SingleRowValue = bStatus
	
End Function

' -------------------------------------------------------------------------------------------------
'                                           chkTBL_MultiRowValue
' -------------------------------------------------------------------------------------------------

'Focus: WebTable

'Desc: Verifys expected value occures multiple times in specified column

'Args -

'	sDataSheetName -  Datasheet populated by the cmnGetTableContent function (Str)
'	sCol -  Column name (Str)
'	sValue - Cell value (Str)

'Req(s) - 

	'Common.vbs
	'cmnGetTableContent function 

' -------------------------------------------------------------------------------------------------
Public Function chkTBL_MultiRowValue(sDataSheetName, sCol, sValue)

	'Get Table row count
	iRowCount = DataTable.GetSheet(sDataSheetName).GetRowCount
	
	'Loop thru table rows seek expected value...
	For iRow = 1 To iRowCount

		'Iterate thru rows in specified column seek expected value
		DataTable.GetSheet(sDataSheetName).SetCurrentRow iRow
		sCelValue = DataTable(sCol, sDataSheetName)
		
		'If expected value found...
		If Trim(LCase(sCelValue)) = Trim(LCase(sValue)) Then
			bStatus = micPass
			sRowFound = sRowFound + ", " +CStr(iRow)
		'If expected value not found...
		ElseIf iRow = iRowCount AND "" = iRows Then
				bStatus = micFail
		End If

		'Remove leading camma if value isn't empty
		If bStatus <> micFail Then
			sRowFound = Trim(Right(sRowFound, Len(sRowFound)-1))
		Else
			sRowFound = "zero"
		End If

	Next

	'Report pass/fail result to test
	sTitle = "Verify Table Value - Multi Rows" 
	sDesc = "Verify " +sValue+ " appears in MULTIPLE ROWS of the " +sCol+ " Column" +vbcr+vbcr+ "Value was found in the following row '" +sRowFound+ "'" +vbcr+vbcr+ "Expected: " + sValue +vbcr+ "Actual: " +sCelValue
	Call Reporter.ReportEvent (bStatus, sTitle, sDesc)	

	chkTBL_MultiRowValue = bStatus
	
End Function

' -------------------------------------------------------------------------------------------------
'                                           chkTBL_AllRowValue
' -------------------------------------------------------------------------------------------------

'Focus: WebTable

'Desc: Verifys expected value occures in all rows of specified column

'Args -

'	sDataSheetName -  Datasheet populated by the cmnGetTableContent function (Str)
'	sCol -  Column name (Str)
'	sValue - Cell value (Str)

'Req(s) - 

	'Common.vbs
	'cmnGetTableContent function 

' -------------------------------------------------------------------------------------------------
Public Function chkTBL_AllRowValue(sDataSheetName, sCol, sValue)

	'Get Table row count
	iRowCount = DataTable.GetSheet(sDataSheetName).GetRowCount
	
	'Loop thru table rows seek expected value...
	For iRow = 1 To iRowCount

		'Iterate thru rows in specified column seek expected value
		DataTable.GetSheet(sDataSheetName).SetCurrentRow iRow
		sCelValue = DataTable(sCol, sDataSheetName)
		
		'If expected value found...
		If Trim(LCase(sCelValue)) = Trim(LCase(sValue)) Then
			bStatus = micPass
			sRowFound = "all" 
		'If expected value not found...
		Else
				bStatus = micFail
				sRowFound = CStr(iRow)	   
				Exit For
		End If
	Next

	'Report pass/fail result to test
	sTitle = "Verify Table Value - All Rows" 
	sDesc = "Verify " +sValue+ " appears in ALL ROWS of the " +sCol+ " Column" +vbcr+vbcr+ "Value was found in the following row '" +sRowFound+ "'" +vbcr+vbcr+ "Expected: " + sValue +vbcr+ "Actual: " +sCelValue
	Call Reporter.ReportEvent (bStatus, sTitle, sDesc)	

	chkTBL_AllRowValue = bStatus
	
End Function

' -------------------------------------------------------------------------------------
'                                             chkLinkExist
' -------------------------------------------------------------------------------------
'Desc: Verifies link exists or not

'Args:

	'sPage =  Place the following text into the arg: 
	
						'Set oPage = Browser("myBrowser").Page("MyPage")
						'Note - that this text can be placed in a datasheet.
						
	'sLinkInnerText = the innettext of the link to be checked
' -------------------------------------------------------------------------------------
Public Function chkLinkExist(sPageRef, sLinkInnerText)
   Call cmnGetGlobalTimeOuts_QC()

	Execute (sPageRef)

	' -----------------------------------------------------------
	'                          Verify Link Exists
	' -----------------------------------------------------------   
	'Verify link exists on page
	Call cmnSetGlobalTimeouts (CInt(DataTable("SpecialTimeOut", dtGlobalSheet)))
	If Not oPage.Link("innertext:=" +sLinkInnerText).Exist Then
		bResult = micFail
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))
		
	Else
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))
		
		'Post verification to test results
		bResult = micPass
		sTitle = "Verify " +sLinkInnerText+ " link exists on page"
		sDesc = sLinkInnerText+ " link was expected on page"
		Call Reporter.ReportEvent (bResult, sTitle, sDesc)
	End If

End Function

' -------------------------------------------------------------------------------------
'                                             chkLinkExist_byPgTitle
' -------------------------------------------------------------------------------------
'Desc: Verifies link exists or not

'Args:

	'sPage =  Place the following text into the arg: 
	
						'Set oPage = Browser("myBrowser").Page("MyPage")
						'Note - that this text can be placed in a datasheet.
						
	'sLinkInnerText = the innettext of the link to be checked
' -------------------------------------------------------------------------------------
Public Function chkLinkExist_byPgTitle(sPageTitle, sLinkInnerText)

   Call cmnGetGlobalTimeOuts_QC()
	' -----------------------------------------------------------
	'                          Verify Link Exists
	' -----------------------------------------------------------   
	'Verify link exists on page
	Call cmnSetGlobalTimeouts (CInt(DataTable("SpecialTimeOut", dtGlobalSheet)))
	If Not Browser("title:=" + sPageTitle).Page("title:=" +sPageTitle).Link("innertext:=" +sLinkInnerText).Exist Then
		bResult = micFail
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))
		
	Else
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))
		
		'Post verification to test results
		bResult = micPass
		sTitle = "Verify " +sLinkInnerText+ " link exists on page"
		sDesc = sLinkInnerText+ " link was expected on page"
		Call Reporter.ReportEvent (bResult, sTitle, sDesc)
	End If

	'return micPass/Fail
	chkLinkExist_byPgTitle = bResult

End Function

' -------------------------------------------------------------------------------------
'                                   chkLinkNavigation
' -------------------------------------------------------------------------------------
'Desc: Verifies link navigation

'Args:

	'oBrowser = Set oBrowser = Browser("myBrowser")
	'oPage =  Set oBrowser = Browser("myBrowser").Page("MyPage")
	'sLinkInnerText = the innettext of the link to be checked
' -------------------------------------------------------------------------------------
Public Function chkLinkNavigation(oBrowser, oPage, sLinkInnerText)

   	Set oBrowser = oBrowser
	Set oPage = oPage

	'Get link info
	 sLinkURL = oPage.Link("innertext:=" +sLinkInnerText).GetROProperty("href")
	'Select link
	oPage.Link("innertext:=" +sLinkInnerText).Click

' -----------------------------------------------------------
'                      Verify Link Redirect
' -----------------------------------------------------------   
	'Verify link navs to correct page
	If oBrowser.Page("href:=" +sLinkURL).Exist Then
		bResult = micPass
	Else
		bResult = micFail
	End If

	'Post verification to test results
	sTitle = "Verify " +sLinkInnerText+ "  link navs to correct page"
	sDesc = sLinkInnerText+ " link expected to nav to URL:  " +sLinkURL
	Call Reporter.ReportEvent (bResult, sTitle, sDesc)

	'Return to start pg
	oBrowser.Back
	
End Function

' -------------------------------------------------------------------------------------
'                                            VerifyHttps
' -------------------------------------------------------------------------------------
' This function verifies that the current browser open is in https mode
Public Function VerifyHttps()

	Dim URL
	
	'URL = Browser("micclass:=Browser").Page("micclass:=Page").GetROProperty("url")
	URL = Browser("Browser").Object.LocationURL
	
	If Mid (URL,1,5) = "https" Then 
		VerifyHttps = True
	Else 
		VerifyHttps = False
	End if
	
End Function

' Function VerifyProperty
' -----------------------
' Verify the value of a specified property
' Parameters:
'	PropertyName - the property name to check
'       ExpectedValue - the expected value of the property
' Returns - True - if the expected value matches the actual value
'
'@Description Checks whether a property value matches its expected value
'@Documentation Check whether  the <Test object name> <test object type> <PropertyName> property value matches the expected value: <ExpectedValue>.
Public Function chkVerifyProperty (obj, PropertyName, ExpectedValue)
	Dim actual_value
	' Get the actual property value
	actual_value = obj.GetROProperty(PropertyName)
	' Compare the actual value to the expected value
	If actual_value = ExpectedValue Then
		Reporter.ReportEvent micPass, "VerifyProperty Succeeded", "The " & PropertyName & " expected value: " & ExpectedValue & " matches the actual value"
		VerifyProperty = True
	Else
		Reporter.ReportEvent micFail, "VerifyProperty Failed", "The " & PropertyName & " expected value: " & ExpectedValue & " does not match the actual value: " & actual_value
		VerifyProperty = False
	End If
End Function

' Function VerifyEnabled
' -------------------------
' Verify whether a specified object is enabled
' Returns - True - if the test object is enabled
'
'@Description Checks whether the specified test object is enabled
'@Documentation Check whether the <Test object name> <test object type> is enabled.
Public Function chkVerifyEnabled (obj)
	Dim enable_property
	' Get the enabled property from the test object
	enable_property = obj.GetROProperty("enabled")
	If enable_property <> 0 Then ' The value is True (anything but 0)
		Reporter.ReportEvent micPass, "VerifyEnabled Succeeded", "The test object is enabled"
		chkVerifyEnabled = True
	Else
		Reporter.ReportEvent micFail, "VerifyEnabled Failed", "The test object is NOT enabled"
		chkVerifyEnabled = False
	End If
End Function


' Function VerifyDisabled
' -------------------------
' Verify whether a specified object is enabled (not disabled)
' Returns - True - if the test object is enabled (not disabled)
'
'@Description Checks whether the specified test object is enabled
'@Documentation Check whether the <Test object name> <test object type> is enabled.
Public Function chkVerifyDisabled (obj)
	Dim enable_property
	' Get the disabled property from the test object
	enable_property = obj.GetROProperty("disabled")
	If enable_property = 0 Then ' The value is False (0) - Enabled
		Reporter.ReportEvent micPass, "VerifyDisabled Succeeded", "The test object is enabled"
		chkVerifyDisabled = True
	Else
		Reporter.ReportEvent micFail, "VerifyDisabled Failed", "The test object is NOT enabled"
		chkVerifyDisabled = False
	End If
End Function

'--------------------------------------------------
' In order to implement the 'VerifyValue' function for all QuickTest Professional test objects,
' there is collection of functions that returns a specific property that represent
' the 'Value' of a test object.

' Function VerifyValueProperty
' --------------------
' Check the value for a specified object
' Returns - True - if the expected Value matches the actual value
'Parameter:
'	ExpectedValue  - the expected value
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedValue>
Public Function chkVerifyValue (obj, ExpectedValue)
	chkVerifyValue = chkVerifyProperty (obj, "value", ExpectedValue)
End Function

' Function VerifyTextProperty
' --------------------
' Check the text for a specified object
' Returns - True - if the expected text matches the actual text
'Parameter:
'	ExpectedText  - the expected text
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedText>.
Public Function chkVerifyText (obj, ExpectedText)
	chkVerifyText = chkVerifyProperty (obj, "text", ExpectedText)
End Function

' Function VerifyDateProperty
' --------------------
' Check the date for a specified object
' Returns - True - if the expected date matches the actual value
'Parameter:
'	ExpectedDate  - the expected date
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedDate>
Public Function chkVerifyDate (obj, ExpectedDate)
	chkVerifyDate = chkVerifyProperty (obj, "date", ExpectedDate)
End Function

' Function VerifyCheckedProperty
' --------------------
' Check the chceked property for a specified object
' Returns - True - if the expected Value matches the actual value
'Parameter:
'	ExpectedChecked  - the expected value
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedChecked>
Public Function chkVerifyChecked (obj, ExpectedChecked)
	chkVerifyChecked = chkVerifyProperty (obj, "checked", ExpectedChecked)
End Function

' Function VerifySelectionProperty
' --------------------
' Check the selection property for a specified object
' Returns - True - if the expected Value matches the actual value
'Parameter:
'	ExpectedSelection  - the expected value
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedSelection>
Public Function chkVerifySelection (obj, ExpectedSelection)
	chkVerifySelection = chkVerifyProperty (obj, "selection", ExpectedSelection)
End Function

' Function VerifyPositionProperty
' --------------------
' Check the position property for a specified object
' Returns - True - if the expected Value matches the actual value
'Parameter:
'	ExpectedPosition  - the expected value
'@Description Checks the object value
'@Documentation Check whether the <Test object name> <test object type> value matches the expected value: <ExpectedPosition>
Public Function chkVerifyPosition (obj, ExpectedPosition)
	chkVerifyPosition = chkVerifyProperty (obj, "position", ExpectedPosition)
End Function

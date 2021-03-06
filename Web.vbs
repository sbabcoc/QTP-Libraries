'********************************************************************************
'	Web Functions
'       -------------------------
'
'   Available Functions:
'	* VerifyProperty - Verifies the value of a specified property (for all Web test objects)
'	* OutputProperty - Returns the value of the specified property (for all Web test objects)
'	* VerifyEnable - Verifies whether a specified object is enabled (for all Web test objects)
'	* VerifyValue - Verifies the value of a specified object (for WebCheckBox, WebEdit, WebFile, WebList, WebRadioGroup)
'	* GetValue - Returns the object value (for WebCheckBox, WebEdit, WebFile, WebList, WebRadioGroup)
'
'   Version: QTP9.0 November 2005
'   ** In order to use the functions in this file, you must also load the "Common.txt" function library file.
'
'********************************************************************************

Option Explicit

' Function Sync
' -----------------------
' Supply default method that returns immediately, to be used in the Frame test object.
' Returns - True.

'@Description Waits for the  test  object to synchronize
'@Documentation Wait for the <Test object name> <Test object type> to synchronize before continuing the run.
Public Function SyncFrame(obj)
	Sync = True
End Function


'@Description Waits for the  test  object to synchronize
'@Documentation Wait for the <Test object name> <Test object type> to synchronize before continuing the run.
'@Note This function performs a standard Sync, then runs recovery scenarios if the current scheme is HTTPS
'@
Public Function Sync(obj)
	obj.Sync
	thisURL = obj.GetROProperty("url")
	sLength = InStr(1, thisURL, ":")
	scheme = Left(thisURL, sLength)
	If (scheme = "https:") Then
		Recovery.Activate
	End If
End Function

' Function VerifyWebEnabled
' -------------------------
' Verify whether a specified object is enabled
' Returns - True - if the test object is enabled
' 
'@Description Checks whether the specified test object is enabled
'@Documentation Check whether the <Test object name> <test object type> is enabled.
Public Function VerifyWebEnabled(obj)
	Dim disable_property
	' Get the enabled property from the test object
	disable_property = obj.GetROProperty("disabled")
	If disable_property = 0 Then ' The value is True (anything but 0)
		Reporter.ReportEvent micPass, "VerifyEnabled Succeeded", "The test object is enabled"
		VerifyWebEnabled = True
	Else
		Reporter.ReportEvent micFail, "VerifyEnabled Failed", "The test object is NOT enabled"
		VerifyWebEnabled = False
	End If
End Function


'@Description 
'@Documentation 
'@Author sbabcoc
'@Date 28-MAR-2011
'@InParameter [in] oRadioGroup, reference, reference to WebRadioGroup
'@InParameter [in] sSelectText, string, text of radio button to select
Public Sub SelectByText ( ByRef oRadioGroup, sSelectText) 

   ' NOTE: This code uses the DOM rather than native QTP objects
   ' TODO: Re-code to use the QTP object model rather than DOM

	Dim oParent, groupName, oRadioButtons, oElement 
	Dim sName 
	
	Set oParent = oRadioGroup.GetTOProperty("parent") 
	groupName = oRadioGroup.GetROProperty("name") 
	Set oRadioButtons = oParent.Object.GetElementsByName( groupName ) 
	
	For Each oElement In oRadioButtons 
		sName = Trim (oElement.getAdjacentText("afterEnd")) 
		If Trim(sSelectText) = Left( sName, Len(Trim (sSelectText)) ) Then 
			oRadioGroup.Select oElement.Value 
		End If 
	Next 
	
	Set oParent = Nothing 
	Set oRadioButtons = Nothing 
	
End Sub


'@Description Select the indicated item(s) in the specified list object
'@Documentation Select <selectSpec> in the <webListObj>
'@Author sbabcoc
'@Date 28-MAR-2011
'@InParameter [in] webListObj, reference, reference to WebList
'@InParameter [in] selectSpec, string/array, item (or list of items) to select
'@InParameter [in] byValue, boolean, 'True' to select by value; 'False' to select by text/index
Public Sub webListSelect(ByRef webListObj, ByRef selectSpec, byValue)

	If (VarType(selectSpec) And vbArray) Then
		selectList = selectSpec
		selectQty = UBound(selectList) + 1
	Else
		selectList = Array(selectSpec)
		selectQty = 1
	End If

	If (byValue) Then
		Set valueDict = webListExtractValues(webListObj)
		listIndex = 0
		Do While (listIndex < selectQty)
			selectSpec = selectList(listIndex)
			If valueDict.Exists(selectSpec) Then
				valueSpec = valueDict.Item(selectSpec)
				valueBits = Split(valueSpec, "|")
				indexSpec = valueBits(0)
			Else
				indexSpec = ""
			End If
			selectList(listIndex) = indexSpec
		Loop
	End If

	listCount = webListObj.GetROProperty("items count")
	listType = webListObj.GetROProperty("select type")

	Select Case listType

		Case "Single Selection"
		Case "ComboBox Select"
			webListObj.Select selectList(0)

		Case "Extended Selection"
			webListObj.Select selectList(0)
			
			listIndex = 1
			Do While (listIndex < selectQty)
				webListObj.ExtendSelect selectList(listIndex)
				listIndex = listIndex + 1
			Loop
	
	End Select

End Sub


'@Description Extract 
'@Documentation Extract 
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] refWebList A reference to a WebList object
'@ReturnValue A dictionary of WebList values
Public Function webListExtractValues(ByRef refWebList)

   ' Declarations
	Dim optionsDict
	Dim regEx, matchList
	Dim offset, listIndex

	' Allocate RegExp object
	Set regEx = New RegExp

	' Create a product options dictionary
	Set optionsDict = CreateObject("Scripting.Dictionary")

	' Get inner HTML of options list
	innerHTML = refColorSize.GetROProperty("innerhtml")
	' Locate start of first option
	offset = InStr(1, innerHTML, "<")
	' Extract leading text (if any)
	listExtra = Left(innerHTML, offset - 1)
	' Trim off leading text (if any)
	innerHTML = Mid(innerHTML, offset)
	' Split into a list of options
	optionList = Split(innerHTML, "</OPTION>")

	' Iterate over options
	For listIndex = 0 To (UBound(optionList) - 1)
		' Extract current option
		thisOption = optionList(listIndex)
		' Locate end of option tag
		offset = InStr(1, thisOption, ">")
		' Extract tag for current option
		optionTag = Left(thisOption, offset)
		' Extract text of current option
		optionText = Mid(thisOption, offset + 1)
		
		' Define pattern to extract value
		regEx.Pattern = "value=(\d+)"
		' Apply pattern to extract value
		Set matchList = regEx.Execute(optionTag)
		' If option value was captured
		If (matchList.Count = 1) Then
			' Extract captured option value
			optionVal = matchList(0).SubMatches(0)
		' Otherwise
		Else
			' No option value
			optionVal = ""
		End If

		' Add option specification to dictionary
		optionsDict.Add optionVal, "#" & listIndex & "|" & optionText
	Next

	' RESULT: Option specifications dictionary
	Set product_ExtractOptions = optionsDict

	' Release objects
	Set optionsDict = Nothing
	Set regEx = Nothing
	Set matchList = Nothing

End Function


' *********************************************************************************************
' *** 			Register the Functions
' *********************************************************************************************

' Register the "VerifyProperty" Function
RegisterUserFunc "Browser" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "Frame" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "Image" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "Link" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "ViewLink" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "Page" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebArea" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebButton" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebCheckBox" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebEdit" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebElement" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebFile" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebList" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebRadioGroup" , "VerifyProperty" , "VerifyProperty"
RegisterUserFunc "WebTable" , "VerifyProperty" , "VerifyProperty"

' Register the "OutputProperty" Function
RegisterUserFunc "Browser" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "Frame" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "Image" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "Link" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "ViewLink" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "Page" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebArea" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebButton" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebCheckBox" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebEdit" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebElement" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebFile" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebList" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebRadioGroup" , "OutputProperty" , "OutputProperty"
RegisterUserFunc "WebTable" , "OutputProperty" , "OutputProperty"

' Register the "VerifyValue" Function
RegisterUserFunc "WebCheckBox" , "VerifyValue" , "VerifyChecked"
RegisterUserFunc "WebEdit" , "VerifyValue" , "VerifyValue"
RegisterUserFunc "WebFile" , "VerifyValue" , "VerifyValue"
RegisterUserFunc "WebList" , "VerifyValue" , "VerifyValue"
RegisterUserFunc "WebRadioGroup" , "VerifyValue" , "VerifyValue"

' Register the "GetValue" Function
RegisterUserFunc "Link" , "GetValue" , "GetTextProperty"
RegisterUserFunc "WebCheckBox" , "GetValue" , "GetCheckedProperty"
RegisterUserFunc "WebEdit" , "GetValue" , "GetValueProperty"
RegisterUserFunc "WebFile" , "GetValue" , "GetValueProperty"
RegisterUserFunc "WebList" , "GetValue" , "GetValueProperty"
RegisterUserFunc "WebRadioGroup" , "GetValue" , "GetValueProperty"

' Register the "VerifyEnable" Function
RegisterUserFunc "WebButton" , "VerifyEnable" , "VerifyWebEnabled"
RegisterUserFunc "WebCheckBox" , "VerifyEnable" , "VerifyWebEnabled"
RegisterUserFunc "WebEdit" , "VerifyEnable" , "VerifyWebEnabled"
RegisterUserFunc "WebFile" , "VerifyEnable" , "VerifyWebEnabled"
RegisterUserFunc "WebList" , "VerifyEnable" , "VerifyWebEnabled"
RegisterUserFunc "WebRadioGroup" , "VerifyEnable" , "VerifyWebEnabled"

' Register the Sync Function
RegisterUserFunc "Frame", "Sync", "SyncFrame", True
RegisterUserFunc "Browser", "Sync", "Sync"
RegisterUserFunc "Page", "Sync", "Sync"

' -------------------------------------------------------------------------------------
'                                    DeleteCookies
' -------------------------------------------------------------------------------------
' DeleteCookies only works with IE
Public Sub DeleteCookies()

	' NOTE: We could use WebUtil.DeleteCookies for this.
	' However, this is an undocumented object/function.

	' Delete Cookies folder items
	DeleteFolderItems 33, "*.txt"
	' Delete Internet Temp Files
	DeleteIETemporaryFiles

End Sub


' -------------------------------------------------------------------------------------
'                                    DeleteIETemporaryFiles()
' -------------------------------------------------------------------------------------
'Desc:
'	Function deletes IE temporary files
'
'Args:
'	None
'
'Usage:
'Call DeleteIETemporaryFiles()
' -------------------------------------------------------------------------------------
Public Sub DeleteIETemporaryFiles()

	' Delete Temp Files items
	DeleteFolderItems 32, "*.*"

End Sub


'@Description Delete matching item(s) from the specified folder
'@Documentation Delete item(s) matching <itemSpec> from <folderSpec>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] folderSpec, string/integer, folder specifier (path or ShellSpecialFolderConstants)
'@InParameter [in] itemSpec, string, item name or pattern
Public Sub DeleteFolderItems(folderSpec, itemSpec)

	' Declarations
	Dim shellObj
	Dim folderObj
	Dim folderItemsObj
	Dim fileSystemObj
	Dim folderPath
	Dim itemPath

	Wait 0
	On Error Resume Next

	' Create a shell object
	Set shellObj = CreateObject("Shell.Application")
	' Get  folder object
	Set folderObj = shellObj.NameSpace(folderSpec)
	' If folder object obtained
	If Not (folderObj Is Nothing) Then
		' Get folder items collection
		Set folderItemsObj = folderObj.Items
		' If folder items collection obtained
		If Not (folderItemsObj Is Nothing) Then
			' If collection is non-empty
			If (folderItemsObj.Count) Then
				' Create a file system object
				Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
				' Get Cookies folder path
				folderPath = folderObj.Self.Path
				' Assemble path to item(s)
				itemPath = folderPath & "\" & itemSpec
				' Delete item(s) (including read-only)
				fileSystemObj.DeleteFile itemPath, True
			End If
		End If
	End If

	' Release objects
	Set shellObj = Nothing
	Set folderObj = Nothing
	Set folderItemsObj = Nothing
	Set fileSystemObj = Nothing

End Sub


'@Description Close browser windows, with option to include Quality Center windows
'@Documentation Close browser windows, including Quality Center windows if <includeQC> is 'True'
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] includeQC, boolean, 'True' to close Quality Center windows; otherwise 'False'
Public Sub CloseBrowsers(includeQC)

	' Declarations
	Dim descBrowser, descPage
	Dim browserList, pageList
	Dim listIndex
	Dim doClose

	' Load browser and page object descriptions
	Set descBrowser = Description.Create()
	descBrowser("micclass").Value = "Browser"
	Set descPage = Description.Create()
	descPage("micclass").Value = "Page"

	' Get list of browser objects
	Set browserList = Desktop.ChildObjects(descBrowser)
   	' Iterate over browser objects
	For listIndex = 0 To (browserList.Count - 1)
		' Assume close
		doClose = True
		' Get list of page objects
		Set pageList = browserList.Item(listIndex).ChildObjects(descPage)

		' If not closing Quality Center pages
		If Not (includeQC) Then
			' If this is a Quality Center page
			If (InStr(1, pageList.Item(0).GetROProperty("title"), QC_PAGE_TAG)) Then
				' Don't close this
				doClose = False
			End If
		End If
		
		' If closing this browser, do so now
		If (doClose) Then browserList.Item(listIndex).Close
	Next

End Sub


' ---------------------------------------------------------------------------------------------------
'                                         cmnOpenBrowser ( )
' ---------------------------------------------------------------------------------------------------

'Desc - 
	'Launches a specified browser type (IE, NS, FF, OP) based with a 
	'specified webPage loaded.

' Args - 

	'sBrowserID -  2-letter code representing the following browser types

		'IE - Internet Explorer
		'NC - Netscape
		'FF - FireFox
		'OP - Opera
		
	'sLoadURL - The URL to the requested webpage
		'Example: "http://www.hotmail.com"

' ---------------------------------------------------------------------------------------------------
Public Function cmnOpenBrowser (sBrowserID, sLoadURL)

	On Error Resume Next

	' Declarations 
	Dim oBrowser

	' Initialization
	Set oBrowser = Nothing

	' Get browser type
	Select Case UCase(sBrowserID)
		Case "IE"
			Set oBrowser = CreateObject("InternetExplorer.Application")
			
		Case "NS"
			Set oBrowser = CreateObject("NetscapeNavigator.Application")
			
		Case "FF"
			Set oBrowser = CreateObject("FireFox.Application")
			
        Case "OP"
			Set oBrowser = CreateObject("Opera.Application")
	End Select

	' If browser object was created
	If Not (oBrowser = Nothing) Then
		'Launch webpage in browser
		oBrowser.Visible = True
		'Navigate to URL
		oBrowser.Navigate sLoadURL 
		' Synchronize with browser
		Browser("name:=.*").Page("name:=.*").Sync
		' Reset error count
		Environment("ErrorCount") = 0
		' Indicate success
		cmnOpenBrowser = True
	Else ' otherwise (no browser)
		' Report the error
		Reporter.ReportEvent micFail, "cmnOpenBrowser", Err.Description
		' Indicate failure
		cmnOpenBrowser = False
	End If
	
End Function


'@Description Split the specified URL into a dictionary of parameter settings
'@Documentation Split <theURL> into a dictionary of parameter settings
'@Author sbabcoc
'@Date 05-APR-2011
'@InParameter [in] theURL, string, The URL from which to extract parameters
'@ReturnValue A dictionary of parameter settings, including the base URL
Public Function SplitURL(theURL)

	' Declarations
	Dim urlDict
	Dim urlBits
	Dim base_url
	Dim parmList
	Dim thisParm
	Dim parmBits
	Dim parmKey
	Dim parmVal

	' Allocate dictionary 
	Set urlDict = CreateObject("Scripting.Dictionary")

	' Split URL from parameters
	urlBits = Split(theURL, "?", -1)
	' If empty URL was specified
	If (UBound(urlBits) = -1) Then
		' Set empty URL
		base_url = ""
	Else
		' Extract base URL
		base_url = urlBits(0)
	End If
	
	' Add base URL to result
	urlDict.Add key_base_url, base_url
	
	' If URL includes parameters
	If (UBound(urlBits) > 0) Then
		' Extract parameter list
		parmList = Split(urlBits(1), "&", -1)
		' Iterate over parameter list
		For Each thisParm in parmList
			' Extract parameter bits
			parmBits = Split(thisParm, "=", -1)
			' Extract parameter key
			parmKey = parmBits(0)

			' Extract parameter value
			Select Case UBound(parmBits)
				Case 0 ' No value specified
					' Set value to ampersand
					parmVal = "&"
					
				Case 1 ' Value without equals
					' Set value as specified
					parmVal = parmBits(1)
					
				Case Else ' Value contains equals
					' Reconstitute specified value
					parmBits(0) = ""
					parmVal = Mid(Join(parmBits, "="), 2)
					
			End Select

			' Add parameter to result
			urlDict.Add parmKey, URLDecode(parmVal)
		Next
	End If

	Set SplitURL = urlDict

End Function


'@Description URL decode the specified string
'@Documentation URL decode <str>
'@Author sbabcoc
'@Date 06-APR-2011
'@InParamater [in] str, string, The URL string to be decoded
'@ResultValue The decoded URL string
Function URLDecode(strEncoded)

	Dim i
	Dim thisChar
	Dim strDecoded

	' Initialize output
	strDecoded = ""

	' If encoded string is non-NULL
	If Not IsNull(strEncoded) Then
		' Iterate over encoded chars
		For i = 1 To Len(strEncoded)
			' Extract current encoded char
			thisChar = Mid(strEncoded, i, 1)
			' If this is a "plus" sign
			If (thisChar = "+") Then
				' Append space character
				strDecoded = strDecoded & " "
			' Otherwise, if this is a percent
			ElseIf (thisChar = "%") Then
				' If there are at least two more chars
				If ((i + 2) <= Len(strEncoded)) Then
					' Append percent-encoded character
					strDecoded = strDecoded & Chr(CLng("&H" & Mid(strEncoded, i + 1, 2)))
					i = i + 2
				End If 
			' Otherwise (un-encoded)
			Else 
				' Append un-encoded character
				strDecoded = strDecoded & thisChar 
			End If 
		Next
	End If

	URLDecode = strDecoded

End Function 


'@Description URL encode the specified string
'@Documentation URL encode <str>
'@Author sbabcoc
'@Date 06-APR-2011
'@InParamater [in] str, string, The URL string to be encoded
'@ResultValue The encoded URL string
Function URLEncode(strDecoded)

	' Declarations
	Dim i
	Dim thisChar
	Dim thisCode
	Dim encoded
	Dim strEncoded

	' Initialize output
	strEncoded = ""
	
	' If decoded string is non-NULL
	If Not IsNull(strDecoded) Then
		' Iterate over decoded chars
		For i = 1 To Len(strDecoded)
			' Extract current decoded char
			thisChar = Mid(strDecoded, i, 1)
			' Get ASCII code for char
			thisCode = Asc(thisChar)

			' If this is a space
			If (thisChar = " ") Then
				' Encode space
				encoded = "+"
			' Otherwise, if this is an un-encoded symbol
			ElseIf (InStr(1, "-_.!~*'()", thisChar) > 0) Then
				' Pass thru as-is
				encoded = thisChar
			' Otherwise, if this is a number (0 thru 9)
			ElseIf ((thisCode > 47) And (thisCode <= 57)) Then
				' Pass thru as-is
				encoded = thisChar
			' Otherwise, if this is an uppercase letter (A thru Z)
			ElseIf ((thisCode > 64) And (thisCode <= 90)) Then
				' Pass thru as-is
				encoded = thisChar
			' Otherwise, if this is a lowercase letter (a thru z)
			ElseIf ((thisCode > 96) And (thisCode <= 122)) Then
				' Pass thru as-is
				encoded = thisChar
			' Otherwise
			Else
				' Get ASCII code in hex, pad left
				encoded = "0" & Hex(thisCode)
				' Percent-encode two-digit hex
				encoded = "%" & Right(encoded, 2)
			End If

			' Append encoded character
			strEncoded = strEncoded & encoded
		Next
	End if

	URLEncode = strEncoded

End Function 


'@Description Compare link URL with expected URL
'@Documentation Compare <lnkURL> with <testURL>
'@Author sbabcoc
'@Date 06-APR-2011
'@InParamater [in] linkURL, string, The actual URL of the link being compared
'@InParameter [in] testURL, string, The expected URL for the link
'@ResultValue 'True' if the actual URL meets with expectation; otherwise 'False'
Public Function CompareLink(linkURL, testURL)

	' Declarations
	Dim thisURL
	Dim schmLen
	Dim scheme

	' Set default result
	CompareLink = False

	' If expected URL is defined
	If Not IsNull(testURL) Then
		' If actual URL meets with expectation
		If (linkURL = testURL) Then
			' RESULT: Success
			CompareLink = True
		' Otherwise, if expecting a URL
		ElseIf (testURL <> "") Then
			' Get URL of current  page
			thisURL = Browser("name:=.*").Page("name:=.*").GetROProperty("url")
			' Get length of URL scheme
			schmLen = InStr(1, thisURL, ":")
			' Get URL scheme string
			scheme = Left(thisURL, schmLen)
		
			' Is current page is secure
			If (scheme = "https:") Then
				' Get length of URL scheme
				schmLen = InStr(1, testURL, ":")
				' Match scheme of test URL with current page
				testURL = scheme & Mid(testURL, schmLen + 1)
	
				' If actual URL meets with expectation
				If (linkURL = testURL) Then
					' RESULT: Success
					CompareLink = True
				End If
			End If
		End If
	End If

End Function


'@Description Extract input data from specified form object
'@Documentation Extract input data from <formObject>
'@Author sbabcoc
'@Date 01-JUN-2011
'@InParameter [in] formObject, reference, The form object from which to extract input data
'@ReturnValue A dictionary of input data items
Public Function ExtractData(formObject)

	' Declarations
	Dim dataDict
	Dim matchExp, matchList, thisMatch
	Dim inputText, inputKeep
	Dim inputName, inputValue
	Dim parmList, thisParm

	' Allocate dictionary 
	Set dataDict = CreateObject("Scripting.Dictionary")
	
	' Allocate RegExp object
	Set matchExp = New RegExp

	' Define pattern to match form input tags
	matchExp.Pattern = "<INPUT [^>]*type=hidden[^>]*>"
	matchExp.IgnoreCase = True
	matchExp.Global = True

	' Extract every input tag from the specified form object
	Set matchList = matchExp.Execute(formObject.GetROProperty("innerhtml"))
	
	' Define pattern to match attrib name/value
	' NOTE: Extracts name and quoted/unquoted value
	matchExp.Pattern = "([^=]+)=((?="")""[^""]+""|(?!"")[^ ]+) "
	
	' Iterate over input tags
	For Each thisMatch In matchList
		' Extract input tag source
		inputText = thisMatch.Value
		' Only retain tag attributes
		inputKeep = Len(inputText) - 8
		inputParms = Mid(inputText, 8, inputKeep) & " "
		' Split into attributes collection
		Set parmsList = matchExp.Execute(inputParms)

		' Iterate over attributes collection
		For Each thisParm In parmsList
			' Differentiate attribute name
			Select Case thisParm.SubMatches(0)

				Case "name"
					' Extract input tag name
					inputName = thisParm.SubMatches(1)

				Case "value"
					' Extract input tag value
					inputValue = thisParm.SubMatches(1)
					' If value is a quoted string
					If (Left(inputValue, 1) = Chr(34)) Then
						' Trim bounding quotation marks
						inputKeep = Len(inputValue) - 2
						inputValue = Mid(inputValue, 2, inputKeep)
					End If
					
			End Select
		Next
		
		' Add current input to data dictionary
		dataDict.Add inputName, inputValue
	Next

	' RESULT: Data dictionary
	Set ExtractData = dataDict

	' Release objects
	Set thisParm = Nothing
	Set parmList = Nothing
	Set thisMatch = Nothing
	Set matchList = Nothing
	Set matchExp = Nothing

End Function


'@Description Remove the indicated parameter from the specified URL
'@Documentation Remove <theParm> from <theURL>
'@Author sbabcoc
'@Date 07-APR-2011
'@InParameter [in] theURL, string, The URL to be processed
'@InParameter [in] theParm, string, The parameter to be removed
'@ReturnValue The specified URL with the indicated parameter removed
Public Function url_RemoveParm(ByVal theURL, ByVal theParm)

	' Declarations
	Dim regEx
	Dim newURL

	' Allocate RegExp object
	Set regEx = New RegExp

	newURL = theURL
	
	' Load pattern for specified parm
	regEx.Pattern = theParm
	' If URL contains spec'd parm
	If (regEx.Test(newURL)) Then
		' Remove specified parm
		newURL = regEx.Replace(newURL, "")
		' Clean up URL tail
		regEx.Pattern = "(\?|&)$"
		newURL = regEx.Replace(newURL, "")
		' Clean up URL body
		regEx.Pattern = "\?&"
		newURL = regEx.Replace(newURL, "?")
		regEx.Pattern = "&&"
		newURL = regEx.Replace(newURL, "&")
	End If

	' RESULT: Processed URL
	url_RemoveParm = newURL

	' Release RegExp
	Set regEx = Nothing

End Function



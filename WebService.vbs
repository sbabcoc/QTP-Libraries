Option Explicit

Public Const XMLNS = "jsxn"
Public Const NODE1 = "//jsxn:obj"
Public Const NSURI = "http://www.rei.com/ns/jsxn"

Public Const TypeOBJ = "obj"
Public Const TypeARR = "arr"
Public Const TypeSTR = "str"
Public Const TypeNUM = "num"
Public Const TypeBOO = "boo"
Public Const TypeNUL = "nul"

' WinHttpRequestOption Constants
Public Const UserAgentString = 0
Public Const URL = 1
Public Const URLCodePage = 2
Public Const EscapePercentInURL = 3
Public Const SslErrorIgnoreFlags = 4
	' 0x0100 : Unknown certification authority (CA) or untrusted root
	' 0x0200 : Wrong usage
	' 0x1000 : Invalid common name (CN)
	' 0x2000 : Invalid date or certificate expired
Public Const SelectCertificate = 5
Public Const EnableRedirects = 6
Public Const UrlEscapeDisable = 7
Public Const UrlEscapeDisableQuery = 8
Public Const SecureProtocols = 9
	' 0x0008 : SSL 2.0
	' 0x0020 : SSL 3.0
	' 0x0080 : Transport Layer Security (TLS) 1.0
Public Const EnableTracing = 10
Public Const RevertImpersonationOverSsl = 11
Public Const EnableHttpsToHttpRedirects = 12
Public Const EnablePassportAuthentication = 13
Public Const MaxAutomaticRedirects = 14
Public Const MaxResponseHeaderSize = 15
Public Const MaxResponseDrainSize = 16
Public Const EnableHttp1_1 = 17
Public Const EnableCertificateRevocationCheck = 18

' JSON Format Constants
Public Const FORMAT_VBS = 0
Public Const FORMAT_XML = 1

' JSON Parser Token Constants
Public Const TOKEN_NONE = 0
Public Const TOKEN_CURLY_OPEN = 1
Public Const TOKEN_CURLY_CLOSE = 2
Public Const TOKEN_SQUARED_OPEN = 3
Public Const TOKEN_SQUARED_CLOSE = 4
Public Const TOKEN_COLON = 5
Public Const TOKEN_COMMA = 6
Public Const TOKEN_STRING = 7
Public Const TOKEN_NUMBER = 8
Public Const TOKEN_TRUE = 9
Public Const TOKEN_FALSE = 10
Public Const TOKEN_NULL = 11

Public Function GetDataFromURL(strURL, strMethod, strPostData)
	Dim lngTimeout
	Dim strUserAgentString
	Dim intSslErrorIgnoreFlags
	Dim blnEnableRedirects
	Dim blnEnableHttpsToHttpRedirects
	Dim strHostOverride
	Dim strLogin
	Dim strPassword
	Dim strResponseText
	Dim objWinHttp
	
	lngTimeout = 59000
	strUserAgentString = "http_requester/0.1"
	intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
	blnEnableRedirects = True
	blnEnableHttpsToHttpRedirects = True
	strHostOverride = ""
	strLogin = ""
	strPassword = ""
	
	Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
	objWinHttp.Open strMethod, strURL
	
	If strMethod = "POST" Then
		objWinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	End If
	
	If strHostOverride <> "" Then
		objWinHttp.SetRequestHeader "Host", strHostOverride
	End If
	
	objWinHttp.Option(UserAgentString) = strUserAgentString
	objWinHttp.Option(SslErrorIgnoreFlags) = intSslErrorIgnoreFlags
	objWinHttp.Option(EnableRedirects) = blnEnableRedirects
	objWinHttp.Option(EnableHttpsToHttpRedirects) = blnEnableHttpsToHttpRedirects
	
	If (strLogin <> "") And (strPassword <> "") Then
		objWinHttp.SetCredentials strLogin, strPassword, 0
	End If
	
	On Error Resume Next
	objWinHttp.Send(strPostData)
	
	If Err.Number = 0 Then
		If objWinHttp.Status = "200" Then
			GetDataFromURL = objWinHttp.ResponseText
		Else
			GetDataFromURL = "HTTP " & objWinHttp.Status & " " & objWinHttp.StatusText
		End If
	Else
		GetDataFromURL = "Error " & Err.Number & " " & Err.Source & " " & Err.Description
	End If
	
	On Error GoTo 0
	
	Set objWinHttp = Nothing
End Function

Public Function validateXML(ByRef xmlDoc)
	Dim isCorrect
	Dim parseErr
	Dim errorList
	Dim thisError
	Dim errorMsg
	Dim errorNum

	' Set to return multiple validation errors
	xmlDoc.setProperty "MultipleErrorMessages", True
	' Validate with attached schemas
	Set parseErr = xmlDoc.validate()
	' Determine validation status
	isCorrect = (parseErr.errorCode = 0)
	
	If (isCorrect) Then
		' Log success message
		Reporter.ReportEvent micPass, "validateXML", "Specified document is valid"
	Else
		' Get list of parsing errors
		Set errorList = parseErr.allErrors
		' Build top-level error message string
		errorMsg =  "Error as returned from validate():" & vbCR & vbCR & _
					"Error code: " & parseErr.errorCode & vbCR & _
					"Error reason: " & parseErr.reason & vbCR & _
					"Error x-path: " & parseErr.errorXPath & vbCR & _
					"Error count: " & errorList.length
					
		' Log failure message
		Reporter.ReportEvent micFail, "validateXML", errorMsg
		
		' Init count
		errorNum = 0
		' iterate over errors
		For Each thisError In errorList
			' Build specific validation failure message
			errorMsg =  "error # " & errorNum & vbCR & _
						"reason: " & thisError.reason & vbCR & _
						"x-path: " & thisError.errorXPath
			' Log information message
			Reporter.ReportEvent micInfo, "validateXML", errorMsg
			' Increment count
			errorNum = errorNum + 1
		Next
	End If
	
	validate = isCorrect

	' Release objects
	Set parseErr = Nothing
	Set errorList = Nothing
	Set thisError = Nothing
End Function

Public Sub addSchemaFromPath(ByRef xmlDoc, ByVal thePath)
	Dim rootNode
	Dim attrNode

	' Select document root element
	Set rootNode = xmlDoc.selectSingleNode(NODE1)

	If IsNull(xmlDoc.schemas) Then
		' Attach schema cache to document
		Set xmlDoc.schemas = CreateObject("Msxml2.XMLSchemaCache.6.0")
		' Create attribute for 'xsi' namespace
		Set attrNode = xmlDoc.createNode(2, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
		' Set value of 'xsi' namespace node
		attrNode.Value = "http://www.w3.org/2001/XMLSchema-instance"
		' Add attribute node to root element
		rootNode.setAttributeNode attrNode
	End If

	' Create attribute for schema location
	Set attrNode = xmlDoc.createNode(2, "xsi:schemaLocation", "http://www.w3.org/2001/XMLSchema-instance")
	' Set location of 'jsxn' schema instance
	attrNode.Value = NSURI & " " & thePath
	' Add attribute node to root element
	rootNode.setAttributeNode attrNode

	' Add specified schema to cache
	xmlDoc.schemas.add NSURI, thePath

	Set rootNode = Nothing
	Set attrNode = Nothing
End Sub

'@Description Returns a new JSON parser object
'@Documentation Returns a new JSON parser object
'@ReturnValue A new JSON parser object
Public Function New_JSON_Parser(format)
	Dim myJSON
	Set myJSON = New JSON_Parser
	Set New_JSON_Parser = myJSON.Init(format)
	Set myJSON = Nothing
End Function

'@Description Returns a new JSON poster object
'@Documentation Returns a new JSON poster object
'@ReturnValue A new JSON poster object
Public Function New_JSON_Poster(format)
	Dim myJSON
	Set myJSON = New JSON_Poster
	Set New_JSON_Poster = myJSON.Init(format)
	Set myJSON = Nothing
End Function

'@Description Returns a new StringBuilder object
'@Documentation Returns a new StringBuilder object
'@ReturnValue A new StringBuilder object
Public Function New_StringBuilder()
	Set New_StringBuilder = New StringBuilder
End Function

'##################
'##### J S O N _ P a r s e r #####
'##################

Class JSON_Parser

	Private m_Poster

	Private Sub Class_Initialize()
		Set m_Poster = Nothing
	End Sub

	Public Function Init(ByVal format)
		Select Case format
			Case FORMAT_VBS
				Set m_Poster = New VBS_Poster
			Case FORMAT_XML
				Set m_Poster = New XML_Poster
		End Select
		Set Init = Me
	End Function

	Private Sub Class_Terminate()
		Set m_Poster = Nothing
	End Sub

	'@Description Parses the specified JSON string into a value
	'@Documentation Parses <json> into a value
	'@InParameter [in] json A JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing an Array, a Dictionary, a number, a string, 'Null', 'True', or 'False'
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Public Function JsonDecode(ByRef json, ByRef success)
		Dim object
		Dim theValue
		' Initialize
		success = True
		' If JSON string isn't NUL
		If Not IsNull(json) Then
			Dim index : index = 1
			
			' Create/initialize a new output object
			Set object = m_Poster.InitObject()
			' Parse the JSON string to rehydrate its data
			theValue = ParseValue(object, "", json, index, success)
			JsonDecode = m_Poster.Finalize(theValue)
		' Otherwise (NUL string)
		Else
			' Return 'Null' for NUL string
			JsonDecode = Array(Null)
		End If

		' Release objects
		Erase theValue
		Set object = Nothing
	End Function

	'@Description Rehydrate the indicated chunk of JSON into a Dictionary object
	'@Documentation Rehydrate the <json> at index <index> into a Dictionary object
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the object within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing a Dictionary (or 'Null' on failure)
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseObject(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
		Dim theObject
		Dim token
		Dim comma
		Dim done
		
		' Get the next token
		token = NextToken(json, index)
		' If start of JSON object
		If (token = TOKEN_CURLY_OPEN) Then
			' Allocate a new object
			theObject = m_Poster.NewObject(parent, elementName)
			' Reset indicator
			comma = False
			' Not done
			done = False
		' Otherwise (not object start)
		Else
			' Indicate failure
			success = False
			' Complete
			done = True
		End If

		' Iterate on JSON
		Do Until (done)
			' Preview the next token
			token = LookAhead(json, index)
			' If JSON string ended abruptly
			If (token = TOKEN_NONE) Then
				' Indicate failure
				success = False
				' Complete
				done = True
			' Otherwise, if the next token is a comma
			ElseIf (token = TOKEN_COMMA) Then
				' If expect comma
				If (comma) Then
					' Consume comma char
					NextToken json, index
					' Reset indicator
					comma = False
				' Otherwise (unexpected)
				Else
					' Indicate failure
					success = False
					' Complete
					done = True
				End If
			' Otherwise, if JSON object complete
			ElseIf (token = TOKEN_CURLY_CLOSE) Then
				' Consume brace char
				NextToken json, index
				' Complete
				done = True
			' Otherwise (next  property)
			Else
				Do ' <== Begin bail-out context
					Dim theLabel
					Dim theValue

					' If expect comma
					If (comma) Then
						' Indicate failure
						success = False
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Get property label string
					theLabel = ParseString(Null, "", json, index, success)
					' If string parse failed
					If Not (success) Then
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Get the next token
					token = NextToken(json, index)
					' If token isn't  label/value delimiter
					If (token <> TOKEN_COLON) Then
						' Indicate failure
						success = False
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Get property value
					theValue = ParseValue(theObject(0), theLabel, json, index, success)
					' If value parse failed
					If Not (success) Then
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Add element to object
					m_Poster.AddElem theObject, theLabel, theValue
					' Release objects
					Erase theValue

					' Expect comma
					comma = True
				Loop Until True ' <== End bail-out context
			End If
		Loop

		' If successful
		If (success) Then
			' Return the object
			ParseObject = m_Poster.EndObject(theObject)
		' Otherwise (failed)
		Else
			' Indicate failure
			ParseObject = Array(Null)
		End If

		' Release objects
		Erase theObject
	End Function

	'@Description Rehydrate the indicated chunk of JSON into a VBScript Array
	'@Documentation Rehydrate the <json> at index <index> into a VBScript Array
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the array within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing an Array (or 'Null' on failure)
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseArray(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
		Dim theArray
		Dim token
		Dim comma
		Dim done
		Dim itemName

		' Get the next token
		token = NextToken(json, index)
		' If start of JSON array
		If (token = TOKEN_SQUARED_OPEN) Then
			' Allocate a new array
			theArray = m_Poster.NewArray(parent, elementName)
			' Reset indicator
			comma = False
			' Not done
			done = False
		' Otherwise (not array start)
		Else
			' Indicate failure
			success = False
			' Complete
			done = True
		End If

		itemName = m_Poster.ItemName(elementName)

		' Iterate on JSON
		Do Until (done)
			' Preview the next token
			token = LookAhead(json, index)
			' If JSON string ended abruptly
			If (token = TOKEN_NONE) Then
				' Indicate failure
				success = False
				' Complete
				done = True
			' Otherwise, if the next token is a comma
			ElseIf (token = TOKEN_COMMA) Then
				' If expect comma
				If (comma) Then
					' Consume comma char
					NextToken json, index
					' Reset indicator
					comma = False
				' Otherwise (unexpected)
				Else
					' Indicate failure
					success = False
					' Complete
					done = True
				End If
			' Otherwise, if JSON array complete
			ElseIf (token = TOKEN_SQUARED_CLOSE) Then
				' Consume bracket char
				NextToken json, index
				' Complete
				done = True
			' Otherwise (next item)
			Else
				Do ' <== Begin bail-out context
					Dim theValue

					' If expect comma
					If (comma) Then
						' Indicate failure
						success = False
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Get item value, identifying container name
					theValue = ParseValue(theArray(0), itemName, json, index, success)
					' If value parse failed
					If Not (success) Then
						' Complete
						done = True
						' Bail out
						Exit Do
					End If

					' Add item to array
					m_Poster.AddItem theArray, theValue
					' Release objects
					Erase theValue

					' Expect comma
					comma = True
				Loop Until True ' <== End bail-out context
			End If
		Loop

		' If successful
		If (success) Then
			' Return VBScript Array
			ParseArray = m_Poster.EndArray(theArray)
		' Otherwise (failed)
		Else
			' Indicate failure
			ParseArray = Array(Null)
		End If

		' Release objects
		Erase theArray
	End Function

	'@Description Rehydrate the indicated chunk of JSON into a native value
	'@Documentation Rehydrate the <json> at index <index> into a native value
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing 'True'
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseTrue(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
				' Consume 'true' token
				NextToken json, index
				ParseTrue = m_Poster.NewBool(parent, elementName, True)
	End Function
	
	'@Description Rehydrate the indicated chunk of JSON into a native value
	'@Documentation Rehydrate the <json> at index <index> into a native value
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing 'False'
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseFalse(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
				' Consume 'false' token
				NextToken json, index
				ParseFalse = m_Poster.NewBool(parent, elementName, False)
	End Function
	
	'@Description Rehydrate the indicated chunk of JSON into a native value
	'@Documentation Rehydrate the <json> at index <index> into a native value
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing 'Null'
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseNull(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
				' Consume 'null' token
				NextToken json, index
				ParseNull = m_Poster.NewNull(parent, elementName)
	End Function
	
	'@Description Rehydrate the indicated chunk of JSON into a native value
	'@Documentation Rehydrate the <json> at index <index> into a native value
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing an Array, a Dictionary, a number, a string, 'Null', 'True', or 'False'
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseValue(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
		Dim theValue

		' Evaluate next token preview
		Select Case LookAhead(json, index)
			Case TOKEN_STRING
				' Get string value
				theValue = ParseString(parent, elementName, json, index, success)
			Case TOKEN_NUMBER
				' Get numeric value
				theValue = ParseNumber(parent, elementName, json, index, success)
			Case TOKEN_CURLY_OPEN
				' Get object value
				theValue = ParseObject(parent, elementName, json, index, success)
			Case TOKEN_SQUARED_OPEN
				' Get array value
				theValue = ParseArray(parent, elementName, json, index, success)
			Case TOKEN_TRUE
				' Get 'true' value
				theValue = ParseTrue(parent, elementName, json, index, success)
			Case TOKEN_FALSE
				' Get 'false' value
				theValue = ParseFalse(parent, elementName, json, index, success)
			Case TOKEN_NULL
				' Get 'null' value
				theValue = ParseNull(parent, elementName, json, index, success)
			Case TOKEN_NONE
				' Indicate failure
				success = False
				theValue = Array(Null)
		End Select

        ParseValue = theValue

		' Release objects
		Erase theValue
	End Function

	'@Description Rehydrate the indicated chunk of JSON into a string value
	'@Documentation Rehydrate the <json> at index <index> into a string value
	'@InParameter [in] parent Reference to parent object (or 'Null' for native string result)
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing a string (or 'Null' on failure)
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseString(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
		Dim theString
		Dim token
		Dim theChar
		Dim done
		Dim remain
		Dim chrNum

		' Allocate string builder object
		Set theString = New StringBuilder

		' Get the next token
		token = NextToken(json, index)
		' If start of JSON string
		If (token = TOKEN_STRING) Then
			' Not done
			done = False
		' Otherwise (not string start)
		Else
			' Indicate failure
			success = False
			' Complete
			done = True
		End If

		' Iterate until string or JSON ends
		Do Until (done Or (index > Len(json)))
			' Get next JSON character
			theChar = Mid(json, index, 1)
			' Increment index
			index = index + 1

			' If quotation mark
			If (theChar = """") Then
				' Complete
				done = True
			' Otherwise, if backslash
			ElseIf (theChar = "\") Then
				' Get count of remaining chars
				remain = Len(json) - index + 1
				' If 1+ chars remain
				If (remain > 0) Then
					' Get next JSON character
					theChar = Mid(json, index, 1)
					' Increment index
					index = index + 1

					' Eval escaped char
					Select Case theChar

						' Quote, slash, slope
						Case """", "\", "/"
							' Append escaped char
							theString.Append theChar

						' b - Backspace
						Case "b"
							' Append backspace
							theString.Append Chr(8)

						' f - Form Feed
						Case "f"
							' Append form feed
							theString.Append Chr(12)

						' n - New Line
						Case "n"
							' Append line feed
							theString.Append vbLF

						' r - Carriage Return
						Case "r"
							' Append carriage return
							theString.Append vbCR

						' t - Horizontal Tab
						Case "t"
							' Append horizontal tab
							theString.Append vbTab

						' x - Hex Char Code
						Case "x"
							' If 3+ chars remain
							If (remain > 2) Then
								' Extract hex char code
								theChar = Mid(json, index, 2)
								' Convert hex string to decimal
								chrNum = CLng("&h" & theChar)
								' Append specified ASCII char
								theString.Append Chr(chrNum)
								' Advance index
								index = index + 2
							End If

						' Octal Char Code
						Case "0", "1", "2", "3"
							' If 4+ chars remain
							If (remain > 3) Then
								' Extract octal char code
								theChar = Mid(json, index, 3)
								' Convert octal string to decimal
								chrNum = CLng("&o" & theChar)
								' Append specified ASCII char
								theString.Append Chr(chrNum)
								' Advance index
								index = index + 3
							End If

						' u - Unicode Chat Code
						Case "u"
							' If 5+ chars remain
							If (remain > 4) Then
								' Extract Unicode char code
								theChar = Mid(json, index, 4)
								' Convert hex string to decimal
								chrNum = CLng("&h" & theChar)
								' Append specified Unicode char
								theString.Append ChrW(chrNum)
								' Advance index
								index = index + 4
							End If

					End Select
				End If
			' Otherwise (normal char)
			Else
				' Append normal char
				theString.Append theChar
			End If
		Loop

		' If normal end
		If (done) Then
			' If parent is not 'Null'
			If Not IsNull(parent) Then
				' Return rehydrated string
				ParseString = m_Poster.NewString(parent, elementName, theString.ToString)
			Else
				' Return native string
				ParseString = theString.ToString
			End If
		' Otherwise (abend)
		Else
			' Indicate failure
			ParseString = Array(Null)
		End If

		' Release objects
		Set theString = Nothing
	End Function

	'@Description Rehydrate the indicated chunk of JSON into a numeric value
	'@Documentation Rehydrate the <json> at index <index> into a numeric value
	'@InParameter [in] parent Reference to parent object
	'@InParameter [in] elementName Name for new object
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@OutParameter [out] success Reference to status output variable
	'@ReturnValue A one-item array containing a number (or 'Null' on failure)
	'@Note Prior to unwrapping the result, use the IsObject() function to determine the result type
	Private Function ParseNumber(ByRef parent, ByVal elementName, ByRef json, ByRef index, ByRef success)
		Dim lastIndex
		Dim charLength
		Dim octStr, hexStr, decStr
		Dim theValue
		Dim regEx, matchList

		' Allocate RegExp object
		Set regEx = New RegExp
		' Match octal, hexadecimal, and decimal (integer/real/scientific) constants
		regEx.Pattern = "^((0[0-7]+)|(0[xX][0-9a-fA-F]+)|([-+]?[0-9]+[.]?[0-9]*([eE][-+]?[0-9]+)?))"

		' Consume whitespace
		EatWhitespace json, index
		' Apply pattern to extract numeric constants
		Set matchList = regEx.Execute(Mid(json, index))

		' If numeric constant extracted
		If (matchList.Count = 1) Then
			' Advance index
			index = index + Len(matchList(0))

			' Extract octal, hex, and decimal
			octStr = matchList(0).SubMatches(1)
			hexStr = matchList(0).SubMatches(2)
			decStr = matchList(0).SubMatches(3)

			' If constant is decimal
			If (Len(decStr) > 0) Then
				' Convert decimal string
				theValue = CDbl(decStr)
			' Otherwise, if constant is hex
			ElseIf (Len(hexStr) > 0) Then
				' Convert hexdecimal string
				theValue = CLng("&h" & hexStr)
			' Otherwise (constant is octal)
			Else
				' Convert octal string
				theValue = CLng("&o" & octStr)
			End If
		' Otherwise (non-numeric)
		Else
			' Indicate failure
			success = False
		End If

		' If successful
		If (success) Then
			' Return rehydrated number
			ParseNumber = m_Poster.NewNumber(parent, elementName, theValue)
		' Otherwise (failed)
		Else
			' Indicate failure
			ParseNumber = Array(Null)
		End If

		' Release objects
		Set matchList = Nothing
		Set regEx = Nothing
	End Function

	'@Description Advance JSON string index to the next non-whitespace char
	'@Documentation Advance JSON string index to the next non-whitespace char
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	Private Sub EatWhitespace(ByRef json, ByRef index)
	   ' Iterate to end of JSON string
		Do Until (index > Len(json))
			' Evaluate current JSON char
			Select Case Mid(json, index, 1)
				' Space, tab, return, linefeed
				Case " ", vbTab, vbCR, vbLF
					' Increment index
					index = index + 1
				' Non-whitespace
				Case Else
					' Done
					Exit Do
			End Select
		Loop
	End Sub
	
	'@Description Preview the next token from the current position of the JSON string
	'@Documentation Preview the next token from index <index> of <json>
	'@InParameter [in] json A JSON string
	'@InParameter  [in] index The offset of the value within the JSON string
	'@ReturnValue One of the defined JSON token constants
	Private Function LookAhead(ByRef json, ByVal index)
		' Get the next token with local index
		LookAhead = NextToken(json, index)
	End Function

	'@Description Get the next token from the current position of the JSON string
	'@Documentation Get the next token from index <index> of <json>
	'@InParameter [in] json A JSON string
	'@InParameter  [in/out] index The offset of the value within the JSON string
	'@ReturnValue One of the defined JSON token constants
	Private Function NextToken(ByRef json, ByRef index)
		Dim token
		Dim theChar
		Dim remain

		' Initialize result
		token = TOKEN_NONE

		' Consume whitespace
		EatWhitespace json, index
		' Get count of remaining chars
		remain = Len(json) - index + 1
		' If 1+ chars remain
		If (remain > 0) Then
			' Get next JSON character
			theChar = Mid(json, index, 1)
			' Increment index
			index = index + 1

			' Evaluate current char
			Select Case theChar

				Case "{"
					token = TOKEN_CURLY_OPEN

				Case "}"
					token = TOKEN_CURLY_CLOSE

				Case "["
					token = TOKEN_SQUARED_OPEN

				Case "]"
					token = TOKEN_SQUARED_CLOSE

				Case ","
					token = TOKEN_COMMA

				Case """"
					token = TOKEN_STRING

				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "+"
					token = TOKEN_NUMBER

				Case ":"
					token = TOKEN_COLON

				Case "n"
					' If 4+ chars remain
					If (remain > 3) Then
						' If JSON token is "null"
						If (Mid(json, index, 3) = "ull") Then
							token = TOKEN_NULL
							' Advance index
							index = index + 3
						End If
					End If

				Case "t"
					' If 4+ chars remain
					If (remain > 3) Then
						' If JSON token is "true"
						If (Mid(json, index, 3) = "rue") Then
							token = TOKEN_TRUE
							' Advance index
							index = index + 3
						End If
					End If

				Case "f"
					' If 5+ chars remain
					If (remain > 4) Then
						' If JSON token is "false"
						If (Mid(json, index, 4) = "alse") Then
							token = TOKEN_FALSE
							' Advance index
							index = index + 4
						End If
					End If

			End Select
		End If

		NextToken = token
	End Function

End Class

'##################
'##### J S O N _ P o s t e r #####
'##################

Class JSON_Poster

	Private m_Parser

	Private Sub Class_Initialize()
		Set m_Parser = Nothing
	End Sub
	
	Public Function Init(format)
		Select Case format
			Case FORMAT_VBS
				Set m_Parser = Me
			Case FORMAT_XML
				Set m_Parser = New XML_Parser
				m_Parser.Init(Me)
		End Select
		Set Init = Me
	End Function

	Private Sub Class_Terminate()
		Set m_Parser = Nothing
	End Sub

	'@Description Converts the specified object into a JSON string
	'@Documentation Converts <object> into a JSON string
	'@InParameter [in] object Dictionary object to be converted
	'@ReturnValue JSON representation of the specified object; 'Null' if encoding fails
	Public Function JsonEncode(ByRef object)
		Dim builder
		Dim success
		' Allocate string builder object
		Set builder = New StringBuilder
		' Serialize the specified object
		success = m_Parser.Serialize(object, builder)

		' If successful
		If (success) Then
			' Return JSON representation
			JsonEncode = builder.ToString
		' Otherwise (failed)
		Else
			' Indicate failure
			JsonEncode = Null
		End If

		' Release objects
		Set builder = Nothing
	End Function

	Public Function Serialize(ByRef object, ByRef builder)
		Serialize = SerializeValue(object, builder)
	End Function

	'@Description Converts the specified VBScript value into a JSON string
	'@Documentation Converts <theValue> into a JSON string
	'@InParameter [in] theValue VBScript value to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeValue(ByRef theValue, ByRef builder)
		Dim strType
		Dim success : success = True

		' If value is an array
		If IsArray(theValue) Then
			' Set type name
			strType = "Array"
		' Otherwise (non-array)
		Else
			' Get actual type name
			strType = TypeName(theValue)
		End If

		' Evaluate value type
		Select Case strType

			Case "String"
				' Convert value to JSON string
				success = SerializeString(theValue, builder)

			Case "Dictionary"
				' Convert value to JSON object
				success = SerializeObject(theValue, builder)

			Case "Array"
				' Convert value to JSON array
				success = SerializeArray(theValue, builder)

			Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal"
				' Convert value to JSON number
				success = SerializeNumber(theValue, builder)

			Case "Boolean"
				' Convert value to JSON boolean
				success = SerializeBool(theValue, builder)

			Case "Null"
				' Convert value to JSON null
				success = SerializeNull(theValue, builder)

			Case Else
				' Indicate failure
				success = False

		End Select

		SerializeValue = success
	End Function

	'@Description Converts the specified Dictionary object into a JSON object
	'@Documentation Converts <theObject> into a JSON object
	'@InParameter [in] theObject Dictionary object to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeObject(ByRef theObject, ByRef builder)
		Dim keyCount, keyIndex
		Dim keyList, thisKey
		Dim success : success = True

		' Start JSON object
		builder.Append "{"
		' Get count of properties
		keyCount = theObject.Count

		' If not an empty object
		If (keyCount > 0) Then
			' Init key index
			keyIndex = 0
			' Get list of prop keys
			keyList = theObject.Keys

			' Interate
			Do
				' Get current key
				thisKey = keyList(keyIndex)

				' Emit property label string
				SerializeString thisKey, builder
				' Emit label/value delimiter
				builder.Append ":"
				' Convert value to a JSON string
				success = SerializeValue(theObject.Item(thisKey), builder)

				' Increment key index
				keyIndex = keyIndex + 1
				' If conversion failed or all keys converted, stop looping
				If Not (success And (keyIndex < keyCount)) Then Exit Do

				' Emit property delimiter
				builder.Append ","
			Loop
		End If

		' End JSON object
		builder.Append "}"
		' Return conversion status
		SerializeObject = success
	End Function

	'@Description Converts the specified VBScript Array into a JSON string
	'@Documentation Converts <theArray> into a JSON string
	'@InParameter [in] theArray VBScript Array to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeArray(ByRef theArray, ByRef builder)
		Dim itemCount, itemIndex
		Dim success : success = True

		' Start JSON array
		builder.Append "["

		' Get array upper bound
		itemCount = -1
		On Error Resume Next
		itemCount = UBound(theArray)
		On Error GoTo 0

		' Convert bound to count
		itemCount = itemCount + 1

		' If not an empty array
		If (itemCount > 0) Then
			' Init item index
			itemIndex = 0

			' Iterate
			Do
				' Convert value to a JSON string
				success = SerializeValue(theArray(itemIndex), builder)

				' Increment item index
				itemIndex = itemIndex + 1
				' If conversion failed or all items converted, stop looping
				If Not (success And (itemIndex < itemCount)) Then Exit Do

				' Emit item delimiter
				builder.Append ","
			Loop
		End If

		' End JSON array
		builder.Append "]"
		' Return conversion status
		SerializeArray = success
	End Function

	'@Description Converts the specified VBScript string into a JSON string
	'@Documentation Converts <theString> into a JSON string
	'@InParameter [in] theString VBScript string to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeString(ByRef theString, ByRef builder)
		Dim charCount, charIndex
		Dim theChar, charStr, charNum

		' Start JSON string
		builder.Append """"

		' Get VBScript string length
		charCount = Len(theString)
		' Iterate over string character
		For charIndex = 1 To charCount
			' Get current string character
			theChar = Mid(theString, charIndex, 1)
			' Evaluate character
			Select Case theChar

				' Quote, slash, slope
				Case """", "\", "/"
					' Emit escaped char
					charStr = "\" & theChar

				' Backspace
				Case Chr(8)
					' Emit JSON backspace
					charStr = "\b"

				' Form feed
				Case Chr(12)
					' Emit JSON form feed
					charStr = "\f"

				' Line Feed
				Case vbLF
					' Emit JSON new line
					charStr = "\n"

				' Carriage Return
				Case vbCR
					' Emit JSON carriage return
					charStr = "\r"

				' Horizontal Tab
				Case vbTab
					' Emit JSON horizontal tab
					charStr = "\t"

				' Non-escaped character
				Case Else
					' Get Unicode char code
					charNum = AscW(theChar)
					' If ASCII control code
					If ((charNum < 32) Or (charNum = 127)) Then
						' Emit octal character constant
						charStr = "\" & Right("00" & Oct(charNum), 3)
					' Otherwise, if not 7-bit ASCII
					ElseIf (charNum > 127) Then
						' Emit Unicode character constant
						charStr = "\u" & Right("00" & Hex(charNum), 4)
					' Otherwise (normal 7-bit ASCII)
					Else
						' Emit normal char
						charStr = theChar
					End If

			End Select

			' Append JSON char
			builder.Append(charStr)
		Next

		' End JSON string
		builder.Append """"
		' Return conversion status
		SerializeString = True
	End Function

	'@Description Converts the specified VBScript number into a JSON string
	'@Documentation Converts <theNumber> into a JSON string
	'@InParameter [in] theNumber VBScript number to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeNumber(ByRef theNumber, ByRef builder)
	   ' Emit JSON number
		builder.Append CStr(CDbl(theNumber))
		' Return conversion status
		SerializeNumber = True
	End Function

	'@Description Converts the specified VBScript boolean into a JSON string
	'@Documentation Converts <theBool> into a JSON string
	'@InParameter [in] theBool VBScript boolean to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeBool(ByVal theBool, ByRef builder)
		' If value is 'True'
		If (theBool) Then
			' Emit JSON "true"
			builder.Append "true"
		' Otherwise (value is 'False')
		Else
			' Emit JSON "false"
			builder.Append "false"
		End If
		' Return conversion status
		SerializeBool = True
	End Function

	'@Description Converts the specified VBScript 'Null' into a JSON string
	'@Documentation Converts <theNull> into a JSON string
	'@InParameter [in] theNull VBScript 'Null' to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeNull(ByVal theNull, ByRef builder)
		' Emit JSON "null"
		builder.Append "null"
		' Return conversion status
		SerializeNull = True
	End Function

End Class

'########################
'##### V B S _ P o s t e r #####
'########################

Class VBS_Poster

	Public Function InitObject()
		Set InitObject = Nothing
	End Function

	Public Function Finalize(ByRef theValue)
		Finalize = theValue
	End Function

	Public Function NewObject(ByRef parent, ByVal elementName)
		Dim theObject
		Set theObject = CreateObject("Scripting.Dictionary")
		NewObject = Array(theObject)
		Set theObject = Nothing
	End Function

	Public Sub AddElem(ByRef theObject, ByRef theLabel, ByRef theValue)
		theObject(0).Add theLabel, theValue(0)
	End Sub

	Public Function EndObject(ByRef theObject)
	   EndObject = theObject
	End Function

	Public Function NewArray(ByRef parent, ByVal elementName)
	   NewArray = Array(Array())
	End Function

	Public Sub AddItem(ByRef theArray, ByRef theValue)
	   Dim arrayRef
	   Dim arrayLen
	   ' Unwrap array
	   arrayRef = theArray(0)
	   ' Get current array length
	   arrayLen = UBound(arrayRef)
	   ' Account for new item
	   arrayLen = arrayLen + 1
	   ' Set new array capacity
	   ReDim Preserve arrayRef(arrayLen)
		' If value is an object
		If IsObject(theValue(0)) Then
			' Add object item to array
			Set arrayRef(arrayLen) = theValue(0)
		' Otherwise (intrinsic)
		Else
			' Add intrinsic item to array
			arrayRef(arrayLen) = theValue(0)
		End If
		' Re-wrap array
		theArray(0) = arrayRef
	End Sub

	Public Function EndArray(ByRef theArray)
	   EndArray = theArray
	End Function

	Public Function NewString(ByRef parent, ByVal elementName, ByRef theString)
		NewString = Array(theString)
	End Function

	Public Function NewNumber(ByRef parent, ByVal elementName, ByVal theNumber)
		NewNumber = Array(theNumber)
	End Function

	Public Function NewBool(ByRef parent, ByVal elementName, ByVal theBool)
		NewBool = Array(CBool(theBool))
	End Function

	Public Function NewNull(ByRef parent, ByVal elementName)
		NewNull = Array(Null)
	End Function

	Public Function ItemName(elementName)
		ItemName = ""
	End Function

End Class

'########################
'##### X M L _ P a r s e r #####
'########################

Class XML_Parser

	Private m_JSON_Poster
	
	Private Sub Class_Initialize()
		Set m_JSON_Poster = Nothing
	End Sub
	
	Public Function Init(poster)
		Set m_JSON_Poster = poster
		Set Init = Me
	End Function
	
	Private Sub Class_Terminate()
		Set m_JSON_Poster = Nothing
	End Sub

	Public Function Serialize(ByRef object, ByRef builder)
		Serialize = SerializeValue(object.selectSingleNode(NODE1), builder)
	End Function

	'@Description Converts the specified JSxN object into a JSON string
	'@Documentation Converts <theValue> into a JSON string
	'@InParameter [in] theValue JSxN object to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeValue(ByRef theValue, ByRef builder)
		Dim nodeName, nameBits
		Dim success : success = True

		' Get node name
		nodeName = theValue.nodeName
		' Split name into bits
		nameBits = splitName(nodeName)
		' If this is a JSxN node
		If (nameBits(0) = XMLNS) Then
			nodeKind = nameBits(1)
		Else
			nodeKind = ""
		End If

		' Evaluate node kind
		Select Case nodeKind

			Case TypeSTR
				' Convert value to JSON string
				success = SerializeString(theValue, builder)
				
			Case TypeOBJ
				' Convert value to JSON object
				success = SerializeObject(theValue, builder)
				
			Case TypeARR
				' Convert value to JSON array
				success = SerializeArray(theValue, builder)
				
			Case TypeNUM
				' Convert value to JSON number
				success = SerializeNumber(theValue, builder)
				
			Case TypeBOO
				' Convert value to JSON boolean
				success = SerializeBool(theValue, builder)
				
			Case TypeNUL
				' Convert value to JSON null
				success = SerializeNull(theValue, builder)

			Case Else
				' Indicate failure
				success = False

		End Select
		
		SerializeValue = success
	End Function

	'@Description Converts the specified JSxN object into a JSON object
	'@Documentation Converts <theObject> into a JSON object
	'@InParameter [in] theObject JSxN object to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeObject(ByRef theObject, ByRef builder)
		Dim propList, thisProp
		Dim propCount, propIndex
		Dim nodeName, nameBits
		Dim success : success = True

		' Start JSON object
		builder.Append "{"
		
		' Get list of property nodes
		Set propList = theObject.childNodes
		' Get count of property nodes
		propCount = propList.length

		' If not an empty object
		If (propCount > 0) Then
			' Init prop index
			propIndex = 0

			' Interate
			Do
				' Get current property node
				Set thisProp = propList(propIndex)
				' Get property node name
				nodeName = thisProp.nodeName
				' Split name into bits
				nameBits = splitName(nodeName)
				' If this is a JSxN node
				If (nameBits(0) = XMLNS) Then
					' Emit property label string
					m_JSON_Poster.SerializeString nameBits(2), builder
					' Emit label/value delimiter
					builder.Append ":"
					' Convert value to a JSON string
					success = SerializeValue(thisProp, builder)
					' Increment prop index
					propIndex = propIndex + 1
				Else
					success = False
				End If
				
				' If conversion failed or all props converted, stop looping
				If Not (success And (propIndex < propCount)) Then Exit Do

				' Emit property delimiter
				builder.Append ","
			Loop
		End If

		' End JSON object
		builder.Append "}"
		' Return conversion status
		SerializeObject = success

		' Release objects
		Set propList = Nothing
		Set thisProp = Nothing
	End Function

	'@Description Converts the specified JSxN Array into a JSON string
	'@Documentation Converts <theArray> into a JSON string
	'@InParameter [in] theArray JSxN Array to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeArray(ByRef theArray, ByRef builder)
		Dim itemList, thisItem
		Dim itemCount, itemIndex
		Dim success : success = True

		' Start JSON array
		builder.Append "["

		' Get list of child nodes
		Set itemList = theArray.childNodes
		' Get count of child nodes
		itemCount = itemList.length

		' If not an empty array
		If (itemCount > 0) Then
			' Init item index
			itemIndex = 0

			' Iterate
			Do
				' Convert value to a JSON string
				success = SerializeValue(itemList(itemIndex), builder)

				' Increment item index
				itemIndex = itemIndex + 1
				' If conversion failed or all items converted, stop looping
				If Not (success And (itemIndex < itemCount)) Then Exit Do

				' Emit item delimiter
				builder.Append ","
			Loop
		End If

		' End JSON array
		builder.Append "]"
		' Return conversion status
		SerializeArray = success

		' Release objects
		Set itemList = Nothing
	End Function
	
	'@Description Converts the specified JSxN string into a JSON string
	'@Documentation Converts <theString> into a JSON string
	'@InParameter [in] theString JSxN string to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeString(ByRef theString, ByRef builder)
		' Emit JSON string
		SerializeString = m_JSON_Poster.SerializeString(theString.text, builder)
	End Function

	'@Description Converts the specified JSxN number into a JSON string
	'@Documentation Converts <theNumber> into a JSON string
	'@InParameter [in] theNumber JSxN number to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeNumber(ByRef theNumber, ByRef builder)
		' Emit JSON number
		SerializeNumber = m_JSON_Poster.SerializeNumber(theNumber.text)
	End Function

	'@Description Converts the specified VBScript boolean into a JSON string
	'@Documentation Converts <theBool> into a JSON string
	'@InParameter [in] theBool VBScript boolean to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeBool(ByVal theBool, ByRef builder)
	   ' Emit JSON boolean
		SerializeBool = m_JSON_Poster.SerializeBool(theBool, builder)
	End Function

	'@Description Converts the specified VBScript 'Null' into a JSON string
	'@Documentation Converts <theNull> into a JSON string
	'@InParameter [in] theNull VBScript 'Null' to be converted
	'@InParameter [in/out] builder StringBuilder object to receive JSON output
	'@ReturnValue 'True' if conversion succeeds; otherwise 'False'
	Public Function SerializeNull(ByVal theNull, ByRef builder)
		' Emit JSON "null"
		SerializeNull = m_JSON_Poster.SerializeNull(theNull, builder)
	End Function

	Private Function splitName(ByVal nodeName)
		Dim nameBits
		Dim nameSpace
		Dim elementType
		Dim elementName
		Dim arrayName

		nodeName = jsonName(nodeName)

		nameBits = Split(nodeName, ".", 2)
		If (UBound(nameBits) = 1) Then
			arrayName = "." & nameBits(1)
			nodeName = nameBits(0)
		Else
			arrayName = ""
		End If

		nameBits = Split(nodeName, "_", 2)
		If (UBound(nameBits) = 1) Then
			elementName = nameBits(1)
			nodeName = nameBits(0)
		Else
			elementName = ""
		End If
		
		nameBits = Split(nodeName, ":", 2)
		If (UBound(nameBits) = 1) Then
			elementType = nameBits(1)
			nameSpace = nameBits(0)
		Else
			elementType = nameBits(0)
			nameSpace = ""
		End If			
		
		splitName = Array(nameSpace, elementType, elementName, arrayName)
	End Function
	
	Private Function jsonName(ByVal xmlName)
		jsonName = Replace(xmlName, "-", "$")
	End Function

End Class

'########################
'##### X M L _ P o s t e r #####
'########################

Class XML_Poster

	Public Function InitObject()
		Dim object
		Dim objIntro
		
		' Create an instance of the Parser object
		Set object = CreateObject("Msxml2.DOMDocument.6.0")

		' Specify synchronous download
		object.async = False
		' Specify no validation
		object.validateOnParse = False
		' Specify resolution of externals
		object.resolveExternals = True
		' Specify whitespace preservation
		object.preserveWhiteSpace = True
		' Set element selection language to XPath
		object.setProperty "SelectionLanguage", "XPath"
		' Set selection namespaces
		object.setProperty "SelectionNamespaces", "xmlns:" & XMLNS & "='" & NSURI & "'"

		' Create processing instruction node
		Set objIntro = object.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
		' Add node to document
		object.appendChild objIntro

		' Return document fragment as seed element
		Set InitObject = object.createDocumentFragment()

		' Release objects
		Set object = Nothing
		Set objIntro = Nothing
	End Function

	Public Function Finalize(ByRef theValue)
		Dim objRoot
		Dim object

		' Unwrap root element
		Set objRoot = theValue(0)
		' Get owner document
		Set object = objRoot.ownerDocument
		' Attach root element to document
		object.appendChild objRoot
		
		' Return wrapped document
		Finalize = Array(object)

		' Release objects
		Set objRoot = Nothing
		Set object = Nothing
	End Function

	Public Function NewObject(ByRef parent, ByVal elementName)
		Dim nodeName
		Dim objObject

		' Assemble node name
		nodeName = joinName(XMLNS, TypeOBJ, elementName)
		' Create new object element
		Set objObject = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Return wrapped object element
		NewObject = Array(objObject)

		' Release objects
		Set document = Nothing
		Set objObject = Nothing
	End Function

	Public Sub AddElem(ByRef theObject, ByRef theLabel, ByRef theValue)
		' Append element to object
		theObject(0).appendChild theValue(0)
	End Sub

	Public Function EndObject(ByRef theObject)
		EndObject = theObject
	End Function

	Public Function NewArray(ByRef parent, ByVal elementName)
		Dim nodeName
		Dim objArray

		' Assemble node name
		nodeName = joinName(XMLNS, TypeARR, elementName)
		' Create new array element
		Set objArray = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Return wrapped array element
		NewArray = Array(objArray)

		' Release objects
		Set objArray = Nothing
	End Function

	Public Sub AddItem(ByRef theArray, ByRef theValue)
		' Append item to array
		theArray(0).appendChild theValue(0)
	End Sub

	Public Function EndArray(ByRef theArray)
		EndArray = theArray
	End Function

	Public Function NewString(ByRef parent, ByVal elementName, ByRef theString)
		Dim nodeName
		Dim objString

		' Assemble node name
		nodeName = joinName(XMLNS, TypeSTR, elementName)
		' Create new string element
		Set objString = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Set specified element value
		objString.text = theString
		' Return wrapped string element
		NewString = Array(objString)

		' Release objects
		Set objString = Nothing
	End Function

	Public Function NewNumber(ByRef parent, ByVal elementName, ByVal theNumber)
		Dim nodeName
		Dim objNumber

		' Assemble node name
		nodeName = joinName(XMLNS, TypeNUM, elementName)
		' Create new number element
		Set objNumber = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Set specified element value
		objNumber.text = theNumber
		' Return wrapped number element
		NewNumber = Array(objNumber)

		' Release objects
		Set objNumber = Nothing
	End Function

	Public Function NewBool(ByRef parent, ByVal elementName, ByVal theBool)
		Dim nodeName
		Dim objBool

		' Assemble node name
		nodeName = joinName(XMLNS, TypeBOO, elementName)
		' Create new number element
		Set objBool = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Set specified element value
		objBool.text = LCase(CStr(theBool))
		' Return wrapped boolean element
		NewBool = Array(objBool)

		' Release objects
		Set objBool = Nothing
	End Function

	Public Function NewNull(ByRef parent, ByVal elementName)
		Dim nodeName
		Dim objNull

		' Assemble node name
		nodeName = joinName(XMLNS, TypeNUL, elementName)
		' Create new number element
		Set objNull = parent.ownerDocument.createNode(1, nodeName, NSURI)
		' Return wrapped boolean element
		NewNull = Array(objNull)

		' Release objects
		Set objNull = Nothing
	End Function

	Private Function joinName(ByVal nameSpace, ByVal elementType, ByVal elementName)
		If (nameSpace <> "") Then nameSpace = nameSpace & ":"
		If (elementName <> "") Then elementName = "_" & xmlName(elementName)
		joinName = nameSpace & elementType & elementName
	End Function
	
	Private Function xmlName(ByVal jsonName)
		xmlName = Replace(jsonName, "$", "-")
	End Function

	Public Function ItemName(elementName)
		ItemName = "." & elementName & "."
	End Function

End Class
	
'###########################
'##### S t r i n g B u i l d e r #####
'###########################

'This string-builder class provides efficient string construction
Class StringBuilder

	Private m_Index
	Private m_Array()
	Private m_Alloc

	' Called at creation of instance
	Private Sub Class_Initialize()
	   m_Alloc = 128
	   Clear
	End Sub

	' Called at destruction of instance
	Private Sub Class_Terminate()
		Erase m_Array
	End Sub

	' Add new string to array
	Public Sub Append(ByRef NewStr)
		m_Array(m_Index) = NewStr
		m_Index = m_Index + 1

		'ReDim array if necessary
		If (m_Index = m_Alloc) Then
			' Double the array size
			m_Alloc = m_Alloc * 2
			ReDim Preserve m_Array(m_Alloc - 1)
		End If
	End Sub

	' Return the concatenated string
	Public Property Get ToString
		ToString = Join(m_Array, "")
	End Property

	' Reset string array
	Public Sub Clear()
		m_Index = 0
		ReDim m_Array(m_Alloc - 1)
	End Sub

End Class 


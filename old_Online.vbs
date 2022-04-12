''-------------------------------------------------------------------------------------
''				dtShoppingBasketLineItems(oLineItemTable)()
''-------------------------------------------------------------------------------------
''Desc:	Returns the number of items in the shopping cart
''
''Args:	
''		sSheetName = Name of the datasheet that will 
''			contain all line items from the shopping basket
''
''Usage:
''		sSheetName = "Line items"
''		Set  oLineItemTable = Browser("Browser").Page("REI.com: Shopping Basket").WebTable("PRODUCT")
'' 		dtShoppingBasketLineItems(sSheetName, oLineItemTable)
''-------------------------------------------------------------------------------------
'Public Function dtShoppingBasketLineItems(sSheetName, oLineItemTable)
'
'	iRow = oLineItemTable.RowCount
'	iCol = oLineItemTable.ColumnCount(1)
'	
'	DataTable.AddSheet(sSheetName)
'	'Get Line item headers and content
'	iQuantityIndex = 0
'	For i = 1 To iRow - 1 Step 2
'		If i = 1 Then
'			DataTable.SetCurrentRow i 
'			For j = 1 To iCol
'				'Line Item Header
'				sCellData = oLineItemTable.GetCellData(i, j)
'				Call DataTable.GetSheet(sSheetName).AddParameter(sCellData, "")
'			Next
'		    i = i + 1
'		End If
'	
'		'Line Item data
'		For j = 2 To iCol + 1 
'			If j <> 4 Then
'				sCellData = oLineItemTable.GetCellData(i, j)
'			Else
'				'Get the quantity
'				sCellData = cmnShoppingcartItemQuantity
'				If sCellData = 0 Then
'                     sCellData = oLineItemTable.GetCellData(i, j)
'				End If
'				iQuantityIndex = iQuantityIndex + 1
'			End If
'			
'			If Not sCellData = "" Then
'				DataTable.SetCurrentRow i/2
'				DataTable.GetSheet(sSheetName).GetParameter(j-1).Value = sCellData
'			End If
'		Next
'	Next
'
'End Function


''-------------------------------------------------------------------------------------
''				dtBillingPageLineItems(oLineItemTable)
''-------------------------------------------------------------------------------------
''Desc:	Returns the number of items in the shopping cart
''
''Args:	
''		sSheetName = Name of the datasheet that will 
''			contain all line items from the billing page
''
''Usage:
''		sSheetName = "Line items"
''		Set  oLineItemTable = Set oTable = Browser("Browser").Page("REI: Checkout: Registered").WebTable("Product")
'' 		dtBillingPageLineItems(sSheetName, oLineItemTable)
''-------------------------------------------------------------------------------------
'Public Function dtBillingPageLineItems(sSheetName, oLineItemTable)
'
'	iRow = oLineItemTable.RowCount
'	iCol = oLineItemTable.ColumnCount(1)
'	
'	DataTable.AddSheet(sSheetName)
'	'Get Line item headers and content
'	iQuantityIndex = 0
'	For i = 1 To iRow 
'		If i = 1 Then
'			DataTable.SetCurrentRow i 
'			For j = 1 To iCol
'				'Line Item Header
'				sCellData = oLineItemTable.GetCellData(i, j)
'				Call DataTable.GetSheet(sSheetName).AddParameter(sCellData, "")
'			Next
'			i = i + 1
'		End If
'	
'		'Line Item data
'		For j = 1 To iCol 
'			sCellData = oLineItemTable.GetCellData(i, j)
'			
'			If Not sCellData = "" Then
'				 DataTable.SetCurrentRow i-1
'				DataTable.GetSheet(sSheetName).GetParameter(j).Value = sCellData
'			End If
'		Next
'	Next
'
'End Function


'' -------------------------------------------------------------------------------------
''                                    Add_to_cart
'' -------------------------------------------------------------------------------------
'Public Function Add_to_cart()
'
'	blnExist = Browser("Browser").Page("Aluminum Hook Tent Stake").Image("Add to Cart").Exist
'	counter = 1 
'
'	While Not blnExist
'		Wait(1)
'		blnExist = Browser("Browser").Page("Aluminum Hook Tent Stake").Image("Add to Cart").Exist
'		counter = counter + 1
'		If counter = 20 then 
'			blnExist = True
'		End if
'	Wend
'	
'End Function


'------------------------------------------------------------------------------------------------------------------------------------
'																cmnSearchREI(sItem)
'------------------------------------------------------------------------------------------------------------------------------------
'  Desc:  
' 		Passes a string to search for
'	 	If string is not found search will send out a warning 
'
' Args:
'		sItem = Product, Sku or Style to search for
'
'Usage:
'		cmnSearchREI(sItem) 
'
'------------------------------------------------------------------------------------------------------------------------------------
 Public Function  cmnSearchREI(sItem)

	Browser("Common").Page("REI Header").WebEdit("Search").Set sItem
	Browser("Common").Page("REI Header").Image("Search").Click
    Browser("Browser").Page("Page").Sync
	If  Browser("Common").Page("REI Header").WebElement("Search Result Msg").Exist(0.5) Then
		 sSearchMessage = Browser("Common").Page("REI Header").WebElement("Search Result Msg").GetROProperty("innertext")
	End If
  
	'Give a warning on a report if error occurs
	If InStr(sSearchMessage, sItem) = 0 Then 
		sError = Browser("Browser").Page("Page").GetROProperty("url")
		Reporter.ReportEvent micWarning, "Search test", "Error in search.  Check  url for search result for item: " & sItem	& ", " & sError
	End If	
 End Function


'------------------------------------------------------------------------------------------------------------------------------------
'												chkVerifyProductImage()
'------------------------------------------------------------------------------------------------------------------------------------
' Pre:
'			Product page is loaded
'Desc:
'			Function gets all the collection of the  product's available color and/or images and
'			loops through all colors/images and verifies that the clicked color shows up on the hero image.
'
'Usage:
'			chkVerifyProductImage()
'
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function chkVerifyProductImage()

		'Get the image description.  This should handle the image being dynamic
	Set oImageDescription = Description.Create()
	oImageDescription("class").Value = "imgSwatch"
	oImageDescription("html id").Value=""
		
	'Set any available product image to iObject
	Set oImageObject = Browser("Browser").Page("Generic Product Page").ChildObjects(oImageDescription)
		
	'Go through all the available image on the product page and do verifications
	With Browser("Browser").Page("Generic Product Page").WebElement("Hero Image Description")
		For i=0 To oImageObject.Count - 1
			oImageObject(i).Click
			Wait(1)
            
			'Verify we got the correct color.  Verification done on available hero image description
			'Get the object's color
			Act_Image_Color = oImageObject(i).GetROProperty("alt")
			'Get the hero image color
			Exp_Image_Color =oImageObject(i).GetROProperty("innertext")
		
			'Checks that the correct color gets displayed in the hero image description.  A warning is given if an error occurs
			sTitle = "Verify Hero Image Changes"
			Call chkVerifyInStrText (Exp_Image_Color, Act_Image_Color, sTitle)
		Next
	End With
		
	'Release the objects used
	Set oImageDescription =Nothing
	Set oImageObject = Nothing

End Function


'------------------------------------------------------------------------------------------------------------------------------------
'																	chkVerifySizingChart
''------------------------------------------------------------------------------------------------------------------------------------
'Pre:
'				Product page is loaded
'Desc:
'				Checks and verifies that sizing chart pops up and loads properly only when 
'				product page sizing chart is applicable
'
'Usage:
'			chkVerifySizingChart()
'
'----------------------------------------------------------------------------------------------------------------------------------------
Public Function chkVerifySizingChart()

   Call cmnGetGlobalTimeOuts_QC()
		
	Call cmnSetGlobalTimeouts (CInt(DataTable("SpecialTimeOut", dtGlobalSheet)))
	
	'check if there is a sizing chart on the product page
	If Browser("Browser").Page("Generic Product Page").Image("Sizing Chart").Exist Then
		Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))
		Browser("Browser").Page("Generic Product Page").Image("Sizing Chart").Click

		'Sync with the sizing info browser/page
		Browser("index:=0").Page("name:=sizecharts").Sync

		'------------------------------------------------------------------------------------------------------------------------------------
		'Verify page is loads
		'------------------------------------------------------------------------------------------------------------------------------------
		If Browser("index:=0").Page("name:=sizecharts").WebElement("html tag:=BODY").Exist Then
			bResult = micPass
		Else
			bResult = micFail
		End If

		sTitle = "Verify Sizing Chart Pops up"
		sDesc = "Sizing Chart appears"

		Reporter.ReportEvent bResult, sTitle,sDesc
	
		'Close the pop up browser
		Browser("index:=0").Close
	End If
		
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))

End Function   


'------------------------------------------------------------------------------------------------------------------------------------
'																chkVerifyProductPageZoom()
'----------------------------------------------------------------------------------------------------------------------------------------
' Pre:
'			Product page is loaded
'Desc:
'			Function checks that all zoom links work and that image zoom frame appears.
'
'Usage:
'			chkVerifyProductPageZoom()
'
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function chkVerifyProductPageZoom()

	Call cmnGetGlobalTimeOuts_QC()   

	'Get all the description of the links from the product page that opens up zoom
	Set oZoomLink = Description.Create()
	oZoomLink("url").Value = "http://qa.rei.com/features/zoom.html.*"
	'Get all the links 
	Set oZoomLinkObj = Browser("Browser").Page("Generic Product Page").ChildObjects(oZoomLink)
	'Go through all the links and verify
	For i =0 To oZoomLinkObj.Count-1
		'click on zoom link
		oZoomLinkObj(i).Click
		sZoomLink=oZoomLinkObj(i).GetROProperty("name")
		Browser("Browser").Page("Generic Product Page").Sync

		'Verify Zoom flash image shows
		Call cmnSetGlobalTimeouts (CInt(DataTable("SpecialTimeOut", dtGlobalSheet)))

		If  Browser("Browser").Page("Generic Product Page").Frame("Zoom Frame").WebElement("Zoom Flash").Exist Then
			bResult =Pass
		Else
			bResult=Fail
		End If

		Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))

		sTitle = "Verify Zoom image appears"
		sDesc="Zoom opened from: " & sZoomLink

		Reporter.ReportEvent bResult, sTitle,sDesc	

		'close
		Browser("Browser").Page("Generic Product Page").Link("Zoom close").Click
	Next

	' Release objects
	Set oZoomLink = Nothing
	Set oZoomLinkObj = Nothing
		
End Function


'------------------------------------------------------------------------------------------------------------------------------------
'																chkWebElemExistByIndex()
'----------------------------------------------------------------------------------------------------------------------------------------
' DESC - iterates thru the web elements on a page searching for the expected innertext.
'				 Returns a micPass/micFail
'
' ARGS - 
	'oPage - page reference of the pg to check
	 'iStartIndex - start index of theelement (where to start)
	 'iEndIndex -  end index of theelement (where to stop)
	 'sInnerText - Inner text expected and searched for
'
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function chkWebElemExistByIndex(oPage, iStartIndex, iEndIndex, sInnerText)

	Call cmnGetGlobalTimeOuts_QC()
  
	Exp_Text = Trim(sInnerText)
	Act_Text = "Failed to find the WebElement"
	iStartIndex = CInt(iStartIndex)
	iEndIndex = CInt(iEndIndex)
	Set oPage = oPage

	Call cmnSetGlobalTimeouts (CInt(DataTable("SpecialTimeOut", dtGlobalSheet)))

	'Find your WebElement by iterating thru Elem indexes and comparing innertext
	For i = iStartIndex To iEndIndex
		'if you find inner text at all
		If "" <> Trim(oPage.WebElement("index:=" & CStr(i)).GetROPRoperty("innertext")) Then
			'Capture the inner text if you got one
			Act_Text = Trim(oPage.WebElement("index:=" & CStr(i)).GetROPRoperty("innertext"))
			If Exp_Text = Act_Text Then
				Exit For
			End If
		End If
	Next

	'Compare Act_InnerText vs Exp_InnerText - Return micPass/micFail
	sTitle = "Verify that a WebElement with an innertext for '" & Exp_Text & " displays on the page"
	Call chkVerifySimpleText (Exp_Text, Act_Text, sTitle)
	
	Call cmnSetGlobalTimeouts (CLng(DataTable("DefaultTimeOut", dtGlobalSheet)))

End Function


' -------------------------------------------------------------------------------------
'                               chkVerifyLinks
' -------------------------------------------------------------------------------------
Public Function chkVerifyLinks(sLinkName, sDestPageTitle)

	Set oPage = Browser("Browser").Page("Common Link Verification Objects")
   
	' -----------------------------------------------------------
	'                            Select Link
	' -----------------------------------------------------------   
	oPage.Link(sLinkName).Click
	oPage.Sync
	
	' -----------------------------------------------------------
	'                      Verify Link Redirect
	' -----------------------------------------------------------  
	Exp_Title = Trim(sDestPageTitle)
	Act_Title = Trim(oPage.WebElement("Generic Page Title").GetROProperty("innertext"))
	
	If Act_Title = Exp_Title Then
		bResult = micPass
		sResult = "PASS"
	Else
		bResult = micFail
		sResult = "FAIL"		
	End If

	sTitle = "Verify Link Navigation - '" & sLinkName & "' Link"
	sDesc = "Expected Page: " & Exp_Title & vbcr & "Actual Page: " & Act_Title

	Call Reporter.ReportEvent (bResult,sTitle, sDesc)
	chkVerifyLinks = sResult

	Browser("Browser").Back
	oPage.Sync

	' Release objects
	Set oPage = Nothing

End Function


' -------------------------------------------------------------------------------------
'                    chkWarmAndFuzzyLinks(sLinkName)
' -------------------------------------------------------------------------------------
Public Function chkWarmAndFuzzyLinks(sLinkName)

	sTitle = "Verify Page " & sLinkName
	sPageTitle = Browser("title:=.*").Page("title:=.*").GetROProperty("title")
	If bFindInText(sLinkName, sPageTitle) Then
		bResult = micPass
		sDetails = "Correct page loads" & vbCr & sPageTitle
	Else
		bResult = micFail
		sDetails = "Incorrect page loads" & vbCr & sPageTitle
	End If
	Reporter.ReportEvent bResult, sTitle, sDetails

End Function


'--------------------------------------------------------------------
'		ActivateNewSearch
'-------------------------------------------------------------------
'Desc: Function just activates new search
'-------------------------------------------------------------------
Public Function ActivateNewSearch()
	SystemUtil.Run "iexplore", ""
	Wait(1)
	
	Browser("Browser").Page("Page").Sync
	Browser("Browser").Navigate("http://psqa.rei.com/siteversion.html")
	Browser("Browser").Page("Page").Sync

	Set oSearchDesc = Description.Create
	oSearchDesc("html tag").Value = "input"
	oSearchDesc("name").Value = "version"
	oSearchDesc("type").Value = "radio"
	Set oSearch = Browser("Browser").Page("Page").ChildObjects(oSearchDesc)

	oSearch(0).Select "#1"

	Browser("Browser").Page("Page").Sync
	Browser("Browser").Close

	' Release objects
	Set oSearchDesc = Nothing
	Set oSearch = Nothing
	
End Function


' --------------------------------------------------------------------------------------
'                                               cmnLogin2
' --------------------------------------------------------------------------------------
'Logs into the online application usine environmental variables 
'(envUID & envPWD) generated from the 'Gen_Register_user [REI.COM]' 
'action from the ONLINE reusable action library.

'Req - Gen_Register_user [REI.COM] must be run first to generate the env vars
' --------------------------------------------------------------------------------------
Public Function cmnLoginREI ()

	sUserID = Environment("envUID")
	sPWD = Environment("envPWD")

	Browser("Browser").Page("Page").Sync
	
	Browser("Account").Page("REI.com: Login").WebEdit("logonId").Set sUserID
	Browser("Account").Page("REI.com: Login").WebEdit("logonPassword").Set 	sPWD
	Browser("Account").Page("REI.com: Login").Link("log in").Click

	Browser("Browser").Page("Page").Sync

End Function


' --------------------------------------------------------------------------------------
'                      cmnSetTestID
' --------------------------------------------------------------------------------------
' Returns a unique test ID between 1000-99999
'Adds a column called TestID to the GlobalSheet & populates it with the ID

' --------------------------------------------------------------------------------------
Public Function cmnSetTestID()
	sTestID = "TestID-" & CStr(Number (1000, 99999) )
	Call dtAddColumn ("Global", "TestID", sTestID)
	cmnSetTestID = sTestID
End Function


' --------------------------------------------------------------------------------------
'                      cmnGetTestID
' --------------------------------------------------------------------------------------
' Returns the previously generated test ID  from the global sheet

' --------------------------------------------------------------------------------------
Public Function cmnGetTestID()
	cmnGetTestID = Trim(DataTable("TestID", dtGlobalSheet))
End Function


' --------------------------------------------------------------------------------------
'                        cmnRemoveShippingAddys
' --------------------------------------------------------------------------------------
'Removes all addresses from shipping addy list

' --------------------------------------------------------------------------------------
Public Function cmnRemoveShippingAddys()

   	Browser("Browser").Page("Address Book").Sync
	Call cmnSetGlobalTimeouts (1000)
	
	Do While Browser("Browser").Page("Address Book").Link("Remove").Exist
		Browser("Browser").Page("Address Book").Link("Remove").Click
        Browser("Browser").Page("Address Book").Sync
	Loop

	Call cmnSetGlobalTimeouts (30000)
	Browser("Browser").Page("Address Book").Sync
	
End Function


'----------------------------------------------------------------------------------------------------------
'							cmnSendEmailToTester
'----------------------------------------------------------------------------------------------------------
' Desc:
'			Send email to tester to verify manual steps
' Args:
'			sMessage: Message in the email to send
' Usage
'			cmnSendEmailToTester( sMessage)
'----------------------------------------------------------------------------------------------------------
Public Function cmnSendEmailToTester(sMessage)

   'Get username
	If Not QCUtil.IsConnected Then
		Reporter.ReportEvent micWarning, "Sending email to tester", "No email sent because test was not run in Quality Center"
	Else
        sUserName = QCUtil.QCConnection.UserName
        sSendTo = sUserName & "@rei.com"
		sTestName = Environment("TestName")'QCUtil.CurrentTest.Name
		QCUtil.QCConnection.SendMail sSendTo, "", "Manual Check for:" & sTestName, sMessage, "", ""
	End If

End Function


'----------------------------------------------------------------------------------------------------------
'						   GetSubtotal
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes the review & pay  or receipt page and returns the current 'Subtotal' of all of the products in the cart
'
'----------------------------------------------------------------------------------------------------------
Public Function GetSubtotal()

	intSubtotalRow = (Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetRowWithCellText("Subtotal"))
	GetSubtotal = Ccur(Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetCellData(intSubtotalRow,2))

End Function


'----------------------------------------------------------------------------------------------------------
'						   GetShipTotal
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes the review & pay  or receipt page and returns the  'Shipping'  price
'  	
'----------------------------------------------------------------------------------------------------------
Public Function GetShipTotal()

	intSubtotalRow = (Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetRowWithCellText("Subtotal"))
	GetShipTotal = CCur(Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetCellData(intSubtotalRow + 1,2))

End Function


'----------------------------------------------------------------------------------------------------------
'						   GetTaxTotal
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes the review & pay  or receipt page and returns the  'Total Tax'  price
'  	
'----------------------------------------------------------------------------------------------------------
Public Function GetTaxTotal()

	intTaxRow = (Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetRowWithCellText("Tax"))
	GetTaxTotal = CCur(Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetCellData(intTaxRow ,2))

End Function


'----------------------------------------------------------------------------------------------------------
'						   GetTotalDue
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes the review & pay  or receipt page and returns the  'Total Due'  price
'  	
'----------------------------------------------------------------------------------------------------------
Public Function GetTotalDue()

	intTotalDueRow = (Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetRowWithCellText("Total Due:"))
	GetTotalDue = CCur(Browser("Browser").Page("REI Payment Information").WebTable("Product table").GetCellData(intTotalDueRow ,2))

End Function


'----------------------------------------------------------------------------------------------------------
'						   GetPrice
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes a string with a $ symbol and returns the price after it in a currency format
' Args:
'			sPrice: string containing $
' Usage
'			productPrice = GetPrice(Browser("Browser").Page("Generic Product Page").WebList("Select Color or Size").GetROProperty ("value"))
'----------------------------------------------------------------------------------------------------------
Public Function GetPrice(sPrice)

	If InStr(sPrice, "$") Then
				GetPrice = CCur(Right(sPrice, Len(sPrice ) - (InStr(sPrice, "$"))))
		Else
			   ' Msgbox("Error!! Price string does not contain a $ symbol!")
	End If
End Function


'----------------------------------------------------------------------------------------------------------
'						   PriceThreshold(minPrice)
'----------------------------------------------------------------------------------------------------------
' Desc:
'			This function takes a minimum price that the shopping cart needs to be at and keeps increasing the qty until it hits that price or higher
' Args:
'			minPrice:  the minimum price you want the cart to be
' Usage
'			must be called on the shopping cart page and must have an item in the cart

				' PriceThreshold(5)
'----------------------------------------------------------------------------------------------------------
Public Function PriceThreshold(minPrice)

	originalprice = Environment.Value("productPrice")
	price = originalprice
	qty = Browser("Checkout").Page("REI.com: Shopping Basket").WebEdit("Quantity").GetROProperty("value")
	
	Do while price < minPrice
		Browser("Checkout").Page("REI.com: Shopping Basket").WebEdit("Quantity").Set(qty + 1)
		Browser("Checkout").Page("REI.com: Shopping Basket").Link("Update").Click
		Browser("Checkout").Page("REI.com: Shopping Basket").Sync
		qty = qty + 1
		price = price + originalprice
	Loop

	Environment.Value("productPrice") = price

End Function


'----------------------------------------------------------------------------------------------------------
'							chkShippingPriceAndEAD
'----------------------------------------------------------------------------------------------------------
' Desc:
'			Checks and verifies shipping price and estimated arrival dates
' Args:
'			oPage: Shipping page
' Usage
'			oPage = Browser("Browser").Page("REI: Shipping Address")
'			chkShippingPriceAndEAD(oPage)
'----------------------------------------------------------------------------------------------------------
Public Function chkShippingPriceAndEAD(oPage)

	Set oRadioGroupDesc = Description.Create
	oRadioGroupDesc("type").Value = "radio"
	oRadioGroupDesc("html tag").Value = "INPUT"
	Set oRadioGroup  = oPage.ChildObjects(oRadioGroupDesc)

	iShipOptionCount = oRadioGroup(0).GetROProperty("items count")
	For i = 0 To iShipOptionCount -1 

		' Verify price shows
		sOption = oPage.WebElement("Innertext:=(.*\$[1-9]+\.[0-9][0-9])","html tag:=label",  "index:=" & i).GetROProperty("innertext")
		sShippingOption = Left(sOption, InStrRev(sOption, ":"))

		If InStr(sOption, "$") > 0 Then
			bResult = micPass
			sPrice = Right(sOption, Len(sOption) - (InStr(sOption, "$")-1))
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If

		Reporter.ReportEvent bResult, sShippingOption & " Shipping price", sDetails
	
		' Verify EAD shows on page
		If i < 3 Then
			sEAD = oPage.WebElement("Innertext:=Estimated.*", "index:=" & i ).GetROProperty("innertext")
			If Not IsEmpty(sEAD) Then
				bResult = micPass
				sDetails = "Verified EAD Appears" & vbcr & sEAD
			Else
				bResult = micFail
				sDetails = "EAD missing!"
			End If
		Else
			bResult = micPass
			sDetails = "No EAD for International shipping"
		End If

		Reporter.ReportEvent bResult, sShippingOption & " Verify EAD shows", sDetails	
	Next

	'Free objects used
	Set oRadioGroupDesc = Nothing
	Set oRadioGroup = Nothing

End Function


Public Function chkShippingPriceAndEAD1(oPage)

	If oPage.WebElement("innertext:=Standard Shipping","index:=0").Exist  Then

		'Browser("Browser").Page("REI: Shipping Address").WEbTable("name:=shipping_option","index:=0").highlight
		iRowStandard = oPage.WEbTable("name:=shipping_option","index:=0").GetRowWithCellText("Standard Shipping")
		sCellDataEAD =oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowStandard,2)
		sCellDataPrice = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowStandard,3)

		sStandardEAD = sCellDataEAD
		sStandardPrice = sCellDataPrice
		'Msgbox sStandardEAD
		
		'Verify EAD shows on page
		If InStr(UCase(sStandardEAD), "MONDAY") > 0 or InStr(UCase(sStandardEAD), "TUESDAY") > 0 or InStr(UCase(sStandardEAD), "WEDNESDAY") > 0 or InStr(UCase(sStandardEAD), "THURSDAY") > 0 or InStr(UCase(sStandardEAD), "FRIDAY") > 0 or InStr(UCase(sStandardEAD), "SATURDAY") > 0 or InStr(UCase(sStandardEAD), "SUNDAY") > 0 Then
			bResult = micPass
			sDetails = "Verified EAD Appears" & vbcr & sStandardEAD
		Else
			bResult = micFail
			sDetails = "EAD missing!"
		End If
		Reporter.ReportEvent bResult, sStandardEAD & " Shipping EAD", sDetails
		
			'Verify Price shows on page
		If InStr(sStandardPrice, "$") > 0 Then
			bResult = micPass
			sPrice = sStandardPrice
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If
		Reporter.ReportEvent bResult, sPrice & " Shipping price", sDetails

	'Else
		'Reporter.ReportEvent micFail, "Satndard Shipping" & "---- Sdandard Shipping option is not available"

	End If


	If oPage.WebElement("innertext:=Two-Day Shipping","index:=0").Exist  Then

	iRowTwoDay = oPage.WEbTable("name:=shipping_option","index:=0").GetRowWithCellText("Two-Day Shipping")
	sCellDataEAD = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowTwoDay,2)
	sCellDataPrice = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowTwoDay,3)

	sTwoDayEAD = sCellDataEAD
	sTwoDayPrice = sCellDataPrice
	'Msgbox sTwoDayEAD
		If InStr(UCase(sTwoDayEAD), "MONDAY") > 0 or InStr(UCase(sTwoDayEAD), "TUESDAY") > 0 or InStr(UCase(sTwoDayEAD), "WEDNESDAY") > 0 or InStr(UCase(sTwoDayEAD), "THURSDAY") > 0 or InStr(UCase(sTwoDayEAD), "FRIDAY") > 0 or InStr(UCase(sTwoDayEAD), "SATURDAY") > 0 or InStr(UCase(sTwoDayEAD), "SUNDAY") > 0 Then
			bResult = micPass
			sDetails = "Verified EAD Appears" & vbcr & sTwoDayEAD
		Else
			bResult = micFail
			sDetails = "EAD missing!"
		End If
		Reporter.ReportEvent bResult, sTwoDayEAD & " Shipping EAD", sDetails

		If InStr(sTwoDayPrice, "$") > 0 Then
			bResult = micPass
			sPrice = sTwoDayPrice
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If
		Reporter.ReportEvent bResult, sPrice & " Shipping price", sDetails

	'Else
		'Reporter.ReportEvent micFail, "TwoDay Shipping" & "---- TwoDay Shipping option is not available"

	End If

	If oPage.WebElement("innertext:=One-Day Express Shipping","index:=0").Exist  Then

	iRowOneDay =oPage.WEbTable("name:=shipping_option","index:=0").GetRowWithCellText("One-Day Express Shipping")
	sCellDataEAD = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowOneDay,2)
	sCellDataPrice = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowOneDay,3)

	sOneDayEAD = sCellDataEAD
	sOneDayPrice = sCellDataPrice
	'Msgbox sOneDayEAD
		If InStr(UCase(sOneDayEAD), "MONDAY") > 0 or InStr(UCase(sOneDayEAD), "TUESDAY") > 0 or InStr(UCase(sOneDayEAD), "WEDNESDAY") > 0 or InStr(UCase(sOneDayEAD), "THURSDAY") > 0 or InStr(UCase(sOneDayEAD), "FRIDAY") > 0 or InStr(UCase(sOneDayEAD), "SATURDAY") > 0 or InStr(UCase(sOneDayEAD), "SUNDAY") > 0 Then
			bResult = micPass
			sDetails = "Verified EAD Appears" & vbcr & sOneDayEAD
		Else
			bResult = micFail
			sDetails = "EAD missing!"
		End If
		Reporter.ReportEvent bResult, sOneDayEAD & " Shipping EAD", sDetails

		If InStr(sOneDayPrice, "$") > 0 Then
			bResult = micPass
			sPrice = sOneDayPrice
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If
		Reporter.ReportEvent bResult, sPrice & " Shipping price", sDetails

	'Else
		'Reporter.ReportEvent micFail, "One-Day Express Shipping" & "---- One-Day Express Shipping option is not available"

	End If


	If oPage.WebElement("innertext:=Canadian Air Parcel Post","index:=0").Exist  Then

	iRowCanada =oPage.WEbTable("name:=shipping_option","index:=0").GetRowWithCellText("Canadian Air Parcel Post")
	sCellDataEAD = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowCanada,2)
	sCellDataPrice = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowCanada,3)

	sCanadaEAD = sCellDataEAD
	sCanadaPrice = sCellDataPrice
	'Msgbox sOneDayEAD
		If InStr(UCase(sCanadaEAD), "MONDAY") > 0 or InStr(UCase(sCanadaEAD), "TUESDAY") > 0 or InStr(UCase(sCanadaEAD), "WEDNESDAY") > 0 or InStr(UCase(sCanadaEAD), "THURSDAY") > 0 or InStr(UCase(sCanadaEAD), "FRIDAY") > 0 or InStr(UCase(sCanadaEAD), "SATURDAY") > 0 or InStr(UCase(sCanadaEAD), "SUNDAY") > 0 Then
			bResult = micFail
			sDetails = "EAD is displayed for Canadian Air Parcel Post!"
		Else
			bResult = micPass
			sDetails = "Verified EAD Appears" & vbcr & "EAD is not displayed for Canadian Air Parcel post as expected"
		End If
		Reporter.ReportEvent bResult, " Shipping EAD", sDetails

		If InStr(sCanadaPrice, "$") > 0 Then
			bResult = micPass
			sPrice = sCanadaPrice
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If
		Reporter.ReportEvent bResult, sPrice & " Shipping price", sDetails

	'Else
		'Reporter.ReportEvent micFail, "One-Day Express Shipping" & "---- One-Day Express Shipping option is not available"

	End If

	If oPage.WebElement("innertext:=International Air Shipping","index:=0").Exist  Then

	iRowInternational =oPage.WEbTable("name:=shipping_option","index:=0").GetRowWithCellText("International Air Shipping")
	sCellDataEAD = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowInternational,2)
	sCellDataPrice = oPage.WEbTable("name:=shipping_option","index:=0").GetCellData(iRowInternational,3)

	sInternationalEAD = sCellDataEAD
	sInternationalPrice = sCellDataPrice
	'Msgbox sOneDayEAD
		If InStr(UCase(sInternationalEAD), "MONDAY") > 0 or InStr(UCase(sInternationalEAD), "TUESDAY") > 0 or InStr(UCase(sInternationalEAD), "WEDNESDAY") > 0 or InStr(UCase(sInternationalEAD), "THURSDAY") > 0 or InStr(UCase(sInternationalEAD), "FRIDAY") > 0 or InStr(UCase(sInternationalEAD), "SATURDAY") > 0 or InStr(UCase(sInternationalEAD), "SUNDAY") > 0 Then
			bResult = micFail
			sDetails = "EAD is displayed for International Air Shipping!"
		Else
			bResult = micPass
			sDetails = "Verified EAD Appears" & vbcr & "EAD is not displayed for International Air Shipping as expected"
		End If
		Reporter.ReportEvent bResult, " Shipping EAD", sDetails

		If InStr(sInternationalPrice, "$") > 0 Then
			bResult = micPass
			sPrice = sInternationalPrice
			sDetails = "Shipping price appears: " & sPrice
		Else
			bResult = micFail
			sDetails = "Shipping price missing!"
		End If
		Reporter.ReportEvent bResult, sPrice & " Shipping price", sDetails

	'Else
		'Reporter.ReportEvent micFail, "One-Day Express Shipping" & "---- One-Day Express Shipping option is not available"

	End If

End Function


'------------------------------------------------------------------------
'                      bFindInText(sSearch, sString)
'------------------------------------------------------------------------
'Desc: Finds a text in a given text using 
'				regular expression in case InStr function
'				is not cutting it
'				Returns true if the search string is found
'			otherwise returns false
'Args
'			sSearch: String to search
'			sString:  Text where to search the
'			string from
'Usage: (this will return false)
' sSearc = "Gift Card"
'sString = "This is a e-Gift Card"
'bFindInText(sSearch, sString)
'------------------------------------------------------------------------
Public Function bFindInText(sSearch, sString)

	Set sRegEx = New RegExp
	'Use a regular expression to get the exact word to look for
	sString = " " & sString
	sRegEx.Pattern = ".*(\s)" & sSearch & "*"
	sRegEx.IgnoreCase = TRUE

	sValue = sRegEx.Test(sString)

	If sValue Then
		bFindInText = sValue
'		Print "Found " & sSearch & " from " & sString
	Else
		bFindInText = sValue
		'Print "Word not found "  & sSearch & " from " & sString
	End If

	sString = Trim(sString)

	' Release objects
	Set sRegEx = Nothing
   
End Function


'------------------------------------------------------------------------
'                      cmnGetScreenShot(oImage)
'------------------------------------------------------------------------
'Desc: Takes a screenshot of the object and 
'		attaches the image to the run instance from 
'		Quality Center.  
'		All images will be saved in the local drive
'Args
'			oImage: Any object(link, image, page, etc)
'		that you want to take a snapshot of
'
'Usage: 
'Set oImage = Browser("browser").Page("page").Image("image")
'Call cmnGetScreenShot(oImage)
'------------------------------------------------------------------------
Public Function cmnGetScreenShot(oImage)
   If QCUtil.IsConnected Then
	   
	   iCurrentRunID = QCUtil.CurrentRun.ID

		iCount = cInt(Environment("ScreenshotCount"))
		iCount = iCount + 1
		
		Environment("ScreenshotCount") = iCount
		sImageName = "ScreenShot" & iCount & ".png"

		sFilePath =  Environment("ResultDir") & "\Report\"
        oImage.CaptureBitmap sFilePath & sImageName		

		Set oTDConnection = QCUtil.TDConnection
		Set oRNFactory = oTDConnection.RunFactory.Item(iCurrentRunID).Attachments
		
		Set oAttachment =oRNFactory.AddItem(Null)
		
		oAttachment.FileName = sFilePath
		oAttachment.Type = 1
		oAttachment.Post()

        sXML = "]]></Disp><BtmPane ><Path><![CDATA[" & sImageName & "]]></Path></BtmPane><Disp>"
		Reporter.ReportEvent micInfo, "QA Screenshot" & sXML ,  "Screenshot captured during the test;"
		
		' Release objects
		Set oTDConnection = Nothing
		Set oRNFactory = Nothing
		Set oAttachment = Nothing

		'Delete the screenshots from local drive
		'Set fso = CreateObject("Scripting.FileSystemObject")
		'set fsoFile = fso.GetFile(sFilePath & sImageName)
		'fsoFile.Delete()

	Else
		Reporter.ReportEvent micDone, "Capture ScreenShot", "You must run test via Quality Center to see the screenshot"
	End If

End Function


'------------------------------------------------------------------------
'                      StopTest()
'------------------------------------------------------------------------
'Desc: Stops the test as if a user would press the 
'			'stop button.  This differs from "ExitTest"
'			in a way that this will give a status of 
'			Not Complete
'Args:
'			NONE
'Usage: 
'Call StopTest()
'------------------------------------------------------------------------
Public Function StopTest()
	sTestStatus = Reporter.RunStatus
	Select Case sTestStatus

		' a failure is found on the test, stop test by exiting to retain falied status
		Case micFail	  
			Reporter.ReportEvent micFail,"QA Stopping Test with a failure", "Please check reporter before updating test status"
			'ExitTest

		Case Else
			Reporter.ReportEvent micInfo,"QA Stopping Test", "Test needs to be manually checked and updated."
			If Browser("Browser").Exist(0.5) Then
				cmnCloseBrowsers ()
			End If
			Set oMercBtn = CreateObject("Mercury.DeviceReplay")
			oMercBtn.PressKey 62
			Set oMercBtn = Nothing

	End Select

End Function


'---------------------------------------------------------------------------------------------------------------------
'											TimerPopup(sTitle, iWaitTime, sMessage)
'---------------------------------------------------------------------------------------------------------------------
'Desc:  Pop up with showing a message and elapsed time when a test is in a 
'	waiting period.  This will also indicate that the test has not frozen not QTP has 
'	crashed
'	
'Args:
'		sTitle: Title of the pop up window
'		iWaitTime: Time of how long the wait time is
'		sMEssage:  Message to show'		
'
'Usage'
'	Call TimerPopup(sTitle, iWaitTime, sMessage)
'---------------------------------------------------------------------------------------------------------------------
Public Function TimerPopup( iWaitTime, sMessage)
	
	Set frmForm = DotNetFactory("System.Windows.Forms.Form", "System.Windows.Forms")
	Set lblLabel1 = DotNetFactory("System.Windows.Forms.Label", "System.Windows.Forms")
	Set lblLabel2 = DotNetFactory("System.Windows.Forms.Label", "System.Windows.Forms")
	
	'Label 1 properties
	 lblLabel1.AutoEllipsis = True
	 lblLabel1.Text  = Environment("TestName")   
	 y = cInt(lblLabel1.Height )

	'Form properties
	frmForm.MaximizeBox = False
	frmForm.MinimizeBox = False
	frmForm.ControlBox = True
	frmForm.Text = "QA Test Timer(Do not close)"

	'add label 1 to the form
    frmForm.Controls.Add(lblLabel1)
	 frmForm.Controls.Item(0).Width = frmForm.Width
	 y = cInt(lblLabel1.Height )
     frmForm.Controls.Item(0).Height = y + 25	 

	 'Label 2 properties
	lblLabel2.Top = frmForm.Controls.Item(0).Height
	
	'add label 2 to the form
 	frmForm.Controls.Add(lblLabel2)
	frmForm.Controls.Item(1).Width = frmForm.Width
	frmForm.Controls.Item(1).Height = cInt( frmForm.Height )  	

	'frmForm.
	frmForm.Activate
	frmForm.TopMost = True
    frmForm.Show()

    For i = 1 To iWaitTime
		iMin = Fix(i/60)
		iSec = i Mod 60
		
		If Len(cStr(iSec)) = 1 Then
			iSec = "0" & cStr(iSec)
		End If
		
		lblLabel2.Text = sMessage & " " & iMin & ":" & iSec
		frmForm.Show()
		frmForm.Refresh
		Wait(1)
	Next
	
	frmForm.Dispose

	' Release objects
	Set frmForm = Nothing
	Set lblLabel1 = Nothing
	Set lblLabel2 = Nothing

End Function


'--------------------------------------------------------------------------------------------
'		cmnIEBrowserDisableJavaScript(oBrowser)
'--------------------------------------------------------------------------------------------
'Description:
'		Disables JavaScript for IE browser
'Args:
'	oBrowser: Object browser to use so QTP knows which of
'					the Browser or Browsers setting to set
'Usage:
'	Set oBrowser = Browser("Browser")
'	Call cmnIEBrowserDisableJavaScript(oBrowser)
'--------------------------------------------------------------------------------------------
Public Function cmnIEBrowserDisableJavaScript(oBrowser)

	sIEVersion = oBrowser.GetROProperty("application version")
	'IE v6
   If sIEVersion = "internet explorer 6" Then
	   oBrowser.WinToolbar("nativeclass:=ToolbarWindow32","location:=0", "window id:=0").Press "&Tools"
	   Wait(1)
		oBrowser.WinMenu("menuobjtype:=3").Select "Internet Options..."
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinTab("nativeclass:=SysTabControl32").Select "Security"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("nativeclass:=Button", "text:=&Custom Level...").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings").WinTreeView("nativeclass:=SysTreeView32").Select "Scripting;Active scripting;Disable"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings").WinButton("text:=OK").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Warning!").WinButton("text:=&Yes").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("text:=OK").Click

	'IE v7	
	Else 
		oBrowser.WinToolbar("nativeclass:=ToolbarWindow32", "location:=1", "window id:=0").Press "&Tools"
		Wait(1)
		oBrowser.WinMenu("menuobjtype:=3").Select "Internet Options"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinTab("nativeclass:=SysTabControl32").Select "Security"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("nativeclass:=Button", "text:=&Custom Level...").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").WinTreeView("nativeclass:=SysTreeView32").Select "Scripting;Active scripting;Disable"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").WinButton("text:=OK").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").Dialog("text:=Warning!").WinButton("text:=&Yes").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("text:=OK").Click

	End If

	'wait until the dialog is closed
	Wait(1)

End Function


'--------------------------------------------------------------------------------------------
'		cmnIEBrowserEnableJavaScript(oBrowser)
'--------------------------------------------------------------------------------------------
'Description:
'		Enables JavaScript for IE browser
'Args:
'	oBrowser: Object browser to use so QTP knows which of
'					the Browser or Browsers setting to set
'Usage:
'	Set oBrowser = Browser("Browser")
'	Call cmnIEBrowserEnableJavaScript(oBrowser)
'--------------------------------------------------------------------------------------------
Public Function cmnIEBrowserEnableJavaScript(oBrowser)

	sIEVersion = oBrowser.GetROProperty("application version")
	'IE v6
   If sIEVersion = "internet explorer 6" Then
	oBrowser.WinToolbar("nativeclass:=ToolbarWindow32","location:=0", "window id:=0").Press "&Tools"
	Wait(1)
	oBrowser.WinMenu("menuobjtype:=3").Select "Internet Options..."
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").WinTab("nativeclass:=SysTabControl32").Select "Security"
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").WinButton("nativeclass:=Button", "text:=&Custom Level...").Click
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings").WinTreeView("nativeclass:=SysTreeView32").Select "Scripting;Active scripting;Enable"
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings").WinButton("text:=OK").Click
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Warning!").WinButton("text:=&Yes").Click
	Wait(1)
	oBrowser.Dialog("text:=(.*)Options").WinButton("text:=OK").Click

	'IE v7
	Else
		oBrowser.WinToolbar("nativeclass:=ToolbarWindow32","location:=1", "window id:=0").Press "&Tools"
		Wait(1)
		oBrowser.WinMenu("menuobjtype:=3").Select "Internet Options"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinTab("nativeclass:=SysTabControl32").Select "Security"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("nativeclass:=Button", "text:=&Custom Level...").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").WinTreeView("nativeclass:=SysTreeView32").Select "Scripting;Active scripting;Enable"
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").WinButton("text:=OK").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").Dialog("text:=Security Settings(.*)").Dialog("text:=Warning!").WinButton("text:=&Yes").Click
		Wait(1)
		oBrowser.Dialog("text:=(.*)Options").WinButton("text:=OK").Click
	End If
	'wait until the dialog is closed
	Wait(1)

End Function


'---------------------------------------------------------------------------------------------------------------------
'                                 cmnVerifyAccountsEmailConfirmation
'---------------------------------------------------------------------------------------------------------------------
'Desc: 
'			Test verification for incoming mail message for:
'			 Admin Your Account
'			Admin Gift Registry
'			Online Your Account
'			Online Gift Registry
'
'Important note:
'			To use this function and its helper function Outlook must be installed.  In addition,
'		there must be a default profile set up.
'
'	Args:
'			sGetTime: used for filtering any emails beginning from the given time
'			sEmailAccount: E-mail account to check  (use: qaauto1@rei.com)
' 			sEmailTestType: checks for Admin, Online or All email
'			iTimeout = in seconds
'
'Usage:
'			For Admin emails
'				cmnVerifyEmailConfirmation (sEmailAccount, sStartTestTime, "Admin")
'			For Online emails
'				cmnVerifyEmailConfirmation (sEmailAccount, sStartTestTime, "Online")
'			For All emails
'				cmnVerifyEmailConfirmation (sEmailAccount, sStartTestTime, "All)"
'---------------------------------------------------------------------------------------------------------------------------
Public Function cmnVerifyEmailConfirmation (sEmailAccount, sStartTestTime,sEmailTestType, iTimeout)
	Call msxOpenEmailDialog()
	Call msxCreateTestEmailAccount(sEmailAccount)
	Call msxConnectToInbox()
	Dim oMailItem
			
	Call msxGetEmailFromInbox(sStartTestTime, oMailItem, iTimeout)
	If Not IsEmpty (oMailItem) Then
	
		Select Case UCase(sEmailTestType)

			Case "ADMIN"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]=Your REI Online Account Setup")
				Call cmnEmailVerification(sStartTestTime, "Admin Your Account", oMails, iTimeout)
	
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Welcome to REI Gift Registry!")
				Call cmnEmailVerification(sStartTestTime, "Admin Gift Registry", oMails, iTimeout)

			Case "ONLINE"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]='REI Customer Service: Your Account is Now Set Up '")
				Call cmnEmailVerification(sStartTestTime, "Your Account", oMails, iTimeout)
				
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Welcome to REI Gift Registry!")
				Call cmnEmailVerification(sStartTestTime, "Gift Registry", oMails, iTimeout)
			
				Set oMails=oMailItem.Restrict("[unread] = true  And [subject]=REI thanks you for shopping online")
				Call cmnEmailVerification(sStartTestTime, "Order Confirmation", oMails, iTimeout)

			Case "ALL"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]=Your REI Online Account Setup")
				Call cmnEmailVerification(sStartTestTime, "Admin Your Account", oMails, iTimeout)
	
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]='REI Customer Service: Your Account is Now Set Up '")
				Call cmnEmailVerification(sStartTestTime, "Your Account", oMails, iTimeout)
	
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Welcome to REI Gift Registry!")
				Call cmnEmailVerification(sStartTestTime, "Admin Gift Registry", oMails, iTimeout)
				Call cmnEmailVerification(sStartTestTime, "Gift Registry", oMails, iTimeout)
	
				Set oMails=oMailItem.Restrict("[unread] = true  And [subject]=REI thanks you for shopping online")
				Call cmnEmailVerification(sStartTestTime, "Order Confirmation", oMails, iTimeout)
	
			End Select
		
		Else
			Reporter.ReportEvent  micWarning, sEmailConfirmationType & " E-mail test", "No Email recieved after waiting  " & CStr(iTimeout) & " seconds.  Please confirm email"
        End If	
	Call msxDisconnectFromInbox()
	Call msxDeleteTestEmailAccount()
	Call msxCloseEmailDialog()

	' Release objects
	Set oMails = Nothing
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'                                      cmnVerifyOneEmailConfirmation
'---------------------------------------------------------------------------------------------------------------------
'Desc: 
'			Test verification for only one mail message for:
'			 Admin Your Account
'			Admin Gift Registry
'			Online Your Account
'			Online Gift Registry
'			Order Confirmation
'
'Important note:
'			To use this function and its helper function Outlook must be installed.  In addition,
'		there must be a default profile set up.
'
'	Args:
'			sGetTime: used for filtering any emails beginning from the given time
'			sEmailAccount: E-mail account to check (use: qaauto1@rei.com)
'			sEmailTestType: type of email to look for (see Case statement)
'			iTimeout = in seconds
'
'Usage:
'			For Admin emails
'				cmnVerifyOneEmailConfirmation (sEmailAccount, sStartTestTime, "Your Account")
'---------------------------------------------------------------------------------------------------------------------------
Public Function cmnVerifyOneEmailConfirmation (sEmailAccount, sStartTestTime, sEmailTestType, iTimeout)
	Call msxOpenEmailDialog()
	Call msxCreateTestEmailAccount(sEmailAccount)
	Call msxConnectToInbox()
	Dim oMailItem

	Call msxGetEmailFromInbox(sStartTestTime, oMailItem, iTimeout)

	If Not IsEmpty (oMailItem) Then
		
		Select Case UCase(sEmailTestType)

			Case "YOUR ACCOUNT"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]='REI Customer Service: Your Account is Now Set Up '")
				Call cmnEmailVerification(sStartTestTime, sEmailTestType, oMails, iTimeout)

			Case "GIFT REGISTRY"
				'Verify Create Gift Registry email
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Welcome to REI Gift Registry!")
				Call cmnEmailVerification(sStartTestTime, sEmailTestType, oMails, iTimeout)

			Case "ADMIN YOUR ACCOUNT"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]=Your REI Online Account Setup")
				Call cmnEmailVerification(sStartTestTime, sEmailTestType, oMails, iTimeout)

			Case "ADMIN GIFT REGISTRY"
				 'Verify Your Account email		
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Welcome to REI Gift Registry!")
				Call cmnEmailVerification(sStartTestTime, sEmailTestType, oMails, iTimeout)

			Case "UPDATE ADMIN GIFT REGISTRY"
				'Verify Update Gift Registry email
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Gift Registry And [subject]=Your REI Gift Registry Has Been Updated")
				If oMails.Count > 0 Then
					For Each Mail In oMails
						Call dtEmailInfotoDataSheet("AGRMailUpdate", Mail.Body)
					Next
				End If
				
			Case "ORDER CONFIRMATION"
				Set oMails=oMailItem.Restrict("[unread] = true  And [From]=REI Customer-Service And [subject]=REI thanks you for shopping online")
				Call cmnEmailVerification(sStartTestTime, sEmailTestType, oMails, iTimeout)
	
		End Select
	Else
		Reporter.ReportEvent  micWarning, sEmailConfirmationType & " E-mail test", "No Email recieved after waiting  " & CStr(iTimeout) & " seconds.  Please confirm email"
     End If	

	Call msxDisconnectFromInbox()
	Call msxDeleteTestEmailAccount()
	Call msxCloseEmailDialog()

	' Release objects
	Set oMails = Nothing
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'                                      cmnEmailVerification
'---------------------------------------------------------------------------------------------------------------------
'Desc: 
'			Verifies information on email for
'			 Admin Your Account
'			Admin Gift Registry
'			Online Your Account
'			Online Gift Registry
'			Order Confirmation
'
'Important note:
'			To use this function and its helper function Outlook must be installed.  In addition,
'		there must be a default profile set up.
'
'	Args:
'			sGetTime: used for filtering any emails beginning from the given time
'			sEmailAccount: E-mail account to check (use: qaauto1)
'			oMailItem: Mail message to verfify
'
'Usage:
'			For Admin emails
'				cmnEmailVerification(Now, "Your Account", oMailMessage)
'---------------------------------------------------------------------------------------------------------------------
Public Function cmnEmailVerification(sGetTime, sEmailConfirmationType, oMailItem, iTimeout)

	sTitle = sEmailConfirmationType & " E-mail Confirmation Test" & vbcr & "E-mail: " & sEmailAccount
    
' 	If Not oMailItem.Count = 0 Then
		For Each Mail In oMailItem

			sMailMessage = Mail.Body

			Select Case UCase(sEmailConfirmationType)

				Case "ADMIN YOUR ACCOUNT"
					iGetUserAccount =InStr( sMailMessage, "Temporary Password: ") - (InStr(sMailMessage, "User Name:") + Len("User Name: "))
					iGetTempPassword =( InStr( sMailMessage, "Please note")-4) - (InStr(sMailMessage, "Temporary Password: ") + Len("Temporary Password: ")) 'Note: -4 to ignore linefeed
					sAccountName = mid(sMailMessage ,InStr( sMailMessage, "User Name:")+Len("User Name: ") , iGetUserAccount)
					sTempPassword = mid(sMailMessage ,InStr( sMailMessage, "Temporary Password: ")+Len("Temporary Password: ") , iGetTempPassword)   
					Call dtEmailInfotoDataSheet("AdminUserName", sAccountName)
					Call dtEmailInfotoDataSheet("UserNameTempPassword", sTempPassword)
	
				Case "ADMIN GIFT REGISTRY"
					iGetUserAccount =(InStr( sMailMessage, "Your temporary password is:") - 4) - InStr( sMailMessage, "GR") ' Note: -4 is to ignore the line feed
					iGetTempPassword = (InStr(1462, sMailMessage, "HYPERLINK")-4) - (InStr(sMailMessage, "Your temporary password is:")+Len("Your temporary password is:  ")) 'Note: -2 to ignore line feed
					sAccountName = mid(sMailMessage ,InStr(sMailMessage, "GR") , iGetUserAccount)
					sTempPassword = mid(sMailMessage , InStr(sMailMessage, "Your temporary password is:") + Len("Your temporary password is:  ") , iGetTempPassword)
					Call dtEmailInfotoDataSheet("AdminGRNumber", sAccountName)
					Call dtEmailInfotoDataSheet("GRTempPassword", sTempPassword)
						
				Case "YOUR ACCOUNT"
					iGetUserAccount =(InStr( sMailMessage, "If you forget") - 5) - (InStr(sMailMessage, "Your User Name is:") + Len("Your User Name is: "))
					sAccountName = mid(sMailMessage ,InStr( sMailMessage, "Your User Name is:")+Len("Your User Name is: ") , iGetUserAccount)
					Call dtEmailInfotoDataSheet("UserName", sAccountName)
						
				Case "GIFT REGISTRY"
'					iGetUserAccount =(InStr(sMailMessage, "Current Registry") - 5) - InStr(1400, sMailMessage, "GR") ' Note: -5 is to ignore the line feed
'					sAccountName = mid(sMailMessage ,InStr(1335, sMailMessage, "GR") , iGetUserAccount)

					iStartGRPhrase = InStr(sMailMessage, "Your Gift Registry Number is")
					iEndGRPhrase = InStr(iStartGRPhrase, sMailMessage, vbcr)
					sGRPhrase = Mid(sMailMessage, iStartGRPhrase, (iEndGRPhrase - iStartGRPhrase))

					iGetUserAccount = InStr( sGRPhrase, "GR") 
					sAccountName = Mid(sGRPhrase, InStr( sGRPhrase, "GR"), Len(sGRPhrase))
					Call dtEmailInfotoDataSheet("GRNumber", sAccountName)
					
				Case "E-MAIL A FRIEND"
						'Part of warm and fuzzy.  currently e-mail a friend is disabled in product page.
		
				Case "ORDER CONFIRMATION"
					iGetOrderNumber = (InStr(sMailMessage, "Sku")-4) -  (InStr(sMailMessage, "Order # :") + Len("Order # : "))
					sOrderNumber = mid(sMailMessage,  InStr(sMailMessage, "Order # :") +Len("Order # : "), iGetOrderNumber)
						
			End Select  	
		
			If Not UCase(sEmailConfirmationType) ="ORDER CONFIRMATION" Then
				'Account verification
				If sAccountName = "" Then
					bResult = micFail
				Else
					bResult = micPass
				End If ' Account Name
				'Send out report for Online Account name
				Reporter.ReportEvent bResult, "Account Name:", sAccountName
				'Password verification 
		
				If UCase(sEmailConfirmationType) = "ADMIN YOUR ACCOUNT" or UCase(sEmailConfirmationType) ="ADMIN GIFT REGISTRY" Then
					If  sTempPassword = "" Then
						bResult = micFail
					Else
						bResult = micPass
					End If 'Temp Password
					'Send out report for Temporary password
					Reporter.ReportEvent bResult, "Temporary Password:",sTempPassword
				End If 'Admin YA/GR
			Else 'Non-YA emails
				If sOrderNumber ="" Then
					bResult = micFail
				Else
					bResult = micPass
				End If 'Order Number
				'Send out report for Order number
				Reporter.ReportEvent bResult, "Order Number: " , sOrderNumber
			End If
		Next
'		Else
'			Reporter.ReportEvent  micWarning, sEmailConfirmationType & " E-mail test", "No Email recieved after waiting  " & CStr(iTimeout) & " seconds.  Please confirm email"
'        End If	
		
End Function


'---------------------------------------------------------------------------------------------------------------------
'                                      dtEmailInfotoDataSheet(sEmailInfoType, sEmailInfo)
'---------------------------------------------------------------------------------------------------------------------
'Desc: 
'			Puts Info from the email to the datasheet
'			 Admin Your Account
'			Admin Gift Registry
'			Online Your Account
'			Online Gift Registry
'			Order Confirmation
'
'	Args:
'			sEmailInfoType:  Type of info to be put in the datasheet
'			sEmailInfo:			Info to be put in the datasheet
'
'Usage:
'				dtEmailInfotoDataSheet(sEmailInfoType, sEmailInfo)
'---------------------------------------------------------------------------------------------------------------------
Public Function dtEmailInfotoDataSheet(sEmailInfoType, sEmailInfo)
	'Create a datasheet containing email info
    sSheetName = "Email Sheet Info"
	Environment("EmailSheet") = sSheetName
	DataTable.AddSheet(sSheetName)
	
	'Check if  headers for Email Info exists or not
	For i = 1 To DataTable.GetSheet(sSheetName).GetParameterCount
	   If DataTable.GetSheet(sSheetName).GetParameter(i).Name = sEmailInfoType Then
		   bColExist = True
	   End If
   Next
   
   If Not bColExist Then
	   Call DataTable.GetSheet(sSheetName).AddParameter(sEmailInfoType, sEmailInfo)
	Else
		dtRow = DataTable.GetSheet(sSheetName).GetRowCount 
		DataTable.GetSheet(sSheetName).SetCurrentRow dtRow
		If Not DataTable.GetSheet(sSheetName).GetParameter(sEmailInfoType) = "" Then
			DataTable.GetSheet(sSheetName).SetCurrentRow dtRow+1
		End If
        DataTable.GetSheet(sSheetName).GetParameter(sEmailInfoType).Value = sEmailInfo
   End If
   
End Function


'---------------------------------------------------------------------------------------------------------------------
'                                                                dtReportSheet
'---------------------------------------------------------------------------------------------------------------------
'DESC: 

	' Adds a datasheet if one does not already exist and populates it with values 
	'to make up a Quick Reference pass/fail report .
	
	' Adds 4 columns:

			' Result
			' Description
			' Expected
			' Actual

	' Can be used iteratively and will keep its place using an Environmental Variable:

			' Environment("envReportSheet")
			' Env Var must be initialized at the beginning of the test: 
				
					' EXAMPLE:  Environment("envReportSheet") = "0"

' REQs: 

		' Initialize Environmental variable at beginning of test
		' All args

' ARGs:

		'sSheetName = Name your new sheet
		'bResult = Returned from the verification method
		'Exp_Text = Same var passed to the verification method
		'Act_Text = Same var passed to the verification method
		'sTitle = Same var passed to the verification method

'---------------------------------------------------------------------------------------------------------------------
Public Function dtReportSheet (bResult, Exp_Text, Act_Text, sTitle)
	' ---------------------------------------------------------------------------------
		 'Determin if the required sheet exists, if not creat it
	' ---------------------------------------------------------------------------------	
'	On Error Resume Next
'	sSheetName = "Results Report"
'	dtSheet =  DataTable.GetSheet(sSheetName).name
'	If "" = dtSheet Then
'	
'		'Create Sheet
'		DataTable.AddSheet(sSheetName)
'		'Add columns
'		Call DataTable.GetSheet(sSheetName).AddParameter ("Result", "")
'		Call DataTable.GetSheet(sSheetName).AddParameter ("Description", "")
'		Call DataTable.GetSheet(sSheetName).AddParameter ("Expected", "")
'		Call DataTable.GetSheet(sSheetName).AddParameter ("Actual", "")
'		
'	End If
'	On Error GoTo 0
'
'	' ---------------------------------------------------------------------------------
'		             'Populate rows with Pass / Fail data
'	' --------------------------------------------------------------------------------- 
'	'Increment Env Var
'	Environment("envReportSheet") = Trim(CInt(Environment("envReportSheet")) +1)
'	iRow = CInt(Environment("envReportSheet"))
'	
'	'Set row in report sheet
'	DataTable.GetSheet(sSheetName).SetCurrentRow iRow
'	'Pass data to report sheet
'	DataTable("Result", sSheetName) = Trim(bResult)
'	DataTable("Description", sSheetName) = Trim(sTitle)
'	DataTable("Expected", sSheetName) = Trim(Exp_Text)
'	DataTable("Actual", sSheetName) = Trim(Act_Text)

End Function


'************************************************************************************************ BEGIN E-MAIL HELPER FUNCTIONS*************************************************************************************************
'---------------------------------------------------------------------------------------------------------------------
'					msxGetEmailFromInbox(sGetTime, byRef oMailItem)
'---------------------------------------------------------------------------------------------------------------------	
'Desc: 
'		E-mail helper function
'			Gets all incoming e-mail from the inbox
'
'Important note:
'
'	Args:
'			sGetTime: used for filtering any emails beginning from the given time
'			sEmailConfirmationType: type of e-mail Confirmation test
'						Your Account
'						Gift Registry
'						Email a friend
'			oMailItem: body of the email message to be checked
'
'Usage:
'
'			cmnEmailConfirmation(Now, "Your Account", MailItem)
'---------------------------------------------------------------------------------------------------------------------------
Public Function msxGetEmailFromInbox(sGetTime, byRef oMailItem, byRef iTimeout)

	iStartTimer =Timer
	If  iTimeout < CInt(Environment("Wait_MEDIUM")/1000) Then
		Set oMSOutlook = CreateObject("Outlook.Application")
		Set oOutlookNameSpace=oMSOutlook.GetNameSpace("MAPI")
		Set oFolder = oOutlookNameSpace.GetDefaultFolder(6)
		Wait(CInt(iTimeout))
		oOutlookNameSpace.SyncObjects
		Set oMailMessage = oFolder.Items

		'Filter message
		sWeekDay=WeekDayName(Weekday(Date), True)
		sDate = CStr(Date)
		sTime=FormatDateTime(sGetTime,"4")
		sMeridian = Right (sGetTime, 2)
		sDateFilter = sWeekDay & " "  & sDate & " " & sTime & " " & sMeridian
		Set oUnreadMails = oMailMessage.Restrict("[LastModificationTime] >" & sDateFilter)

		If  oUnreadMails.Count = 0 Then
			iEndTimer = Timer
			iTimeout = iTimeout + (iEndTimer - iStartTimer)		
			Call msxGetEmailFromInbox(sGetTime, oMailItem, iTimeout)
			Exit Function
		End If
		Set oMailItem = oUnreadMails
	End If
		
	' Release objects
	Set oMSOutlook = Nothing
	Set oOutlookNameSpace = Nothing
	Set oFolder = Nothing
	Set oMailMessage = Nothing
	Set oUnreadMails = Nothing

End Function


'---------------------------------------------------------------------------------------------------------------------
'				 msxOpenEmailDialog()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Open email set up dialog.
'
'Args:
'		None
'
'Usage:
'		msxOpenEmailDialog()
'---------------------------------------------------------------------------------------------------------------------
Public Function msxOpenEmailDialog()
	SystemUtil.Run "C:\WINNET\system32\rundll32.exe","C:\WINNET\system32\shell32.dll,Control_RunDLL ""C:\PROGRA~1\COMMON~1\System\MSMAPI\1033\MLCFG32.CPL"",@0","",""
End Function


'---------------------------------------------------------------------------------------------------------------------
'					 msxCloseEmailDialog()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Close email set up dialog.
'
'Args:
'		None
'
'Usage:
'		msxCloseEmailDialog()
'---------------------------------------------------------------------------------------------------------------------
Public Function msxCloseEmailDialog()
		Dialog("text:=Mail").WinButton("text:=OK").Click
End Function


'---------------------------------------------------------------------------------------------------------------------
'				 msxCreateTestEmailAccount(sEmailAccount)
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Creates a new test email account  
'
'Args:
'	  sEmailAccount - Email account to create
'
'Usage:
'		CreateTestEmailAccount(sEmailAccount)
'---------------------------------------------------------------------------------------------------------------------
Public Function msxCreateTestEmailAccount(sEmailAccount)
	'Email server to use
	sEmailServer = "ahqvexdb1.reicorpnet.com"
    sTestEmailProfile = "Test Email Profile"
	sEmail = Left(sEmailAccount, InStr(sEmailAccount, "@")-1)

	'---------------------------------------------------------------
	'ud: Added 10/09/2008
    sEmailLastChar = Right(sEmail, 1)
	Do While IsNumeric(sEmailLastChar)
		sEmail = Left(sEmail, Len(sEmail)-1)
		sEmailLastChar = Right(sEmail, 1)
		'--------------------Begin Debug---------------------
		'Print sEmail
		'Print sEmailLastChar
		'--------------------End Debug---------------------
	Loop
	'---------------------------------------------------------------

	Call cmnSetGlobalTimeouts(1000)
	If Dialog("text:= Mail Setup.*").Exist Then
		cmnSetGlobalTimeouts(10000)
		Dialog("text:=Mail Setup.*").WinButton("text:=&Show Profiles...").Click
		Call msxDeleteTestEmailProfile(sTestEmailProfile)
	End If
	Call cmnSetGlobalTimeouts(10000)
    With Dialog("text:=Mail") 'Start Dialog Mail
		.WinButton("text:=A&dd...").Click
        .Dialog("text:=New Profile").WinEdit("attached text:=Profile &Name:").Set sTestEmailProfile
		.Dialog("text:=New Profile").WinButton("text:=Ok").Click
		With .Dialog("text:=E-mail Accounts") 'Start Dialog Email Accounts
			.WinRadioButton("text:=Add a new &e-mail account").Click
			.WinButton("text:=&Next >").Click
			.WinRadioButton("text:=&Microsoft Exchange Server").Click
			.WinButton("text:=&Next >").Click
			.WinEdit("attached text:=Microsoft &Exchange Server:").Set sEmailServer
			.WinEdit("attached text:=&User Name:").Set sEmail
			.WinButton("text:=&Next >").Click
		End With 'End Dialog Email Accounts
	End With
    Call msxLoginToEmailServer()
	With Dialog("text:=Mail") 'Start Dialog Mail
		.Dialog("text:=E-mail Accounts").WinButton("text:=Finish").Click
		.WinRadioButton("text:=Always &use this profile").Click
		.WinComboBox("nativeclass:=ComboBox").Select (sTestEmailProfile)
		.WinButton("text:=&Apply").Click
	End with 'End Dialog Mail
End Function


'---------------------------------------------------------------------------------------------------------------------
'						 msxDeleteTestEmailAccount()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Deletes the created test account
'
'Args:
'		None
'
'Usage:
'		msxDeleteTestEmailAccount()
'---------------------------------------------------------------------------------------------------------------------
Public Function msxDeleteTestEmailAccount()
   With Dialog("text:=Mail")
		.WinButton("text:=R&emove").Click
		.Dialog("text:=Microsoft Office Outlook").WinButton("text:=&Yes").Click
		'Dialog("text:=Mail").WinButton("text:=&Apply").Click
	End with
End Function


'---------------------------------------------------------------------------------------------------------------------
'						 msxDeleteTestEmailProfile(sEmailProfile)
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'       Checks if there is already a test account on the inbox.  If there is delete that test
'		account.
'
'Args:
'		sEmailProfile: Target e-mail profile
'
'Usage:
'		msxDeleteTestEmailProfile("Default Outlook Profile")
'---------------------------------------------------------------------------------------------------------------------
Public Function msxDeleteTestEmailProfile(sEmailProfile)
	With Dialog("text:=Mail").WinList("attached text:=The following pr&ofiles are set up on this computer:")
		For i=0 To .GetItemsCount-1
			If .GetItem(i) = sEmailProfile Then
				.Select sEmailProfile
			    Call msxDeleteTestEmailAccount()
		   End If
	Next
   End with
End Function


'---------------------------------------------------------------------------------------------------------------------
'								 msxLoginToEmailServer()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Logs in to the server if functional box is not connected to the server
'
'Args:
'		None
'
'Usage:
'		msxLoginToEmailServer()
'---------------------------------------------------------------------------------------------------------------------
Public Function msxLoginToEmailServer()
   Set oNetwork = CreateObject("wscript.network")
   If Not oNetwork.UserDomain = "REICORPNET" Then
	   With Dialog("text:=Connect to.*") 'Start Login Dialog
			.WinEdit("nativeclass:=Edit",  "attached text:=&User name:").Set "reicorpnet\" & QCUtil.QCConnection.UserName
			.WinEdit("nativeclass:=Edit",  "attached text:=&Password:").SetSecure Crypt.Encrypt(QCUtil.QCConnection.Password)
			.WinButton("nativeclass:=Button",  "text:=OK").Click
		End with 'End Login Dialog
	End If

End Function


'---------------------------------------------------------------------------------------------------------------------
'											 msxConnectToInbox()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Connects created account to the inbox
'
'Args:
'		None
'
'Usage:
'		msxConnectToInbox()
'---------------------------------------------------------------------------------------------------------------------
Public Function msxConnectToInbox()
	Dialog("text:=Mail").WinButton("text:=P&roperties").Click
    Dialog("text:=Mail").Dialog("text:=Mail Setup.*").WinButton("text:=&E-mail Accounts...").Click
	Call cmnSetGlobalTimeouts(1000)
	If Dialog("text:=E-mail Accounts").Exist Then
			Call cmnSetGlobalTimeouts(10000)
			Dialog("text:=E-mail Accounts").WinButton("text:=&Next >").Click
			Call msxLoginToEmailServer()
		Else
			Call cmnSetGlobalTimeouts(10000)
			Call msxLoginToEmailServer()
		   Dialog("text:=E-mail Accounts").WinButton("text:=&Next >").Click
		End If
   Call cmnSetGlobalTimeouts(10000)

End Function


'---------------------------------------------------------------------------------------------------------------------
'														msxDisconnectFromInbox()
'---------------------------------------------------------------------------------------------------------------------
'Desc:
'		E-mail helper function
'		Disconnects test account from the inbox
'
'Args:
'		None
'
'Usage:
'		msxDisconnectFromInbox
'---------------------------------------------------------------------------------------------------------------------
Public Function msxDisconnectFromInbox()
	 Dialog("text:=E-mail Accounts").WinButton("text:=Cancel").Click
	 Dialog("text:=Mail").Dialog("text:=Mail Setup.*").WinButton("text:=&Close").Click
End Function

'*************************************************************************************************** END E-MAIL HELPER FUNCTIONS***************************************************************************************************
'
''--------------------------------------------------------
'	' Name: Action Pay With Online Coupone
'	' Remarks: N/A
'	' Purpose:  Verify if the website exist, enter online coupone and click continue  
'	' Arguments: N/a
'	' Return: Integer
'	' Author: Mohamed  Ramadin 
'	' Date: 04/27/2010
'	' References: N/A
''--------------------------------------------------------
'Function PayWithOnlineCoupone(sText) ' pass the datatable as parameter and  make this as function 
'				 ' assign Environment parameter to datatable parameter 
'		'sCouponeCode =  Trim(DataTable("CouponeCode",dtLocalSheet)) 
'		With Browser("Title:= .*REI: .*").Page("Title:= .*REI: .*")      'Browser("Title:= REI: Checkout: Registered").Page("Title:= REI: Checkout: Registered")
'				 If .Exist(2) Then 
'					  ' .WebEdit("name:=coupon").Set  sCouponeCode
'						 Browser("Title:= .*REI: .*").Page("Title:= .*REI: .*").WebEdit("name:=coupon").Set sText
'						 .Sync
'						'.Image("Continue").Click ' call this in the action level 
'				  End If
'		End with
'End Function 
'
'''--------------------------------------------------------
''	' Name:  Private Sub GetSavedText()
''	' Remarks: N/A
''	' Purpose:  
''	' Arguments: N/a
''	' Return: Integer
''	' Author: Mohamed  Ramadin 
''	' Date: 04/27/2010
''	' References: N/A
'''--------------------------------------------------------
'Function CouponSavedMoney(StringText)
'		   If  InStr(StringText,"$")Then
'				  Textstring = Mid(StringText,InStr(1, StringText,"$"))
'				   Textstring2 =Left(Textstring,InStr(1,Textstring," "))		   
'				 
'						 Reporter.ReportEvent micPass, "The free  bar is " & StringText, "  " 
'						 Reporter.ReportEvent micPass, "Amount  you saved is " & Textstring2, " " 			
'			 Else 
'						Reporter.ReportEvent micFail, "The free bar statemant fail" & StringText, " " 
'						Reporter.ReportEvent micFail,  "Amount saved fail" & Textstring2, " " 
'			End If
'   										   
'End Function


'-------------------------------------------------------------------------------------
'				dtComparisonChartSpecTable(oSpecTable)
'-------------------------------------------------------------------------------------
'Desc:	Returns the number of items in the shopping cart
'
'Args:	
'		sSheetName = Name of the datasheet that will 
'			contain all line items from the billing page
'
'Usage:
'		sSheetName = "Line items"
'      Set oComparisonTable=Browser("Browser").Page("Comparisons").WebTable("Comparison Chart")
' 		dtComparisonChartSpecTable(sSheetName, oSpecTable)
'-------------------------------------------------------------------------------------
Public Function dtComparisonChartSpecTable(sSheetName, oSpecTable)

	iRow = oSpecTable.RowCount
	iCol = oSpecTable.ColumnCount(1)
	
   DataTable.AddSheet(sSheetName)
	'Get Line item headers and content
	iQuantityIndex = 0
	For i = 1 To iRow 
		
		If i = 1 Then
			DataTable.SetCurrentRow i 
			For j = 1 To iCol
				'Line Item Header
				sCellData = oSpecTable.GetCellData(i, j)
				Call DataTable.GetSheet(sSheetName).AddParameter(sCellData, "")
			Next
			i = i + 1
		End If

		'Line Item data
		For j = 1 To iCol 
			'If j <> 3Then
				sCellData = oSpecTable.GetCellData(i, j)
			'Else
				'Get the quantity
				'sCellData = cmnShoppingcartNumberOfItems()
			'	sCellData = cmnShoppingcartItemQuantity
				'If sCellData = 0 Then
                 '    sCellData = oLineItemTable.GetCellData(i, j)
				'End If
				'iQuantityIndex = iQuantityIndex + 1
			'End If
			
			If Not sCellData = "" Then
				 DataTable.SetCurrentRow i-1
				DataTable.GetSheet(sSheetName).GetParameter(j).Value = sCellData
			End If
		Next
	Next
End Function


'-------------------------------------------------------------------------------------
'				dtComparisonChartItems(oComparisonChartTable)
'-------------------------------------------------------------------------------------
'Desc:	Returns the number of items in the shopping cart
'
'Args:	
'		sSheetName = Name of the datasheet that will 
'			contain all comparison chart items from the comparison chart page
'
'Usage:
'		sSheetName = "Comparison Chart Items"
'		Set  oComparisonChartTable = Set oTable = Browser("Browser").Page("title:=.*from REI.*\.com").WebTable("class:=compareTbl")
' 		dtComparisonChartItems(sSheetName, oComparisonChartTable)
'-------------------------------------------------------------------------------------
Public Function dtComparisonChartItems(sSheetName, oComparisonChartTable)

	iRow = oComparisonChartTable.RowCount
	iCol = oComparisonChartTable.ColumnCount(1)
	
   DataTable.AddSheet(sSheetName)
	'Get Comparison Chart item headers and content
	For i = 1 To iRow 
		
		DataTable.SetCurrentRow i 

		'Assign Comparison Chart Item Header
		If i=1 Then
			k=1
			For j =2 To iCol
				Call DataTable.GetSheet(sSheetName).AddParameter("Item " & k, "")
				k=k+1
			Next
			i=i+1
		End If

		'Grab Comparison Chart Data
		For j=2 To iCol
			sCellData = oComparisonChartTable.GetCellData(i, j)
			DataTable.GetSheet(sSheetName).GetParameter(j-1).Value = sCellData
		Next
	Next

End Function


' --------------------------------------------------------------------------------------
'                        cmnSecurityAlert
' --------------------------------------------------------------------------------------
' Dismisses the dialog if and only if it displays
' --------------------------------------------------------------------------------------
Public Function cmnSecurityAlert()

	cmnSetGlobalTimeouts 3000

	If Dialog("regexpwndtitle:=Security Information","regexpwndclass:=#32770","visible:=True","index:=0").Exist(5) Then
		Dialog("regexpwndtitle:=Security Information","regexpwndclass:=#32770","visible:=True","index:=0").Activate
		Dialog("regexpwndtitle:=Security Information","regexpwndclass:=#32770","visible:=True","index:=0").Highlight
		Dialog("regexpwndtitle:=Security Information","regexpwndclass:=#32770","visible:=True","index:=0").WinButton("text:=&Yes").Highlight
		Dialog("regexpwndtitle:=Security Information","regexpwndclass:=#32770","visible:=True","index:=0").WinButton("text:=&Yes").Click
	End If

	cmnSetGlobalTimeouts 3000

End Function



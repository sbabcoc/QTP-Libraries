'@Description Extract detail dictionary from billing page item
'@Documentation Extract detail dictionary from <productObj>
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] productObj A reference to a billing page item object
'@ReturnValue The detail dictionary for the specified billing page item
Public Function billing_ExtractItemDetails(ByRef productObj)

   ' Declarations
	Dim detailsDict
	Dim regEx, matchList, thisMatch
	Dim listIndex

	' Allocate RegExp object
	Set regEx = New RegExp
	' Define pattern to extract item details
	regEx.Pattern = ".*?<(DIV|SPAN) class=(h3|detail)>([^<]*)<\/\1>"
	regEx.Global = True

	' Apply pattern to extract list of item detail objects
	Set matchList = regEx.Execute(productObj.GetROProperty("innerhtml"))
	' Get count of detail objects
	matchCount = matchList.Count
	' Set detail assumptions
	Select Case matchCount

		Case 6
			detailKeys = Array(key_ProdName, key_ProdOpt, key_ProdNum, "_", key_ProdCost, key_ProdQty)

		Case 7
			detailKeys = Array(key_ProdPrev, key_ProdName, key_ProdOpt, key_ProdNum, "_", key_ProdCost, key_ProdQty)

		Case Else

	End Select

	' Create a item detail dictionary
	Set detailsDict = CreateObject("Scripting.Dictionary")

	' Iterate over item detail objects
	For listIndex = 0 To (matchCount - 1)
		' Get current item detail
		Set thisMatch = matchList(listIndex)
		' Add item detail to details dictionary
		detailsDict.Add detailKeys(listIndex), thisMatch.SubMatches(2)
	Next

	' RESULT: Item detail dictionary
	Set billing_ExtractItemDetails = detailsDict

	' Release objects
	Set regEx = Nothing
	Set detailsDict = Nothing
	Set matchList = Nothing
	
End Function


'@Description Get detailed list of  billing page contents
'@Documentation Get detailed list of  billing contents
'@Author sbabcoc
'@Date 23-MAR-2011
'@ReturnValue An array of  item detail dictionaries
Public Function billing_GetInventory()

   ' Declarations
   Dim descItemWrap
   Dim itemWrapList
   Dim itemWrapCount, listIndex
   Dim inventory()

	With Browser("Checkout").Page("REI: Checkout: Common")
		' Get description of billing page item wrap objects
		Set descItemWrap = .WebElement("ItemWrap").GetTOProperties
		' Get list of item wrap objects
		Set itemWrapList = .ChildObjects(descItemWrap)
	End With

	' Get count of item wrap objects
	itemWrapCount = itemWrapList.Count
	' If item wrap objects were found
	If (itemWrapCount > 0) Then
		' Allocate space for detail dictionaries
		ReDim inventory(itemWrapCount - 1)
		' Iterate over item wrap objects
		For listIndex = 0 to (itemWrapCount - 1)
			' Store details of current item in the inventory
			Set inventory(listIndex) = billing_ExtractItemDetails(itemWrapList(listIndex))
		Next
	End If

	' RESULT: Detailed inventory
	billing_GetInventory = inventory

	' Release objects
	Set descItemWrap = Nothing
	Set itemWrapList = Nothing

End Function


'@InParameter [in] checkFlags
'		^^^^^^^# : 0 = Guest user checkout; 1 = Registered checkout
'		^^^^^^#^ : 0 = Start as guest user; 1 = Start as registered
'		^^^^##^^ : (reserved)
'		####^^^^ : Target stage:
'							: 0 = Billing Information
'							: 1 = Shipping Information
'							: 2 = Review & Pay
'							: 3 = Order Confirmation
'@InParameter [in] isBoxable
'@InParameter [in] shipMode
'		0 = Normal shipping
'		1 = RSPU only (bicycles)
'		2 = Ship only (plastic gift card)
'		3 = Not shipped (e-gift card)
Public Function billing_VerifyCheckoutFlow(checkFlags, isBoxable, shipMode)

	' Split flag components
	logInState = checkFlags And 3
	targetStage = checkFlags \ 16
	
	' Get log-in state
	loggedIn = IsLoggedIn()
	' If logged in, set logged-in state
	If (loggedIn) Then logInState = logInState Or 4
	
	Select Case logInState
	
		 ' 0 = checkout as guest, not logged in yet
		' 2 = checkout as guest, log in first, not logged in yet
		Case 0, 2
			asGuest = True
			doLogOut = False
	
		' 4 = checkout as guest, already logged in
		' 6 = checkout as guest, log in first, already logged in
		Case 4, 6
			asGuest = True
			doLogOut = True
	
		' 1 = registered checkout, not logged in yet
		' 7 = registered checkout, log in first, already logged in
		Case 1, 7
			asGuest = False
			doLogOut = False
	
		' 3 = registered checkout, log in first, not logged in yet
		Case 3
			asGuest = False
			doLogOut = False
	
		' 5 = registered checkout, already logged in
		Case 5
			asGuest = False
			doLogOut = True
	
	End Select
	
	' Initialize result
	isCorrect = False
	
	Do ' <== Begin bail-out context
	
		If (doLogOut) Then
			' Get reference to 'Log Out' link
			Set refObject = Browser("Common").Page("REI Header").Link("Log Out")
			' Load description of account login page
			Set descTarget = Browser("PageStub").Page("REI.com: Login").GetTOProperties
	
			' Verify that 'Log Out' link lands on Login page
			chk_href = chkVerifyLinkTarget(refObject, descTarget)
			' If navigation fails, bail out
			If IsNull(chk_href) Then Exit Do
	
			' Clear indicator
			loggedIn = False
		End If
	
		' Navigate to the Shopping Cart
		chk_href = cart_NavigateToPage()
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do
	
		' Get reference to 'CHECKOUT NOW' link
		Set refObject = Browser("Checkout").Page("REI.com: Shopping Basket").Link("CHECKOUT NOW")
	
		' If logged in
		If (loggedIn) Then
			' Load description of generic checkout page
			Set descTarget = Browser("Checkout").Page("REI: Checkout: Common").GetTOProperties
		' Otherwise (not logged in)
		Else
			' Load description of account login page
			Set descTarget = Browser("Checkout").Page("REI.com: Login").GetTOProperties
		End If
				
		' Verify that 'CHECKOUT' link lands on expected page
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do
	
		' If not logged in
		If Not (loggedIn) Then
			' Verify login page, proceeding as registered user or guest
			chk_href = billing_VerifyLoginPage(asGuest)
			' If navigation fails, bail out
			If IsNull(chk_href) Then Exit Do
		End If
	
		' Verify the checkout billing page
		isCorrect = billing_VerifyBillingPage(isBoxable, shipMode)
		If Not (isCorrect Or (targetState > 0)) Then Exit Do

		' Get reference to 'Continue' button
'		Set refObject = Browser("Checkout").Page("REI.com: Checkout: Common").WebButton("Continue")
		Set refObject = Browser("Checkout").Page("REI.com: Checkout: Common").WebButton("continue (plain)")

		' No shipping (e-gift card) => REI.com Shipping Checkout (neither address nor method)
		' Same as billing => REI.com Shipping Checkout (method only)
		' Different than billing => REI.com Shipping Checkout (address and method)

		' REI Store Pickup => REI: Store Pickup (still Shipping) => REI Payment Information (RSPU)
	
		' If logged in
		If (loggedIn) Then
			' Load description of generic checkout page
			Set descTarget = Browser("Checkout").Page("REI: Checkout: Common").GetTOProperties
		' Otherwise (not logged in)
		Else
			' Load description of account login page
			Set descTarget = Browser("Checkout").Page("REI.com: Login").GetTOProperties
		End If
				
		' Verify that 'Continue' button lands on expected page
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do
	
	Loop Until True ' <== End bail-out context
	
	billing_VerifyCheckoutFlow = isCorrect

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing
	
End Function


'@Description Verify the content and target of the billing-mode login page
'@Documentation  Verify the content and target of the billing-mode login page
'@Author sbabcoc
'@Date 11-JUL-2011
'@InParameter [in] asGuest 'True' to proceed as guest user; 'False' to proceed as registered
'@ReturnValue If the billing page opens, the 'href' property of the related link; otherwise, 'Null'
Public Function billing_VerifyLoginPage(asGuest)
	
	' Verify header content
	chkVerifyPageHeader()
	' Verify footer content
	chkVerifyPageFooter()
	' Verify breadcrumbs
	chkVerifyBreadcrumbs()
	
	' If guest checkout
	If (asGuest) Then
		' Get reference to 'proceed as guest' link
'		Set refObject = Browser("Checkout").Page("REI.com: Login").Link("proceed as guest")
		Set refObject = Browser("Checkout").Page("REI.com: Login").Link("proceed as guest (plain)")
	' Otherwise (registered user checkout)
	Else
		' Get REI Online account credentials
		acctCreds = account_GetCredentials()
		' Extract account user name
		userName = acctCreds(0)
		' Extract account password
		password = acctCreds(1)
	
		With Browser("Checkout").Page("REI.com: Login")
			' Populate account user name field
			.WebEdit("logonId").Set userName
			' Populate account password field
			.WebEdit("password").Set password
		End With
	
		' Get reference to 'log in' button
		Set refObject = Browser("Checkout").Page("REI.com: Login").WebButton("log in")
	End If
	
	' Load description of generic checkout page
	Set descTarget = Browser("Checkout").Page("REI: Checkout: Common").GetTOProperties
	' Verify that 'log in' button lands on checkout page
	chk_href = chkVerifyLinkTarget(refObject, descTarget)
	
	billing_VerifyLoginPage = chk_href

End Function


'@Description Verify elements of billing page
'@Documentation Verify elements of billing page
'@Author sbabcoc
'@Date 25-JUL-2011
'@InParameter [in] isBoxable 'False' if product is oversize; otherwise 'True'
'@InParameter [in] shipMode Product-related shipping mode
'		0 = Normal shipping
'		1 = RSPU only (bicycles)
'		2 = Ship only (plastic gift card)
'		3 = Not shipped (e-gift card)
Public Function billing_VerifyBillingPage(isBoxable, shipMode)

	' Declarations
	Dim isCorrect

	' Initialize result
	isCorrect = True

	' Verify header content
	chkFlags = chkVerifyPageHeader()
	isCorrect = isCorrect And ((chkFlags And 32768) = 32768)
	
	' Verify footer content
	chkFlags = chkVerifyPageFooter()
	isCorrect = isCorrect And ((chkFlags And 32768) = 32768)

	' Verify breadcrumbs
	chkState = chkVerifyBreadcrumbs()
	isCorrect = isCorrect And chkState
	
	' Verify existence of heading element
	Set refObject = Browser("Checkout").Page("REI: Checkout: Common").WebElement("Heading")
	chkState = chkVerifyExistence(refObject, "Billing Page Heading", EXPECT_EXISTS, "verifyBillingPage 01")
	isCorrect = isCorrect And chkState
	
	' Verify existence of "Edit Order" link
	Set refObject = Browser("Checkout").Page("REI: Checkout: Common").Link("Edit Order")
	chkState = chkVerifyExistence(refObject, "Edit Order Link", EXPECT_EXISTS, "verifyBillingPage 02")
	isCorrect = isCorrect And chkState
	
	' Verify billing address section
	chkState = billing_VerifyBillingAddress()
	isCorrect = isCorrect And chkState
	
	' Verify shipping mode section
	chkState = billing_VerifyShippingMode(shipMode, 1)
	isCorrect = isCorrect And chkState
	
	' Verify membership section
	chkState = billing_VerifyREIMembership()
	isCorrect = isCorrect And chkState
	
	' Verify online coupon section
	chkState = billing_VerifyOnlineCoupons()
	isCorrect = isCorrect And chkState
	
	' Verify TAM tenders section
	chkState = billing_VerifyTAMTenders()
	isCorrect = isCorrect And chkState
	
	' Verify gift options section
	chkState = billing_VerifyGiftOptions(isBoxable, shipMode)
	isCorrect = isCorrect And chkState
	
	' Verify existence of "Continue" button
'	Set refObject = Browser("Checkout").Page("REI: Checkout: Common").WebButton("Continue")
	Set refObject = Browser("Checkout").Page("REI: Checkout: Common").WebButton("continue (plain)")
	chkState = chkVerifyExistence(refObject, "Continue button", EXPECT_EXISTS, "verifyBillingPage 09")
	isCorrect = isCorrect And chkState

	billing_VerifyBillingPage = isCorrect
	
	' Release objects
	Set refObject = Nothing
	
End Function


Public Function billing_VerifyBillingAddress()

	' Initialize result
	isCorrect = True

	' Get log-in state
	loggedIn = IsLoggedIn()
	
	' Get account first name
	firstName = MyGetParameter("createUserAccount [Online]", "FirstName", 1)
	' Get account lastr name
	lastName = MyGetParameter("createUserAccount [Online]", "LastName", 1)
	' Get account ZIP code
	zipCode = MyGetParameter("createUserAccount [Online]", "ZipCode", 1)
	' Get daytime phone
	AM_Phone = MyGetParameter("createUserAccount [Online]", "AM_Phone", 1)
	' Get night phone
	PM_Phone = MyGetParameter("createUserAccount [Online]", "PM_Phone", 1)
	' Get address line 1
	address1 = MyGetParameter("createUserAccount [Online]", "Address1", 1)
	' Get address line 2
	address2 = MyGetParameter("createUserAccount [Online]", "Address2", 1)
	' Get address line 3
	address3 = MyGetParameter("createUserAccount [Online]", "Address3", 1)
	' Get city name
	cityName = MyGetParameter("createUserAccount [Online]", "CityName", 1)
	' Get state name
	stateName = MyGetParameter("createUserAccount [Online]", "StateName", 1)
	' Get country name
	country = MyGetParameter("createUserAccount [Online]", "Country", 1)
	
	' Get REI Online account credentials
	acctCreds = account_GetCredentials()
	' Extract account user name
	emailAddr = acctCreds(0)
	
	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' ##### BILLING ADDRESS #####
	
		' Verify existence of "BillTo" sub-head
		Set refObject = .WebElement("Sub1: BillTo")
		chkState = chkVerifyExistence(refObject, "Billing Address Sub-head", EXPECT_EXISTS, "verifyBillingPage 03")
		isCorrect = isCorrect And chkState
	
		' Verify existence of First Name label
		Set refObject = .WebElement("First Name:*")
		chkState = chkVerifyExistence(refObject, "First Name Label", EXPECT_EXISTS, "verifyBillingPage 03.01")
		isCorrect = isCorrect And chkState
	
		' Verify existence of First Name field
		Set refObject = .WebEdit("bill_safname")
		chkState = chkVerifyExistence(refObject, "First Name Field", EXPECT_EXISTS, "verifyBillingPage 03.02")
		isCorrect = isCorrect And chkState

		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), firstName, STR_EQUAL, "verifyBillingPage 03.02.01", "First Name Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set firstName
			End If
		End If
	
		' Verify existence of Middle Initial label
		Set refObject = .WebElement("Middle Initial:")
		chkState = chkVerifyExistence(refObject, "Middle Initial Label", EXPECT_EXISTS, "verifyBillingPage 03.03")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Middle Initial field
		Set refObject = .WebEdit("bill_samname")
		chkState = chkVerifyExistence(refObject, "Middle Initial Field", EXPECT_EXISTS, "verifyBillingPage 03.04")
		isCorrect = isCorrect And chkState
	
		' NOTE: Skipping verify/populate of Middle Initial
	
		' Verify existence of Last Name label
		Set refObject = .WebElement("Last Name:*")
		chkState = chkVerifyExistence(refObject, "Last Name Label", EXPECT_EXISTS, "verifyBillingPage 03.05")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Last Name field
		Set refObject = .WebEdit("bill_salname")
		chkState = chkVerifyExistence(refObject, "Last Name Field", EXPECT_EXISTS, "verifyBillingPage 03.06")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), lastName, STR_EQUAL, "verifyBillingPage 03.06.01", "Last Name Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set lastName
			End If
		End If
	
		' Verify existence of Address 1 label
		Set refObject = .WebElement("Address 1:*")
		chkState = chkVerifyExistence(refObject, "Address 1 Label", EXPECT_EXISTS, "verifyBillingPage 03.07")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Address 1 field
		Set refObject = .WebEdit("bill_saaddr1")
		chkState = chkVerifyExistence(refObject, "Address 1 Field", EXPECT_EXISTS, "verifyBillingPage 03.08")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), address1, STR_EQUAL, "verifyBillingPage 03.08.01", "Address 1 Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set address1
			End If
		End If
	
		' Verify existence of Address 2 label
		Set refObject = .WebElement("Address 2:")
		chkState = chkVerifyExistence(refObject, "Address 2 Label", EXPECT_EXISTS, "verifyBillingPage 03.09")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Address 2 field
		Set refObject = .WebEdit("bill_saaddr2")
		chkState = chkVerifyExistence(refObject, "Address 2 Field", EXPECT_EXISTS, "verifyBillingPage 03.10")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), address2, STR_EQUAL, "verifyBillingPage 03.10.01", "Address 2 Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set address2
			End If
		End If
	
		' Verify existence of Address 3 label
		Set refObject = .WebElement("Company or C/O:")
		chkState = chkVerifyExistence(refObject, "Address 3 Label", EXPECT_EXISTS, "verifyBillingPage 03.11")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Address 3 field
		Set refObject = .WebEdit("bill_saaddr3")
		chkState = chkVerifyExistence(refObject, "Address 3 Field", EXPECT_EXISTS, "verifyBillingPage 03.12")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), address3, STR_EQUAL, "verifyBillingPage 03.12.01", "Address 3 Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set address3
			End If
		End If
	
		' Verify existence of City label
		Set refObject = .WebElement("City:*")
		chkState = chkVerifyExistence(refObject, "City Label", EXPECT_EXISTS, "verifyBillingPage 03.13")
		isCorrect = isCorrect And chkState
	
		' Verify existence of City field
		Set refObject = .WebEdit("bill_sacity")
		chkState = chkVerifyExistence(refObject, "City Field", EXPECT_EXISTS, "verifyBillingPage 03.14")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), cityName, STR_EQUAL, "verifyBillingPage 03.14.01", "Address 1 Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set cityName
			End If
		End If
	
		' Verify existence of State label
		Set refObject = .WebElement("State:*")
		chkState = chkVerifyExistence(refObject, "State Label", EXPECT_EXISTS, "verifyBillingPage 03.15")
		isCorrect = isCorrect And chkState
	
		' Verify existence of State list
		Set refObject = .WebList("bill_sastate")
		chkState = chkVerifyExistence(refObject, "State List", EXPECT_EXISTS, "verifyBillingPage 03.16")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), stateName, STR_EQUAL, "verifyBillingPage 03.16.01", "State List")
				isCorrect = isCorrect And chkState
			Else
				refObject.Select stateName
			End If
		End If
	
		' Verify existence of Country label
		Set refObject = .WebElement("Country:*")
		chkState = chkVerifyExistence(refObject, "Country Label", EXPECT_EXISTS, "verifyBillingPage 03.17")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Country list
		Set refObject = .WebList("bill_sacntry")
		chkState = chkVerifyExistence(refObject, "Country List", EXPECT_EXISTS, "verifyBillingPage 03.18")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), country, STR_EQUAL, "verifyBillingPage 03.18.01", "Country List")
				isCorrect = isCorrect And chkState
			Else
				refObject.Select country
			End If
		End If
	
		' Verify existence of ZIP Code label
		Set refObject = .WebElement("Zip (Postal) Code:*")
		chkState = chkVerifyExistence(refObject, "ZIP Code Label", EXPECT_EXISTS, "verifyBillingPage 03.19")
		isCorrect = isCorrect And chkState
	
		' Verify existence of ZIP Code field
		Set refObject = .WebEdit("bill_sazipc")
		chkState = chkVerifyExistence(refObject, "ZIP Code Field", EXPECT_EXISTS, "verifyBillingPage 03.20")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), zipCode, STR_EQUAL, "verifyBillingPage 03.20.01", "ZIP Code Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set zipCode
			End If
		End If
	
		' Verify existence of Day Phone label
		Set refObject = .WebElement("Day Phone:*")
		chkState = chkVerifyExistence(refObject, "Day Phone Label", EXPECT_EXISTS, "verifyBillingPage 03.21")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Daytime Phone field
		Set refObject = .WebEdit("bill_saphone1")
		chkState = chkVerifyExistence(refObject, "Day Phone Field", EXPECT_EXISTS, "verifyBillingPage 03.22")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), AM_Phone, STR_EQUAL, "verifyBillingPage 03.22.01", "Day Phone Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set AM_Phone
			End If
		End If
	
		' Verify existence of Night Phone label
		Set refObject = .WebElement("Night Phone:")
		chkState = chkVerifyExistence(refObject, "Night Phone Label", EXPECT_EXISTS, "verifyBillingPage 03.23")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Night Phone field
		Set refObject = .WebEdit("bill_saphone2")
		chkState = chkVerifyExistence(refObject, "Night Phone Field", EXPECT_EXISTS, "verifyBillingPage 03.24")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), PM_Phone, STR_EQUAL, "verifyBillingPage 03.24.01", "Night Phone Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set PM_Phone
			End If
		End If
	
		' Verify existence of E-mail Address label
		Set refObject = .WebElement("E-mail Address:*")
		chkState = chkVerifyExistence(refObject, "E-mail Address Label", EXPECT_EXISTS, "verifyBillingPage 03.25")
		isCorrect = isCorrect And chkState
	
		' Verify existence of E-mail Address field
		Set refObject = .WebEdit("bill_saemail1")
		chkState = chkVerifyExistence(refObject, "E-mail Address Field", EXPECT_EXISTS, "verifyBillingPage 03.26")
		isCorrect = isCorrect And chkState
	
		If (chkState) Then
			If (loggedIn) Then
				chkState = chkVerifyParity(refObject.GetROProperty("value"), emailAddr, STR_EQUAL, "verifyBillingPage 03.26.01", "E-mail Address Field")
				isCorrect = isCorrect And chkState
			Else
				refObject.Set emailAddr
			End If
		End If
	
	End With
	
	With Browser("Checkout").Page("REI: Checkout: Non-Registered")
	
		' Verify existence of Gearmail checkbox
		Set refObject = .WebCheckBox("gearmail")
		chkState = chkVerifyExistence(refObject, "Gearmail Checkbox Field", Not loggedIn, "verifyBillingPage 03.27")
		isCorrect = isCorrect And chkState
	
		If (chkState And Not loggedIn) Then
			refObject.Set "Off"
		End If
	
		' Verify existence of Privacy Policy button
		Set refObject = .WebButton("Privacy Policy")
		chkState = chkVerifyExistence(refObject, "Privacy Policy Button", Not loggedIn, "verifyBillingPage 03.28")
		isCorrect = isCorrect And chkState
	
	End With

	billing_VerifyBillingAddress = isCorrect
	
	' Release objects
	Set refObject = Nothing

End Function

'@InParameter [in] shipMode
'		0 = Normal shipping
'		1 = RSPU only (bicycles)
'		2 = Ship only (plastic gift card)
'		3 = Not shipped (e-gift card)
'@InParameter [in] shipPick
'		1 = Ship to billing address
'		2 = Ship to another address
'		3 = Ship to local REI Store
Public Function billing_VerifyShippingMode(shipMode, shipPick)

	' Initialize result
	isCorrect = True

	' Get log-in state
	loggedIn = IsLoggedIn()
	
	' ##### SHIPPING MODE #####
	
	Select Case shipMode
	
		' Normal shipping
		Case 0
			doExpect = True
			rspuOnly = False
			hasShipSameAs = True
			hasShipAnother = True
			hasShipToREI = True
	
		' RSPU only (bicycles)
		Case 1
			doExpect = True
			rspuOnly = True
			hasShipSameAs = False
			hasShipAnother = False
			hasShipToREI = True
			
			shipPick = 3
			
		' Ship only (plastic gift card)
		Case 2
			doExpect = True
			rspuOnly = False
			hasShipSameAs = True
			hasShipAnother = True
			hasShipToREI = False

			' If RSPU picked, use billing addr
			If (shipPick = 3) Then shipPick = 1
	
		' Not shipped (e-gift card)
		Case 3
			doExpect = False
			rspuOnly = False
			hasShipSameAs = False
			hasShipAnother = False
			hasShipToREI = False
	
	End Select
	
	' Verify existence of "ShipTo" sub-head
	Set refObject = Browser("Checkout").Page("REI: Checkout: Common").WebElement("Sub2: ShipTo")
	chkState = chkVerifyExistence(refObject, "Shipping Mode Sub-head", doExpect, "verifyBillingPage 04")
	isCorrect = isCorrect And chkState
	
	' Verify existence of "Add Shipping Address" link
	Set refObject = Browser("Checkout").Page("REI: Checkout: Registered").Link("Add new shipping address")
	chkState = chkVerifyExistence(refObject, "Add Shipping Address link", loggedIn And Not rspuOnly, "verifyBillingPage 04.01")
	isCorrect = isCorrect And chkState

	' NOTE: Registered checkout adds the possibility of one or more shipping addresses
	' Each shipping address adds a radio button to the shipping mode radio group
	' Under each radio button line is a DIV with the address and an "Edit Address" link
	
	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' Verify existence of RSPU notice
		Set refObject = .WebElement("Note: RSPU Only")
		chkState = chkVerifyExistence(refObject, "RSPU notice", rspuOnly, "verifyBillingPage 04.02")
		isCorrect = isCorrect And chkState

		' Initialize ship flag
		hasShipPick = False
	
		' Verify existence of shipping mode radio group
		Set refRadGrp = .WebRadioGroup("shipping_method")
		chkState = chkVerifyExistence(refRadGrp, "Shipping Mode Radio Group", doExpect, "verifyBillingPage 04.03")
		isCorrect = isCorrect And chkState
	
		' If expected radio group exists
		If (doExpect And chkState) Then
			' Allocate RegExp object
			Set regEx = New RegExp
	
			' Verify existence of option 1
			Set refObject = .WebElement("method1")
			chkState = chkVerifyExistence(refObject, "Method 1: Shipping Same As Billing", hasShipSameAs, "verifyBillingPage 04.03.01")
			isCorrect = isCorrect And chkState

			' If expected option 1 exists
			If (hasShipSameAs And chkState) Then
				' Extract option 1 source
				optionText = refObject.GetROProperty("innerhtml")
	
				' Verify option 1 ID
				regEx.Pattern = "\bid=shipBilling\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.02", "Method 1 ID")
				isCorrect = isCorrect And chkState

				' Verify option 1 value
				regEx.Pattern = "\bvalue=1\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.03", "Method 1 Value")
				isCorrect = isCorrect And chkState

				' If option 1 specified, set ship flag
				If (shipPick = 1) Then hasShipPick = True
			End If
			
			' Verify existence of option 2
			Set refObject = .WebElement("method2")
			chkState = chkVerifyExistence(refObject, "Method 2: Different Shipping Address", hasShipAnother, "verifyBillingPage 04.03.04")
			isCorrect = isCorrect And chkState
	
			' If expected option 2 exists
			If (hasShipAnother And chkState) Then
				' Extract option 2 source
				optionText = refObject.GetROProperty("innerhtml")
	
				' Verify option 2 ID
				regEx.Pattern = "\bid=shipDifferent\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.05", "Method 2 ID")
				isCorrect = isCorrect And chkState
				
				' Verify option 2 value
				regEx.Pattern = "\bvalue=2\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.06", "Method 2 Value")
				isCorrect = isCorrect And chkState
	
				' If option 2 specified, set ship flag
				If (shipPick = 2) Then hasShipPick = True
			End If
			
			' Verify existence of option 3
			Set refObject = .WebElement("method3")
			chkState = chkVerifyExistence(refObject, "Method 3: REI Store Pickup", hasShipToREI, "verifyBillingPage 04.03.07")
			isCorrect = isCorrect And chkState
	
			' If expected option 3 exists
			If (hasShipToREI And chkState) Then
				' Extract option 3 source
				optionText = refObject.GetROProperty("innerhtml")
	
				' Verify option 3 ID
				regEx.Pattern = "\bid=rspu\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.08", "Method 3 ID")
				isCorrect = isCorrect And chkState
			
				' Verify option 3 value
				regEx.Pattern = "\bvalue=3\b"
				chkState = chkVerifyParity(regEx.Test(optionText), "correct", CMP_ASSERT, "verifyBillingPage 04.03.09", "Method 3 Value")
				isCorrect = isCorrect And chkState
	
				' If option 3 specified, set ship flag
				If (shipPick = 3) Then hasShipPick = True
			End If

			' Verify that specified shipping method is one of the available options
			chkState = chkVerifyParity(hasShipPick, "available", CMP_ASSERT, "verifyBillingPage 04.03.10", "Specified Shipping Method")
			isCorrect = isCorrect And chkState
			
			' If has spec'd option
			If (chkState) Then
				' Get value of every shipping option
				everyItem = refRadGrp.GetROProperty("all items")
				' Split option string into array
				itemList = Split(everyItem, ";")
				' Iterate over shipping option values
				For index = 0 to UBound(itemList)
					' If this value matches specified option
					If (itemList(index) = CStr(shipPick)) Then
						' Select this shipping option
						refRadGrp.Select "#" & index
						Exit For
					End If
				Next
			End If
			
			' Release objects
			Set regEx = Nothing
		End If
	
		' Verify existence of RSPU state list
		Set refObject = .WebList("pickupStoreState")
		chkState = chkVerifyExistence(refObject, "REI Store Pickup State list", hasShipToREI, "verifyBillingPage 04.04")
		isCorrect = isCorrect And chkState

		' If has RSPU option and state list exists
		If (hasShipToREI And chkState) Then
			' If RSPU option spec'd
			If (shipPick = 3) Then
				' Get state name
				stateName = MyGetParameter("createUserAccount [Online]", "StateName", 1)
				' Select state name
				refObject.Select stateName
			End If
		End If

		' Verify existence of RSPU info button
		Set refObject = getButtonLink("storepickupinfo")
		chkState = chkVerifyParity((refObject Is Nothing) <> hasShipToREI, "present", CMP_ASSERT, "verifyBillingPage 04.05", "REI Store Pickup Info button")
		isCorrect = isCorrect And chkState
	
	End With

	billing_VerifyShippingMode = isCorrect
	
	' Release objects
	Set refObject = Nothing
	Set refRadGrp = Nothing

End Function


Public Function billing_VerifyREIMembership()

	' Initialize result
	isCorrect = True

	' Get log-in state
	loggedIn = IsLoggedIn()
	
	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' ##### REI MEMBERSHIP #####
	
		' Verify existence of "Member" sub-head
		Set refObject = .WebElement("Sub3: Member")
		chkState = chkVerifyExistence(refObject, "REI Membership Sub-head", EXPECT_EXISTS, "verifyBillingPage 05")
		isCorrect = isCorrect And chkState
	
		' Verify existence of "Use Dividend" checkbox
		Set refObject = .WebCheckBox("use_dividend")
		chkState = chkVerifyExistence(refObject, "Use Dividend checkbox", EXPECT_EXISTS, "verifyBillingPage 05.01")
		isCorrect = isCorrect And chkState
	
	End With
	
	' Verify existence of "Edit Member Number" link
	Set refObject = Browser("Checkout").Page("REI: Checkout: Registered").Link("Edit membership number")
	chkState = chkVerifyExistence(refObject, "Edit Member Number link", loggedIn, "verifyBillingPage 05.02")
	isCorrect = isCorrect And chkState
	
	With Browser("Checkout").Page("REI: Checkout: Non-Registered")
	
		' Verify existence of Membership Number label
		Set refObject = .WebElement("Member Number")
		chkState = chkVerifyExistence(refObject,"Membership Number label", Not loggedIn, "verifyBillingPage 05.03")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Membership Number field
		Set refObject = .WebEdit("member_number")
		chkState = chkVerifyExistence(refObject,"Membership Number field", Not loggedIn, "verifyBillingPage 05.04")
		isCorrect = isCorrect And chkState
		
		' Verify existence of "Buy Membership" checkbox
		Set refObject = .WebCheckBox("buy_membership")
		chkState = chkVerifyExistence(refObject,"Buy Membership checkbox", Not loggedIn, "verifyBillingPage 05.05")
		isCorrect = isCorrect And chkState
	
		' Verify existence of "Member Info" button
		Set refObject = getButtonLink("ab_pop_onlineaccount")
		chkState = chkVerifyParity((refObject Is Nothing) = loggedIn, "present", CMP_ASSERT, "verifyBillingPage 05.06", "Member Info button")
		isCorrect = isCorrect And chkState
		
	End With

	billing_VerifyREIMembership = isCorrect
	
	' Release objects
	Set refObject = Nothing

End Function


Public Function billing_VerifyOnlineCoupons()

	' Initialize result
	isCorrect = True

	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' ##### ONLINE COUPONS #####
	
		' Verify existence of "Coupon" sub-head
		Set refObject = .WebElement("Sub4: Coupon")
		chkState = chkVerifyExistence(refObject, "Online Coupons Sub-head", EXPECT_EXISTS, "verifyBillingPage 06")
		isCorrect = isCorrect And chkState
	
		' Verify existence of coupon code label
		Set refObject = .WebElement("Enter Coupon Code:")
		chkState = chkVerifyExistence(refObject, "Coupon Code label", EXPECT_EXISTS, "verifyBillingPage 06.01")
		isCorrect = isCorrect And chkState
	
		' Verify existence of coupon code field
		Set refObject = .WebEdit("coupon")
		chkState = chkVerifyExistence(refObject, "Coupon Code field", EXPECT_EXISTS, "verifyBillingPage 06.02")
		isCorrect = isCorrect And chkState
		
	End With

	billing_VerifyOnlineCoupons = isCorrect
	
	' Release objects
	Set refObject = Nothing

End Function


Public Function billing_VerifyTAMTenders()

	' Initialize result
	isCorrect = True

	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' ##### TAM TENDERS #####
	
		' Verify existence of "TAM" sub-head
		Set refObject = .WebElement("Sub5: TAM")
		chkState = chkVerifyExistence(refObject, "TAM Tenders Sub-head", EXPECT_EXISTS, "verifyBillingPage 07")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Number label
		Set refObject = .WebElement("Number:")
		chkState = chkVerifyExistence(refObject, "Number label", EXPECT_EXISTS, "verifyBillingPage 07.01")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Number field
		Set refObject = .WebEdit("gc1")
		chkState = chkVerifyExistence(refObject, "Number field", EXPECT_EXISTS, "verifyBillingPage 07.02")
		isCorrect = isCorrect And chkState
		
		' Verify existence of PIN label
		Set refObject = .WebElement("PIN:")
		chkState = chkVerifyExistence(refObject, "PIN label", EXPECT_EXISTS, "verifyBillingPage 07.03")
		isCorrect = isCorrect And chkState
		
		' Verify existence of PIN field
		Set refObject = .WebEdit("gcnum1")
		chkState = chkVerifyExistence(refObject, "PIN field", EXPECT_EXISTS, "verifyBillingPage 07.04")
		isCorrect = isCorrect And chkState
		
		' Verify existence of TAM Info button
		Set refObject = .WebButton("TAM Info")
		chkState = chkVerifyExistence(refObject, "TAM Info button", EXPECT_EXISTS, "verifyBillingPage 07.05")
		isCorrect = isCorrect And chkState
		
	End With

	billing_VerifyTAMTenders = isCorrect
	
	' Release objects
	Set refObject = Nothing

End Function


'@InParameter [in] isBoxable
'@InParameter [in] shipMode
'		0 = Normal shipping
'		1 = RSPU only (bicycles)
'		2 = Ship only (plastic gift card)
'		3 = Not shipped (e-gift card)
Public Function billing_VerifyGiftOptions(isBoxable, shipMode)

	' Initialize result
	isCorrect = True

	With Browser("Checkout").Page("REI: Checkout: Common")
	
		' ##### GIFT OPTIONS #####
	
		doExpect = (shipMode < 2)
	
		' Verify existence of "Gift Options" sub-head
		Set refObject = .WebElement("Sub6: Gift")
		chkState = chkVerifyExistence(refObject, "Gift Options Sub-head", doExpect, "verifyBillingPage 08")
		isCorrect = isCorrect And chkState
	
		' Verify existence of Gift checkbox
		Set refObject = .WebCheckBox("gift")
		chkState = chkVerifyExistence(refObject, "Gift checkbox", doExpect, "verifyBillingPage 08.01")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Learn More button
		Set refObject = .WebButton("Learn more")
		chkState = chkVerifyExistence(refObject, "Learn More button", doExpect And isBoxable, "verifyBillingPage 08.02")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Not Boxable notice
		Set refObject = .WebElement("Note: Not Boxable")
		chkState = chkVerifyExistence(refObject, "Not Boxable notice", doExpect And Not isBoxable, "verifyBillingPage 08.03")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Gift Message label
		Set refObject = .WebElement("Gift Message")
		chkState = chkVerifyExistence(refObject, "Gift Message label", doExpect And Not isBoxable, "verifyBillingPage 08.04")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Gift Message field 1
		Set refObject = .WebEdit("gift_message")
		chkState = chkVerifyExistence(refObject, "Gift Message field 1", doExpect And Not isBoxable, "verifyBillingPage 08.04")
		isCorrect = isCorrect And chkState
		
		' Verify existence of Gift Message field 2
		Set refObject = .WebEdit("gift_message2")
		chkState = chkVerifyExistence(refObject, "Gift Message field 2", doExpect And Not isBoxable, "verifyBillingPage 08.05")
		isCorrect = isCorrect And chkState
		
	End With

	billing_VerifyGiftOptions = isCorrect
	
	' Release objects
	Set refObject = Nothing

End Function

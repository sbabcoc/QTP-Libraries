' Store detail dictionary keys
Const key_StorLink = "LINK"
Const key_StorName = "NAME"
Const key_StorNum = "NUM"
Const key_StorAddr = "ADDR"
Const key_StorCity = "CITY"
Const key_StorStat = "STAT"
Const key_StorZip = "ZIP"
Const key_StorPhon = "PHON"
Const key_StorMap = "MAP"
Const key_StorDir = "DIR"
Const key_StorBike = "BIKE"
Const key_StorSkis = "SKIS"
Const key_StorRent = "RENT"

'@Description Extract services from store services object
'@Documentation Extract  services from <servicesObj>
'@Author sbabcoc
'@Date 22-JUL-2011
'@InParameter [in] servicesObj A reference to a store services object
'@ReturnValue The service items for the specified services object
Public Function store_ExtractServices(ByRef servicesObj)

	' Declarations
	Dim childList, thisChild, listIndex
	Dim hasBikeShop, hasSkiShop, hasRentals

	' Initialize results
	hasBikeShop = False
	hasSkiShop = False
	hasRentals = False

	' Get child objects for services group
	Set childList = servicesObj.ChildObjects()
	' Iterate over child object list
	For listIndex = 0 To (childList.Count - 1)
		' Extract current child object
		Set thisChild = childList(listIndex)

		' Determine service from image "alt" text
		Select Case thisChild.GetROProperty("alt")

			Case "Bike Shop"
				hasBikeShop = True

			Case "Ski Shop"
				hasSkiShop = True

			Case "Store Rentals"
				hasRentals = True

		End Select
	Next
	
	' RESULT: Store services
	store_ExtractServices = Array(hasBikeShop, hasSkiShop, hasRentals)

	' Release objects
	Set childList = Nothing
	Set thisChild = Nothing
	
End Function


'@Description Extract detail dictionary from store result item
'@Documentation Extract detail dictionary from <storeObj>
'@Author sbabcoc
'@Date 22-MAR-2011
'@Libraries Global
'@Repositories Locator
'@InParameter [in] storeObj A reference to a store result item object
'@ReturnValue The detail dictionary for the specified store result item
Public Function store_ExtractItemDetails(ByRef storeObj)

   ' Declarations
	Dim detailsDict
	Dim descDivision
	Dim descStoreLink, descStoreAddr, descStoreCity
	Dim descStorePhone, descStoreMap, descStoreDir
	Dim descStoreSvcs
	Dim divisionList
	Dim thisDivision
	Dim listIndex
	Dim linkObj

	' Create a store detail dictionary
	Set detailsDict = CreateObject("Scripting.Dictionary")

	' Define generic division element description
	Set descDivision = Description.Create()
	descDivision("micclass").Value = "WebElement"
	descDivision("html tag").Value = "DIV"
	
	With Browser("Locator").Page("REI Stores").WebElement("StoreList").WebElement("Store")
		' Load store detail object descriptions
		Set descStoreName = .WebElement("StoreName").GetTOProperties
		Set descStoreAddr = .WebElement("StoreAddr").GetTOProperties
		Set descStoreCity = .WebElement("StoreCity").GetTOProperties
		Set descStorePhone = .WebElement("StorePhone").GetTOProperties
		Set descStoreMap = .WebElement("StoreMap").GetTOProperties
		Set descStoreDir = .WebElement("StoreDir").GetTOProperties
		Set descStoreSvcs = .WebElement("StoreSvcs").GetTOProperties
	End With

	' Get list of store detail objects
	Set divisionList = storeObj.ChildObjects(descDivision)
	' Iterate over store detail objects
	For listIndex = 0 To (divisionList.Count - 1)
		' Get current store detail object
		Set thisDivision = divisionList(listIndex)
		' If this is the store name
		If (isDescribedObject(thisDivision, descStoreName)) Then
			' Extract link object
			Set linkObj = thisDivision.ChildObjects()(0)
			' Get store page link
			pageLink = linkObj.GetROProperty("href")
			' Add store page link to detail dictionary
			detailsDict.Add key_StorLink, pageLink
			' Add store name to detail dictionary
			detailsDict.Add key_StorName, linkObj.GetROProperty("innertext")
			' Add store number to detail dictionary
			detailsDict.Add key_StorNum, store_GetNumberFromURL(pageLink)
		' Otherwise, if this is the store address
		ElseIf (isDescribedObject(thisDivision, descStoreAddr)) Then
			' Add cost to detail dictionary
			detailsDict.Add key_StorAddr, thisDivision.GetROProperty("innertext")
		' Otherwise, if this is the store city/state/zip
		ElseIf (isDescribedObject(thisDivision, descStoreCity)) Then
			' Get store city/state/zip
			cityBits = Split(thisDivision.GetROProperty("innertext"))
			' Add city to detail dictionary (minus trailing comma)
			detailsDict.Add key_StorCity, Left(cityBits(0), Len(cityBits(0)) - 1)
			' Add state to detail dictionary
			detailsDict.Add key_StorStat, cityBits(1)
			' Add ZIP to detail dictionary
			detailsDict.Add key_StorZip, cityBits(2)
		' Otherwise, if this is the store phone
		ElseIf (isDescribedObject(thisDivision, descStorePhone)) Then
			' Add store phone to detail dictionary
			detailsDict.Add key_StorPhon, thisDivision.GetROProperty("innertext")
		' Otherwise, if this is the store map link
		ElseIf (isDescribedObject(thisDivision, descStoreMap)) Then
			' Extract link object
			Set linkObj = thisDivision.ChildObjects()(0)
			' Get store map link
			mapLink = linkObj.GetROProperty("href")
			' Add store map link to detail dictionary
			detailsDict.Add key_StorMap, mapLink
		' Otherwise, if this is the directions link
		ElseIf (isDescribedObject(thisDivision, descStoreDir)) Then
			' Extract link object
			Set linkObj = thisDivision.ChildObjects()(0)
			' Get directions link
			directLink = linkObj.GetROProperty("href")
			' Add directions link to detail dictionary
			detailsDict.Add key_StorDir, directLink
		' Otherwise, if this is the store services group
		ElseIf (isDescribedObject(thisDivision, descStoreSvcs)) Then
			' Extract indicated services from group
			serviceList = store_ExtractServices(thisDivision)
			' Add bike shop to details dictionary
			detailsDict.Add key_StorBike, serviceList(0)
			' Add ski shop to details dictionary
			detailsDict.Add key_StorSkis, serviceList(1)
			' Add store rentals to details dictionary
			detailsDict.Add key_StorRent, serviceList(2)
		End If
	Next

	' RESULT: Item detail dictionary
	Set store_ExtractItemDetails = detailsDict

	' Release objects
	Set detailsDict = Nothing
	Set descDivision = Nothing
	Set descStoreLink = Nothing
	Set descStoreAddr = Nothing
	Set descStoreCity = Nothing
	Set descStorePhone = Nothing
	Set descStoreMap = Nothing
	Set descStoreDir = Nothing
	Set descStoreSvcs = Nothing
	Set divisionList = Nothing
	Set thisDivision = Nothing
	Set linkObj = Nothing
	
End Function


'@Description Get detailed list of  stores result contents
'@Documentation Get detailed list of  stores result contents
'@Author sbabcoc
'@Date 22-JUL-2011
'@Libraries Global
'@Repositories Locator
'@ReturnValue An array of  store detail dictionaries
Public Function store_GetResults()

   ' Declarations
   Dim descStoreDetail
   Dim storeDetailList, storeDetailCount, listIndex
   Dim results()

	With Browser("Locator").Page("REI Stores").WebElement("StoreList")
		' Load description of search result store detail objects
		Set descStoreDetail = .WebElement("Store").GetTOProperties
		' Get list of store detail objects
		Set storeDetailList = .ChildObjects(descStoreDetail)
	End With
	
	' Get count of store detail objects
	storeDetailCount = storeDetailList.Count
	' If store detail objects were found
	If (storeDetailCount > 0) Then
		' Allocate space for detail dictionaries
		ReDim results(storeDetailCount - 1)
		' Iterate over store detail objects
		For listIndex = 0 to (storeDetailCount - 1)
			' Store details of current store in the results array
			Set results(listIndex) = store_ExtractItemDetails(storeDetailList(listIndex))
		Next
	End If

	' RESULT: Detailed results
	store_GetResults = results

	' Release objects
	Set descStoreDetail = Nothing
	Set storeDetailList = Nothing
	
End Function


'@Description Extract number from store page URL
'@Documentation Extract number from <theURL>
'@Author sbabcoc
'@Date 22-JUL-2011
'@InParameter [in] theURL A reference to a store page URL
'@ReturnValue The number from the specified store page URL
Public Function store_GetNumberFromURL(ByRef theURL)

	' Declarations
	Dim regEx, matchList

	' Allocate RegExp object
	Set regEx = New RegExp

	' Set match pattern
	regEx.Pattern = ".+\.rei\.com/stores/([^/]+).*"
	' Apply pattern to extract product SKU
	Set matchList = regEx.Execute(theURL)
	' RESULT: Product SKU
	store_GetNumberFromURL = matchList(0).Submatches(0)

	' Release objects
	Set regEx = Nothing
	Set matchList = Nothing

End Function


'@Description 
'@Documentation 
'@Author sbabcoc
'@Date 11-JUL-2011
'@InParameter [in] 
'@InParameter [in] 
'@InParameter [in] 
'@ReturnValue 
Public Function store_SubmitQuery(city, state, zip)

	' Declarations
	Dim refObject
	Dim descTarget
	Dim chk_href
	Dim storeList

	' Initialize result
	storeList = Null

	Do ' <== Begin bail-out context
	
		' Navigate to the Store Locator
		chk_href = store_NavigateToPage()
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do

		With Browser("Locator").Page("REI Store Locator")
	
			' Set the "City" field
			.WebEdit("cityInput").Set city
			' Set the "State" field
			.WebEdit("stateInput").Set state
			' Set the "City" field
			.WebEdit("zipInput").Set zip
	
		End With
	
		' Get reference to "find a store" button
		Set refObject = Browser("Locator").Page("REI Store Locator").WebButton("find a store")
		' Load description of "REI Stores" page
		Set descTarget = Browser("Locator").Page("REI Stores").GetTOProperties
		' Verify target of "find a store" button
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do
		
		storeList = store_GetResults()

	Loop Until True ' <== End bail-out context
	
	store_SubmitQuery = storeList
	
	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function


Public Function store_TestLocator(city, state, zip, store)

   Dim storeList

	' Initialize result
	storeList = Null

	Do ' <== Begin bail-out context
	
		' Navigate to the Store Locator
		chk_href = store_NavigateToPage()
		' If navigation fails, bail out
		If IsNull(chk_href) Then Exit Do

		' Verify Store Locator page
		store_VerifyLocatorPage

		' Submit specified query
		storeList = store_SubmitQuery(city, state, zip)
		' If query failed, bail out
		If IsNull(storeList) Then Exit Do

		wantNone = isNull(store)
		haveNone = (safeUBound(storeList) = -1)

		chkVerifyParityEx (haveNone = True), "empty", CMP_ASSERT, wantNone, "store_TestLocator 01", "Store Locator Result"

		If Not (wantNone Or haveNone) Then
			chkVerifyParity storeList(0).Item(key_StorNum), store, STR_EQUAL, "store_TestLocator 02", "First Store Number"
		End If

		retraceStep

	Loop Until True ' <== End bail-out context
	
	store_TestLocator = storeList

	' Release objects
	If Not (safeUBound(storeList) = -1) Then Erase storeList
	
End Function


'@Description Navigate to the Store Locator
'@Documentation Navigate to the Store Locator
'@Author sbabcoc
'@Date 23-MAR-2011
'@Libraries Global, Verifications
'@Repositories Common, Web
'@ReturnValue If navigation succeeds, the URL of the Store Locator; otherwise 'Null'
Public Function store_NavigateToPage()

	' Declarations
	Dim refObject
	Dim descTarget
	Dim chk_href

	' Initialize result
	chk_href = Null

	' Get reference to current page
	Set refObject = Browser("Browser").Page("Page")
	' Load description of Store Locator page
	Set descTarget = Browser("PageStub").Page("REI Store Locator").GetTOProperties

	' If current page is the Store Locator
	If (isDescribedObject(refObject, descTarget)) Then
		' Get URL of Store Locator
		chk_href = refObject.GetROProperty("url")
	' Otherwise (not Store Locator)
	Else
		' Get reference to 'Stores' link
		Set refObject = Browser("Common").Page("REI Header").Link("Stores")
		' Verify target of 'Stores' link
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
	End If

	store_NavigateToPage = chk_href

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function


'@Description Verify elements of the REI Store Locator page
'@Documentation Verify elements of the REI Store Locator page
'@Author sbabcoc
'@Date 25-JUL-2011
Public Function store_VerifyLocatorPage()

	' Declarations
	Dim refObject

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
	
	'##### STEP 09: MAIN BODY 09 #####
	
	' Verify 'City' element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("City")
	chkState = chkVerifyExistence(refObject, "City", EXPECT_EXISTS, "Main Body 09")
	isCorrect = isCorrect And chkState
	
	'##### STEP 10: MAIN BODY 10 #####
	
	' Verify 'City' field exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebEdit("cityInput")
	chkState = chkVerifyExistence(refObject, "City", EXPECT_EXISTS, "Main Body 10")
	isCorrect = isCorrect And chkState
	
	'##### STEP 11: MAIN BODY 11 #####
	
	' Verify 'State' element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("State")
	chkState = chkVerifyExistence(refObject, "State", EXPECT_EXISTS, "Main Body 11")
	isCorrect = isCorrect And chkState
	
	'##### STEP 12: MAIN BODY 12 #####
	
	' Verify 'State' field exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebEdit("stateInput")
	chkState = chkVerifyExistence(refObject, "State", EXPECT_EXISTS, "Main Body 12")
	isCorrect = isCorrect And chkState
	
	'##### STEP 13: MAIN BODY 13 #####
	
	' Verify 'Zip' element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("Zip")
	chkState = chkVerifyExistence(refObject, "Zip", EXPECT_EXISTS, "Main Body 13")
	isCorrect = isCorrect And chkState
	
	'##### STEP 14: MAIN BODY 14 #####
	
	' Verify 'Zip' field exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebEdit("zipInput")
	chkState = chkVerifyExistence(refObject, "Zip", EXPECT_EXISTS, "Main Body 14")
	isCorrect = isCorrect And chkState
	
	'##### STEP 17: MAIN BODY 17 #####
	
	' Verify 'find a store' button exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebButton("find a store")
	chkState = chkVerifyExistence(refObject, "find a store", EXPECT_EXISTS, "Main Body 17")
	isCorrect = isCorrect And chkState
	
	'##### STEP 18: MAIN BODY 18 #####
	
	' Verify legend element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("Legend")
	chkState = chkVerifyExistence(refObject, "Legend", EXPECT_EXISTS, "Main Body 18")
	isCorrect = isCorrect And chkState
	
	'##### STEP 19: MAIN BODY 19 #####
	
	' Verify bike shop element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("Bike Shop")
	chkState = chkVerifyExistence(refObject, "Bike Shop", EXPECT_EXISTS, "Main Body 19")
	isCorrect = isCorrect And chkState
	
	'##### STEP 20: MAIN BODY 20 #####
	
	' Verify bike shop image exists
	Set refObject = Browser("Locator").Page("REI Store Locator").Image("Bike Shop")
	chkState = chkVerifyExistence(refObject, "Bike Shop", EXPECT_EXISTS, "Main Body 20")
	isCorrect = isCorrect And chkState
	
	'##### STEP 21: MAIN BODY 21 #####
	
	' Verify ski shop element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("Ski Shop")
	chkState = chkVerifyExistence(refObject, "Ski Shop", EXPECT_EXISTS, "Main Body 21")
	isCorrect = isCorrect And chkState
	
	'##### STEP 22: MAIN BODY 22 #####
	
	' Verify ski shop image exists
	Set refObject = Browser("Locator").Page("REI Store Locator").Image("Ski Shop")
	chkState = chkVerifyExistence(refObject, "Ski Shop", EXPECT_EXISTS, "Main Body 22")
	isCorrect = isCorrect And chkState
	
	'##### STEP 23: MAIN BODY 23 #####
	
	' Verify store rentals element exists
	Set refObject = Browser("Locator").Page("REI Store Locator").WebElement("Store Rentals")
	chkState = chkVerifyExistence(refObject, "Store Rentals", EXPECT_EXISTS, "Main Body 23")
	isCorrect = isCorrect And chkState
	
	'##### STEP 24: MAIN BODY 24 #####
	
	' Verify store rentals image exists
	Set refObject = Browser("Locator").Page("REI Store Locator").Image("Store Rentals")
	chkState = chkVerifyExistence(refObject, "Store Rentals", EXPECT_EXISTS, "Main Body 24")
	isCorrect = isCorrect And chkState

	store_VerifyLocatorPage = isCorrect

	' Release objects
	Set refObject = Nothing

End Function


'@Description Verify elements of the REI Stores page
'@Documentation Verify elements of the REI Stores page
'@Author sbabcoc
'@Date 25-JUL-2011
Public Function store_VerifyREIStoresPage()

	' Declarations
	Dim refObject
	Dim descTarget

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
	
	'##### STEP 01: MAIN BODY 01 #####
	
	' Verify 'Heading' element exists
	Set refObject = Browser("Locator").Page("REI Stores").WebElement("Heading")
	chkState = chkVerifyExistence(refObject, "Heading", EXPECT_EXISTS, "Main Body 01")
	isCorrect = isCorrect And chkState
	
	'##### STEP 02: MAIN BODY 02 #####
	
	' Verify 'City' field exists
	Set refObject = Browser("Locator").Page("REI Stores").Link("Search Again")
	Set descTarget = Browser("Locator").Page("REI Store Locator").GetTOProperty
	chkState = chkVerifyPageLink(refObject, descTarget, NULL)
	isCorrect = isCorrect And chkState

	store_VerifyREIStoresPage = isCorrect
	
	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function

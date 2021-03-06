Option Explicit

'@Description Navigate to the main page of the specified product
'@Documentation Navigate to the main page of product <partNumber>
'@Author sbabcoc
'@Date 23-MAR-2011
'@Libraries Global, Verifications
'@Repositories Common, Web, Product
'@InParameter [in] partNumber The SKU of the product whose main page is requested
'@ReturnValue If navigation succeeds, the URL of the specified product; otherwise 'Null'
Public Function product_NavigateToPage(partNumber)

	' Declarations
	Dim refObject
	Dim descTarget
	Dim pageSKU
	Dim chk_href

	' Initialize result
	chk_href = Null

	' Get reference to current page
	Set refObject = Browser("Browser").Page("Page")
	' Load description of generic product page
	Set descTarget = Browser("Product").Page("ProductCommon").GetTOProperties

	' If current page is a product page
	If (isDescribedObject(refObject, descTarget)) Then
		' Get URL of current page
		chk_href = refObject.GetROProperty("url")
		' Extract SKU from product page URL
		pageSKU = product_GetSKUFromURL(chk_href)
		' If current page is for specified product
		If (pageSKU = Left(partNumber, 6)) Then
			' Get reference to 'Product Page' link
			Set refObject = Browser("Product").Page("ProductCommon").Link("Product Page")
			' If 'Product Page' link exists
			If (refObject.Exist(0)) Then
				' Click 'Product Page' link
				followLink(refObject)
				' Get URL of current page
				chk_href = refObject.GetROProperty("url")
			End If
		' Otherwise (not specified product)
		Else
			' Reset result
			chk_href = Null
		End If
	End If

	' If navigation required
	If IsNull(chk_href) Then
		' Load description of product main page
		Set descTarget = Browser("Product").Page("ProductMain").GetTOProperties
		' Search for specified product
		chk_href = search_SubmitQuery(partNumber, descTarget)
	End If

	product_NavigateToPage = chk_href

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function


'@Description Extract SKU from product page URL
'@Documentation Extract SKU from <theURL>
'@Author sbabcoc
'@Date 05-MAY-2011
'@InParameter [in] theURL A reference to a product page URL
'@ReturnValue The SKU from the specified product page URL
Public Function product_GetSKUFromURL(ByRef theURL)

	' Declarations
	Dim regEx, matchList

	' Allocate RegExp object
	Set regEx = New RegExp

	' Set match pattern
	regEx.Pattern = ".+\.rei\.com/product/([^/]+).*"
	' Apply pattern to extract product SKU
	Set matchList = regEx.Execute(theURL)
	' RESULT: Product SKU
	product_GetSKUFromURL = matchList(0).Submatches(0)

	' Release objects
	Set regEx = Nothing
	Set matchList = Nothing

End Function


'@Description Process elements of a product price group into a details dictionary
'@Documentation Process elements of <priceGroupObj> into <detailsDictObj>
'@Author sbabcoc
'@Date 05-MAY-2011
'@InParameter [in] priceGroupObj A reference to a product page price group
'@OutParameter [out] detailsDictObj A reference to a product details dictionary
'@ReturnValue Count of processed product group elements
Public Function product_ProcessPriceGroup(ByRef priceGroupObj, ByRef detailsDictObj)

	' Declarations
	Dim descDivision
	Dim descOrigPrice, descProdPrice, descSalePrice
	Dim divisionList
	Dim thisDivision
	Dim listIndex
	
	' Load generic division element description
	Set descDivision = Browser("Browser").Page("Page").WebElement("Division").GetTOProperties

	With Browser("Product").Page("ProductMain").WebElement("ProdRight").WebElement("PriceGroup")
		' Load product price object descriptions
		Set descOrigPrice = .WebElement("OrigPrice").GetTOProperties
		Set descProdPrice = .WebElement("ProdPrice").GetTOProperties
		Set descSalePrice = .WebElement("SalePrice").GetTOProperties
	End With

	' Get list of price group objects
	Set divisionList = priceGroupObj.ChildObjects(descDivision)
	' Iterate over price group objects
	' NOTE: Skip nested container object
	For listIndex = 1 To (divisionList.Count - 1)
		' Get current price group object
		Set thisDivision = divisionList(listIndex)

		' If this is the original price
		If (isDescribedObject(thisDivision, descOrigPrice)) Then
			' Add original price to detail dictionary
			detailsDictObj.Add key_ProdOrig, search_ExtractPrice(thisDivision)
		' Otherwise, if this is the current price
		ElseIf (isDescribedObject(thisDivision, descProdPrice)) Then
			' Add original price to detail dictionary
			detailsDictObj.Add key_ProdCost, search_ExtractPrice(thisDivision)
		' Otherwise, if this is the sale price
		ElseIf (isDescribedObject(thisDivision, descSalePrice)) Then
			' Add sale price to detail dictionary
			detailsDictObj.Add key_ProdCost, search_ExtractPrice(thisDivision)
		End If
	Next

	' RESULT: Count of price group objects
	product_ProcessPriceGroup = divisionList.Count

	' Release objects
	Set descDivision = Nothing
	Set descOrigPrice = Nothing
	Set descProdPrice = Nothing
	Set descSalePrice = Nothing
	Set priceGroupObj = Nothing
	Set divisionList = Nothing
	Set thisDivision = Nothing
	
End Function


'@Description Extract value from percent savings element
'@Documentation Extract value from <pctSavingObj>
'@Author sbabcoc
'@Date 05-MAY-2011
'@InParameter [in] pctSavingObj A reference to a percent saving object
'@ReturnValue The value of the specified percent saving object
Public Function product_ExtractPctSaving(ByRef pctSavingObj)

	' Declarations
	Dim regEx, matchList
	Dim descPctSaving
	
	' Allocate RegExp object
	Set regEx = New RegExp

	With Browser("Product").Page("ProductMain").WebElement("ProdRight")
		Set descPctSaving = .WebElement("PctSaving").GetTOProperties
	End With

	' Set match pattern
	regEx.Pattern = descPctSaving("innertext").Value
	' Apply pattern to extract percent saving
	Set matchList = regEx.Execute(pctSavingObj.GetROProperty("innertext"))
	' RESULT: Percent saving
	product_ExtractPctSaving  = CInt(matchList(0).Submatches(0))

	' Release objects
	Set regEx = Nothing
	Set descPctSaving = Nothing
	Set matchList = Nothing
	
End Function


'@Description Extract detail dictionary from product page
'@Documentation Extract product detail dictionary
'@Author sbabcoc
'@Date 22-MAR-2011
'@ReturnValue The detail dictionary for the current product page
Public Function product_ExtractItemDetails()

   ' Declarations
	Dim detailsDict
	Dim regEx, matchList
	Dim descElement, descPctSaving, descPriceGroup
	Dim descProdNum, descProdTitle, descProcRating
	Dim prodRightObj
	Dim elementList
	Dim thisElement
	Dim listIndex

	' Allocate RegExp object
	Set regEx = New RegExp

	' Create a product detail dictionary
	Set detailsDict = CreateObject("Scripting.Dictionary")

	' Define generic web element  description
	Set descElement = Description.Create()
	descElement("micclass").Value = "WebElement"
	
	With Browser("Product").Page("ProductMain").WebElement("ProdRight")
		' Load product detail object descriptions
		Set descPctSaving = .WebElement("PctSaving").GetTOProperties
		Set descProdNum = .WebElement("ProdNum").GetTOProperties
		Set descPriceGroup = .WebElement("PriceGroup").GetTOProperties
		Set descProdTitle = .WebElement("ProdTitle").GetTOProperties
		Set descProdRating = .Image("ProdRating").GetTOProperties
	End With

	' Get reference to "ProdRight" container object
	Set prodRightObj = Browser("Product").Page("ProductMain").WebElement("ProdRight")

	' Get list of elements in "ProdRight" container object
	Set elementList = prodRightObj.ChildObjects(descElement)
	' Iterate over "ProdRight" elements
	For listIndex = 0 To (elementList.Count - 1)
		' Get current element
		Set thisElement = elementList(listIndex)

		' If this element contains text
		If Len(thisElement.GetROProperty("innertext")) Then
			' If this is the product title
			If (isDescribedObject(thisElement, descProdTitle)) Then
				' Add product name to detail dictionary
				detailsDict.Add key_ProdName, thisElement.GetROProperty("innertext")
			' Otherwise, if this is the product number
			ElseIf (isDescribedObject(thisElement, descProdNum)) Then
				' Set match pattern
				regEx.Pattern = descProdNum("innertext").Value
				' Apply pattern to extract product number
				Set matchList = regEx.Execute(thisElement.GetROProperty("innertext"))
				' Add product number to detail dictionary
				detailsDict.Add key_ProdNum, matchList(0).SubMatches(0)
			' Otherwise, if this is the product price group
			ElseIf (isDescribedObject(thisElement, descPriceGroup)) Then
				' Process price group objects
				groupCount = product_ProcessPriceGroup(thisElement, detailsDict)
				' Account for processed objects
				listIndex = listIndex + groupCount
    		' Otherwise, if this is the percent saving
			ElseIf (isDescribedObject(thisElement, descPctSaving)) Then
				' Add percent saving to detail dictionary
				detailsDict.Add key_ProdSave, product_ExtractPctSaving(thisElement)
			End If
		' Otherwise, if this is the product rating
		ElseIf (isDescribedObject(thisElement, descProdRating)) Then
			' Add product rating to detail dictionary
			detailsDict.Add key_ProdRate, search_ExtractRating(thisElement)
		End If
	Next

	' RESULT: Item detail dictionary
	Set product_ExtractItemDetails = detailsDict

	' Release objects
	Set regEx = Nothing
	Set detailsDict = Nothing
	Set descElement = Nothing
	Set descPctSaving = Nothing
	Set descProdNum = Nothing
	Set descPriceGroup = Nothing
	Set descProdTitle = Nothing
	Set descProdRating = Nothing
	Set prodRightObj = Nothing
	Set elementList = Nothing
	Set thisElement = Nothing
	Set matchList = Nothing
	
End Function


'@Description Extract product specifications
'@Documentation Extract product specifications
'@Author sbabcoc
'@Date 22-MAR-2011
'@Repositories Product
'@ReturnValue The specifications for the current product
Public Function product_ExtractItemSpecs()

   ' Declarations
	Dim prodSpecsDict
	Dim descSpecRow, descSpecLabel, descSpecValue
	Dim specRowList, specLabelList, specValueList
	Dim strSpecLabel, strSpecValue
	Dim thisSpecRow
	Dim listIndex

	' Create a product specifications dictionary
	Set prodSpecsDict = CreateObject("Scripting.Dictionary")

	With Browser("Product").Page("ProductSpecs")
		' Load descriptions of specification label and value
		Set descSpecLabel = .WebElement("Row").WebElement("Specification").GetTOProperties
		Set descSpecValue = .WebElement("Row").WebElement("Description").GetTOProperties

		' Load product specification row object description
		Set descSpecRow = .WebElement("Row").GetTOProperties
		' Get list of rows on product specification page
		Set specRowList = .ChildObjects(descSpecRow)
	End With
	
	' Iterate over product specification rows
	For listIndex = 0 To (specRowList.Count - 1)
		' Get current product specification row
		Set thisSpecRow = specRowList(listIndex)

		' Get lists to product specification label and value elements
		Set specLabelList = thisSpecRow.ChildObjects(descSpecLabel)
		Set specValueList = thisSpecRow.ChildObjects(descSpecValue)

		' If both product specification elements exist
		If ((specLabelList.Count > 0) And (specValueList.Count > 0)) Then
			' Extract text from product specification label
			strSpecLabel = specLabelList(0).GetROProperty("innertext")
			' Extract text from product specification value
			strSpecValue = specValueList(0).GetROProperty("innertext")
			' Add product specification to dictionary
			prodSpecsDict.Add strSpecLabel, strSpecValue
		End If
	Next

	' RESULT: Product specifications dictionary
	Set product_ExtractItemSpecs = prodSpecsDict

	' Release objects
	Set prodSpecsDict = Nothing
	Set descSpecRow = Nothing
	Set descSpecLabel = Nothing
	Set descSpecValue = Nothing
	Set specRowList = Nothing
	Set specLabelList = Nothing
	Set specValueList = Nothing
	Set thisSpecRow = Nothing
	
End Function


'@Description Extract product color/size options
'@Documentation Extract product color/size options
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] refColorSize A reference to a color/size WebList object
'@ReturnValue A dictionary of color/size options
'		Keys are 10-digit product SKUs
'		Values are <index>|<label>|<status>|<value>
'			<index>: WebList option index
'			<label>: WebList option label (color/size/price)
'			<status>: 'True' if on backorder; otherwise 'False"
'			<value>: WebList option value
Public Function product_ExtractOptions(ByRef refColorSize)

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
		
		' Define pattern to extract backorder status
		regEx.Pattern = "( \* )?(.*)"
		' Apply pattern to extract backorder status
		Set matchList = regEx.Execute(optionText)
		' Extract option backorder status
		optionStat = (Len(matchList(0).SubMatches(0)) > 0)
		' Extract option color/size/price
		optionDesc = matchList(0).SubMatches(1)
		
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

		' Define pattern to extract product number
		regEx.Pattern = "productnumber=""(.*)"""
		' Apply pattern to extract product number
		Set matchList = regEx.Execute(optionTag)
		' If product number was captured
		If (matchList.Count = 1) Then
			' Extract captured product number
			optionNum = matchList(0).SubMatches(0)
		' Otherwise
		Else
			' No product number
			optionNum = "_none_"
		End If

		' Add option specification to dictionary
		optionsDict.Add optionNum, "#" & listIndex & "|" & optionDesc & "|" & optionStat & "|" & optionVal
	Next

	' RESULT: Option specifications dictionary
	Set product_ExtractOptions = optionsDict

	' Release objects
	Set optionsDict = Nothing
	Set regEx = Nothing
	Set matchList = Nothing

End Function


'@Description Verify the target of the indicated product page tab link
'@Documentation Verify the target of the indicated product page tab link
'@Author sbabcoc
'@Date 22-MAR-2011
'@Libraries Verifications
'@InParameter [in] refObject A reference to the tab link being verified
'@InParameter [in] tabFlag A bit-mapped flag value indicating the tab being verified
'		prod_GeneralTab: Product general information tab
'		prod_DescripTab: Product description tab
'		prod_ImagesTab: Product images tab
'		prod_SpecsTab: Product specifications tab
'		prod_ReviewsTab: Product reviews tab
'@InParameter [in] doExpect Existential expectation:
'		EXPECT_EXISTS: Expect indicated tab to be present
'		EXPECT_ABSENT: Expect indicated tab to be absent
'@InParameter [in] baseURL Base URL of current product page
'@ReturnValue 'True' if  verification passes; otherwise 'False'
Public Function product_VerifyTabLink(ByRef refObject, ByVal tabFlag, ByVal doExpect, ByRef baseURL)

	' Declarations
	Dim tabName
	Dim tabParm
	Dim isCorrect

	Select Case tabFlag

		Case prod_GeneralTab
			tabName = "GENERAL"
			tabParm = "?un_jtt_v_product=yes"

		Case prod_DescripTab
			tabName = "DESCRIPTION"
			tabParm = "?un_jtt_v_anchor=prodInfor"

		Case prod_ImagesTab
			tabName = "IMAGES"
			tabParm = "?un_jtt_v_anchor=tabsMedia"

		Case prod_SpecsTab
			tabName = "SPECS"
			tabParm = "?un_jtt_v_anchor=prodSpecs"

		Case prod_ReviewsTab
			tabName = "REVIEWS"
			tabParm = "?un_jtt_v_anchor=customerReview"

	End Select

	' If presence/absence of specified tab is as expected
	If chkVerifyExistence(refObject, tabName & " tab", doExpect, "product_VerifyTabLink EXISTS") Then
		' If expecting to exist
		If (doExpect) Then
			' Indicate parity between actual and expected tab link targets
			isCorrect = chkVerifyParity(refObject.GetROProperty("href"), baseURL & tabParm, STR_EQUAL, "product_VerifyTabLink TARGET", "Link target")
		Else
			' Indicate success
			isCorrect = True
		End If
	Else
		' Indicate failure
		isCorrect = False
	End If

	product_VerifyTabLink = isCorrect

End Function


'@Description Verify presence/absence/targets of product page tab links
'@Documentation Verify presence/absence/targets of product page tab links
'@Author sbabcoc
'@Date 22-MAR-2011
'@Libraries Verifications
'@Repositories Web, Product
'@ReturnValue 'True' if  verification passes; otherwise 'False'
Public Function product_VerifyPageTabs()

	' Declarations
	Dim resultVal
	Dim isCorrect
	Dim refObject
	Dim descTarget
	Dim baseURL
	Dim offset
	Dim viewFlag
	Dim viewHead
	Dim doExpect
	Dim listIndex
	Dim tabValue

	' Initialize results
	resultVal = 0
	isCorrect = True

	' If current page is a product page view
	Set refObject = Browser("Browser").Page("Page")
	Set descTarget = Browser("Product").Page("ProductCommon").GetTOProperties
	If (isDescribedObject(refObject, descTarget)) Then

		'##### EXTRACT PRODUCT BASE URL #####

		baseURL = refObject.GetROProperty("url")
		offset = InStr(1, baseURL, "?")
		If (offset > 0) Then
			baseURL = Left(baseURL, offset - 1)
		End If

		'##### DETERMINE CURRENT VIEW #####

		' Get reference to page heading object
		Set refObject = Browser("Product").Page("ProductCommon").WebElement("Heading")
		' If page heading exists
		If (refObject.Exist(0)) Then
			' If this is a tab header
			If (refObject.GetROProperty("height") = prod_TabHeadSize) Then
				' Get heading text
				viewHead = refObject.GetROProperty("innertext")
	
				Select Case Trim(viewHead)
	
					Case "Description"
						' This is the Description tab
						viewFlag = prod_DescripTab
	
					Case "Item Specifications"
						' This is the Specs tab
						viewFlag = prod_SpecsTab
	
					Case Else
						' This is the Images tab
						viewFlag = prod_ImagesTab
	
				End Select
			' Otherwise (not a tab header)
			Else
				' This is the product main tab
				viewFlag = prod_GeneralTab
			End If
		' Otherwise (page heading absent)
		Else
			' This is the Reviews tab
			viewFlag = prod_ReviewsTab
		End If

		With Browser("Product").Page("ProductCommon")
			' Initialize list of tab values
			tabsList = Array(prod_GeneralTab, prod_DescripTab, prod_ImagesTab, prod_SpecsTab, prod_ReviewsTab)
			' Initialize list of tab references
			refsList = Array(.Link("Product Page"), .Link("Product Description"), .Link("Product Images"), .Link("Product Specs"), .Link("Customer Reviews"))
		End With

		' Iterate over product page tabs
		For listIndex = 0 To UBound(tabsList)
			' Get current tab value
			tabValue = tabsList(listIndex)
			' Get current tab reference
			Set refObject = refsList(listIndex)
			
			' Set expectation for presence/absence
			doExpect = Not (viewFlag = tabValue)

			' If current product page tab matches expectations
			If product_VerifyTabLink(refObject, tabValue, doExpect, baseURL) Then
				' Indicate current  tab is correct
				resultVal = resultVal Or tabValue
			Else
				' Indicate failure
				isCorrect = False
			End If
		Next
	Else
		' Indicate failure
		isCorrect = False
		Reporter.ReportEvent micWarning, "product_VerifyPageTabs", "Current page is not a product page"
	End If

	If (isCorrect) Then resultVal = resultVal Or prod_VerifyOK
	product_VerifyPageTabs = resultVal

	' Release objects
	Erase refsList
	Set descTarget = Nothing
	Set refObject = Nothing
	
End Function


'@Description Select the indicated item in the Color/Size list
'@Documentation Select <selectSpec> in the Color/Size list
'@Author sbabcoc
'@Date 28-MAR-2011
'@Libraries Verifications
'@InParameter [in] selectSpec, string, item to select (SKU/label/index)
'@ReturnValue 'True' is selection succeeds; otherwise 'False'
Public Function product_SelectColorSize(ByVal selectSpec)

   didSelect = False

	' Get reference to color/size options list
	Set refColorSize = Browser("Product").Page("ProductMain").WebList("ColorSize")
	' If color/size options list exists
	If chkVerifyExistence(refColorSize, "Color/Size List", EXPECT_EXISTS, "product_SelectColorSize") Then
		' If selection specifier is numeric
		If IsNumeric(selectSpec) Then
			' If specifier is a 10-digit SKU
			If (Len(selectSpec) = 10) Then
				' Extract product color/size options
				Set optionsDict = product_ExtractOptions(refColorSize)

				' If SKU matches a color/size option
				If (optionsDict.Exists(selectSpec)) Then
					' Extract option properties
					optionSpec = optionsDict.Item(selectSpec)
					' Split (index, label, status, value)
					optionBits = Split(optionSpec, "|")
					' Get index of spec'd SKU
					selectSpec = optionBits(0)
				End If

				' Release object
				Set optionsDict = Nothing
			End If
		End If
	
		Err.Clear
		On Error Resume Next
	
		' Try specifier as-is and backordered
		selectList = Array(selectSpec, " * " & selectSpec)
		' Iterate over option states
		For Each thisSelect In selectList
			' Try to select this option
			refColorSize.Select thisSelect
			' If selection succeeded
			If (Err.Number = 0) Then
				' Indicate success
				didSelect = True
				' Stop iteration
				Exit For
			' Otherwise (selection failed)
			Else
				' Extract error information
				errNum = Err.Number
				errMsg = Err.Description
				errSrc = Err.Source
				' Reset error
				Err.Clear
			End If
		Next

		' Verify result of selection process
		chkVerifyParity didSelect, "selected", CMP_ASSERT, "product_SelectColorSize", "Specified color/size"
	End If

	product_SelectColorSize = didSelect

	' Release objects
	Set refColorSize = Nothing

End Function


'@Description Add indicated quantity of specified product to shopping catr
'@Documentation Add <partQuantity> units of product <partNumber> to shopping cart
'@Author sbabcoc
'@Date 29-JUN-2011
'@Libraries Verifications
'@Repositories Product, Checkout
'@InParameter [in] partNumber, string, SKU of product to be added
'@InParameter [in] partQuantity, number, Quantity of product to add
'@InParameter [in] doEmptyCart, boolean, 'True' to purge cart; otherwise 'False'
'@ReturnValue 'True' is addition succeeds; otherwise 'False'
Public Function product_AddToCart(partNumber, partQuantity, doEmptyCart)

	' Declarations
	Dim isCorrect
	Dim chk_href

	' Initialize result
	isCorrect = False

	Do ' <== Begin bail-out context

		' If cart purge requested
		If (doEmptyCart) Then
			' Purge the shopping cart
			isCorrect = cart_RemoveItem(Null)
			' If purge fails, bail out
			If Not (isCorrect) Then Exit Do
		End If

		' Navigate to main page of specified product
		chk_href = product_NavigateToPage(partNumber)
		' Check navigation status
		isCorrect = Not IsNull(chk_href)
		' If navigation fails, bail out
		If Not (isCorrect) Then Exit Do
	
		' Select specific color/size option
		isCorrect = product_SelectColorSize(partNumber)
		' If selection fails, bail out
		If Not (isCorrect) Then Exit Do

		' Set content of Quantity field to specified quantity
		Browser("Product").Page("ProductMain").WebEdit("Quantity").Set partQuantity
		
		' Get reference to "Add to Cart" button
		Set refObject = Browser("Product").Page("ProductMain").WebButton("Add to Cart")

		splashPage = query_GetSplashPage(partNumber)
		' If product has associated splash
		If (Len(splashPage) > 0) Then
			' Assemble semaphore name
			splashFlag = "SPLASH_" & UCase(splashPage)
			' Determine if we've accepted this already
			splashTest = MyGetParameter(dtLocalSheet, splashFlag, 1)

			' If we haven't accepted yet
			If IsNull(splashTest) Then
				Select Case splashPage

					Case "boats"
						' Load description of Paddle Sports Gear terms page
						Set descTarget = Browser("Checkout").Page("PaddleGearTerms").GetTOProperties

					Case "cycle"
						' Load description of Wheeled Sports Gear terms page
						Set descTarget = Browser("Checkout").Page("WheeledGearTerms").GetTOProperties

					Case "skis"
						' Load description of Snow Sports Gear terms page
						Set descTarget = Browser("Checkout").Page("SnowGearTerms").GetTOProperties

				End Select

				' If we know what to expect
				If Not IsEmpty(descTarget) Then
					' Verify click of "Add to Cart" shows expected splash
					chk_href = chkVerifyLinkTarget(refObject, descTarget)
					' Check navigation status
					isCorrect = Not IsNull(chk_href)
					' If navigation fails, bail out
					If Not (isCorrect) Then Exit Do

					' Get reference to "agree" button
					Set refObject = Browser("Checkout").Page("TermsPageCommon").WebButton("accept")
				End If
			End If
		End If

		' Load description of Shopping Cart page
		Set descTarget = Browser("Checkout").Page("REI.com: Shopping Basket").GetTOProperties
		
		' Verify button click lands on Shopping Cart page
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' Check navigation status
		isCorrect = Not IsNull(chk_href)
		' If navigation fails, bail out
		If Not (isCorrect) Then Exit Do

		' Verify addition of product/quantity
		isCorrect = cart_VerifyAddition(partNumber, partQuantity)

	Loop Until True ' <== End bail-out context
	
	product_AddToCart = isCorrect

	' Release objects
	Set descTarget = Nothing
	Set refObject = Nothing

End Function

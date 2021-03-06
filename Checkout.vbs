Option Explicit

'@Description Update splash page agreement flags from cart contents
'@Documention Update splash page agreement flags from cart contents
'@Author sbabcoc
'@Date 11-JUL-2011
Public Sub cart_UpdateSplashFlags()

   ' Declarations
	Dim descItemNum
	Dim itemNumList
	Dim itemNumCount
	Dim listIndex
	Dim regEx, matchList
	Dim thisItemNum
	Dim splashPage
	Dim splashFlag
	
	On Error Resume Next

	Set sheetRef = MyGetSheetRef(dtLocalSheet)
	sheetRef.DeleteParameter("SPLASH_BOATS")
	sheetRef.DeleteParameter("SPLASH_CYCLE")
	sheetRef.DeleteParameter("SPLASH_SKIS")
	Set sheetRef = Nothing

	On Error GoTo 0

	With Browser("Checkout").Page("REI.com: Shopping Basket")
		' Get description of shopping cart item number objects
		Set descItemNum = .WebElement("ItemNum").GetTOProperties
		' Get list of item number objects
		Set itemNumList = .ChildObjects(descItemNum)
	End With

	' Get count of item number objects
	itemNumCount = itemNumList.Count
	' If item number objects were found
	If (itemNumCount > 0) Then
		' Allocate RegExp object
		Set regEx = New RegExp
		' Set pattern to extract item number
		regEx.Pattern = descItemNum("innertext")
		
		' Iterate over item number objects
		For listIndex = 0 to (itemNumCount - 1)
			' Get current item number object
			Set thisItemNum = itemNumList(listIndex)
			' Apply pattern to extract item number
			Set matchList = regEx.Execute(thisItemNum.GetROProperty("innertext"))
			' Query for item splash page
			splashPage = query_GetSplashPage(matchList(0).SubMatches(0))
			' If product has associated splash
			If (Len(splashPage) > 0) Then
				' Assemble semaphore name
				splashFlag = "SPLASH_" & UCase(splashPage)
				' Indicate acceptance of terms
				MySetParameter dtLocalSheet, splashFlag, 1, True, flagAddParam
			End If
		Next

		' Release objects
		Set thisItemNum = Nothing
		Set matchList = Nothing
		Set regEx = Nothing
	End If

	' Release objects
	Set itemNumList = Nothing
	Set descItemNum = Nothing

End Sub


'@Description Remove all items from the shopping cart
'@Documentation Remove all items from the shopping cart
'@Note This function uses the "Remove" button to remove items one at a time
'@Note The cart can also be emptied with cart_RemoveItem(Null)
Public Function EmptyShoppingCart()

   ' Declarations
   Dim isCorrect
   Dim chk_href

   ' Initialize result
   isCorrect = False

	' Navigate to Shopping Cart page
	chk_href = cart_NavigateToPage()
	' If navigation succeeds
	If Not IsNull(chk_href) Then
		cmnSetGlobalTimeouts 2000
	
		Do While Browser("Checkout").Page("REI.com: Shopping Basket").WebButton("Remove").Exist
			Browser("Checkout").Page("REI.com: Shopping Basket").WebButton("Remove").Click
			Browser("Browser").Page("Page").Sync
		Loop
	
		cmnSetGlobalTimeouts 30000
	
		' Update splash flags
		cart_ UpdateSplashFlags
		' Indicate success
		isCorrect = True
	End If

	EmptyShoppingCart = isCorrect
		
End Function


'@Description Remove specified part number from the shopping cart
'@Documentation Remove path <partNumber> from the shopping cart
'@Author sbabcoc
'@Date 11-JUL-2011
'@InParameter [in] partNumber Part number to be removed
'		Every item beginning with the specified number will be removed
'		Specify 'Null' to remove all items from the shopping cart
Public Function cart_RemoveItem(partNumber)

   ' Declarations
   Dim isCorrect
   Dim chk_href

   ' Initialize result
   isCorrect = False

	' Navigate to Shopping Cart page
	chk_href = cart_NavigateToPage()
	' If navigation succeeds
	If Not IsNull(chk_href) Then
		' Purge the shopping cart
		isCorrect = cart_UpdateQuantity(partNumber, updt_FixedQuantity) ' quantity = 0
	End If

	' Update splash flags
	cart_UpdateSplashFlags

	cart_RemoveItem = isCorrect
	
End Function


'@Description Navigate to the Shopping Cart
'@Documentation Navigate to the Shopping Cart
'@Author sbabcoc
'@Date 23-MAR-2011
'@Libraries Global, Verifications
'@Repositories Common, Web
'@ReturnValue If navigation succeeds, the URL of the Shopping Cart; otherwise 'Null'
Public Function cart_NavigateToPage()

	' Declarations
	Dim refObject
	Dim descTarget
	Dim chk_href

	' Initialize result
	chk_href = Null

	' Get reference to current page
	Set refObject = Browser("Browser").Page("Page")
	' Load description of Shopping Cart page
	Set descTarget = Browser("PageStub").Page("REI.com: Shopping Basket").GetTOProperties

	' If current page is the Shopping Cart
	If (isDescribedObject(refObject, descTarget)) Then
		' Get URL of Shopping Cart
		chk_href = refObject.GetROProperty("url")
	' Otherwise (not Shopping Cart)
	Else
		' Get reference to 'Cart' link
		Set refObject = Browser("Common").Page("REI Header").Link("Cart")
		' Verify target of 'Cart' link
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
	End If

	cart_NavigateToPage = chk_href

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function


'@Description Extract detail dictionary from shopping cart item
'@Documentation Extract detail dictionary from <productObj>
'@Author sbabcoc
'@Date 22-MAR-2011
'@InParameter [in] productObj A reference to a shopping cart item object
'@ReturnValue The detail dictionary for the specified shopping cart item
Public Function cart_ExtractItemDetails(ByRef productObj)

   ' Declarations
	Dim detailsDict
	Dim regEx, matchList
	Dim descDivision, descItemImage, descItemDetail
	Dim descItemLink, descItemNum, descItemCost, descItemQty
	Dim descItemLinkObj, descItemQtyObj
	Dim itemImageList, itemDetailList, divisionList, itemLinkList, itemQtyList
	Dim thisDivision
	Dim listIndex

	' Allocate RegExp object
	Set regEx = New RegExp

	' Create a product detail dictionary
	Set detailsDict = CreateObject("Scripting.Dictionary")

	' Load generic division element description
	Set descDivision = Browser("Browser").Page("Page").WebElement("Division").GetTOProperties
	
	With Browser("Checkout").Page("REI.com: Shopping Basket")
		' Load product detail object descriptions
		Set descItemImage = .Image("ItemImage").GetTOProperties
		Set descItemDetail = .WebElement("ItemDetail").GetTOProperties
		Set descItemLink = .WebElement("ItemLink").GetTOProperties
		Set descItemNum = .WebElement("ItemNum").GetTOProperties
		Set descItemCost = .WebElement("ItemCost").GetTOProperties
		Set descItemQty = .WebElement("ItemQty").GetTOProperties

		' Load product link object description
		Set descItemLinkObj = .Link("ItemLink").GetTOProperties
		' Load product quantity object description
		Set descItemQtyObj = .WebEdit("ItemQty").GetTOProperties
	End With

	' Get list of product image objects
	Set itemImageList = productObj.ChildObjects(descItemImage)
	' If product image is present
	If (itemImageList.Count = 1) Then
		' Add product image link to detail dictionary
		detailsDict.Add key_ProdPict, itemImageList(0).GetROProperty("href")
	End If

	' Get list of product detail objects
	Set itemDetailList = productObj.ChildObjects(descItemDetail)
	' If product detail is present
	If (itemDetailList.Count = 1) Then
		' Get list of product detail objects
		Set divisionList = itemDetailList(0).ChildObjects(descDivision)
		' Iterate over product detail objects
		For listIndex = 0 To (divisionList.Count - 1)
			' Get current product detail object
			Set thisDivision = divisionList(listIndex)
	
			' If this is the product link
			If (isDescribedObject(thisDivision, descItemLink)) Then
				' Get list of product link objects
				Set itemLinkList = thisDivision.ChildObjects(descItemLinkObj)
				' Add product link to detail dictionary
				detailsDict.Add key_ProdLink, itemLinkList(0).GetROProperty("href")
				' Add product name to detail dictionary
				detailsDict.Add key_ProdName, itemLinkList(0).GetROProperty("innertext")
			' Otherwise, if this is the product number
			ElseIf (isDescribedObject(thisDivision, descItemNum)) Then
				' Set match pattern
				regEx.Pattern = descItemNum("innertext").Value
				' Apply pattern to extract product number
				Set matchList = regEx.Execute(thisDivision.GetROProperty("innertext"))
				' Add product number to detail dictionary
				detailsDict.Add key_ProdNum, matchList(0).SubMatches(0)
			' Otherwise, if this is the product cost
			ElseIf (isDescribedObject(thisDivision, descItemCost)) Then
				' Set match pattern
				regEx.Pattern = descItemCost("innertext").Value
				' Apply pattern to extract product cost
				Set matchList = regEx.Execute(thisDivision.GetROProperty("innertext"))
				' Add product cost to detail dictionary
				detailsDict.Add key_ProdCost, CCur(matchList(0).SubMatches(0))
			' Otherwise, if this is the product quantity
			ElseIf (isDescribedObject(thisDivision, descItemQty)) Then
				' Get list of product quantity objects
				Set itemQtyList = thisDivision.ChildObjects(descItemQtyObj)
				' Add product quantity to detail dictionary
				detailsDict.Add key_ProdQty, CInt(itemQtyList(0).GetROProperty("value"))
			End If
		Next
	End If

	' RESULT: Item detail dictionary
	Set cart_ExtractItemDetails = detailsDict

	' Release objects
	Set regEx = Nothing
	Set detailsDict = Nothing
	Set descDivision = Nothing
	Set descItemImage = Nothing
	Set descItemDetail = Nothing
	Set descItemLink = Nothing
	Set descItemNum = Nothing
	Set descItemCost = Nothing
	Set descItemQty = Nothing
	Set descItemLinkObj = Nothing
	Set descItemQtyObj = Nothing
	Set divisionList = Nothing
	Set thisDivision = Nothing
	Set itemDetailList = Nothing
	Set itemImageList = Nothing
	Set itemLinkList = Nothing
	Set matchList = Nothing
	Set itemQtyList = Nothing
	
End Function


'@Description Get detailed list of  shopping cart contents
'@Documentation Get detailed list of  shopping cart contents
'@Author sbabcoc
'@Date 23-MAR-2011
'@ReturnValue An array of  item detail dictionaries
Public Function cart_GetInventory()

	' Declarations
	Dim descItemWrap
	Dim itemWrapList
	Dim itemWrapCount, listIndex
	Dim inventory()

	With Browser("Checkout").Page("REI.com: Shopping Basket")
		' Get description of shopping cart item wrap objects
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
			Set inventory(listIndex) = cart_ExtractItemDetails(itemWrapList(listIndex))
		Next
	End If

	' RESULT: Detailed inventory
	cart_GetInventory = inventory

	' Release objects
	Set descItemWrap = Nothing
	Set itemWrapList = Nothing

End Function


'@Description Update quantity of specified shopping cart item
'@Documentation Update quantity of specified shopping cart item
'@Author sbabcoc
'@Date 11-JUL-2011
'@InParameter [in] productObj A reference to a shopping cart item object
'@InParameter [in] partNumber Part number of product to be updated
'		Set to 'Null' to request update of every product
'@InParameter [in] partQuantity Quantity to set if update indicated
'		Set to 'Empty' to request use of random quantity
'@ReturnValue 'True' if update was performed; otherwise 'False'
Public Function cart_UpdateItemQuantity(ByRef productObj, partNumber, partQuantity)

	' Declarations
	Dim regEx, matchList
	Dim descItemNum, itemNumList
	Dim descItemQty, itemQtyList
	Dim doUpdate

	' If every product is indicated
	If IsNull(partNumber) Then
		doUpdate = True
	Else
		' Load description of item number object
		Set descItemNum = Browser("Checkout").Page("REI.com: Shopping Basket").WebElement("ItemNum").GetTOProperties
		' Allocate RegExp object
		Set regEx = New RegExp
		' Set pattern to extract item number
		regEx.Pattern = descItemNum("innertext")

		' Get list of item quantity fields
		Set itemNumList = productObj.ChildObjects(descItemNum)
		' Apply pattern to extract product number
		Set matchList = regEx.Execute(itemNumList(0).GetROProperty("innertext"))
		' Update if item part number begins with specified number
		doUpdate = (InStr(1, matchList(0).SubMatches(0), partNumber) = 1)

		' Release objects
		Set matchList = Nothing
		Set itemNumList = Nothing
		Set rexEx = Nothing
		Set descItemNum = Nothing
	End If

	If (doUpdate) Then
		' Load description of item quantity field
		Set descItemQty = Browser("Checkout").Page("REI.com: Shopping Basket").WebEdit("ItemQty").GetTOProperties
		' Get list of item quantity fields
		Set itemQtyList = productObj.ChildObjects(descItemQty)

		' If no fixed quantity specified
		If IsEmpty(partQuantity) Then
			' Get current item quantity
			oldQuantity = CInt(itemQtyList(0).GetROProperty("value"))
			' Get  new quantity from 1 to 10
			Randomize
			newQuantity = Int((10 * Rnd) + 1)
			' If new matches old, turn it upside down
			If (newQuantity = oldQuantity) Then newQuantity = 11 - newQuantity
		Else
			' Use specified quantity
			newQuantity = partQuantity
		End If
		
		' Set item quantity to new value
		itemQtyList(0).Set CStr(newQuantity)

		' Release objects
		Set itemQtyList = Nothing
		Set descItemQty = Nothing
	End If

	cart_UpdateItemQuantity = doUpdate

End Function


'@Description Update quantity of each shopping cart item
'@Documentation Update quantity of each shopping cart item
'@Author sbabcoc
'@Date 23-MAR-2011
'@InParameter [in] partNumber Part number of product to be updated
'		Set to 'Null' to request update of every product
'@InParameter [in] updateFlags Option flags and fixed quantity specification
'		updt_VerifySubtotal: Set to request validation of cart subtotal
'		updt_FixedQuantity: Set to indicate specification of fixed quantity
'		updt_QuantityMask: Bitfield reserved for fixed quantity specification
'@ReturnValue 'True' if operation completed successfully; otherwise 'False'
Public Function cart_UpdateQuantity(partNumber, updateFlags)

	' Declarations
	Dim descItemWrap, itemWrapList, itemWrapCount
	Dim descUpdate, updateList
	Dim listIndex, oldQuantity, newQuantity
	Dim partQuantity, didUpdate

	' Initialize status
	didUpdate = False

	' If fixed quantity specified
	If (updateFlags And updt_FixedQuantity) Then
		' Extract fixed quantity from update flags
		partQuantity = (updateFlags And updt_QuantityMask)
		' Limit quantity to defined maximum value
		If (partQuantity > 9999) Then partQuantity = 9999
	End If

	With Browser("Checkout").Page("REI.com: Shopping Basket")
		' Get description of shopping cart item wrap objects
		Set descItemWrap = .WebElement("ItemWrap").GetTOProperties
		' Get list of item wrap objects
		Set itemWrapList = .ChildObjects(descItemWrap)
	End With

	' Get count of item wrap objects
	itemWrapCount = itemWrapList.Count
	' If item wrap objects were found
	If (itemWrapCount > 0) Then
		' Iterate over item wrap objects
		For listIndex = 0 to (itemWrapCount - 1)
			' If update indicated for this product, revise item quantity
			didUpdate = didUpdate Or cart_UpdateItemQuantity(itemWrapList(listIndex), partNumber, partQuantity)
		Next
	End If

	' If any item changed
	If (didUpdate) Then
		With Browser("Checkout").Page("REI.com: Shopping Basket")
			' Click on the first update button
			.WebButton("Update").Click
			' Synchronize with the browser
			.Sync
		End With
	End If

	' If subtotal validation requested
	If (updateFlags And updt_VerifySubtotal) Then
		' Validate cart subtotal
		isCorrect = cart_VerifySubtotal()
	Else
		isCorrect = True
	End If

	cart_UpdateQuantity = isCorrect

	' Release objects
	Set descItemWrap = Nothing
	Set itemWrapList = Nothing

End Function


'@Description Verify that computed subtotal matches displayed subtotal
'@Documentation Verify that computed subtotal matches displayed subtotal
'@Author sbabcoc
'@Date 02-JUN-2011
'@ReturnValue 'True' if computed subtotal matches displayed subtotal; otherwise 'False'
Public Function cart_VerifySubtotal()

	' Declarations
	Dim isCorrect
	Dim refSubtotal
	Dim cart_items, this_item
	Dim item_count, item_index
	Dim regEx, matchList
	Dim displayedSubtotal
	Dim computedSubtotal

	' Get reference to shopping cart subtotal element
	Set refSubtotal = Browser("Checkout").Page("REI.com: Shopping Basket").WebElement("Subtotal")

	' Get current cart inventory
	cart_items = cart_GetInventory()
	' Get count of items
	item_count = safeUBound(cart_items)
	' If cart contains items
	If (item_count >= 0) Then
		' If  subtotal element is present
		If chkVerifyExistence(refSubtotal, "Cart Subtotal", EXPECT_EXISTS, "cart_VerifySubtotal EXISTS") Then
			' Allocate RegExp object
			Set regEx = New RegExp
			' Define pattern to extract subtotal
			regEx.Pattern = "subtotal: \$((\d{1,3}(\,\d{3})*|(\d+))\.\d{2})"
			regEx.IgnoreCase = True
	
			' Apply pattern to described shopping cart subtotal
			Set matchList = regEx.Execute(refSubtotal.GetROProperty("innertext"))
			' Extract displayed subtotal from description
			displayedSubtotal = CCur(matchList(0).SubMatches(0))

			' Init computed value
			computedSubtotal = 0
			' Iterate over shopping cart items
			For item_index = 0 to item_count
				' Get current item 
				Set this_item = cart_items(item_index)
				' Add extended item cost to computed subtotal
				computedSubtotal = computedSubtotal + (this_item.Item(key_ProdCost) * this_item.Item(key_ProdQty))
			Next

			' Verify that computed subtotal matches displayed subtotal
			isCorrect = chkVerifyParity(computedSubtotal, displayedSubtotal, CMP_EQUAL, "Displayed Subtotal Consistency", "Computed subtotal")

			' Release objects
			Set this_item = Nothing
			Set matchList = Nothing
			Set regEx = Nothing
		' Otherwise (subtotal element is absent)
		Else
			' Indicate failure
			isCorrect = False
		End If
	' Otherwise (cart is empty)
	Else
		' Verify that subtotal element is absent
		isCorrect = chkVerifyExistence(refSubtotal, "Cart Subtotal", EXPECT_ABSENT, "cart_VerifySubtotal EXISTS")
	End If

	cart_VerifySubtotal = isCorrect

	' Release objects
	Erase cart_items
	Set refSubtotal = Nothing

End Function


'@Description Verify that product links are internally consistent
'@Documentation Verify that product links are internally consistent
'@Author sbabcoc
'@Date 02-JUN-2011
Public Sub cart_VerifyProdLinks()

	' Declarations
	Dim isCorrect
	Dim cart_items, this_item
	Dim item_count, item_index
	Dim regEx
	Dim prodNum, prodName
	Dim linkTail

	' Get current cart inventory
	cart_items = cart_GetInventory()
	' Get count of items
	item_count = safeUBound(cart_items)
	' If cart contains items
	If (item_count >= 0) Then
		' Allocate RegExp object
		Set regEx = New RegExp
		regEx.Global = True
		
		' Iterate over shopping cart items
		For item_index = 0 to item_count
			' Get current item 
			Set this_item = cart_items(item_index)
			
			' Get product number
			prodNum = this_item.Item(key_ProdNum)
			' Only retain first 6 characters
			prodNum = Left(prodNum, 6)
			
			' Get product name
			prodName = this_item.Item(key_ProdName)
			' Match non-alphanumerics
			regEx.Pattern = "[^-A-Za-z0-9 ]"
			' Remove all non-alphanumerics
			prodName = regEx.Replace(prodName, "")
			' Match spaces
			regEx.Pattern = "( |-)+"
			' Replace spaces with hyphen
			prodName = regEx.Replace(prodName, "-")
			' Convert to lowercase
			prodName = LCase(prodName)

			' Synthesize link tail from munged number and name
			linkTail = ".rei.com/product/" & prodNum & "/" & prodName
			
			' Verify internal consistency of text link
			chkVerifyParity this_item.Item(key_ProdLink), linkTail, STR_TAIL, "Synthesized Product Link", "Actual product link"

			' Verify consistency of text and image links
			chkVerifyParity this_item.Item(key_ProdPict), this_item.Item(key_ProdLink), STR_EQUAL, "Consistent With Text Link", "Product image link"
		Next

		' Release objects
		Set this_item = Nothing
		Set regEx = Nothing
	End If

	' Release objects
	Erase cart_items

End Sub


'@Description Verify addition of specified quantity of specified product to shopping cart
'@Documentation Verify addition of <prodQty> units of product <prodNum> to shopping cart
'@Author sbabcoc
'@Date 26-MAY-2011
'@Libraries Global, Verifications
'@InParameter [in] prodNum Product number of newly-added item
'@InParameter [in] prodQty Specified quantity of newly-added item
'@ReturnValue 'True' if operation completed successfully; otherwise 'False'
Public Function cart_VerifyAddition(prodNum, prodQty)

	' Get current cart inventory
	cart_items = cart_GetInventory()
	' Get count of items
	item_count = safeUBound(cart_items)
	' If items exist
	If (item_count >= 0) Then
		' Get first item
		Set first_item = cart_items(0)
		
		' Confirm that first item number is as specified
		If chkVerifyParity(first_item.Item(key_ProdNum), prodNum, STR_HEAD, "SKU check", "SKU of first item") Then
			' Confirm that first item quantity is as specified
			isCorrect = chkVerifyParity(first_item.Item(key_ProdQty), prodQty, CMP_EQUAL, "Quantity check", "Quantity of first item")
		Else
			isCorrect = False
		End If
	
		' Release objects
		Set first_item = Nothing
	Else
		isCorrect = False
		Reporter.ReportEvent micFail, "cart_VerifyAddition", "Shopping cart is empty"
	End If

	cart_VerifyAddition = isCorrect

	cart_UpdateSplashFlags

	' Release objects
	Erase cart_items
	
End Function

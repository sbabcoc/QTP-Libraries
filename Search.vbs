Option Explicit

' FUTURE SCENARIOS
' * Brand Search
' * Acronym Search
' * Keyword Search
' * Sort Order
' * Category
' * Refinements
' * Result Set Tally
' * Sale & Clearance
' * More/Less Filters
' * Result Orphans

	'Orphan Rules -  Outlet page may or may not show 4 extra products on Page 1. Item has to exceed items per page +  8 to show a next page
	'2 criterias to see we have the correct # of items in the page given the items per page
	'	1) Product count needs to be less than or equal to Items Per page
	'	2) Orphan Rule: has to exceed Items per page + 8 to show another page

'@Description Submit the specified search term, evaluating landing page with specified description
'@Documentation Submit <searchTerm>, evaluating landing page with <descTarget>
'@Author sbabcoc
'@Date 11-JUL-2011
'@InParameter [in] searchTerm Search term to be submitted
'@InParameter [in] descTarget Description of expected landing page
'@ReturnValue If search lands on the expected page, the URL of the page; otherwise 'Null'
Public Function search_SubmitQuery(searchTerm, descTarget)

	' Declarations
	Dim refObject
	Dim chk_href

	' Initialize result
	chk_href = Null

	' Get reference to search field
	Set refObject = Browser("Common").Page("REI Header").WebEdit("Search REI")
	
	' If search field doesn't exist
	If Not (refObject.Exist(0)) Then
		' Get reference to REI logo image
		Set refObject = Browser("Common").Page("REI Header").Image("REI.com")
		' Click REI logo
		followLink(refObject)
		' Get reference to search field
		Set refObject = Browser("Common").Page("REI Header").WebEdit("Search REI")
	End If
	
	' Set content of search field
	refObject.Set searchTerm
	
	' Set reference and target description for 'GO' button
	Set refObject = Browser("Common").Page("REI Header").WebButton("GO")

	' Verify target of 'GO' button
	chk_href = chkVerifyLinkTarget(refObject, descTarget)

	search_SubmitQuery = chk_href

	' Release objects
	Set refObject = Nothing

End Function


'@Description Get search result match count
'@Documentation Get search result match count
'@Author sbabcoc
'@Date 05-MAY-2011
'@Repositories Search
'@ReturnValue The search result match count
Public Function search_GetResultCount()

	' Declarations
	Dim regEx, matchList
	Dim resultSummary
	
	' Allocate RegExp object
	Set regEx = New RegExp

	' Get search result summary
	resultSummary = Browser("Search").Page("Page").WebElement("Summary").GetROProperty("innertext")
	' Define regex pattern
	regEx.Pattern = "\d+"
	' Apply pattern to search result summary
	Set matchList = regEx.Execute(resultSummary)
	' RESULT: Match count
	search_GetResultCount = CInt(matchList(0))

	' Release objects
	Set regEx = Nothing
	Set matchList = Nothing

End Function


'@Description Extract value or range from product price object
'@Documentation Extract value or range from <priceObj>
'@Author sbabcoc
'@Date 05-MAY-2011
'@InParameter [in] priceObj A reference to a product price object
'@ReturnValue The value/range for the specified price object
Public Function search_ExtractPrice(ByRef priceObj)

	' Declarations
	Dim regEx, matchList
	Dim priceLower, priceUpper
	
	' Allocate RegExp object
	Set regEx = New RegExp

	' Set match pattern
	regEx.Pattern = ".*?(\$[0-9]+\.[0-9]{2})( - (\$[0-9]+\.[0-9]{2}))?(.*)"
	' Apply pattern to extract product price(s)
	Set matchList = regEx.Execute(priceObj.GetROProperty("innertext"))
	' Extract product lower (or only) price
	priceLower = CCur(matchList(0).SubMatches(0))
	' If product upper price was captured
	If Len(matchList(0).Submatches(2)) Then
		' Extract product upper price
		priceUpper = CCur(matchList(0).SubMatches(2))
		' RESULT: Price range
		search_ExtractPrice = Array(priceLower, priceUpper)
	Else
		' RESULT: Extracted price
		search_ExtractPrice = priceLower
	End If

	' Release objects
	Set regEx = Nothing
	Set matchList = Nothing
	
End Function


'@Description Extract value from product rating object
'@Documentation Extract value from <ratingObj>
'@Author sbabcoc
'@Date 05-MAY-2011
'@InParameter [in] ratingObj A reference to a product rating object
'@ReturnValue The value for the specified rating object
Public Function search_ExtractRating(ByRef ratingObj)

	Dim regEx, matchList
	Dim prodRating
	
	' Allocate RegExp object
	Set regEx = New RegExp

	' Set match pattern
	regEx.Pattern = ".*/img\?url=images%2Fstars([0-5])(point5)?\.gif.*"
	' Apply pattern to extract rating
	Set matchList = regEx.Execute(ratingObj.GetROProperty("src"))
	' Get rating whole-number value
	prodRating = CDbl(matchList(0).Submatches(0))
	' If rating includes semi-star, add half a point
	If Len(matchList(0).Submatches(1)) Then prodRating = prodRating + 0.5

	' RESULT: Extracted product rating
	search_ExtractRating = prodRating

	' Release objects
	Set regEx = Nothing
	Set matchList = Nothing
	
End Function


'@Description Extract value from product review count object
'@Documentation Extract value from <reviewCountObj>
'@Author sbabcoc
'@Date 05-MAY-2011
'@Repositories Search
'@InParameter [in] reviewCountObj A reference to a product review count object
'@ReturnValue The value for the specified review count object
Public Function search_ExtractReviewCount(ByRef reviewCountObj)

	Dim regEx, matchList
	Dim descProdRevu
	
	' Allocate RegExp object
	Set regEx = New RegExp

	With Browser("Search").Page("Page").WebElement("ProdList").WebElement("Product").WebElement("ProdInfo")
		Set descProdRevu = .WebElement("ProdRevu").GetTOProperties
	End With

	' Set match pattern
	regEx.Pattern = descProdRevu("innertext").Value
	' Apply pattern to extract review count
	Set matchList = regEx.Execute(reviewCountObj.GetROProperty("innertext"))
	' RESULT: Product review count
	search_ExtractReviewCount  = CInt(matchList(0).Submatches(0))

	' Release objects
	Set regEx = Nothing
	Set descProdRevu = Nothing
	Set matchList = Nothing
	
End Function


'@Description Extract detail dictionary from search result item
'@Documentation Extract detail dictionary from <productObj>
'@Author sbabcoc
'@Date 22-MAR-2011
'@Libraries Global
'@Repositories Search
'@InParameter [in] productObj A reference to a search result item object
'@ReturnValue The detail dictionary for the specified search result item
Public Function search_ExtractItemDetails(ByRef productObj)

   ' Declarations
	Dim detailsDict
	Dim descElement
	Dim descProdInfo, descProdLink, descProdPrice
	Dim descSalePrice, descProdRating, descProdRevu
	Dim prodInfoObj
	Dim elementList
	Dim thisElement
	Dim listIndex
	Dim prodRating

	' Create a product detail dictionary
	Set detailsDict = CreateObject("Scripting.Dictionary")

	' Define generic web element  description
	Set descElement = Description.Create()
	descElement("micclass").Value = "WebElement"
	
	With Browser("Search").Page("Page").WebElement("ProdList").WebElement("Product").WebElement("ProdInfo")
		' Load product detail object descriptions
		Set descProdInfo = .GetTOProperties
		Set descProdLink = .Link("ProdLink").GetTOProperties
		Set descProdPrice = .WebElement("ProdPrice").GetTOProperties
		Set descSalePrice = .WebElement("SalePrice").GetTOProperties
		Set descProdRating = .Image("ProdRating").GetTOProperties
		Set descProdRevu = .WebElement("ProdRevu").GetTOProperties
	End With

	' Get  product information object
	Set prodInfoObj = productObj.ChildObjects(descProdInfo)(0)

	' Get list of product detail objects
	Set elementList = prodInfoObj.ChildObjects(descElement)
	' Iterate over product detail objects
	For listIndex = 0 To (elementList.Count - 1)
		' Get current product detail object
		Set thisElement = elementList(listIndex)

		' If this element contains text
		If Len(thisElement.GetROProperty("innertext")) Then
			' If this is the product link
			If (isDescribedObject(thisElement, descProdLink)) Then
				' Get product link
				productLink = thisElement.GetROProperty("href")
				' Add product link to detail dictionary
				detailsDict.Add key_ProdLink, productLink
				' Add product name to detail dictionary
				detailsDict.Add key_ProdName, thisElement.GetROProperty("innertext")
				' Add product number to detail dictionary
				detailsDict.Add key_ProdNum, product_GetSKUFromURL(productLink)
			' Otherwise, if this is the sale price
			ElseIf (isDescribedObject(thisElement, descSalePrice)) Then
				' Add cost to detail dictionary
				detailsDict.Add key_ProdCost, search_ExtractPrice(thisElement)
			' Otherwise, if this is the product price
			ElseIf (isDescribedObject(thisElement, descProdPrice)) Then
				' Add cost to detail dictionary
				detailsDict.Add key_ProdCost, search_ExtractPrice(thisElement)
			' Otherwise, if this is the review count
			ElseIf (isDescribedObject(thisElement, descProdRevu)) Then
				' Add review count to detail dictionary
				detailsDict.Add key_ProdRevu, search_ExtractReviewCount(thisElement)
			End If
		' Otherwise, if this is the product rating
		ElseIf (isDescribedObject(thisElement, descProdRating)) Then
			' Add product rating to detail dictionary
			detailsDict.Add key_ProdRate, search_ExtractRating(thisElement)
		End If
	Next

	' RESULT: Item detail dictionary
	Set search_ExtractItemDetails = detailsDict

	' Release objects
	Set detailsDict = Nothing
	Set descElement = Nothing
	Set descProdInfo = Nothing
	Set descProdLink = Nothing
	Set descProdPrice = Nothing
	Set descSalePrice = Nothing
	Set descProdRating = Nothing
	Set descProdRevu = Nothing
	Set prodInfoObj = Nothing
	Set elementList = Nothing
	Set thisElement = Nothing
	
End Function


'@Description Get detailed list of  search result contents
'@Documentation Get detailed list of  search result contents
'@Author sbabcoc
'@Date 23-MAR-2011
'@Libraries Global
'@Repositories Search
'@ReturnValue An array of  product detail dictionaries
Public Function search_GetResults()

   ' Declarations
   Dim descProdDetail
   Dim prodDetailList, prodDetailCount, listIndex
   Dim results()

	With Browser("Search").Page("Page").WebElement("ProdList")
		' Load description of search result product detail objects
		Set descProdDetail = .WebElement("Product").GetTOProperties
		' Get list of product detail objects
		Set prodDetailList = .ChildObjects(descProdDetail)
	End With
	
	' Get count of product detail objects
	prodDetailCount = prodDetailList.Count
	' If product detail objects were found
	If (prodDetailCount > 0) Then
		' Allocate space for detail dictionaries
		ReDim results(prodDetailCount - 1)
		' Iterate over product detail objects
		For listIndex = 0 to (prodDetailCount - 1)
			' Store details of current product in the results array
			Set results(listIndex) = search_ExtractItemDetails(prodDetailList(listIndex))
		Next
	End If

	' RESULT: Detailed results
	search_GetResults = results

	' Release objects
	Set descProdDetail = Nothing
	Set prodDetailList = Nothing
	
End Function


'@Description Get list of product page URLs from search results
'@Documentation Get list of product page URLs from search results
'@Author sbabcoc
'@Date 23-MAR-2011
'@Repositories Search
'@ReturnValue An array of  product page URLs
Public Function search_GetResultLinks()

	' Declarations
	Dim resultSummary, resultCount
	Dim descProdLink, prodLinkList
	Dim linkCount, linkIndex
	Dim resultList(), listIndex
	Dim pageCount

	' Get search result match count
	resultCount = search_GetResultCount()

	' Set result array size
	ReDim resultList(resultCount - 1)

	' Load description of search result product page links
	Set descProdLink = Browser("Search").Page("Page").WebElement("ProdList").WebElement("Product").WebElement("ProdInfo").Link("ProdLink").GetTOProperties

	' Init index
	listIndex = 0
	' Init page count
	pageCount = 0

	' Iterate
	Do
		' Get every product link on the current search results page
		Set prodLinkList = Browser("Search").Page("Page").ChildObjects(descProdLink)

		' Get current page link count
		linkCount = prodLinkList.Count
		' Iterate over page links
		For linkIndex = 0 To (linkCount - 1)
			' Save URL of current  link in output array
			resultList(listIndex) = prodLinkList(linkIndex).GetROProperty("href")
			' Increment array index
			listIndex = listIndex + 1
		Next

		' If done processing results, exit loop
		If (listIndex = resultCount) Then Exit Do

		' Multiple pages
		multipage = True
		' Click the "next page" search results navigation link
		Browser("Search").Page("Page").WebElement("PageNav").Link(">>").Click
		' Allow page to finish loading
		Browser("Search").Page("Page").Sync
		' Increment page count
		pageCount = pageCount + 1
	Loop

	While (pageCount > 0)
		' Navigate to prior page
		Browser("Search").Back
		' Allow page to finish loading
		Browser("Search").Sync
		' Decrement page count
		pageCount = pageCount - 1
	Wend

	' RESULT: Array of product page URLs
	search_GetResultLinks = resultList
	
	' Release objects
	Set prodLinkList = Nothing
	Set descProdLink = Nothing

End Function

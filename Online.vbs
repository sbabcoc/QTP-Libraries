Option Explicit

' Initialize breadcrumb array
Dim breadcrumbs()

' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
'                                    Online Function Library
'                                Recreational Equipment Inc.
' -------------------------------------------------------------------------------------
' ------------------------------- ATTENTION --------------------------------
' The contents of this file is the solely property of REI. 
' Unauthorized use of this resource is expressly forbidden.
' -------------------------------------------------------------------------------------

'@Description Display a message in the page Search field
'@Documentation Display <sMsg> in the page Search field
'@InParameter [in] sMsg The message to display;
' 		Specify empty string to display current action
'		Specify non-string (e.g. - 'Null") to clear field
Public Sub cmnSetVisualCueInSearchField(sMsg)

	' If a string was specified
	If (VarType(sMsg) = vbString) Then
		' If string has content
		If (Len(sMsg) > 0) Then
			' Display specified string in search field
			Browser("Common").Page("REI Header").WebEdit("Search REI").Set sMsg
		' Otherwise (empty string)
		Else
			' Display current action in search field
			Browser("Common").Page("REI Header").WebEdit("Search REI").Set "Running '" & Environment("ActionName") & "' action"
		End If
	' Otherwise (non-string)
	Else
		' Clear content of search field
		Browser("Common").Page("REI Header").WebEdit("Search REI").Set ""
	End If

End Sub


'@Description Add the current page to the breadcrumb trail and click the specified link
'@Documentation Add the current page to the breadcrumb trail and click <objLink>
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Global
'@Repositories Web
'@InParameter [in] objLink, reference, A reference to a Link runtime object
Public Sub followLink(objLink)

	' Declarations
	Dim crumbCount

	' If specified link exists
	If (objLink.Exist(0)) Then
		' Get new breadcrumb count
		crumbCount = safeUBound(breadcrumbs) + 1
		' Allocate new breadcrumb
		ReDim Preserve breadcrumbs(crumbCount)
		' Push current page URL
		breadcrumbs(crumbCount) = Browser("Browser").Page("Page").GetROProperty("url")
		' Click specified link
		objLink.Click
		' Allow page to finish loading
		Browser("Browser").Page("Page").Sync
	Else
		' Report this unexpected condition
		Reporter.ReportEvent micFail, "followLink", "Link [" & objLink.ToString & "] doesn't exist"
		' Exit the test
		ExitTest ("Link doesn't exist")
	End If

End Sub


'@Description Retrace last navigation step, discarding the coresponding breadcrumb
'@Documentation  Retrace last navigation step, discarding the coresponding breadcrumb
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Global
'@Repositories Web
Public Sub retraceStep()

	' Declarations
	Dim crumbBound
	Dim expectURL
	Dim actualURL

	' Get current breadcrumb count
	crumbBound = safeUBound(breadcrumbs)
	' If breadcrumbs exist
	If (crumbBound >= 0) Then
		' Navigate to prior page
		Browser("Browser").Back
		' Allow page to finish loading
		Browser("Browser").Page("Page").Sync
		' Get current page URL
		actualURL = Browser("Browser").Page("Page").GetROProperty("url")

		' Get expected URL
		expectURL = breadcrumbs(crumbBound)
		
		' If this is the last crumb
		If (crumbBound = 0) Then
			' Deallocate array
			Erase breadcrumbs
		' Otherwise (crumbs remain)
		Else
			' Discard the last breadcrumb
			ReDim Preserve breadcrumbs(crumbBound - 1)
		End If
		
		' If actual URL isn't expected URL
		If (actualURL <> expectURL) Then
			' Report this unexpected condition
			Reporter.ReportEvent micWarning, "retraceStep", _
				"Retrace page doesn't match breadcrumb" & vbCR & _
				"EXPECT: " & expectURL & vbCR & "ACTUAL: " & actualURL
		End If
	Else
		' Report this unexpected condition
		Reporter.ReportEvent micWarning, "retraceStep", "Breadcrumb array is empty"
		' Navigate to prior page
		Browser("Browser").Back
		' Allow page to finish loading
		Browser("Browser").Page("Page").Sync
	End If

End Sub


'@Description Return the URL of the last breadcrumb
'@Documentation Return the URL of the last breadcrumb
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Global
'@ReturnValue If crumbs exist, the URL of the last crumb; otherwise, an empty string
Public Function getLastCrumb()

	' Declarations
	Dim crumbBound

	' Get current breadcrumb count
	crumbBound = safeUBound(breadcrumbs)
	' If breadcrumbs exist
	If (crumbBound >= 0) Then
		' RESULT: Last crumb
		getLastCrumb = breadcrumbs(crumbBound)
	Else
		' Report this unexpected condition
		Reporter.ReportEvent micWarning, "getLastCrumb", "Breadcrumb array is empty"
		' RESULT: Empty string
		getLastCrumb = ""
	End If

End Function


'@Description Determine if the current page is a production site page
'@Documentation Determine if the current page is a production site page
'@Author sbabcoc
'@Date 08-JUL-2011
'@ReturnValue 'True' if the current page is a production site page; otherwise, 'False'
Public Function isProduction()

   ' Declarations
   Dim regEx
   Dim startURL

   ' Allocate RegExp object
   Set regEx = New RegExp
   ' Define pattern to match production domain (either deployed or pre-release)
   regEx.Pattern = "^https?:\/\/((m\.rei\.com|demomobile\.usablenet\.com)\/mt\/)?www\.rei\.com\b"

	' Get URL for site under test
	startURL = Environment("envStartURL")
	isProduction = regEx.Test(startURL)

	' Release objects
	Set regEx = Nothing


End Function


'@Description Determine if the body of the current page has the specified class
'@Documentation Determine if the body of the current page has class <theClass>
'@Author sbabcoc
'@Date 08-JUL-2011
'@Repositories Common
'@ReturnValue 'True' if the body of the current page has the specified class; otherwise, 'False'
Public Function hasBodyClass(theClass)

   ' Declarations
   Dim regEx
   Dim classStr

   ' Allocate RegExp object
   Set regEx = New RegExp
   ' Define pattern to match spec'd class
   regEx.Pattern = "\b" & theClass & "\b"

	classStr = Browser("Browser").page("Page").WebElement("Body").GetROProperty("class")
	hasBodyClass = regEx.Test(classStr)

	' Release objects
	Set regEx = Nothing

End Function


'@Description Determine if the current page is an REI-OUTLET page
'@Documentation Determine if the current page is an REI-OUTLET page
'@Author sbabcoc
'@Date 07-JUL-2011
'@Repositories Common
'@ReturnValue 'True' if the current page is an REI-OUTLET page; otherwise, 'False'
Public Function isREI_Outlet()

	isREI_Outlet = hasBodyClass("outlet")

End Function


'@Description Determine if the current page is a checkout page
'@Documentation Determine if the current page is a checkout page
'@Author sbabcoc
'@Date 07-JUL-2011
'@Repositories Common
'@ReturnValue 'True' if the current page is a checkout page; otherwise, 'False'
Public Function isCheckout()

	isCheckout = hasBodyClass("checkout")

End Function


'@Description Determine if the current page is the Home page
'@Documentation Determine if the current page is the Home page
'@Author sbabcoc
'@Date 04-APR-2011
'@Repositories Common
'@ReturnValue 'True' if the current page is the Home page; otherwise, 'False'
Public Function isHomePage()

	isHomePage = Browser("PageStub").Page("REI.com Home").Exist(0)

End Function


'@Description Determine if a user is logged in to the site
'@Documentation Determine if a user is logged in to the site
'@Author sbabcoc
'@Date 04-APR-2011
'@Repositories Common
'@ReturnValue 'True' is a user is logged in to the site; otherwise; 'False'
Public Function isLoggedIn()

	isLoggedIn = Browser("Common").Page("REI Header").Link("Your Account").Exist(0)

End Function


'@Description Verify common elements of page headers
'@Documentation Verify common elements of page headers
'@Author sbabcoc
'@Date 01-APR-2011
'@Libraries Verifications, DataTable
'@Repositories Web, Common
'@ReturnValue 'True' if  common page header elements are correct; otherwise, 'False'
Public Function chkVerifyPageHeader()

	' Declarations
	Dim resultVal
	Dim isCorrect
	Dim refObject
	Dim descTarget

	' Initialize results
	resultVal = 0
	isCorrect = True

	' Verify the REI Logo image link
	Set refObject = Browser("Common").Page("REI Header").Image("REI.com")
	' If this is an REI-OUTLET page
	If isREI_Outlet() Then
		' Cache REI-OUTLET links
		link_Logo = link_Logo2
		link_LogIn = link_LogIn2
		link_Cart = link_Cart2
		' Load description of the REI-OUTLET Store home page
		Set descTarget = Browser("PageStub").Page("REI-OUTLET").GetTOProperties
	' Otherwise (an REI Online page)
	Else
		' Cache REI Online links
		link_Logo = link_Logo1
		link_LogIn = link_LogIn1
		link_Cart = link_Cart1
		' Load description of the REI Online Store home page
		Set descTarget = Browser("PageStub").Page("REI.com Store").GetTOProperties
'		' Load description of the REI .com home page  (QuickBase ticket #158)
'		Set descTarget = Browser("PageStub").Page("REI.com Home").GetTOProperties
	End If

	If chkVerifyPageLink(refObject, descTarget, link_Logo) Then
		resultVal = resultVal Or head_ImageLink
	Else
		isCorrect = False
	End If

	' Verify the Home link
	Set refObject = Browser("Common").Page("REI Header").Link("Home")
	Set descTarget = Browser("PageStub").Page("REI.com Home").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_Home) Then
		resultVal = resultVal Or head_HomeLink
	Else
		isCorrect = False
	End If

	' Verify the Stores link
	Set refObject = Browser("Common").Page("REI Header").Link("Stores")
	Set descTarget = Browser("PageStub").Page("REI Store Locator").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_Stores) Then
		resultVal = resultVal Or head_StoresLink
	Else
		isCorrect = False
	End If

	' Verify the Cart link
	Set refObject = Browser("Common").Page("REI Header").Link("Cart")
	Set descTarget = Browser("PageStub").Page("REI.com: Shopping Basket").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_Cart) Then
		resultVal = resultVal Or head_CartLink
	Else
		isCorrect = False
	End If

	' If not a checkout page
	If Not isCheckout() Then
		' Verify existence of Search group
		Set refObject = Browser("Common").Page("REI Header").WebElement("SearchGroup")
		If chkVerifyExistence(refObject, "Search Group", EXPECT_EXISTS, "chkVerifyPageHeader") Then
			resultVal = resultVal Or head_SearchGroup
		Else
			isCorrect = False
		End If

		' Verify existence of Search field
		Set refObject = Browser("Common").Page("REI Header").WebEdit("Search REI")
		If chkVerifyExistence(refObject, "Search", EXPECT_EXISTS, "chkVerifyPageHeader") Then
			resultVal = resultVal Or head_SearchField
		Else
			isCorrect = False
		End If

		' Verify existence of GO button
		Set refObject = Browser("Common").Page("REI Header").WebButton("GO")
		If chkVerifyExistence(refObject, "Search", EXPECT_EXISTS, "chkVerifyPageHeader") Then
			resultVal = resultVal Or head_GoButton
		Else
			isCorrect = False
		End If
	End If

	' If logged in
	If (isLoggedIn()) Then
		' Verify existence of  Log Out link
		Set refObject = Browser("Common").Page("REI Header").Link("Log Out")
		If chkVerifyExistence(refObject, "Log Out", EXPECT_EXISTS, "chkVerifyPageHeader") Then
			resultVal = resultVal Or head_LogOutLink
		Else
			isCorrect = False
		End If

		' Verify the Your Account link
		Set refObject = Browser("Common").Page("REI Header").Link("Your Account")
		Set descTarget = Browser("PageStub").Page("Your Account").GetTOProperties

		If chkVerifyPageLink(refObject, descTarget, link_YourAcct) Then
			resultVal = resultVal Or head_YourAcctLink
		Else
			isCorrect = False
		End If
	Else
		' Verify the Log In link
		Set refObject = Browser("Common").Page("REI Header").Link("Log In")
		Set descTarget = Browser("PageStub").Page("REI.com: Login").GetTOProperties

		If chkVerifyPageLink(refObject, descTarget, link_LogIn) Then
			resultVal = resultVal Or head_LogInLink
		Else
			isCorrect = False
		End If
	End If

	If (isCorrect) Then resultVal = resultVal Or head_VerifyOK
	chkVerifyPageHeader = resultVal

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing
	
End Function


'@Description Verify common elements of page footers
'@Documentation Verify common elements of page footers
'@Author sbabcoc
'@Date 01-APR-2011
'@Libraries Verifications, DataTable
'@Repositories Web, Common
'@ReturnValue 'True' if  common page footer elements are correct; otherwise, 'False'
Public Function chkVerifyPageFooter()

	' Declarations
	Dim resultVal
	Dim isCorrect
	Dim refObject
	Dim descTarget

	' Initialize results
	resultVal = 0
	isCorrect = True

	' Verify existence of Call REI element
	Set refObject = Browser("Common").Page("REI Footer").WebElement("Call REI")
	If chkVerifyExistence(refObject, "Call REI", EXPECT_EXISTS, "chkVerifyPageFooter") Then
		resultVal = resultVal Or foot_CallREIText
	Else
		isCorrect = False
	End If

	' Verify the Help link
	Set refObject = Browser("Common").Page("REI Footer").Link("Help")
	Set descTarget = Browser("PageStub").Page("REI Help Section").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_Help) Then
		resultVal = resultVal Or foot_HelpLink
	Else
		isCorrect = False
	End If

	' Verify the Privacy link
	Set refObject = Browser("Common").Page("REI Footer").Link("Privacy")
	Set descTarget = Browser("PageStub").Page("REI Privacy Policy").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_Privacy) Then
		resultVal = resultVal Or foot_PrivacyLink
	Else
		isCorrect = False
	End If

	' Verify the Go to REI.com link
	Set refObject = Browser("Common").Page("REI Footer").Link("Go to REI.com")
	Set descTarget = Browser("PageStub").Page("REI Online").GetTOProperties

	If chkVerifyPageLink(refObject, descTarget, link_GoToREI) Then
		resultVal = resultVal Or foot_GoToREILink
	Else
		isCorrect = False
	End If

	' Verify the Feedback link
	If chkVerifyFeedbackLink() Then
		resultVal = resultVal Or foot_FeedbackLink
	Else
		isCorrect = False
	End If

	If (isCorrect) Then resultVal = resultVal Or foot_VerifyOK
	chkVerifyPageFooter = resultVal

	' Release objects
	Set refObject = Nothing
	Set descTarget = Nothing

End Function


'@Description Verify the page footer Feedback link
'@Documentation Verify the page footer Feedback link
'@Author sbabcoc
'@Date 01-APR-2011
'@Libraries Verifications, Web
'@Repositories Web
'@ReturnValue 'True' if  page footer Feedback link is correct; otherwise, 'False'
Public Function chkVerifyFeedbackLink()

	' Declarations
	Dim isCorrect
	Dim objLink
	Dim linkDict
	Dim thisURL
	Dim lastURL
	Dim thisKey
	Dim thisVal

	' Initialize result
	isCorrect = True

	' Get reference to Feedback link
	Set objLink = Browser("Common").Page("REI Footer").Link("Feedback")

	' If Feedback link exists
	If chkVerifyExistence(objLink, "Feedback", EXPECT_EXISTS, "chkVerifyFeedbackLink EXISTS") Then
		' Extract parameters dictionary from Feedback link
		Set linkDict = SplitURL(objLink.GetROProperty("href"))

		' Get expected referring page
		thisURL = Browser("Browser").Page("Page").GetROProperty("url")
		' Get expected previous page
		lastURL = getLastCrumb()

		' Report OpinionLab page URL
		Reporter.ReportEvent micInfo,"chkVerifyFeedbackLink", "OpinionLab URL = " & linkDict.Item(key_base_url)

		' If [referer] parameter exists
		If (linkDict.Exists(key_referer)) Then
			' Verify referring page URL
			isCorrect = isCorrect And chkVerifyParity(DemobilizeURL(linkDict.Item(key_referer)), DemobilizeURL(thisURL), STR_EQUAL, "chkVerifyFeedbackLink [referer]", "[referer] parameter")
		' Otherwise ([referer] parameter absent)
		Else
			' RESULT: Failure
			isCorrect = False
			Reporter.ReportEvent micFail, "chkVerifyFeedbackLink", "[referer] parameter is absent"
		End If

		' If [prev] parameter exists
		If (linkDict.Exists(key_prev)) Then
			' Verify previous page URL
			isCorrect = isCorrect And chkVerifyParity(DemobilizeURL(linkDict.Item(key_prev)), DemobilizeURL(lastURL), STR_EQUAL, "chkVerifyFeedbackLink [prev]", "[prev] parameter")
		' Otherwise ([prev] parameter absent)
		Else
			' RESULT: Failure
			isCorrect = False
			Reporter.ReportEvent micFail, "chkVerifyFeedbackLink", "[prev] parameter is absent"
		End If

		' If [time1] parameter exists
		If (linkDict.Exists(key_time1)) Then
			' Report value in human-readable format
			Reporter.ReportEvent micInfo, "chkVerifyFeedbackLink", "Timestamp [time1] = " & epoch2date(linkDict.Item(key_time1))
		' Otherwise ([time1] parameter absent)
		Else
			' RESULT: Failure
			isCorrect = False
			Reporter.ReportEvent micFail, "chkVerifyFeedbackLink", "[time1] parameter is absent"
		End If

		' If [time2] parameter exists
		If (linkDict.Exists(key_time2)) Then
			' Report value in human-readable format
			Reporter.ReportEvent micInfo, "chkVerifyFeedbackLink", "Timestamp [time2] = " & epoch2date(linkDict.Item(key_time2))
		' Otherwise ([time2] parameter absent)
		Else
			' RESULT: Failure
			isCorrect = False
			Reporter.ReportEvent micFail, "chkVerifyFeedbackLink", "[time2] parameter is absent"
		End If
	' Otherwise (Feedback link absent)
	Else
		' RESULT: Failure
		isCorrect = False
	End If
	
	chkVerifyFeedbackLink = isCorrect

	' Release objects
	Set objLink = Nothing
	Set linkDict = Nothing

End Function


'@Description Evaluate link URL against "pass" and "fail" URLs
'@Documentation Evaluate <lnkURL> against <passURL> and <failURL>
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Web
'@InParamater [in] linkURL, string, The actual URL of the link being compared
'@InParameter [in] passURL, string, The "passing" URL for the link
'@InParameter [in] failURL, string, The "failing" URL for the link
'@ResultValue Evaluation result: -1 = PASS; 0 = FAIL; 1 = no match
Public Function chkEvaluateLink(linkURL, passURL, failURL)

	' If actual URL matches "passing" URL
	If CompareLink(linkURL, passURL) Then
		' RESULT: Success
		chkEvaluateLink = EVAL_PASS
	' Otherwise, if actual URL matches "failing" URL
	ElseIf CompareLink(linkURL, failURL) Then
		' RESULT: Success
		chkEvaluateLink = EVAL_FAIL
	' Otherwise
	Else
		' RESULT: No match
		chkEvaluateLink = EVAL_NONE
	End If

End Function


'@Description Verify that the specified link targets the indicated page
'@Documentation Verify that <objLink> targets <descTarget>
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Global
'@Repositories Web
'@InParameter [in] objLink, reference, A reference to a Link runtime object
'@InParameter [in] descTarget, reference, A reference to a Description test object
'@ReturnValue If the link checks out, the 'href' property of the link; otherwise, 'Null'
Public Function chkVerifyLinkTarget(ByRef objLink, ByRef descTarget)

	' Declarations
	Dim link_href
	Dim parmValue

	' If specified link exists
	If (objLink.Exist(0)) Then
		' Get link object target href
		link_href = objLink.GetROProperty("href")
		' Click link
		followLink objLink
		
		' If landing page matches expectation
		If (isDescribedObject(Browser("Browser").Page("Page"), descTarget)) Then
			' If link target href is empty
			' NOTE: Occurs when "link" is a button
			If IsEmpty(link_href) Then
				' RESULT: Current page URL
				chkVerifyLinkTarget = Browser("Browser").Page("Page").GetROProperty("url")
			Else
				' RESULT: Link object target href
				chkVerifyLinkTarget = link_href
			End If
			' Report the success
			Reporter.ReportEvent micPass, "chkVerifyLinkTarget TARGET Passed", "Target of " & objLink.ToString & " is correct"
		Else
			' RESULT: Bad link target
			chkVerifyLinkTarget = Null
			' Report the failure
			Reporter.ReportEvent micFail, "chkVerifyLinkTarget TARGET Failed", "Target of " & objLink.ToString & " is incorrect"
		End If
	Else
		' RESULT: Link absent
		chkVerifyLinkTarget = Empty
		' Report the failure
		Reporter.ReportEvent micWarning, "chkVerifyLinkTarget", "Specified link object doesn't exist"
	End If

End Function


'@Description Using 'href' caching, verify that the specified link targets the indicated page
'@Documentation Using 'href' caching, verify that the specified link targets the indicated page
'@Author sbabcoc
'@Date 06-APR-2011
'@Libraries Verifications, DataTable
'@InParameter [in] objLink, reference, A reference to a Link runtime object
'@InParameter [in] descTarget, reference, A reference to a Description test object
'@InParameter [in] parmName, string, Name of parameter used for 'href' caching, or 'Null' to decline caching
'@ReturnValue 'True' if the link checks out; otherwise, 'False'
Public Function chkVerifyPageLink(ByRef objLink, ByRef descTarget, ByVal parmName)

	' Declarations
	Dim link_href
	Dim chk_href
	Dim href_Correct
	Dim href_Wrong

	' If specified link is present
	If chkVerifyExistence(objLink, objLink.ToString, EXPECT_EXISTS, "chkVerifyPageLink EXISTS") Then
		' If 'href' caching requested
		If Not IsNull(parmName) Then
			' Get cached link target correct href
			href_Correct = MyGetParameter(dtGlobalSheet, parmName, 1)
			' Get cached link target wrong href
			href_Wrong = MyGetParameter(dtGlobalSheet, parmName, 2)
		' Otherwise (caching declined)
		Else
			' Set default values
			href_Correct = Null
			href_Wrong = Null
		End If

		' Get link object target href
		link_href = objLink.GetROProperty("href")

		' Evaluate link against cached "passing" and "failing" URLs
		Select Case chkEvaluateLink(link_href, href_Correct, href_Wrong)

			' If link points to correct href 
			Case EVAL_PASS
				' RESULT: Success
				chkVerifyPageLink = True
				' Report the success
				Reporter.ReportEvent micPass, "chkVerifyPageLink TARGET Passed", "Target of " & objLink.ToString & " is correct"

			' If link points to wrong href
			Case EVAL_FAIL
				' RESULT: Failure
				chkVerifyPageLink = False
				' Report the failure
				Reporter.ReportEvent micFail, "chkVerifyPageLink TARGET Failed", "Target of " & objLink.ToString & " is incorrect"

			' If no match
			Case Else
				' Verify link target
				chk_href = chkVerifyLinkTarget(objLink, descTarget)
				' Return to previous page
				retraceStep

				' If landing page matches expectation
				If Not IsNull(chk_href) Then
					' RESULT: Success
					chkVerifyPageLink = True
					' If 'href' caching requested, cache link target correct href
					If Not IsNull(parmName) Then MySetParameter dtGlobalSheet, parmName, 1, link_href, flagAddParam
				Else
					' RESULT: Failure
					chkVerifyPageLink = False
					' If 'href' caching requested, cache link target wrong href
					If Not IsNull(parmName) Then MySetParameter dtGlobalSheet, parmName, 2, link_href, flagAddParam
				End If

		End Select
	' Otherwise (specified link is absent)
	Else
		' RESULT: Failure
		chkVerifyPageLink = False
	End If

End Function


'@Description Verify actual breadcrumbs against saved breadcrumbs
'@Documentation Verify actual breadcrumbs against saved breadcrumbs
'@Author sbabcoc
'@Date 07-APR-2011
'@ResultValue 'True' if verification succeeds; otherwise 'False'
Public Function chkVerifyBreadcrumbs()

   On Error Resume Next

	' Declarations
	Dim isCorrect
	Dim arrayCount
	Dim array_href
	Dim arrayIndex
	Dim descCrumb
	Dim crumbList
	Dim thisCrumb
	Dim link_text
	Dim link_href
	Dim crumbCount
	Dim crumbIndex
	Dim crumbStart

	' Initialize result
	isCorrect = True

	' Initialize array count
	arrayCount = 0
	' Get count of saved breadcrumbs
	arrayCount = UBound(breadcrumbs) + 1

	' Get description of page footer breadcrumb link
	Set descCrumb = Browser("Common").Page("REI Footer").Link("Breadcrumb").GetTOProperties
	' Get current list of page footer breadcrumb links
	Set crumbList = Browser("Browser").Page("Page").ChildObjects(descCrumb)
	' Get count of page footer breadcrumb links
	crumbCount = crumbList.Count

	' If actual crumbs outnumber saved crumbs
	If (crumbCount  > arrayCount) Then
		isCorrect = False
		crumbStart = crumbCount - arrayCount
		Reporter.ReportEvent micWarning, "chkVerifyBreadcrumbs", "Actual breadcrumbs outnumber saved breadcrumbs"
	Else
		crumbStart = 0
	End If

	arrayIndex = arrayCount - 1
	For crumbIndex = crumbStart To (crumbCount - 1)
		Set thisCrumb = crumbList(crumbIndex)
		link_text = thisCrumb.GetROProperty("text")
		link_href = NormalizeURL(thisCrumb.GetROProperty("href"))
		
		If (arrayIndex >= 0) Then
			array_href = NormalizeURL(breadcrumbs(arrayIndex))
			If CompareLink(link_href, array_href) Then
				Reporter.ReportEvent micInfo, "chkVerifyBreadcrumbs", "Breadcrumb [" & link_text & "] matches saved breadcrumb"
				' Decrement array index
				arrayIndex = arrayIndex - 1
			Else
				isCorrect = False
				Reporter.ReportEvent micWarning, "chkVerifyBreadcrumbs", _
					"Breadcrumb [" & link_text & "]  doesn't match saved breadcrumb" & vbCR & _
					"EXPECT: " & array_href & vbCR & "ACTUAL: " & link_href
				' NOTE: Assume links were clicked directly, not via the followLink() method. Don't decrement array index.
			End If
		Else
			isCorrect = False
			Reporter.ReportEvent micWarning, "chkVerifyBreadcrumbs", "Breadcrumb [" & link_text & "] found with no remaining saved breadcrumb"
		End If
	Next

	chkVerifyBreadcrumbs = isCorrect

End Function


'@Description Normalize the specified mobile site URL
'@Documentation Normalize <theURL>
'@Author sbabcoc
'@Date 07-APR-2011
'@InParameter [in] theURL, string, The mobile site URL to be normalized
'@ReturnValue The normalized form of the specified URL
Public Function NormalizeURL(ByVal theURL)

	' Declarations
	Dim regEx
	Dim newURL

	' Allocate RegExp object
	Set regEx = New RegExp

	newURL = theURL

	' Load pattern for "store" parm
	regEx.Pattern = "un_jtt_v_store=(y(es)?|rei)"
	' Normalize "store" parm
	newURL = regEx.Replace(newURL, "un_jtt_v_store=yes")

	' Load pattern for "storeId" parm
	regEx.Pattern = "storeId=800(0|1)"
	' Normalize "store" parm
	newURL = regEx.Replace(newURL, "storeId=800x")

	' Remove "home" parm
	newURL = url_RemoveParm(newURL, "un_jtt_v_home=yes")
	' Remove "redirect" parm
	newURL = url_RemoveParm(newURL, "un_jtt_redirect")
	' Remove "product" parm
	newURL = url_RemoveParm(newURL, "un_jtt_v_product=yes")

	' RESULT: Normalized URL
	NormalizeURL = newURL

	' Release RegExp
	Set regEx = Nothing
	
End Function


'@Description Extract the original URL from the specified mobile site URL
'@Documentation Exract the original URL from <theURL>
'@Author sbabcoc
'@Date 07-APR-2011
'@InParameter [in] theURL, string, The mobile site URL from which to extract the original URL
'@ReturnValue The original URL from which the specified mobile site URL was derived
Public Function DemobilizeURL(ByVal theURL)


	Dim theDict
	Dim offset
	Dim scheme
	Dim thePath
	Dim newURL
	Dim delimiter
	Dim thisKey
	Dim thisVal

	' Split specified URL into base path and parameters
	Set theDict = SplitURL(theURL)
	' Extract URL base path
	thePath = theDict.Item(key_base_url)
	' Remove base path parm
	theDict.Remove(key_base_url)
	' Locate end of scheme specifier
	offset = InStr(1, thePath, ":")
	' If scheme spec'd
	If (offset > 0) Then
		' Extract scheme from URL
		scheme = Left(thePath, offset)
		' Locate end of mobile site URL prefix
		offset = InStr(offset + 3, thePath, "/mt/")
		' If URL prefixed
		If (offset > 0)  Then
			' Assemble original URL path
			thePath = scheme & "//" & Mid(thePath, offset + 4)
		End If
	End If

	' Initialize new URL
	newURL = thePath
	' Set parm list delimiter
	delimiter = "?"

	' Iterate over parameters
	For Each thisKey In theDict.Keys
		' If this isn't a mobile-specific parameter
		If (InStr(1, thisKey, "un_") = 0) Then
			' Append parameter key
			newURL = newURL & delimiter & thisKey
			' Set parm item delimiter
			delimiter = "&"
			' Extract parameter value
			thisVal = theDict.Item(thisKey)
			' If this isn't the "no value" value
			If  (thisVal <> "&") Then
				' Append parameter value
				newURL = newURL & "=" & URLEncode(thisVal)
			End If
		End If
	Next

	' RESULT: Original URL
	DemobilizeURL = newURL

	' Release dictionary
	Set theDict = Nothing

End Function


'@Description Correlate state name to its abbreviation
'@Documentation Correlate <stateSpec> to its abbreviation/name
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] stateSpec, string, state name/abbreviation
'@ReturnValue The abbreviation/name corresponding to the specified state
Public Function correlateStateToAbbr(stateSpec)

	' Declarations
	Dim sheetRef
	Dim originParm
	Dim resultParm
	Dim rowIndex

	' Get reference to US States sheet, loading if needed
	Set sheetRef = loadSheetFromPath("US State List", "State Abbr.xls", 1, 0)

	' If abbreviation specified
	If (stateSpec.Len = 2) Then
		originParm = "Abbr"
		resultParm = "State"
	Else ' otherwise (name specified)
		originParm = "State"
		resultParm = "Abbr"
	End If

	' Get row index of specified state name/abbreviation
	rowIndex = getParamValueRowIndex(sheetRef.Name, originParm, stateSpec, 0)
	' If specified state was found
	If Not IsNull(rowIndex) Then
		' RESULT: Corresponding name/abbreviation
		correlateStateToAbbr = sheetRef.GetParameter(resultParm).ValueByRow(rowIndex)
	Else ' otherwise (state not found)
		' Log the failure
		Reporter.ReportEvent micFail, "correlateStateToAbbr", "Name/abbreviation [" & stateSpec & "] not found. Compare spelling with 'State Abbr.xls' workbook."
		' Exit the test
		ExitTest ("State not found")
	End If

	' Release sheet object
	Set sheetRef = Nothing
   
End Function


Public Function getButtonLink(linkSpec)

	' Declarations
	Dim regEx
	Dim descButtonLink
	Dim buttonList
	Dim buttonCount
	Dim buttonIndex
	Dim thisButton
	Dim buttonHTML
	Dim buttonFound

	' Initialize result
	buttonFound = False

	' Allocate RegExp object
	Set regEx = New RegExp
	' Set button link pattern
	regEx.Pattern = linkSpec

	Set descButtonLink = Description.Create()
	descButtonLink("micclass").Value = "WebElement"
	descButtonLink("html tag").Value = "SPAN"
	descButtonLink("class").Value = "un_buttonLink"

	' Get current list of button links
	Set buttonList = Browser("Browser").Page("Page").ChildObjects(descButtonLink)
	' Get count of button links
	buttonCount = buttonList.Count
	' Iterate over button links
	For buttonIndex = 0 To (buttonCount - 1)
		Set thisButton = buttonList(buttonIndex)
		buttonHTML = thisButton.GetROProperty("innerhtml")
		buttonFound = regEx.Test(buttonHTML)
		If (buttonFound) Then Exit For
	Next

	If (buttonFound) Then
		Set getButtonLink = thisButton
	Else
		Set getButtonLink = Nothing
	End If

	' Release objects
	Set thisButton = Nothing
	Set buttonList = Nothing
	Set descButtonLink = Nothing
	Set regEx = Nothing
	
End Function


' --------------------------------------------------------------------------------------
'                        cmnDismissSecurityAlert
' --------------------------------------------------------------------------------------
' Dismisses the dialog if and only if it displays
' --------------------------------------------------------------------------------------
Public Function cmnDismissSecurityAlert()
	Browser("Browser").Page("Page").Sync			
End Function

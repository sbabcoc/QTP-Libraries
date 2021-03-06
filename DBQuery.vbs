Option Explicit

'@Description Get connection string for the front-end database
'@Documentation Get connection string for the front-end database
'@Author sbabcoc
'@Date 08-JUL-2011
'@ReturnValue Front-end database connection string
Public Function query_GetConnectStr()

	' Declarations
	Dim driver
	Dim server

	' Get ADODB driver specification
	driver = Environment("dbDriver")

	' If testing in production
	If isProduction() Then
		' Use production database server
		server = Environment("dbProdHost")
	Else
		' Use QA database server
		server = Environment("dbQAHost")
	End If

	query_GetConnectStr = "DRIVER=" & driver & "; DBQ=" & server & "; DBA=R"

End Function


'@Description Get credentials for the front-end database
'@Documentation Get credentials for the front-end database
'@Author sbabcoc
'@Date 08-JUL-2011
'@ReturnValue Front-end database credentials as (<userName>, <password>)
Public Function query_GetCredentials()

	' Declarations
	Dim userName
	Dim password

	' Get database user name
	userName = Environment("dbUserName")
	' Get database password
	password = Environment("dbPassword")

	query_GetCredentials = Array(userName, password)

End Function


'@Description Execute the specified query and return array of rows
'@Documentation Execute <queryStr> and return array of rows
'@Author sbabcoc
'@Date 08-JUL-2011
'@InParameter [in] queryStr, string, T-SQL query string
'@InParameter [in] selectFlags, bitmap, Selection/quantity specification:
'		[^###############] qry_RandOrder : 1 = order rows randomly; 0 = order rows sequentially
'		[#^^^^^^^^^^^^^^^] qry_RowsMask : number of rows to return; 0 = return all rows
'@ReturnValue If query returns rows, a rows array; otherwise 'Null'
Public Function query_Execute(queryStr, selectFlags)

	' Declarations
	Dim connectStr
	Dim credentials
	Dim connection
	Dim recordSet
	Dim setLength
	Dim selectCount

	' If random row order specified, tell the server to randomize the row order
	If (selectFlags And qry_RandOrder) Then queryStr = queryStr & " ORDER BY dbms_random.value "

	' Extract fixed quantity from select flags
	selectCount = (selectFlags And qry_RowsMask)
	' If fixed row count spec'd, tell the server how many rows to return
	If (selectCount > 0) Then queryStr = "SELECT * FROM ( " & queryStr & ") WHERE rownum <= " & selectCount
	
	' Get database connection string
	connectStr = query_GetConnectStr()
	' Get database credentials
	credentials = query_GetCredentials()

	' Create connection and record set
	Set connection = CreateObject("ADODB.Connection") 
	Set recordSet = CreateObject("ADODB.RecordSet")
	' Set record set to use client-side cursor
	recordSet.CursorLocation = 3 ' => adUseClient
	
	' Open database connection
	connection.Open connectStr , credentials(0), credentials(1)
	' Execute specified query
	recordSet.Open queryStr, connection, 3, 1, 1 ' => adOpenStatic, adLockReadOnly, adCmdText
	' Get result record count
	setLength = recordSet.RecordCount

	' If rows were returned
	If (setLength > 0) Then
		' Extract all rows into result array
		query_Execute = recordSet.GetRows()
		' If fixed row count spec'd
		If (selectCount > 0) Then
			' If fewer rows rec'd than req'd
			If (setLength < selectCount) Then
				Reporter.ReportEvent micWarning, "query_Execute", "Returning fewer rows than requested for query: " & vbCR & _
					queryStr & vbCR & "EXPECT: " & selectCount & vbCR & "ACTUAL: " & setLength
			End If
		End If
	Else
		' Indicate no records
		query_Execute = Null
		Reporter.ReportEvent micWarning, "query_Execute", "No rows returned for query: " & vbCR & queryStr
	End If
    
	' Close database
	recordSet.Close
	connection.Close

	' Release objects
	Set connection = Nothing
	Set recordSet = Nothing

End Function


'@Description Get brand name for specified part number
'@Documentation Get brand name for <partNumber>
'@Author sbabcoc
'@Date 07-SEP-2011
'@InParameter [in] partNumber, string, REI part number (SKU)
'@ReturnValue Brand name if specified; or empty string
Public Function query_GetBrandName(partNumber)
	' Declarations
	Dim queryStr
	Dim recordArray
	Dim brandName
	Dim brandTest
	Dim styleNumber
	
	' Initialize result
	brandName = ""

	If IsNumeric(partNumber) Then
		' Extract style number from part number
		styleNumber = Mid(CStr(partNumber), 1, 6)
		
		' Assemble query string to get product brand page
		queryStr = "SELECT mfname FROM catentry WHERE partnumber = '" & styleNumber & "'"
		' Execute assembled query
		recordArray = query_Execute(queryStr, qry_FixedRowCnt Or 1)
	
		' If records were returned
		If Not IsNull(recordArray) Then
			' Extract brand name
			brandTest = recordArray(0, 0)
			' If brand name is defined
			If Not IsNull(brandTest) Then
				' Use defined value
				brandName = brandTest
			End If
		End If
	End If

	query_GetBrandName = brandName
		
End Function


'@Description Get splash page property for specified part number
'@Documentation Get splash page property for <partNumber>
'@Author sbabcoc
'@Date 08-JUL-2011
'@InParameter [in] partNumber, string, REI part number (SKU)
'@ReturnValue Value of splash page property (boats, cycle, skis); or empty string
Public Function query_GetSplashPage(partNumber)

	' Declarations
	Dim queryStr
	Dim recordArray
	Dim splashPage
	Dim splashTest
	Dim styleNumber
	
	' Initialize result
	splashPage = ""

	If IsNumeric(partNumber) Then
		' Extract style number from part number
		styleNumber = Mid(CStr(partNumber), 1, 6)
		
		' Assemble query string to get product splash page
		queryStr = _
			"SELECT ca.attribute_value " & _
			"FROM catentry ce " & _
			"JOIN catgrprel cgr ON ce.FIELD3 = cgr.CATGROUP_ID_CHILD " & _
			"JOIN category_attributes ca ON (cgr.CATGROUP_ID_CHILD = ca.CATGROUP_ID or cgr.CATGROUP_ID_PARENT = ca.CATGROUP_ID) " & _
			"WHERE attribute_name = 'splash' AND partnumber = '" & styleNumber & "'"
			
		' Execute assembled query
		recordArray = query_Execute(queryStr, qry_FixedRowCnt Or 1)
	
		' If records were returned
		If Not IsNull(recordArray) Then
			' Extract splash page
			splashTest = recordArray(0, 0)
			' If splash page is defined
			If Not IsNull(splashTest) Then
				' Use defined value
				splashPage = splashTest
			End If
		End If
	End If

	query_GetSplashPage = splashPage
		
End Function


'@Description Get store and backorder specifications from product type flags
'@Documentation Get store and backorder specifications from <prodTypeFlags>
'@Author sbabcoc
'@Date 26-AUG-2011
'@InParameter [in] prodTypeFlags, bitmap, Product type flags:
'		qry_REI_Online_StockGone_CanBakOrd - Product from REI Online, out of stock, back-order-able
'		qry_REI_Online_HaveStock_NotBakOrd - Product from REI Online, stock on hand, not back-order-able
'		qry_REI_Online_HaveStock_CanBakOrd - Product from REI Online, stock on hand, back-order-able
'		qry_REI_Online_HaveStock_DoNotCare - Product from REI Online, stock on hand, ignore BO status
'		qry_REI_Outlet_HaveStock_NotBakOrd - Product from REI-OUTLET, stock on hand, not back-order-able
'		qry_Any_Store_HaveStock_DoNotCare - Product from Any store, stock on hand, ignore BO status 
'@ReturnValue Array of T-SQL qualifiers, as (<store-spec>, <store-ent-spec>, <back-ord-spec>)
Private Function query_DecodeTypeFlags(prodTypeFlags)

	Dim storeSpec
	Dim storeEntSpec
	Dim backOrdSpec

	' Set default store and quantity specs
	storeSpec = "AND ca.store_id = 8000 "
	storeEntSpec = "AND sc.storeent_id = 8000 "

	Select Case prodTypeFlags

		' REI Online, sold out, backorders taken
		Case qry_REI_Online_StockGone_CanBakOrd
			backOrdSpec = "AND ce.field4 = 0 "

		' REI Online, in stock, cannot backorder
		Case qry_REI_Online_HaveStock_NotBakOrd
			backOrdSpec = "AND ce.field4 > 0 "

		' REI Online, in stock, backorders taken
		Case qry_REI_Online_HaveStock_CanBakOrd
			backOrdSpec = "AND ce.field4 = 0 "

		' REI Online, in stock, don't check BO status
		Case qry_REI_Online_HaveStock_DoNotCare
			backOrdSpec = ""

		' REI Outlet, in stock, cannot backorder
		Case qry_REI_Outlet_HaveStock_NotBakOrd
			storeSpec = "AND ca.store_id = 8001 "
			storeEntSpec = "AND sc.storeent_id = 8001 "
			backOrdSpec = "AND ce.field4 > 0 "

		' None of the above
		' Online or Outlet, in stock, don't check BO status
		Case Else
			storeSpec = ""
			storeEntSpec = "AND sc.storeent_id != 8002 "
			backOrdSpec = ""

	End Select

	query_DecodeTypeFlags = Array(storeSpec, storeEntSpec, backOrdSpec)

End Function


'@Description Get RSPU stores authorized for specified part number
'@Documentation Get RSPU stores authorized for <partNumber>
'@Author sbabcoc
'@Date 09-SEP-2011
'@InParameter [in] partNumber, string, REI part number (SKU)
'@ReturnValue Array of authorized stores; unrestricted SKUs return 'vbUseDefault'
Public Function query_GetRspuStores(partNumber)

	' Declarations
	Dim storeSpec
	Dim styleNumber
	Dim styleSplash
	Dim styleBrand
	Dim rspuParms
	Dim brandStores
	
	' Initialize result
	storeSpec = Array()

	' If part number is numeric
	If IsNumeric(partNumber) Then
		' Extract style number from part number
		styleNumber = Mid(CStr(partNumber), 1, 6)

		' If product isn't a giftcard
		If Not (styleNumber = "999049") Then
			' Set default store spec
			storeSpec = vbUseDefault
			' Get product splash page
			styleSplash = query_GetSplashPage(styleNumber)
			' If product is a bicycle
			If (styleSplash = "cycle") Then
				' Get product brand name
				styleBrand = query_GetBrandName(styleNumber)
				' Get brands and stores for restricted RSPU-only bikes
				Set rspuParms = query_GetRspuParms(CAT_SELECT)
				' If this is a restricted RSPU-only brand
				If (rspuParms.Exists(styleBrand)) Then
					' Extract list of authorized stores
					brandStores = rspuParms.Item(styleBrand)
					' Convert string-format list to array
					storeSpec = Split(brandStores, ",")
				End If
			End If
		End If
	End If

	query_GetRspuStores = storeSpec

	' Release objects
	Set rspuParms = Nothing

End Function


'@Description Get brand names for specified classes of RSPU-only bicycles
'@Documentation Get brand names for specified classes of RSPU-only bicycles
'@Author sbabcoc
'@Date 05-SEP-2011
'@InParameter [in] prodSelectSpec, constant, Product category specification
'		Flags:
'			CAT_GLOBAL - RSPU Global; return RSPU-only product without store limitations (implies CAT_WHEEL)
'			CAT_SELECT - RSPU Select; return RSPU-only product with limited store selections (implies CAT_WHEEL) 
'			CAT_RSPU - RSPU product; return RSPU-only product with or without limitations (implies CAT_WHEEL)
'@ReturnValue Array of brand names for RSPU-only products
Public Function query_GetRspuBrands(prodSelectSpec)

	' Declarations
	Dim rspuParms

	Set rspuParms = query_GetRspuParms(prodSelectSpec)
	query_GetRspuBrands = rspuParms.Keys

	' Release objects
	Set rspuParms = Nothing

End Function


'@Description Get brand names for specified classes of RSPU-only bicycles
'@Documentation Get brand names for specified classes of RSPU-only bicycles
'@Author sbabcoc
'@Date 05-SEP-2011
'@InParameter [in] prodSelectSpec, constant, Product category specification
'		Flags:
'			CAT_GLOBAL - RSPU Global; return RSPU-only product without store limitations (implies CAT_WHEEL)
'			CAT_SELECT - RSPU Select; return RSPU-only product with limited store selections (implies CAT_WHEEL) 
'			CAT_RSPU - RSPU product; return RSPU-only product with or without limitations (implies CAT_WHEEL)
'@ReturnValue Array of brand names for RSPU-only products
Public Function query_GetRspuParms(prodSelectSpec)

	' Declarations
	Dim couponSpec
	Dim rspuQuery
	Dim parmArray
	Dim parmCount
	Dim parmIndex
	Dim parmDict

	Dim isDone, thisIdent
	Dim nextIdent, thisValue
	Dim valueDelim, parmValue
	Dim labelDelim, parmLabel
	Dim brandDelim, brandName
	Dim brandCount, brandList()
	Dim storeSpec

	' Initialize result
	Set parmDict = CreateObject("Scripting.Dictionary")
	
	' Discriminate RSPU class
	Select Case prodSelectSpec

		' RSPU (all stores)
		Case CAT_GLOBAL
			couponSpec = "WHERE co.name = 'RSPU BIKES' "

		' RSPU (select stores)
		Case CAT_SELECT
			couponSpec = "WHERE co.name = 'RSPU_RESTRICT_BIKES' "

		' RSPU (don't care)
		Case CAT_RSPU
			couponSpec = "WHERE co.name IN ('RSPU BIKES', 'RSPU_RESTRICT_BIKES') "
   
	End Select

	parmQuery = _
		"SELECT cr1.child_cpobj_id AS Ident, cp.parameter_value AS Value " & _
		"FROM coupon_objects co " & _
		"JOIN coupon_relations cr1 ON cr1.parent_cpobj_id = co.id " & _
		"JOIN coupon_relations cr2 ON cr2.parent_cpobj_id = cr1.child_cpobj_id " & _
		"JOIN coupon_parameters cp ON cp.cpobj_id = cr2.child_cpobj_id " & _
		couponSpec & _
		"ORDER BY cr2.parent_cpobj_id, cr2.child_cpobj_id "

	' Get RSPU parameter records
	parmArray = query_Execute(parmQuery, 0)
	' If records were returned
	If Not IsNull(parmArray) Then
		' Get count of RSPU parm records
		parmCount = UBound(parmArray, 2)

		'##### INIT #####
		' ... parm index
		parmIndex = 0
		' ... 'done' flag
		isDone = False
		' ... brand count
		brandCount = -1
		' ... store specification
		storeSpec = vbUseDefault
		' ... coupon identifier
		thisIdent = CStr(parmArray(0, 0))
		
		' Iterate over RSPU parm records
		Do
			' Extract parameter label + value
			thisValue = parmArray(1, parmIndex)

			' Locate final '=' character
			valueDelim = InStrRev(thisValue, "=")
			' Extract parameter value
			parmValue = Mid(thisValue, valueDelim + 1)
			' Locate prior '=' character
			labelDelim = InStrRev(thisValue, "=", valueDelim - 1)
			' Extract  parameter label
			parmLabel = Mid(thisValue, labelDelim + 1, valueDelim - labelDelim - 1)

			' Differentiate parm label
			Select Case parmLabel

				'##### INCLUDED BRAND #####
				Case "Include"
					' Locate space char
					brandDelim = InStr(parmValue, " ")
					' Extract brand name
					brandName = Left(parmValue, brandDelim - 1)

					' Increment brand count
					brandCount = brandCount + 1
					' Expand brand list for new entry
					ReDim Preserve brandList(brandCount)
					' Add new entry to list of brands
					brandList(brandCount) = brandName

				'##### INCLUDED STORES #####
				Case "InclStores"
					' Set store specification
					storeSpec = parmValue

			End Select
			
			' Increment parm index
			parmIndex = parmIndex + 1
			' If process is complete
			If (parmIndex > parmCount) Then
				' Indicate done
				isDone = True
			Else
				' Extract next coupon identifier
				nextIdent = CStr(parmArray(0, parmIndex))
			End If

			' If process complete or coupon complete
			If (isDone Or (nextIdent <> thisIdent)) Then
				' Iterate over brand names
				For Each brandName in brandList
					' Add brand with allowed stores
					parmDict.Add brandName, storeSpec
				Next
				
				'##### RESET #####
				' ... brand list
				Erase brandList
				' ... brand count
				brandCount = -1
				' ... store specification
				storeSpec = vbUseDefault

				' Update coupon ID
				thisIdent = nextIdent
			End If
		Loop Until (isDone)
	End If
	
	Set query_GetRspuParms = parmDict

	' Release objects
	Set parmDict = Nothing

End Function


'@Description Build a T-SQL query string based on specified parameters
'@Documentation Build a T-SQL query string based on specified parameters
'@Author sbabcoc
'@Date 26-AUG-2011
'@InParameter [in] prodTypeFlags, bitmap, Product type flags:
'		qry_REI_Online_StockGone_CanBakOrd - Product from REI Online, out of stock, back-order-able
'		qry_REI_Online_HaveStock_NotBakOrd - Product from REI Online, stock on hand, not back-order-able
'		qry_REI_Online_HaveStock_CanBakOrd - Product from REI Online, stock on hand, back-order-able
'		qry_REI_Online_HaveStock_DoNotCare - Product from REI Online, stock on hand, ignore BO status
'		qry_REI_Outlet_HaveStock_NotBakOrd - Product from REI-OUTLET, stock on hand, not back-order-able
'		qry_Any_Store_HaveStock_DoNotCare - Product from Any store, stock on hand, ignore BO status 
'@InParameter [in] prodSelectSpec, constant, Product selection specification
'		Constants:
'			CAT_UNSPEC - Unspecified; return any product regardless of category
'			CAT_NOSPL - No splash; return product with no "Terms and Conditions" page
'			CAT_WHEEL - Wheeled; return product with "Wheeled Sports Gear" T&C page
'			CAT_SNOW - Snow; return product with "Snow Sports Gear" T&C page
'			CAT_PADDLE - Paddle; return product with "Paddle Sports Gear" T&C page
'			CAT_SPLASH - Splash; return product with any of the three T&C pages
'		Flags:
'			CAT_GLOBAL - RSPU Global; return RSPU-only product without store limitations (implies CAT_WHEEL)
'			CAT_SELECT - RSPU Select; return RSPU-only product with limited store selections (implies CAT_WHEEL) 
'			CAT_RSPU - RSPU product; return RSPU-only product with or without limitations (implies CAT_WHEEL)
'			CAT_LARGE - Oversize; return product with oversize shipping requirements
'			CAT_RISKY - Hazardous; return product with hazardous material restrictions
'@InParameter [in] prodMinQuantity, number, If not requesting out-of-stock product, minimum on-hand quantity
'@ReturnValue T-SQL query string based on specified parameters
Public Function query_ByCatConst(prodTypeFlags, prodSelectSpec, prodMinQuantity)

	' NOTE: All of the queries produced by this function include a JOIN to the storecent table.
	'                This ensures that all selected products are actually available for purchase.

	' Declarations
	Dim specArray
	Dim storeSpec
	Dim storeEntSpec
	Dim backOrdSpec
	Dim quantitySpec
	Dim sizeSpec
	Dim riskSpec
	Dim selectClause
	Dim fromClause
	Dim joinClause
	Dim whereClause
	Dim subQuery
	Dim splashSpec
	Dim rspuBrands
	Dim rspuQuery
	Dim rspuJoin

	' Get SQL qualifiers for store spec and backorder status
	specArray = query_DecodeTypeFlags(prodTypeFlags)
	' Extract qualifiers
	storeSpec = specArray(0)
	storeEntSpec = specArray(1)
	backOrdSpec = specArray(2)

	'##### DECODE SELECTION CRITERIA #####

	' If selected products should be in stock
	If (prodMinQuantity > 0) Then
		quantitySpec = "AND (i1.quantity + i2.quantity) BETWEEN " & prodMinQuantity & " AND 9000 "
	Else
		quantitySpec = "AND (i1.quantity + i2.quantity) = 0 "
	End If

	'##### CHECK CATEGORY SPECIFICATION #####

	' If RSPU-only product is specified
	If (prodSelectSpec And CAT_RSPU) Then
		rspuBrands = query_GetRspuBrands(prodSelectSpec And CAT_RSPU)

		rspuQuery = _
			"SELECT cer.catentry_id_child as catentry_id " & _
			"FROM catentry ce " & _
			"JOIN catentrel cer ON cer.catentry_id_parent = ce.catentry_id " & _
			"WHERE ce.mfname IN ('" & Join(rspuBrands, "', '") & "') "

		rspuJoin = "JOIN ( " & rspuQuery & ") rspu ON rspu.catentry_id = sc.catentry_id "

		prodSelectSpec = CAT_WHEEL
	Else
		rspuJoin = ""
	End If

	' If oversize product is specified
	If (prodSelectSpec And CAT_LARGE) Then
		sizeSpec = "AND cee.oversize_flag = 'Y' "
		prodSelectSpec = prodSelectSpec And (Not CAT_LARGE)
	Else
		sizeSpec = "AND cee.oversize_flag = 'N' "
	End If

	' If hazardous product is specified
	If (prodSelectSpec And CAT_RISKY) Then
		riskSpec = "AND cee.hazardous_flag = 'Y' "
		prodSelectSpec = prodSelectSpec And (Not CAT_RISKY)
	Else
		riskSpec = "AND cee.hazardous_flag = 'N' "
	End If

	' If product from any category is indicated
	If (prodSelectSpec = CAT_UNSPEC) Then

		selectClause = _
			"SELECT ce.partnumber as SKU, (i1.quantity + i2.quantity) as Qty "

		fromClause = _
			"FROM catentry ce "

		joinClause = _
			"JOIN catentry_extension cee ON cee.catentry_id = ce.catentry_id " & _
			"JOIN inventory i1 ON i1.catentry_id = cee.catentry_id " & _
			"JOIN inventory i2 ON i2.catentry_id = i1.catentry_id " & _
			"JOIN storecent sc ON sc.catentry_id = i2.catentry_id "

		whereClause = _
			"WHERE ce.catenttype_id = 'ItemBean' " & _  
			backOrdSpec & _
			sizeSpec & _
			riskSpec & _
			"AND i1.ffmcenter_id = 1 " & _
			"AND i2.ffmcenter_id = 2 " & _
			quantitySpec & _
			storeEntSpec

	' Otherwise, if product with no "splash" is spec'd
	ElseIf (prodSelectSpec = CAT_NOSPL) Then

		subQuery = _
			"SELECT ce.partnumber " & _
			"FROM catentry ce " & _
			"JOIN catgrprel cgr ON ce.field3 = cgr.catgroup_id_child " & _
			"JOIN category_attributes ca ON (cgr.catgroup_id_child = ca.catgroup_id OR cgr.catgroup_id_parent = ca.catgroup_id) " & _
			"WHERE ca.attribute_name = 'splash' "

		selectClause = _
			"SELECT DISTINCT ce.partnumber as SKU, (i1.quantity + i2.quantity) as Qty "

		fromClause = _
			"FROM catentry ce2 "

		joinClause = _
			"LEFT JOIN ( " & subQuery & ") temp ON temp.partnumber = ce2.partnumber " & _
			"JOIN catentrel cer ON cer.catentry_id_parent = ce2.catentry_id " & _
			"JOIN catentry ce ON ce.catentry_id = cer.catentry_id_child " & _
			"JOIN catentry_extension cee ON cee.catentry_id = ce.catentry_id " & _
			"JOIN inventory i1 ON i1.catentry_id = cee.catentry_id " & _
			"JOIN inventory i2 ON i2.catentry_id = i1.catentry_id " & _
			"JOIN storecent sc ON sc.catentry_id = i2.catentry_id "

		whereClause = _
			"WHERE ce2.catenttype_id = 'ProductBean' " & _
			"AND temp.partnumber IS null " & _
			backOrdSpec & _
			sizeSpec & _
			riskSpec & _
			"AND i1.ffmcenter_id = 1 " & _
			"AND i2.ffmcenter_id = 2 " & _
			quantitySpec & _
			storeEntSpec

	' Otherwise (product with "splash")
	Else

		' Decode category spec
		Select Case prodSelectSpec

			Case CAT_WHEEL ' Wheeled Sports Gear
				splashSpec = "AND ca.attribute_value = 'cycle' "

			Case CAT_SNOW ' Snow Sports Gear
				splashSpec = "AND ca.attribute_value = 'skis' "

			Case CAT_PADDLE ' Paddle Sports Gear 
				splashSpec = "AND ca.attribute_value = 'boats' "

			Case Else ' Every item with "splash" page
				splashSpec = ""

		End Select

		selectClause = _
			"SELECT DISTINCT ce.partnumber as SKU, (i1.quantity + i2.quantity) as Qty "

		fromClause = _
			"FROM category_attributes ca "

		joinClause = _
			"JOIN catgrprel cgr ON (cgr.catgroup_id_parent = ca.catgroup_id or cgr.catgroup_id_child = ca.catgroup_id) " & _
			"JOIN catgpenrel cger ON cger.catgroup_id = cgr.catgroup_id_child " & _
			"JOIN catentrel cer ON cer.catentry_id_parent = cger.catentry_id " & _
			"JOIN catentry ce ON ce.catentry_id = cer.catentry_id_child " & _
			"JOIN inventory i1 ON i1.catentry_id = ce.catentry_id " & _
			"JOIN inventory i2 ON i2.catentry_id = i1.catentry_id " & _
			"JOIN storecent sc ON sc.catentry_id = i2.catentry_id " & _
			rspuJoin

		whereClause = _
			"WHERE ca.attribute_name = 'splash' " & _
			splashSpec & _
			storeSpec & _
			"AND cger.catalog_id != 40000008002 " & _
			backOrdSpec & _
			"AND i1.ffmcenter_id = 1 " & _
			"AND i2.ffmcenter_id = 2 " & _
			quantitySpec & _
			storeEntSpec

	End If

	' Assemble query string from component clauses
	query_ByCatConst = selectClause & fromClause & joinClause & whereClause

End Function


'@Description Acquire a random product number matching specified criteria
'@Documentation Acquire a random product number matching specified criteria
'@Author sbabcoc
'@Date 29-JUN-2011
'@InParameter [in] prodTypeFlags, bitmap, Product type flags:
'		qry_REI_Online_StockGone_CanBakOrd - Product from REI Online, out of stock, back-order-able
'		qry_REI_Online_HaveStock_NotBakOrd - Product from REI Online, stock on hand, not back-order-able
'		qry_REI_Online_HaveStock_CanBakOrd - Product from REI Online, stock on hand, back-order-able
'		qry_REI_Online_HaveStock_DoNotCare - Product from REI Online, stock on hand, ignore BO status
'		qry_REI_Outlet_HaveStock_NotBakOrd - Product from REI-OUTLET, stock on hand, not back-order-able
'		qry_Any_Store_HaveStock_DoNotCare - Product from Any store, stock on hand, ignore BO status 
'@InParameter [in] prodSelectSpec, number/string, Product selection specification
'		##### CONSTANTS #####
'		CAT_UNSPEC - Unspecified; return any product regardless of category
'		CAT_NOSPL - No splash; return product with no "Terms and Conditions" page
'		CAT_WHEEL - Wheeled; return product with "Wheeled Sports Gear" T&C page
'		CAT_SNOW - Snow; return product with "Snow Sports Gear" T&C page
'		CAT_PADDLE - Paddle; return product with "Paddle Sports Gear" T&C page
'		CAT_SPLASH - Splash; return product with any of the three T&C pages
'		##### FLAGS #####
'		CAT_GLOBAL - RSPU Global; return RSPU-only product without store limitations (implies CAT_WHEEL)
'		CAT_SELECT - RSPU Select; return RSPU-only product with limited store selections (implies CAT_WHEEL) 
'		CAT_RSPU - RSPU product; return RSPU-only product with or without limitations (implies CAT_WHEEL)
'		CAT_LARGE - Oversize; return product with oversize shipping requirements
'		CAT_RISKY - Hazardous; return product with hazardous material restrictions
'@InParameter [in] prodMinQuantity, number, If not requesting out-of-stock product, minimum on-hand quantity
'@ReturnValue If found, result as (<product-number>, <product-quantity>); otherwise 'Null' 
Public Function query_GetRandomProduct(prodTypeFlags, prodSelectSpec, prodMinQuantity)

	' Declarations
	Dim recordArray
	
	' Execute assembled query
	recordArray = query_GetProduct(prodTypeFlags, prodSelectSpec, qry_Type_CatConstant, prodMinQuantity, qry_RandOrder Or 1)

	' If records were returned
	If Not IsNull(recordArray) Then
		' RESULT: First record returned
		query_GetProduct = recordArray(0)
	Else
		query_GetProduct = Null
	End If

End Function


'@Description Acquire product number(s) matching specified criteria
'@Documentation Acquire product number(s) matching specified criteria
'@Author sbabcoc
'@Date 29-JUN-2011
'@InParameter [in] prodTypeFlags, bitmap, Product type flags:
'		qry_REI_Online_StockGone_CanBakOrd - Product from REI Online, out of stock, back-order-able
'		qry_REI_Online_HaveStock_NotBakOrd - Product from REI Online, stock on hand, not back-order-able
'		qry_REI_Online_HaveStock_CanBakOrd - Product from REI Online, stock on hand, back-order-able
'		qry_REI_Online_HaveStock_DoNotCare - Product from REI Online, stock on hand, ignore BO status
'		qry_REI_Outlet_HaveStock_NotBakOrd - Product from REI-OUTLET, stock on hand, not back-order-able
'		qry_Any_Store_HaveStock_DoNotCare - Product from Any store, stock on hand, ignore BO status 
'@InParameter [in] prodSelectSpec, number/string, Product selection specification
'		##### CONSTANTS #####
'		CAT_UNSPEC - Unspecified; return any product regardless of category
'		CAT_NOSPL - No splash; return product with no "Terms and Conditions" page
'		CAT_WHEEL - Wheeled; return product with "Wheeled Sports Gear" T&C page
'		CAT_SNOW - Snow; return product with "Snow Sports Gear" T&C page
'		CAT_PADDLE - Paddle; return product with "Paddle Sports Gear" T&C page
'		CAT_SPLASH - Splash; return product with any of the three T&C pages
'		##### FLAGS #####
'		CAT_GLOBAL - RSPU Global; return RSPU-only product without store limitations (implies CAT_WHEEL)
'		CAT_SELECT - RSPU Select; return RSPU-only product with limited store selections (implies CAT_WHEEL) 
'		CAT_RSPU - RSPU product; return RSPU-only product with or without limitations (implies CAT_WHEEL)
'		CAT_LARGE - Oversize; return product with oversize shipping requirements
'		CAT_RISKY - Hazardous; return product with hazardous material restrictions
'@InParameter [in] selectSpecType, number, Selection specification type
'		qry_Type_CatConstant - Specification is a category constant
'		qry_Type_CatEntry_ID - Specification is a CatEntry_ID value
'		qry_Type_CatGroup_ID - Specification is a CatGroup_ID value
'		qry_Type_StyleNumber - Specification is an REI style number
'		qry_Type_PartNumber - Specification is an REI part number (SKU)
'		qry_Type_Brand_ID - Specification is a brand identifier
'		qry_Type_CatKeyword - Specification is a category keyword
'		qry_Type_ProdKeyword - Specification is a product keyword
'		qry_Type_CatDescrip - Specification is a category description
'		qry_Type_ProdDescrip - Specification is a product description
'		qry_Type_BrandName - Specification is a brand name
'@InParameter [in] prodMinQuantity, number, If not requesting out-of-stock product, minimum on-hand quantity
'@InParameter [in] selectFlags, bitmap, Selection/quantity specification:
'		[^###############] qry_RandOrder : 1 = order rows randomly; 0 = order rows sequentially
'		[#^^^^^^^^^^^^^^^] qry_RowsMask : number of rows to return; 0 = return all rows
'@ReturnValue If found, result as (<product-number>, <product-quantity>); otherwise 'Null' 
Public Function query_GetProduct(prodTypeFlags, prodSelectSpec, selectSpecType, prodMinQuantity, selectFlags)

	' Declarations
	Dim queryStr, recordArray

	Select Case selectSpecType

		Case qry_Type_CatConstant
			queryStr = query_ByCatConst(prodTypeFlags, prodSelectSpec, prodMinQuantity)

		Case qry_Type_CatEntry_ID
			' Fixed row count of 1 is implied
			selectFlags = 1

		Case qry_Type_CatGroup_ID

		Case qry_Type_StyleNumber

		Case qry_Type_PartNumber
			' Fixed row count of 1 is implied
			selectFlags = 1

		Case qry_Type_Brand_ID

		Case qry_Type_CatKeyword

		Case qry_Type_ProdKeyword

		Case qry_Type_CatDescrip

		Case qry_Type_ProdDescrip

		Case qry_Type_BrandName

		Case Else

	End Select

	' Execute assembled query
	query_GetProduct = query_Execute(queryStr, selectFlags)

End Function

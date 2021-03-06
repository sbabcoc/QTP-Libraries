Option Explicit

'@Description Navigate to the order form for indicated gift card type
'@Documentation Navigate to the order form for indicated gift card type
'@Author sbabcoc
'@Date 11-JUL-2011
'@Libraries Global, Verifications
'@Repositories Common, Web, GiftCard
'@InParameter [in] sendECard 'True' for e-gift card form; 'False' for plastic card form
'@ReturnValue If navigation succeeds, the URL of the indicated order form; otherwise 'Null'
Public Function giftcard_NavigateToPage(sendECard)

	' Declarations
	Dim refObject
	Dim descTarget
	Dim chk_href

	' Load description of top-level Gift Cards page
	Set descTarget = Browser("GiftCard").Page("REI Gift Cards").GetTOProperties
	' Search for REI gift card SKU
	chk_href = search_SubmitQuery("9990490006", descTarget)

	' If navigation was successful
	If Not IsNull(chk_href) Then
		' If sending an e-card
		If (sendECard) Then
			' Get reference to "Buy E-Gift Card" link
			Set refObject = Browser("GiftCard").Page("REI Gift Cards").Link("Buy E-Gift Cards")
			' Load description of e-gift card order form
			Set descTarget = Browser("GiftCard").Page("Buy E-Gift Card").GetTOProperties
		' Otherwise (plastic gift card)
		Else
			' Get reference to "Buy Gift Card" link
			Set refObject = Browser("GiftCard").Page("REI Gift Cards").Link("Buy Gift Cards")
			' Load description of gift card order form
			Set descTarget = Browser("GiftCard").Page("Buy Gift Card").GetTOProperties
		End If

		' Verify target of "Buy Gift Card" link
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
	End If

	giftcard_NavigateToPage = chk_href

End Function


'@Description Add indicated type of gift card to shopping catr
'@Documentation Add indicated type of gift card to shopping cart
'@Author sbabcoc
'@Date 13-JUL-2011
'@Libraries Verifications
'@Repositories GiftCard, Checkout
'@InParameter [in] sendECard, boolean, 'True' for e-card; 'False' for plastic
'@InParameter [in] cardAmount, number, Amount for gift card in dollars
'@InParameter [in] cardQuantity, number, Quantity of gift cards to purchase
'@InParameter [in] doEmptyCart, boolean, 'True' to purge cart; otherwise 'False'
'@ReturnValue 'True' is addition succeeds; otherwise 'False'
Public Function giftcard_AddToCart(sendECard, cardAmount, cardQuantity, doEmptyCart)

	' Declarations
	Dim isCorrect
	Dim refObject
	Dim descTarget
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

		' Navigate to order form of specified gift card type
		chk_href = giftcard_NavigateToPage(sendECard)
		' Check navigation status
		isCorrect = Not IsNull(chk_href)
		' If navigation fails, bail out
		If Not (isCorrect) Then Exit Do

		With Browser("GiftCard").Page("GC Common")
			' Populate the "To" field
			.WebEdit("gc_to").Set "You"
			' Populate the "From" field
			.WebEdit("gc_from").Set "Me"
			' Populate the "Amount" field
			.WebEdit("gc_amount").Set cardAmount
			' Populate the "Message" field
			.WebEdit("gc_message").Set "Enjoy!"
			' Select first card media type
			.WebRadioGroup("gc_media").Select "#0"
		End With
		
		' If sending an e-card
		If (sendECard) Then
			' Always buy 1
			cardQuantity = 1
			
			With Browser("GiftCard").Page("Buy E-Gift Card")
				' Populate the target e-mail field
				.WebEdit("gc_email1").Set "you@hoo.com"
				' Populate the verify e-mail field
				.WebEdit("gc_email2").Set "you@hoo.com"
			End With
		' Otherwise (plastic)
		Else
			With Browser("GiftCard").Page("Buy Gift Card")
				' Populate the quantity field
				.WebEdit("gc_quantity").Set cardQuantity
			End With
		End If

		' Get reference to "preview your card" button
		Set refObject = Browser("GiftCard").Page("GC Common").WebButton("preview your card")
		' Load definition of gift card "Preview" page
		Set descTarget = Browser("GiftCard").Page("Preview Common").GetTOProperties

		' Verify target of "preview your card" button
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' Check navigation status
		isCorrect = Not IsNull(chk_href)
		' If navigation fails, bail out
		If Not (isCorrect) Then Exit Do

		' Get reference to "add to cart" link
		Set refObject = Browser("GiftCard").Page("Preview Common").Link("add to cart")
		' Load description of Shopping Cart page
		Set descTarget = Browser("Checkout").Page("REI.com: Shopping Basket").GetTOProperties
		
		' Verify target of "add to cart" link
		chk_href = chkVerifyLinkTarget(refObject, descTarget)
		' Check navigation status
		isCorrect = Not IsNull(chk_href)
		' If navigation fails, bail out
		If Not (isCorrect) Then Exit Do

		' Verify addition of product/quantity
		isCorrect = cart_VerifyAddition("9990490006", cardQuantity)

	Loop Until True ' <== End bail-out context
	
	giftcard_AddToCart = isCorrect

	' Release objects
	Set descTarget = Nothing
	Set refObject = Nothing

End Function

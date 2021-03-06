' Define prefix of Quality Center windows
Const QC_PAGE_TAG = "Quality Center"

' Define constants for test set user fields
Const FAST_TIMEOUT = "CY_USER_05"
Const RUN_BY_PRIORITY = "CY_USER_06"
Const SITE_UNDER_TEST = "CY_USER_03"
Const SLOW_TIMEOUT = "CY_USER_01"
Const TEST_ENVIRONMENT = "CY_USER_07"

' Define constants for testing resources
Const URL_CMC_TOOL = "http://ahqcsnuc2:8080/qa5"

' Product detail dictionary keys
Const key_ProdPict = "PICT"				' Product page image link
Const key_ProdLink = "LINK"			' Product page link
Const key_ProdName = "NAME"	' Product name
Const key_ProdNum = "NUM"		' Product catalog number
Const key_ProdCost = "COST"		' Product unit cost
Const key_ProdOrig = "ORIG"			' Original product cost
Const key_ProdSave = "SAVE"		' Savings (percentage)
Const key_ProdQty = "QTY"			' Quantity of this product
Const key_ProdRate = "RATE"		' Product rating
Const key_ProdRevu = "REVU"		' Product review count
Const key_ProdCat = "CAT"			' Product category number
Const key_ProdOpt = "OPT"			' Product color/size option
Const key_ProdPrev = "PREV"		' Product from previous visit

' REI Header link parameters
Const link_Logo1 = "LOGO1_LINK"
Const link_Logo2 = "LOGO2_LINK"
Const link_Home = "HOME_LINK"
Const link_Stores = "STORES_LINK"
Const link_LogIn1 = "LOGIN1_LINK"
Const link_LogIn2 = "LOGIN2_LINK"
Const link_LogOut = "LOGOUT_LINK"
Const link_YourAcct = "YOURACCT_LINK"
Const link_Cart1 = "CART1_LINK"
Const link_Cart2 = "CART2_LINK"

' REI Footer link parameters
Const link_CallREI = "CALL_REI_LINK"
Const link_Help = "HELP_LINK"
Const link_Feedback = "FEEDBACK_LINK"
Const link_Privacy = "PRIVACY_LINK"
Const link_GoToREI = "GO_TO_REI_LINK"

' Feedback link parameter keys
Const key_time1 = "time1"
Const key_time2 = "time2"
Const key_prev = "prev"
Const key_referer = "referer"

' Page header verification flags
Const head_ImageLink = 1
Const head_HomeLink = 2
Const head_StoresLink = 4
Const head_CartLink = 8
Const head_SearchGroup = 16
Const head_SearchField = 32
Const head_GoButton = 64
Const head_SearchSet = 112 ' head_SearchGroup Or head_SearchField Or head_GoButton
Const head_LogInLink = 128
Const head_LogOutLink = 256
Const head_YourAcctLink = 512
Const head_VerifyOK = 32768

' Page footer verification flags
Const foot_CallREIText = 1
Const foot_HelpLink = 2
Const foot_PrivacyLink = 4
Const foot_GoToREILink = 8
Const foot_FeedbackLink = 16
Const foot_VerifyOK = 32768

' Product page tab verification flags
Const prod_GeneralTab = 1
Const prod_DescripTab = 2
Const prod_ImagesTab = 4
Const prod_SpecsTab = 8
Const prod_ReviewsTab = 16
Const prod_VerifyOK = 32768

' Product page tab heading height
Const prod_TabHeadSize = 26

' chkEvaluateLink return values
Const EVAL_PASS = -1
Const EVAL_FAIL = 0
Const EVAL_NONE = 1

' query_GetRandomProduct flags
'===== Store Specification
Const qry_StoreSpec = 4
Const qry_REI_Online = 0
Const qry_REI_Outlet = 4
'===== Stock / BackOrder
Const qry_StockSpec = 3
Const qry_StockGone_CanBakOrd = 0
Const qry_HaveStock_NotBakOrd = 1
Const qry_HaveStock_CanBakOrd = 2
Const qry_HaveStock_DoNotCare = 3
'===== Product Type Mask
Const qry_QueryMask = 7
'===== Supported Types
Const qry_REI_Online_StockGone_CanBakOrd = 0
Const qry_REI_Online_HaveStock_NotBakOrd = 1
Const qry_REI_Online_HaveStock_CanBakOrd = 2
Const qry_REI_Online_HaveStock_DoNotCare = 3
Const qry_REI_Outlet_HaveStock_NotBakOrd = 5
Const qry_Any_Store_HaveStock_DoNotCare = 7

' query_GetRandomProduct select spec types
Const qry_Type_CatConstant = 1
Const qry_Type_CatEntry_ID = 2
Const qry_Type_CatGroup_ID = 3
Const qry_Type_StyleNumber = 4
Const qry_Type_PartNumber = 5
Const qry_Type_Brand_ID = 6
Const qry_Type_CatKeyword = 7
Const qry_Type_ProdKeyword = 8
Const qry_Type_CatDescrip = 9
Const qry_Type_ProdDescrip = 10
Const qry_Type_BrandName = 11

' query_GetRandomProduct category values
Const CAT_UNSPEC = 0
Const CAT_NOSPL = 1
Const CAT_WHEEL = 2
Const CAT_SNOW = 4
Const CAT_PADDLE = 8
Const CAT_SPLASH = 14
Const CAT_GLOBAL = 16
Const CAT_SELECT = 32
Const CAT_RSPU = 48
Const CAT_LARGE = 64
Const CAT_RISKY = 128

' cart_UpdateQuantity flags
Const updt_QuantityMask = 16383
Const updt_FixedQuantity = 16384
Const updt_VerifySubtotal = 32768

' query_Execute flags
Const qry_RowsMask = 16383
Const qry_RandOrder = 32768

' CloseBrowsers Constants
Const EXCLUDE_QC = False
Const INCLUDE_QC = True

' SplitURL parameter keys
Const key_base_url = "base_url"

Option Explicit

'@Description Generate an REI Online user name for the specified user name
'@Documentation Generate an REI Online user name for <userName>
'@Author sbabcoc
'@Date 29-JUN-2011
'@InParameter [in] userName The name of the user for whom the account is requested
'@ReturnValue REI Online user name as "<user-name><epoch-seconds>@rei.com"
Public Function account_GenerateUserName(userName)

	' Get current time as epoch
	theEpoch = date2epoch(Now())
	' Convert epoch to seconds
	theSuffix = CStr(theEpoch / 1000)
	' Assemble test user e-mail address
	testUser = userName & theSuffix & "@rei.com"

	account_GenerateUserName = testUser

End Function


'@Description Generate an REI Online password for the specified user name
'@Documentation Generate an REI Online password for <userName>
'@Author sbabcoc
'@Date 29-JUN-2011
'@InParameter [in] userName The name of the user for whom the password is requested
'@ReturnValue REI Online password for the specified user
Public Function account_GeneratePassword(userName)

	' Init password
	password = ""
	' Init byte mask
	byteMask = 165
	' Start with padded username
	seedword = userName & "bOgUs"
	
	' Process 8 chars
	For index = 1 To 8
		' Extract current character
		thisChar = Mid(seedword, index, 1)
		' Get ASCII code for char
		thisCode = Asc(thisChar)
		' Apply current byte mask
		thisCode = thisCode Xor byteMask
		' Create next byte mask
		byteMask = (thisCode + 37) And 255
		' Scale masked result
		thisCode = Int(thisCode * 94 / 256)
		' Ensure visible 7-bit ASCII result
		thisCode = ((thisCode + 34) And 127) - 1
		' Append coded char to password
		password = password & Chr(thisCode)
	Next

	account_GeneratePassword = password

End Function


'@Description Get REI Online account credentials for the current user
'@Documentation Get REI Online account credentials for the current user
'@Author sbabcoc
'@Date 29-JUN-2011
'@ReturnValue REI Online credentials as (<userName>, <password>)
Public Function account_GetCredentials()

	' Declarations
	Dim testUser
	Dim password
	Dim userName

	On Error Resume Next

	' Get account user name
	testUser = Environment.Value("rei_UserName")
	' Get account password
	password = Environment.Value("rei_Password")

	On Error GoTo 0

	If (IsEmpty(testUser) Or IsEmpty(password)) Then
		' Get name of current user
		userName = getUserName()
		' Generate account user name
		testUser = account_GenerateUserName(userName)
		' Generate account password
		password = account_GeneratePassword(userName)

		' Store account credential in the environment
		Environment.Value("rei_UserName") = testUser
		Environment.Value("rei_Password") = password
	End If

	account_GetCredentials = Array(testUser, password)

End Function



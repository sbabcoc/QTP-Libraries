Option Explicit

' Bitmapped flag constants
Const flagAddSheet = 1
Const flagAddParam = 2
Const flagReplace = 4
Const flagEveryRow = 8

'@Description Get a reference to the specified sheet
'@Documentation Get a reference to <theSheet>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@ReturnValue On success, a reference to the specified sheet; otherwise 'Nothing'
Public Function MyGetSheetRef(ByVal theSheet)

	' Declarations
	Dim sheetSpec

	Wait 0
	On Error Resume Next

	' Initialize result
	Set MyGetSheetRef = Nothing

	' If global sheet is specified
	If (theSheet = dtGlobalSheet) Then
		' Spec is name of global sheet
		sheetSpec = DataTable.GlobalSheet.Name
	' otherwise, if local sheet is specified
	ElseIf (theSheet = dtLocalSheet) Then
		' Spec is name of local sheet
		sheetSpec = DataTable.LocalSheet.Name
	' otherwise
	Else
		' Use spec as-is
		sheetSpec = theSheet
	End If

	' Get reference to the specified sheet
	' NOTE: If sheet is absent, 'Nothing' is returned
	Set MyGetSheetRef = DataTable.GetSheet(sheetSpec)

End Function


'@Description Get a reference to the specified parameter
'@Documentation Get a reference to parameter <theCol> on <theSheet>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theCol, string/integer, parameter specifier (name or 1-based index)
'@ReturnValue On success, a reference to the specified parameter; otherwise 'Nothing
Public Function MyGetParamRef(ByVal theSheet, ByVal theCol)

	' Declarations
	Dim sheetRef

	Wait 0
	On Error Resume Next

	' Initialize result
	Set MyGetParamRef = Nothing

	' Get reference to specified sheet
	Set sheetRef = MyGetSheetRef(theSheet)
	' If specified sheet exists
	If Not (sheetRef Is Nothing) Then
		' Get reference to the specified parameter
		Set MyGetParamRef = sheetRef.GetParameter(theCol)
	End If

	' Release sheet object
	Set sheetRef = Nothing

End Function


'@Description Get the value of the specified parameter
'@Documentation Get parameter <theCol> at <theRow> on <theSheet>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theCol, string/integer, parameter specifier (name or 1-based index)
'@InParameter [in] theRow, integer/Null, row index (1-based); 'Null' specifies current row
'@ReturnValue On success, the value of the specified parameter; otherwise 'Null'
Public Function MyGetParameter(ByVal theSheet, ByVal theCol, ByVal theRow)

	' Declarations
	Dim paramRef

	Wait 0
	On Error Resume Next

	' Initialize result
	MyGetParameter = Null

	' Get reference to specified parameter
	Set paramRef = MyGetParamRef(theSheet, theCol)
	' If specified parameter exists
	If Not (paramRef Is Nothing) Then
		' If default row indicated
		If IsNull(theRow) Then
			' RESULT: Value of specified parameter at current row
			MyGetParameter = paramRef.Value
		' otherwise
		Else
			' RESULT: Value of specified parameter at specified row
			MyGetParameter = paramRef.ValueByRow(theRow)
		End If
	End If

	' Release parameter object
	Set paramRef = Nothing

End Function


'@Description Set the value of the specified parameter
'@Documentation Set parameter <theCol> at <theRow> on <theSheet> to <theValue>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theCol, string/integer, parameter specifier (name or 1-based index)
'@InParameter [in] theRow, integer/Null, row index (1-based); 'Null' specifies current row
'@InParameter [in] theValue, intrinsic, parameter value
'@InParameter [in] theFlags, bitmap, flag specification
'		Set flagAddSheet to authorize creation of specified sheet
'		Set flagAddParam to authorize creation of specified parameter
'@ReturnValue If the specified parameter already existed, the previous value; otherwise 'Null'
Public Function MySetParameter(ByVal theSheet, ByVal theCol, ByVal theRow, ByVal theValue, ByVal theFlags)

	' Declarations
	Dim sheetRef
	Dim paramRef

	Wait 0
	On Error Resume Next

	' Initialize result
	MySetParameter = Null

	' Get reference to specified parameter
	Set paramRef = MyGetParamRef(theSheet, theCol)
	' If parameter doesn't exist
	If (paramRef Is Nothing) Then
		' If authorized to add parameter
		If (theFlags And flagAddParam) Then
			' Get reference to specified sheet
			Set sheetRef = MyGetSheetRef(theSheet)
			' If sheet doesn't exist
			If (sheetRef Is Nothing) Then
				' If authorized to add sheet
				If (theFlags And flagAddSheet) Then
					' Add specified sheet
					Set sheetRef = DataTable.AddSheet(theSheet)
				Else ' otherwise (can't add sheet)
					' Log the failure
					Reporter.ReportEvent micFail, "MySetParameter", "Sheet [" & theSheet & "] is absent and caller didn't authorize adding it"
					' Exit the test
					ExitTest ("Sheet absent")
				End If
			End If

			' Add new parameter and set row 1 to specified value
			Set paramRef = sheetRef.AddParameter(theCol, theValue)
			' If default row indicated, get current row
			If IsNull(theRow) Then theRow = sheetRef.GetCurrentRow

			' If not setting row 1
			If (theRow > 1) Then
				' Clear out row 1
				paramRef.ValueByRow(1) = Null
				' Set parameter value at indicated row
				paramRef.ValueByRow(theRow) = theValue
			End If
		Else ' otherwise (can't add parameter)
			' Log the failure
			Reporter.ReportEvent micFail, "MySetParameter", "Parameter [" & theCol & "] is absent and caller didn't authorize adding it"
			' Exit the test
			ExitTest ("Param absent")
		End If
	Else ' otherwise (parameter exists)
		' If default row indicated
		If IsNull(theRow) Then
			' RESULT: Parameter value at current row
			MySetParameter = paramRef.Value
			' Set parameter value at current row
			paramRef.Value = theValue
		Else ' otherwise (specific row)
			' RESULT: Parameter value at specified row
			MySetParameter = paramRef.ValueByRow(theRow)
			' Set parameter value at specified row
			paramRef.ValueByRow(theRow) = theValue
		End If
	End If

	' Release objects
	Set sheetRef = Nothing
	Set paramRef = Nothing

End Function


'@Description Set the current row of the specified data sheet
'@Documentation Set the current row of <theSheet> to <theRow>
'@Author sbabcoc
'@Date 16-AUG-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theRow, integer, row index (1-based)
Public Sub MySetCurrentRow(ByVal theSheet, ByVal theRow)

	' Declarations
	Dim sheetRef

	Wait 0
	On Error Resume Next
	Err.Clear

	' Get reference to specified sheet
	Set sheetRef = MyGetSheetRef(theSheet)
	' If specified sheet exists
	If Not (sheetRef Is Nothing) Then
		' Set current row as specified
		sheetRef.SetCurrentRow(theRow)
		' If unable to set current row
		If (Err.Number <> 0) Then
			' Log the failure
			Reporter.ReportEvent micFail, "MySetCurrentRow", "Failed to set current row of sheet [" & theSheet & "] to [" & theRow & "]" & vbCR & Err.Description
			' Exit the test
			ExitTest ("SetCurrentRow failed")
		End If
	Else ' otherwise (sheet absent)
		' Log the failure
		Reporter.ReportEvent micFail, "MySetCurrentRow", "Sheet [" & theSheet & "] is absent"
		' Exit the test
		ExitTest ("Sheet absent")
	End If

	' Release sheet object
	Set sheetRef = Nothing

End Sub


'@Description Load the indicated data table sheet from the specified sheet of the Excel workbook at the specified path
'@Documentation Load data table sheet <theSheet> from sheet <theIndex> of the Excel workbook at <thePath>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] thePath, string, path to Excel workbook (relative or absolute)
'@InParameter [in] theIndex, integer, index of workbook sheet to load (1-based)
'@InParameter [in] theFlags, bitmap, flag specification
'		Set flagReplace to authorize replacement of existing sheet
'@ReturnValue A reference to the loaded sheet
Public Function loadSheetFromPath(ByVal theSheet, ByVal thePath, ByVal theIndex, ByVal theFlags)

	' Declarations
	Dim fileSysObj
	Dim pathSpec
	Dim sheetRef
	Dim doLoad

	Wait 0
	On Error Resume Next

	' Get reference to specified sheet
	Set sheetRef = MyGetSheetRef(theSheet)
	' If sheet doesn't exist
	If (sheetRef = Nothing) Then
		' Set 'load' flag
		doLoad = True
		' Add specified sheet
		Set sheetRef = DataTable.AddSheet(theSheet)
	Else ' otherwise (sheet exists)
		' Set 'load' flag if authorized to replace
		doLoad = (theFlags And flagReplace)
	End If

	' If 'load' indicated
	If (doLoad) Then
		' Create file system object
		Set fileSysObj = CreateObject("Scripting.FileSystemObject")
		' If specified path is absolute
		If (fileSysObj.FileExists(thePath)) Then
			' Use path as-is
			pathSpec = thePath
		Else ' otherwise (relative path)
			' Resolve relative path to absolute path
			pathSpec = PathFinder.Locate(thePath)
			' If path resolution fails
			If (pathSpec = "") Then
				' Log the failure
				Reporter.ReportEvent micFail, "loadSheetFromPath", "Path [" & thePath & "] not found. Check QTP Folders tab search list"
				' Exit the test
				ExitTest ("Path not found")
			End If
		End If
	
		' Reset error object
		Err.Clear
		' Import indicated workbook sheet to spec'd data table sheet
		DataTable.ImportSheet pathSpec, theIndex, sheetRef.Name
		' If import fails
		If (Err.Number) Then
			' Log the failure
			Reporter.ReportEvent micFail, "loadSheetFromPath", Err.Description
			' Exit the test
			ExitTest ("Import failed")
		End If
	End If

	' RESULT: Reference to loaded sheet
	Set loadSheetFromPath = sheetRef

	' Release objects
	Set fileSysObj = Nothing
	Set sheetRef = Nothing

End Function


'@Description Get the row index of the indicated parameter containing the specified value
'@Documentation Get the row index of parameter <theCol> on <theSheet> containing <theValue>
'@Author sbabcoc
'@Date 14-MAR-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theCol, string/integer, parameter specifier (name or 1-based index)
'@InParameter [in] theValue, intrinsic, parameter value
'@InParameter [in] theFlags, bitmap, flag specification
'		Set flagEveryRow to request a list of rows containing the specified value
'@ReturnValue If the specified value is found, its row index or list of indexes; otherwise 'Null'
Public Function getParamValueRowIndex(ByVal theSheet, ByVal theCol, ByVal theValue, ByVal theFlags)

	' Declarations
	Dim sheetRef
	Dim paramRef
	Dim rowIndex
	Dim rowArray()
	Dim arrayLen

	Wait 0
	On Error Resume Next

	' Initialize result
	getParamValueRowIndex = Null

	' Get reference to specified sheet
	Set sheetRef = MyGetSheetRef(theSheet)
	' If specified sheet exists
	If Not (sheetRef Is Nothing) Then
		' Initialize parameter reference
		Set paramRef = Nothing
		' Get reference to specified parameter
		Set paramRef = sheetRef.GetParameter(theCol)
		' If specified parameter exists
		If Not (paramRef Is Nothing) Then
			' Initialize array length
			arrayLen = 0
			' redimension row array
			ReDim rowArray(sheetRef.GetRowCount - 1)

			' Iterate over parameter rows
			For rowIndex = 1 to sheetRef.GetRowCount
				' If current row contains desired value
				If (paramRef.ValueByRow(rowIndex) = theValue) Then
					' Add current index to row array
					rowArray(arrayLen) = rowIndex
					' Increment array length
					arrayLen = arrayLen + 1
					' If not returning a list of row indexes, stop looking
					If Not (theFlags And flagEveryRow) Then Exit For
				End If
			Next

			' If value was found
			If (arrayLen) Then
				' If returning a list of row indexes
				If (theFlags And flagEveryRow) Then
					' Discard unused items
					ReDim Preserve rowArray(arrayLen - 1)
					' RESULT: List of row indexes
					getParamValueRowIndex = rowArray
				Else ' otherwise (single index)
					' RESULT: First row index
					getParamValueRowIndex = rowArray(0)
				End If
			End If
		End If
	End If

	' Release objects
	Set sheetRef = Nothing
	Set paramRef = Nothing
	
End Function


'@Description Get parameter dictionary for indicated row of specified sheet
'@Documentation Get parameter dictionary for row <theRow> of sheet <theSheet>
'@Author sbabcoc
'@Date 16-AUG-2011
'@InParameter [in] theSheet, string/integer, sheet specifier (name or 1-based index)
'@InParameter [in] theRow, integer/Null, row index (1-based); 'Null' specifies current row
Public Function getRowRecord(ByVal theSheet, ByVal theRow)

	' Declarations
	Dim sheetRef
	Dim paramDict
	Dim parmCount
	Dim parmIndex
	Dim thisParam
	Dim parmName
	Dim parmValue

	Wait 0
	On Error Resume Next
	Err.Clear

	' Get reference to specified sheet
	Set sheetRef = MyGetSheetRef(theSheet)
	' If specified sheet exists
	If Not (sheetRef Is Nothing) Then
		' Get current row
		oldRow = sheetRef.GetCurrentRow
		' If default row indicated, use current row
		If IsNull(theRow) Then theRow = oldRow
		' If indicated row isn't current
		If (theRow <> oldRow) Then
			 ' Set current row as specified
			sheetRef.SetCurrentRow(theRow)
			' If unable to set current row
			If (Err.Number <> 0) Then
				' Log the failure
				Reporter.ReportEvent micFail, "getRowRecord", "Failed to set current row of sheet [" & theSheet & "] to [" & theRow & "]" & vbCR & Err.Description
				' Exit the test
				ExitTest ("SetCurrentRow failed")
			End If
		End If

		' Create parameter dictionary
		Set paramDict = CreateObject("Scripting.Dictionary")
		' Get count of sheet parameters
		parmCount = sheetRef.GetParameterCount
		' Iterate over parameters
		For parmIndex = 1 to parmCount
			' Get current parameter object
			Set thisParam = sheetRef.GetParameter(parmIndex)
			' Get parameter name
			parmName = thisParam.Name
			' Get parameter value
			parmValue = thisParam.Value
			' Add parameter to dictionary
			paramDict.Add parmName, parmValue
		Next

		' RESULT: Parameter dictionary
		Set getRowRecord = paramDict

		' If current row was altered
		If (oldRow <> theRow) Then
			 ' Restore current row setting
			sheetRef.SetCurrentRow(oldRow)
		End If
	Else ' otherwise (sheet absent)
		' Log the failure
		Reporter.ReportEvent micFail, "getRowRecord", "Sheet [" & theSheet & "] is absent"
		' Exit the test
		ExitTest ("Sheet absent")
	End If

	' Release objects
	Set sheetRef = Nothing
	Set paramDict = Nothing
	Set thisParam = Nothing

End Function

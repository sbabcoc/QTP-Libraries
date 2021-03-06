'-------------------------------------------------------SITE ERROR HANDLER--------------------------------------------------------
Function UnexpectedPopUp(Object)
	If (Object.GetROProperty("micclass") = "Dialog") Then
		If (Object.GetROProperty("IsPopupWindow")) Then
			Set desc = Description.Create
			desc("text").RegularExpression = False
			sDialogName = Object.GetROProperty("text")
			If (InStr(sDialogName, "Internet") > 0) Then
				desc("text").Value = "OK"
				sDetails = sDialogName & " popup encountered.  'OK' button clicked"
			Else
				desc("text").Value = "&Yes"
				sDetails = sDialogName & " popup encountered.  'Yes' button clicked"
			End If

			Object.Activate

			Set kids = Object.ChildObjects(desc)
			If (kids.Count = 1) Then
				kids(0).Click
			Else
				sDetails = sDialogName & " popup encountered.  'Return' key typed"
				Object.Type micReturn
			End If
			
			Reporter.ReportEvent micWarning, "Pop up", sDetails

			Set desc = Nothing
			Set kids = Nothing
		End If
	End If
End Function

'-------------------------------------------------------------------------------------------------------
'ErrRecoveryScenarioHandler(Object, Method, Arguments, retVal)
'-------------------------------------------------------------------------------------------------------
'Desc: The function is called when a tests hits an error and the scenario 
'				recovery is called.  The Function counts the number of error and if it hits the 
'				set limit for error, an error message is reported and tests is stopped
'
'Args:	The following objects are needed but not really used
'			Object
'			Method,
'			Arguments
'			retVal
'----------------------------------------------------------------------------------------------------------
Function ErrRecoveryScenarioHandler(Object, Method, Arguments, retVal)

    iErrorNumber =  Environment("ErrorCount") 

	iErrorNumber = iErrorNumber + 1
	iErrorCountLimit = cInt(Environment("ErrorCountLimit") )
	bExitAction = False

    sPath = Environment("ResultDir") & "\Report\"
	sImageName = "Error" & iErrorNumber & ".png"

	If Browser("index:=0").Exist(0.5) Then
		Browser("index:=0").CaptureBitmap sPath & sImageName, True
	Else
		Desktop.CaptureBitmap sPath & sImageName, True
	End If

	sXML = "]]></Disp><BtmPane ><Path><![CDATA[" & sImageName & "]]></Path></BtmPane><Disp>"

    If iErrorNumber = iErrorCountLimit Then
		Reporter.ReportEvent micFail, "QA Found An Error" & sXML, "Encountered " & iErrorNumber & ", stopping Test" 
		ExitTest
   End If

   If bExitAction Then
	   ExitAction
	Else
		Reporter.ReportEvent micFail, "QA Found An Error" & sXML, "Encountered " & iErrorNumber & " errors on the test" 
   End If
	Environment("ErrorCount") = iErrorNumber

End Function 

#$Language="VBScript"
#$Interface="1.0"
' Connect-DetectErrorConnecting.vbs

Dim g_strError

Sub Main()
	Dim pass1, pass2 , IP
	strExcelPath = "E:\TEST.xlsx"
	Set objExcel = CreateObject("Excel.Application")
	objExcel.WorkBooks.Open strExcelPath
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	
	Set re1 = New RegExp
	re1.Pattern = "INTEGER:(.*)"
	re1.IgnoreCase = True
	re1.Global = True
	
	If nResult <> 0 Then
		objExcel.Visible = True		
	Else
		crt.Screen.WaitForString "#"
		crt.Screen.Send chr(13)
		'crt.Screen.WaitForString "#"
		'crt.Screen.Send "cd /usr/local/nagios/libexec/" & chr(13)
		'crt.Screen.WaitForString "libexec]#"
		For intRow = 2 to 2000
			IP = objExcel.Cells(intRow,1).Value
			pass1 = objExcel.Cells(intRow,2).Value
			pass2 = objExcel.Cells(intRow,3).Value
			strConnectInfo = "/SSH2 /L root /PASSWORD " & pass1 & IP  
			nResult = Connect(strConnectInfo)
			
			
			IP = objExcel.Cells(intRow,2).Value
			For i = 0 to 15
				crt.Screen.Send "snmpwalk -v 1 -c FPTHCM " & IP & " " & currentandvoltagealams & i & chr(13)
				strcompleteoutput = ""
				strcompleteoutput = crt.Screen.ReadString("#")
				Set matches1 = re1.Execute(strcompleteoutput)
				For Each match In matches1
					temp1 = match.SubMatches(0)
					objExcel.Visible = True
					objExcel.Cells(intRow,i+23).Value = temp1	
				Next
				
				crt.Screen.Send "snmpwalk -v 1 -c FPTHCM " & IP & " " & acalarms & i & chr(13)
				strcompleteoutput = ""
				strcompleteoutput = crt.Screen.ReadString("#")
				Set matches1 = re1.Execute(strcompleteoutput)
				For Each match In matches1
					temp1 = match.SubMatches(0)
					objExcel.Visible = True
					objExcel.Cells(intRow,i+39).Value = temp1	
				Next
				
				crt.Screen.Send "snmpwalk -v 1 -c FPTHCM " & IP & " " & monitoralarms & i & chr(13)
				strcompleteoutput = ""
				strcompleteoutput = crt.Screen.ReadString("#")
				Set matches1 = re1.Execute(strcompleteoutput)
				For Each match In matches1
					temp1 = match.SubMatches(0)
					objExcel.Visible = True
					objExcel.Cells(intRow,i+55).Value = temp1	
				Next
				
				crt.Screen.Send "snmpwalk -v 1 -c FPTHCM " & IP & " " & rectifierconverterinverterfailalarms & i & chr(13)
				strcompleteoutput = ""
				strcompleteoutput = crt.Screen.ReadString("#")
				Set matches1 = re1.Execute(strcompleteoutput)
				For Each match In matches1
					temp1 = match.SubMatches(0)
					objExcel.Visible = True
					objExcel.Cells(intRow,i+71).Value = temp1	
				Next
				
				crt.Screen.Send "snmpwalk -v 1 -c FPTHCM " & IP & " " & voltagealarms & i & chr(13)
				strcompleteoutput = ""
				strcompleteoutput = crt.Screen.ReadString("#")
				Set matches1 = re1.Execute(strcompleteoutput)
				For Each match In matches1
					temp1 = match.SubMatches(0)
					objExcel.Visible = True
					objExcel.Cells(intRow,i+87).Value = temp1	
				Next
			Next
		Next
	End If

	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Quit
End Sub


Function Connect(strConnectInfo)

    g_strError = ""
    On Error Resume Next
        crt.Session.Connect strConnectInfo
        nError = Err.Number
        strErr = Err.Description
    On Error Goto 0
    Connect = nError
    If nError <> 0 Then
        g_strError = strErr
    End If
End Function
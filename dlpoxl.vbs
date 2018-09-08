Option Explicit

Dim goFS    : Set goFS    = CreateObject("Scripting.FileSystemObject")
Dim goWAN   : Set goWAN   = WScript.Arguments.Named
Dim goExcel : Set goExcel = Nothing
Dim goWBook : Set goWBook = Nothing
Dim gnRet   : gnRet       = 1
Dim gaErr   : gaErr       = Array(0, "", "")

Dim WSh     : Set WSh = WScript.CreateObject("WScript.Shell")



On Error Resume Next
    WScript.Echo "Beginning"
    gnRet = Main()
    WScript.Echo "End"
    gaErr = Array(Err.Number, Err.Source, Err.Description)
    If Not goExcel Is Nothing Then
        goExcel.Quit
        Set goExcel = nothing
    End If
On Error GoTo 0
WScript.Echo gaErr(0) & " " & gaErr(1) & " " & gaErr(2)
WScript.Quit gnRet

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SendKeysTo (keys, wait)
	WSh.SendKeys keys
	Wscript.Sleep wait
End Function
'ABOVE IS THE ALL-MIGHTY FUNCTION!!!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getData (workbook, ByRef poLinesArr)
	Dim size: size = UBound(poLinesArr, 2)
    Dim length: length = workbook.Sheets(1).UsedRange.Rows.Count-1
    ReDim Preserve poLinesArr(15, length + size)

    If poLinesArr(0,0) = "" Then
        Dim j, k
        For j = 0 To length
			Dim r = 0
			For k = 65 To 79
				
				If workbook.Sheets(1).Range(Chr(k) & j) Is Not "" Then
					poLinesArr(r,j) = workbook.Sheets(1).Range("A" & j) '1
				else
					poLinesArr(r,j) = "null"
				End If
				r = r + 1
			Next
		Next
	else
		Dim j, k
        For j = size To size + length+1
			Dim r = 0
			For k = 65 To 79
				
				If workbook.Sheets(1).Range(Chr(k) & j) Is Not "" Then
					poLinesArr(r,j) = workbook.Sheets(1).Range("A" & j) '1
				else
					poLinesArr(r,j) = "null"
				End If
				r = r + 1
			Next
		Next
    End If
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Main()
  Main = 1

Dim poLinesArr(15,1)
Dim RowCount
Dim Always  : Always = true
ReDim poArr(1)

	'open POList.xlsx
	Set goExcel = CreateObject("Excel.Application")
	Set goWBook = goExcel.Workbooks.Open("C:\Users\Christopher.Sutton\Desktop\Assignment\purchaseOrders\POList.xlsx")
	
	'check length and redim poArr to fit file length
	RowCount = goWBook.Sheets(1).UsedRange.Rows.Count
	ReDim poArr(RowCount)
	'copy contents to poArr
	Dim i
	For i = 0 To RowCount-1
            With goWBook.Worksheets(1)
		poArr(i) = .Range("A" & i+1).Value
		WScript.Echo poArr(i)
            End With
	Next
	goWBook.Close False
	''''Array of POs have been acquired
	
	'make a for loop with int size of poArr
	For i = 0 To RowCount - 1
		'before while loop, use SendKeysTo to maneuver to, write
		'the PO # into the filter, enter, export PO lines
		'add a while loop to catch when the excel file opens with GetObject
		Always = true
		While (Always)
			Set goExcel = GetObject( , "Excel.Application")
			If goExcel Is Nothing Then
				WScript.Sleep 1000
				Set goExcel = GetObject( , "Excel.Application")
			else
				Always = false
			End If
			
		Wend
		Always = true
		
		'get the new workbook (Book1) with po lines use initial check to 
		Set goWBook = goExcel.Workbooks("Book1") 'hopefully get Book1...
		'redim poLinesArr columns (as rows) 
		'if it's the first, take headers too but exclude last row
		'else take 2nd row down -- make sure to transpose rows to columns for
		'dynamic memory allocation

	Next
		
	
		

	'create excel file and transpose and fill it with poLinesArr
	'save and close

  Main = 0
End Function ' Main  
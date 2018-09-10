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
	getData = 1
	Dim size: size = UBound(poLinesArr, 2)
    Dim length: length = workbook.Sheets(1).UsedRange.Rows.Count-1
    ReDim Preserve poLinesArr(15, length + size)
	Dim j, k, r
    If poLinesArr(0,0) = "" Then
        
        For j = 0 To length
			r = 0
			For k = 65 To 79
				
				If workbook.Sheets(1).Range(Chr(k) & j) Is Not "" Then
					poLinesArr(r,j) = workbook.Sheets(1).Range(Chr(k) & j)
				else
					poLinesArr(r,j) = "null"
				End If
				r = r + 1
				k = k + 1
			Next
		Next
	else
        For j = size To size + length+1
			r = 0
			For k = 65 To 79
				
				If workbook.Sheets(1).Range(Chr(k) & j) Is Not "" Then
					poLinesArr(r,j) = workbook.Sheets(1).Range(Chr(k) & j)
				else
					poLinesArr(r,j) = "null"
				End If
				r = r + 1
				k = k + 1
			Next
		Next
    End If
    getData = 0
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Main()
  Main = 1

Dim poLinesArr(15,1)
Dim RowCount
Dim i, j
Dim Always  : Always = true
ReDim poArr(1)

	'open POList.xlsx
	Set goExcel = CreateObject("Excel.Application")
	Set goWBook = goExcel.Workbooks.Open("C:\Users\Christopher.Sutton\Desktop\Assignment\purchaseOrders\POList.xlsx")
	
	'check length and redim poArr to fit file length
	RowCount = goWBook.Sheets(1).UsedRange.Rows.Count
	ReDim poArr(RowCount)
	'copy contents to poArr
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

		SendKeysTo "%{TAB}" 5000
		If i = 0 Then
			SendKeysTo poArr(i) 100
			SendKeysTo "{ENTER}" 3000
		else
			SendKeysTo "^{PgUp}" 100
			SendKeysTo "^{UP}" 100
			SendKeysTo "+{END}" 100
			SendKeysTo "{ENTER}" 3000
		End If
		SendKeysTo "^E" 3000







		'the PO # into the filter, enter, export PO lines
		'add a while loop to catch when the excel file opens with GetObject
		Always = true
		While (Always)
			Set goExcel = GetObject( , "Excel.Application")
			If goExcel Is Nothing Then
				WScript.Sleep 1000
			else
				Always = false
			End If
			
		Wend
		Always = true
		
		'get the new workbook (Book1) with po lines use initial check to 
		Set goWBook = goExcel.Workbooks("Book1") 'hopefully get Book1...
		'redim poLinesArr columns (as rows)
		Dim rowCount: rowCount = goWBook.sheets(1).UsedRange.Row.Count -1
		'if it's the first, take headers too but exclude last row
		Dim result: result = getData(goWBook, poLinesArr)
		'else take 2nd row down -- make sure to transpose rows to columns for
		'dynamic memory allocation
		goWBook.close false
	Next
		
	Dim finalSize: finalSize = UBound(poLinesArr, 2)
	'TESTING ARRAY
	For i = 0 To finalSize
		For j = 0 To 15
			WScript.StdOut.write(poLinesArr(j,i))
		Next
		Wscript.StdOut.write("\n")
	Next
		

	'create excel file and transpose and fill it with poLinesArr
	'save and close

  Main = 0
End Function ' Main  
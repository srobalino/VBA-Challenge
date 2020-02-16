Attribute VB_Name = "Module11"
'Create a script that will loop through all the stocks for one year for each run and take the following information.

'Code must check the ticker value to collect the open and close years for each ticker. EX= ticket value= first_Open_Tciker (first row of ticker value) -Close_of_Ticker (last row of ticket value)

Sub RemoveHiddenRows()
Dim ws As Worksheet
Dim j As Integer

j = 0

For Each ws In ActiveWorkbook.Worksheets

   For iCntr = lastrow To 1 Step -1
    If Rows(iCntr).Hidden = True Then
        Rows(iCntr).EntireRow.Delete
    End If
    Next
Next ws
End Sub

Sub AddTitles()
For Each ws In ThisWorkbook.Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

Next

End Sub

Sub Multiloop()

'Dim ws As Worksheet
'set the variables

Dim ws As Worksheet
Dim ticker As String
Dim Yearly_Change As Integer
Dim totalVol As Double
Dim percent_change As Double
Dim yearly_open As Double
Dim yearly_close As Double
Set ws = ActiveSheet
Dim Summary_Table_Row As Integer
Dim i As Long
Dim TickStartRow As Long
Dim TickEndRow As Long
Summary_Table_Row = 2


'define the last row
lastrow = Cells(Rows.Count, 7).End(xlUp).Row

'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Work through the loop
For Each ws In ThisWorkbook.Worksheets
For i = 2 To lastrow   'why is this only going to 100? What wrong with you! Hidden Cells?

    'checks if stil within the ticker value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    ws.Cells(i, 1).Value = ticker
 
    'Define yearly open and close
    yearly_open = yearly_open + ws.Cells(i, 3).Value
    yearly_close = yearly_close + ws.Cells(i, 6).Value
    
    'get the yearly cahange and the percent change
    Yearly_Change = yearly_close - yearly_open  'check the math-yearly close should be the end of each ticker
    
    percent_change = Yearly_Change / yearly_open * 100
    
    'get the total volume
   
      totalVol = totalVol + Cells(i, 7).Value
    'print result to each column
    
   
    ws.Range("I" & Summary_Table_Row).Value = ticker
    
    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
     
    ws.Range("K" & Summary_Table_Row).Value = percent_change
    ws.Range("L" & Summary_Table_Row).Value = totalVol
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    

    Yearly_Change = 0
    percent_change = 0
    totalVol = 0
    


End If

Next i

'location of the Column 10
changeyear_color = ws.Cells(Rows.Count, 10).End(xlUp).Row

'loop though the year change and add conditional formatting

For x = 2 To changeyear_color

    If ws.Cells(x, 10) <= 0 Then
    ws.Cells(x, 10).Interior.ColorIndex = 3
    
    ElseIf Cells(x, 10) > 1 Then
    ws.Cells(x, 10).Interior.ColorIndex = 4
    
End If



Next x
 
Next ws

'Didnt have time to add the challenge
'' Dim percent_increase As Long_
' Dim percent_decrease As Long
' 'Dim great_vol As Float
'
'
'    ws.Range("P1") = "Ticker"
'    ws.Range("Q1") = "Value"
'
'    ws.Range("P" & Summary_Table_Row).Value = percent_increase
'
'    ws.Range("Q" & Summary_Table_Row).Value = percent_decrease

End Sub

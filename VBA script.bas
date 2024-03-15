Attribute VB_Name = "Module1"
Sub tracker():

'Sheets.Add.Name = "combined_data"
'Sheets("combined_data").Move before:=Sheets(1)

'Set combined_sheet = Worksheets("combined_data")
For Each ws In Worksheets
Dim WorksheetName As String

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name



Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Variant
Dim total As Integer
Dim ticker_counter As Double
Dim summary As Integer
Dim total_counter As LongLong


yearly_change = 0
yearly_counter = 1
percent_change = 0
total_counter = 0
summary = 2


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"



For i = 2 To LastRow

    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        total_counter = total_counter + Cells(i, 7).Value
       'need to keep the first value below
       yearly_change = Cells(i, 3).Value - Cells(i, 6).Value
        percent_change = Cells(i, 6).Value / Cells(i, 3).Value - 1
        
        Range("I" & summary).Value = ticker
       Range("J" & summary).Value = yearly_change
        Range("L" & summary).Value = total_counter
       Range("K" & summary).Value = percent_change
        
        summary = summary + 1
        
        total_counter = 0
        
        Else
       
        
        ticker = Cells(i, 1).Value
        total_counter = total_counter + Cells(i, 7).Value
        
        
        End If
        
        If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        ElseIf Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4

        End If
        
    Next i

Dim increase_vol As LongLong
Dim max_vol As Long
max_vol = 0
Dim increase_per As Long
Dim max_per As Long
max_per = 0
Dim decrease_per As Long
Dim min_per As Long
min_per = 999999

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


For j = 2 To LastRow
    increase_vol = Cells(j, 12).Value
    increase_per = Cells(j, 11).Value
    decrease_per = Cells(j, 10).Value
        If increase_vol > max_num Then
            max_num = increase_vol
            Cells(4, 17).Value = max_num
            Cells(4, 16).Value = Cells(j, 9).Value
            ElseIf increase_vol < max_num Then
            max_num = max_num
        End If
        
                If increase_per > max_per Then
            max_per = increase_per
            Cells(2, 17).Value = max_per
            Cells(4, 16).Value = Cells(j, 9).Value
            ElseIf increase_per < max_per Then
            max_per = max_per
        End If
                
                If decrease_per < min_per Then
            min_per = decrease_per
            Cells(3, 17).Value = min_per
            Cells(4, 16).Value = Cells(j, 9).Value
            ElseIf decrease_per < min_per Then
            min_per = min_per
        End If
        
Next j

Next ws

End Sub








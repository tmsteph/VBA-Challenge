Sub loopThroughStocks()

Dim Ticker As String

Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Single
Dim summary_Table_Row As Integer
summary_Table_Row = 2
Dim Opening_Price As Double
Dim Closing_Price As Double
''' Make column headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 11).Value = "Percentage Change"

'''Logic Time!!
    For i = 2 To 1000000
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then ' if new section
        
            Ticker = Cells(i, 1).Value 'Assign ticker Value
            Total_Stock_Volume = Cells(i, 7).Value 'assign and reset first volume of new group
            Range("I" & summary_Table_Row).Value = Ticker 'print Ticker
            Opening_Price = Cells(i, 3).Value 'set opening price of group
            
           
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'if at end of section
            
            Closing_Price = Cells(i, 6).Value
            '''Yearly Change Logic
            Yearly_Change = Closing_Price - Opening_Price
            If Yearly_Change = 0 Then
                Percent_Change = 0
            ElseIf Opening_Price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Opening_Price * 100
            End If
            
            
            Range("L" & summary_Table_Row).Value = Total_Stock_Volume  'printTotal stock Volume
            Range("K" & summary_Table_Row).Value = Round(Percent_Change, 2) & "%" ' Print Percent Change
            Range("J" & summary_Table_Row).Value = Yearly_Change 'print yearly Change
            If Yearly_Change > 0 Then
                Range("J" & summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            ElseIf Yearly_Change < 0 Then
                Range("J" & summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If
            summary_Table_Row = summary_Table_Row + 1 'advance row count
        
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value 'add i stock to total
            
            
        End If
        
          
    Next i
    
End Sub

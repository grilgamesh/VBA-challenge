Attribute VB_Name = "Module1"
Sub VBA_homework()
    Dim ticker As String
    Dim openingvalue As Double
    Dim closingvalue As Double
    Dim annualchange As Double
    Dim percentchange As String
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
   
   'create blank space to put summary data into
    Sheets.Add(Before:=Sheets(1)).Name = "Summary output"
    Sheets("Summary output").Cells(1, 1).Value = "Ticker"
    Sheets("Summary output").Cells(1, 2).Value = "Yearly Change"
    Sheets("Summary output").Cells(1, 3).Value = "Percentage Change"
    Sheets("Summary output").Cells(1, 4).Value = "Total Stock Volume"
 
                 

    For Each ws In Worksheets
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openingvalue = ws.Cells(2, 3).Value
        For i = 2 To lastRow
            volume = volume + ws.Cells(i, 7).Value
            ' Check if we are on the last line of the current stock,
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'update variables
                ticker = ws.Cells(i, 1).Value
                closingvalue = ws.Cells(i, 6).Value
                annualchange = openingvalue - closingvalue
                percentchange = FormatPercent((annualchange / openingvalue), 2)
                
                'export summary data to final page
                'MsgBox (ticker + ", " + Str(annualchange) + ", " + Str(percentchange) + ", " + Str(volume))
                 Sheets("Summary output").Cells(Summary_Table_Row, 1).Value = ticker
                 Sheets("Summary output").Cells(Summary_Table_Row, 2).Value = annualchange
                 Sheets("Summary output").Cells(Summary_Table_Row, 3).Value = percentchange
                 Sheets("Summary output").Cells(Summary_Table_Row, 4).Value = volume
                              

                ' increment the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'reset variables
                volume = 0
                openingvalue = ws.Cells(i + 1, 3).Value
            End If
        Next i
    Next
    
    'set conditional formatting for yearly change column
    Sheets("Summary output").Range("B:B").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Sheets("Summary output").Range("B:B").FormatConditions(1).Interior.Color = vbRed
    Sheets("Summary output").Range("B:B").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    Sheets("Summary output").Range("B:B").FormatConditions(2).Interior.Color = vbGreen
    Sheets("Summary output").Range("B1").FormatConditions.Delete
    
    'calculate superstars and pooperstars
    
    Dim increasest As String
    Dim maxIncrease As Double
        
    Dim decreasest As String
    Dim maxDecrease As Double
    
    Dim volumest As String
    Dim maxVolume As Double
    
    summaryLastRow = Sheets("Summary output").Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To summaryLastRow
        If Sheets("Summary output").Cells(i, 3).Value > maxIncrease Then
            increasest = Sheets("Summary output").Cells(i, 1).Value
            maxIncrease = Sheets("Summary output").Cells(i, 3).Value
        ElseIf Sheets("Summary output").Cells(i, 3).Value < maxDecrease Then
            decreasest = Sheets("Summary output").Cells(i, 1).Value
            maxDecrease = Sheets("Summary output").Cells(i, 3).Value
        End If

        If Sheets("Summary output").Cells(i, 4).Value > maxVolume Then
            volumest = Sheets("Summary output").Cells(i, 1).Value
            maxVolume = Sheets("Summary output").Cells(i, 4).Value
        End If
    Next i
    

    
    'set up super-summary
    
    Sheets("Summary output").Cells(1, 7).Value = "Ticker"
    Sheets("Summary output").Cells(1, 8).Value = "Value"
    
    Sheets("Summary output").Cells(2, 6).Value = "Greatest % increase"
    Sheets("Summary output").Cells(2, 7).Value = increasest
    Sheets("Summary output").Cells(2, 8).Value = FormatPercent(maxIncrease, 2)
    
    
    Sheets("Summary output").Cells(3, 6).Value = "Greatest % decrease"
    Sheets("Summary output").Cells(3, 7).Value = decreasest
    Sheets("Summary output").Cells(3, 8).Value = FormatPercent(maxDecrease, 2)


    Sheets("Summary output").Cells(4, 6).Value = "Greatest total volume"
    Sheets("Summary output").Cells(4, 7).Value = volumest
    Sheets("Summary output").Cells(4, 8).Value = maxVolume
    
End Sub





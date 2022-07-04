Attribute VB_Name = "VbaChallenge"
Sub YearlyStocks()

    Dim WS_Count As Integer
    Dim w As Integer

    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.count

    ' Begin to loop through the worksheets.
    For w = 1 To WS_Count

        Dim openPrice As Double
        Dim closePrice As Double
        Dim change As Double
        Dim changeRatio As Double
        Dim count As Long
        Dim totalVolume As LongLong
        Dim iRange As String
        Dim i As Long
        Dim Worksheet As String
        Dim openRow As Long
        
        'Getting the name of the worksheet and activating it to run the code below
        ' on each worksheet one by one
        Worksheet = Worksheets(w).Name
        Worksheets(Worksheet).Activate
                
        
        'Constant for Conditional Formating
        Dim rg As Range
        Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    
        'Making Header for the results
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        Columns("I:L").AutoFit
        
        count = 2
        openRow = 2
        For i = 2 To Range("B1").End(xlDown).Row
            
            'Checking for the change in ticker name
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Taking Ticker Name and Opening price of the stock
                ticker = Cells(i, 1).Value
                closePrice = Cells(i, 6).Value
                openPrice = Cells(openRow, 3).Value
                
                
                change = closePrice - openPrice
                changeRatio = change / openPrice
            
                'Adding values to the cell
                Cells(count, 9).Value = ticker
                Cells(count, 10).Value = change
                Cells(count, 11).Value = FormatPercent(changeRatio, 2)
                
                'Calculating Total Stock Volume by adding the Cell range from open to close
                closeCellNum = i
                iRange = "G" & openRow & ":" & "G" & closeCellNum
                Cells(count, 12).Value = Application.WorksheetFunction.Sum(Range(iRange))
                
                openRow = i + 1
                count = count + 1
            
            End If
                            
        Next i
        
        
        'Conditional Formating
        Set rg = Range("J2", Range("J2").End(xlDown))

        'Deleting Any existing Conditional Formating
        rg.FormatConditions.Delete

        'Adding the rules for each conditional formating
        Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)

        With cond1
        .Interior.Color = vbGreen
        End With

        With cond2
        .Interior.Color = vbRed
        End With
        
        
        '******************************************Bonus*******************************************
        
        Dim numRows As Long
        Dim gPerctIncrease As Double
        Dim gPerctDecrease As Double
        Dim gTotalVol As LongLong

        'Creating the label
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

        Cells(2, 15).Value = "Greatest % increase"
        Cells(3, 15).Value = "Greatest % decrease"
        Cells(4, 15).Value = "Greatest total volume"

        Columns("O").AutoFit

        'Adding Greatest Increase, Greatest Decrease and Greatest Volume Values

        numRows = Range("K1").End(xlDown).Row

        gPerctIncrease = Application.WorksheetFunction.Max(Range("K2:K" & numRows))
        gPerctDecrease = Application.WorksheetFunction.Min(Range("K2:K" & numRows))
        gTotalVol = Application.WorksheetFunction.Max(Range("L2:L" & numRows))

        'Adding Values in the cells
        Cells(2, 17).Value = FormatPercent(gPerctIncrease, 2)
        Cells(3, 17).Value = FormatPercent(gPerctDecrease, 2)
        Cells(4, 17).Value = gTotalVol

        'Getting the row numbers for the Greatest Increase, Greatest Decrease and Greatest Volume Values
        gPerctIncRow = Application.WorksheetFunction.Match(gPerctIncrease, Range("K2:K" & numRows), 0)
        gPerctDecRow = Application.WorksheetFunction.Match(gPerctDecrease, Range("K2:K" & numRows), 0)
        gTotalVolRow = Application.WorksheetFunction.Match(gTotalVol, Range("L2:L" & numRows), 0)

        'Reading Ticker Values and Adding it in the cells
        Cells(2, 16).Value = Cells(gPerctIncRow + 1, 9).Value
        Cells(3, 16).Value = Cells(gPerctDecRow + 1, 9).Value
        Cells(4, 16).Value = Cells(gTotalVolRow + 1, 9).Value


    Next w
    
End Sub





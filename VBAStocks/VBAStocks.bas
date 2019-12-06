Attribute VB_Name = "Module1"
Sub runWorksheet()

Dim xSheet As Worksheet
Application.ScreenUpdating = False
For Each xSheet In Worksheets
    xSheet.Select
    Call stockAnalysis
Next
Application.ScreenUpdating = True

End Sub
Sub stockAnalysis()

' add header titles for summary
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
counter = 1
uniqueCount = 0
totalUniqueVolume = 0

' populate ticker symbol
For i = 2 To lastrow

    ' count unique tickers
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        uniqueCount = uniqueCount + 1
    End If
    
    ' find next ticker symbol that doesn't match
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        counter = counter + 1 ' increase counter to place unique value in proper row
        lastUniqueTickerRow = i ' assign last unique ticker row to current index
        firstUniqueTickerRow = lastUniqueTickerRow - uniqueCount ' assign first unique ticker row
        
        ' place unique tickers and stock analysis values in proper columns
        Cells(counter, 9).Value = Cells(i, 1).Value
        Cells(counter, 10).Value = Cells(lastUniqueTickerRow, 6).Value - Cells(firstUniqueTickerRow, 3).Value
        
        If Cells(firstUniqueTickerRow, 3).Value <> 0 Then
            Cells(counter, 11).Value = (Cells(lastUniqueTickerRow, 6).Value - Cells(firstUniqueTickerRow, 3).Value) / Cells(firstUniqueTickerRow, 3).Value
        Else
            Cells(counter, 11).Value = 0
        End If
        Cells(counter, 11).NumberFormat = "0.00%"
        
        For j = firstUniqueTickerRow To lastUniqueTickerRow
            totalUniqueVolume = totalUniqueVolume + Cells(j, 7).Value
        Next j
        
        Cells(counter, 12).Value = totalUniqueVolume
    
        uniqueCount = 0 ' reset for next unique ticker
        totalUniqueVolume = 0 ' reset for next unique ticker
    End If
    
Next i

Call formatting

End Sub

Sub formatting()

Dim rng As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition

Set rng = Range("J2", Range("J2").End(xlDown))

' clear any existing conditional formatting
rng.FormatConditions.Delete

' define the rule for each conditional format
Set condition1 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
Set condition2 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")

' define the format applied for each conditional format
With condition1
   .Interior.Color = vbRed
End With

With condition2
   .Interior.Color = vbGreen
End With

End Sub

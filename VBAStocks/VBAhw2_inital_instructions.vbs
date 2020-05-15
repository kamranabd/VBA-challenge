Attribute VB_Name = "Module1"
Sub VBAhw2()
    Dim Ticker_Name As String
    Dim Summary_Table_Row As Integer
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Percent_Change As Double
    Dim Year_Change As Double
    Dim Vol_Total As Variant
    
    Summary_Table_Row = 2
    Year_Open = Cells(2, 3).Value
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Vol_Total = 0
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("K2:K" & LastRow).NumberFormat = "0.00%"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To LastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            Year_Close = Cells(i, 6).Value
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Year_Change = Year_Close - Year_Open
            
            Range("J" & Summary_Table_Row).Value = Year_Change
            If Percent_Change >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            Percent_Change = Year_Change / Year_Open
            Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Vol_Total = Vol_Total + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = Vol_Total
            
            Summary_Table_Row = Summary_Table_Row + 1
            Year_Open = Cells(i + 1, 3).Value
            Vol_Total = 0
        Else
            Vol_Total = Vol_Total + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub

Attribute VB_Name = "Module1"
Sub VBAhw2Challenge1()
    Dim Ticker_Name As String
    Dim Summary_Table_Row As Integer
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Percent_Change As Double
    Dim Year_Change As Double
    Dim Vol_Total As Variant
    
    Dim Largest_Percent_Increase As Double
    Dim Largest_Percent_Decrease As Double
    Dim Ticker_Value_Increase As String
    Dim Ticker_Value_Decrease As String
    Dim Greatest_Stock_Vol As Variant
    Dim Ticker_Value_Vol As String
    
    Largest_Percent_Increase = 0
    Largest_Percent_Decrease = 0
    Greatest_Stock_Vol = 0
    
    Summary_Table_Row = 2
    Year_Open = Cells(2, 3).Value
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Vol_Total = 0
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("K2:K" & LastRow).NumberFormat = "0.00%"
    Range("L1").Value = "Total Stock Volume"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    For i = 2 To LastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            Year_Close = Cells(i, 6).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Year_Change = Year_Close - Year_Open
            
            Range("J" & Summary_Table_Row).Value = Year_Change
            
            Percent_Change = Year_Change / Year_Open
            Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Vol_Total = Vol_Total + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = Vol_Total
            
            If Percent_Change >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            Year_Open = Cells(i + 1, 3).Value
            Vol_Total = 0
        Else
            Vol_Total = Vol_Total + Cells(i, 7).Value
            
        End If
        
    Next i
    
    For j = 2 To LastRow
        If Cells(j, 11).Value > Largest_Percent_Increase Then
            Largest_Percent_Increase = Cells(j, 11).Value
            Ticker_Value_Increase = Cells(j, 9).Value
        ElseIf Cells(j, 11).Value < Largest_Percent_Decrease Then
            Largest_Percent_Decrease = Cells(j, 11).Value
            Ticker_Value_Decrease = Cells(j, 9).Value
            
        End If
    Next j
    
    For k = 2 To LastRow
        If Cells(k, 12).Value > Greatest_Stock_Vol Then
            Greatest_Stock_Vol = Cells(k, 12).Value
            Ticker_Value_Vol = Cells(k, 9).Value
        End If
    Next k
    
    Range("Q2").Value = Largest_Percent_Increase
    Range("P2").Value = Ticker_Value_Increase
    Range("Q3").Value = Largest_Percent_Decrease
    Range("P3").Value = Ticker_Value_Decrease
    Range("P4").Value = Ticker_Value_Vol
    Range("Q4").Value = Greatest_Stock_Vol
     
End Sub

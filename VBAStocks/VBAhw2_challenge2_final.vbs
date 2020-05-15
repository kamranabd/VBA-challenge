Attribute VB_Name = "Module1"
Sub VBAhw2Challenge2()
    Dim ws As Variant
    For Each ws In Worksheets
    
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
    Year_Open = ws.Cells(2, 3).Value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Vol_Total = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker_Name = ws.Cells(i, 1).Value
            Year_Close = ws.Cells(i, 6).Value
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            Year_Change = Year_Close - Year_Open
            
            ws.Range("J" & Summary_Table_Row).Value = Year_Change
            
            If Year_Open <> 0 Then
                Percent_Change = Year_Change / Year_Open
            Else
                Percent_Change = 0
            End If
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Vol_Total
            
            If Percent_Change >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            Year_Open = ws.Cells(i + 1, 3).Value
            Vol_Total = 0
        Else
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    For j = 2 To LastRow
        If ws.Cells(j, 11).Value > Largest_Percent_Increase Then
            Largest_Percent_Increase = ws.Cells(j, 11).Value
            Ticker_Value_Increase = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value < Largest_Percent_Decrease Then
            Largest_Percent_Decrease = ws.Cells(j, 11).Value
            Ticker_Value_Decrease = ws.Cells(j, 9).Value
            
        End If
    Next j
    
    For k = 2 To LastRow
        If ws.Cells(k, 12).Value > Greatest_Stock_Vol Then
            Greatest_Stock_Vol = ws.Cells(k, 12).Value
            Ticker_Value_Vol = ws.Cells(k, 9).Value
        End If
    Next k
    
    ws.Range("Q2").Value = Largest_Percent_Increase
    ws.Range("P2").Value = Ticker_Value_Increase
    ws.Range("Q3").Value = Largest_Percent_Decrease
    ws.Range("P3").Value = Ticker_Value_Decrease
    ws.Range("P4").Value = Ticker_Value_Vol
    ws.Range("Q4").Value = Greatest_Stock_Vol
  
    Next ws
  
End Sub

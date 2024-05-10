Attribute VB_Name = "Module1"
Sub Stock_Data()
  
    Dim ws_num As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim Total_stock_volume As Double
    Dim percentageChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim i As Long
    Dim Stock_summary_table As Integer
    Dim summaryTableStartRow As Long
    
    ws_num = ThisWorkbook.Worksheets.Count
    
    For i = 1 To ws_num
      Set ws = ThisWorkbook.Worksheets(i)
      ws.Activate
      
      Stock_summary_table = 2
      Total_stock_volume = 0
      greatestIncrease = -999999999999#
      greatestDecrease = 999999999999#
      greatestVolume = 0
    

        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        summaryTableStartRow = 1
        
        ws.Cells(summaryTableStartRow, "I").Value = "Ticker"
        ws.Cells(summaryTableStartRow, "J").Value = "Quarterly Change"
        ws.Cells(summaryTableStartRow, "K").Value = "Percent Change"
        ws.Cells(summaryTableStartRow, "L").Value = "Total Stock Volume"
        
        ' Loop through all rows of data
        For j = 2 To lastRow
            
            ' Check if the ticker symbol changes
            If ws.Cells(j, 1).Value <> ws.Cells(j - 1, 1).Value Then
                ' If it's a new ticker symbol
                    
                ' Output the previous ticker symbol and its total stock volume to the summary table
                If ticker <> "" Then
                    ws.Range("I" & Stock_summary_table).Value = ticker
                    ws.Range("J" & Stock_summary_table).Value = closingPrice - openingPrice
                    ws.Range("K" & Stock_summary_table).Value = Format((closingPrice - openingPrice) / openingPrice, "0.00%")
                    ws.Range("L" & Stock_summary_table).Value = Total_stock_volume
                    Stock_summary_table = Stock_summary_table + 1
                    
                    If percentageChange > greatestIncrease Then
                        greatestIncrease = percentageChange
                        greatestIncreaseTicker = ticker
                    End If
                    
                    If percentageChange < greatestDecrease Then
                        greatestDecrease = percentageChange
                        greatestDecreaseTicker = ticker
                    End If
                    
                    If Total_stock_volume > greatestVolume Then
                        greatestVolume = Total_stock_volume
                        greatestVolumeTicker = ticker
                    End If
                End If
                
                ' Set the ticker symbol for the current row
                ticker = ws.Cells(j, 1).Value
                
                ' Reset the total stock volume for the new ticker symbol
                Total_stock_volume = 0
                
                openingPrice = ws.Cells(j, 3).Value
            End If
            
            ' Accumulate the total stock volume for the current ticker symbol
            Total_stock_volume = Total_stock_volume + ws.Cells(j, 7).Value
            
            closingPrice = ws.Cells(j, 6).Value
            
            percentageChange = ((closingPrice - openingPrice) / openingPrice)
            
        Next j
        
        ' Output the last ticker symbol and its total stock volume to the summary table
        If ticker <> "" Then
            ws.Range("I" & Stock_summary_table).Value = ticker
            ws.Range("J" & Stock_summary_table).Value = closingPrice - openingPrice
            ws.Range("K" & Stock_summary_table).Value = Format((closingPrice - openingPrice) / openingPrice, "0.00%")
            ws.Range("L" & Stock_summary_table).Value = Total_stock_volume
            
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = ticker
            End If
            
            If Total_stock_volume > greatestVolume Then
                greatestVolume = Total_stock_volume
                greatestVolumeTicker = ticker
            End If
        End If
        
        ' Output the results to the specified cells
        ws.Range("N2").Value = "Greatest%Increase"
        ws.Range("N3").Value = "Greatest%Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        ' Output ticker and value pairs to Columns O and P with headers
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("O2").Value = greatestIncreaseTicker
        ws.Range("P2").Value = Format(greatestIncrease, "0.00%")
        ws.Range("O3").Value = greatestDecreaseTicker
        ws.Range("P3").Value = Format(greatestDecrease, "0.00%")
        ws.Range("O4").Value = greatestVolumeTicker
        ws.Range("P4").Value = greatestVolume

    ' Apply conditional formatting
    With ws.Range("J2:J" & lastRow)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
    End With
   Next i
End Sub


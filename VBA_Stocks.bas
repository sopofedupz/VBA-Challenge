Attribute VB_Name = "Module5"
Sub Summary_WallStreetStock()
    
    For Each ws In Worksheets
    
            '----------------------------------------------------------------------------------------------------------------------------
            ' Prepping the dataset - Start by sorting the columns based on Ticker and Date
            '----------------------------------------------------------------------------------------------------------------------------
            
            ' Look for the last row and last column
            Dim LastRow As Long
            Dim LastCol As Long
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
            
            ' Converting the column index number to the column letter
            ' Source = TheSpreadsheetGuru
            ' (https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number)
            Dim ColumnNumber As Long
            Dim ColumnLetter As String
        
            'Input Column Number
            ColumnNumber = LastCol
        
            'Convert To Column Letter
            ColumnLetter = Split(ws.Cells(1, ColumnNumber).Address, "$")(1)
            
            ' Sort the columns
            ws.Range("A1:" & ColumnLetter & LastRow).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Key2:=ws.Range("B1") _
                , Order2:=xlAscending, Header:=xlYes
                                            
            '--------------------------------------------------------------------------------------------------------------------------
            ' Summarize the data
            '--------------------------------------------------------------------------------------------------------------------------
        
            ' Start by creating the column headers of the summary table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
        
            ' Set an initial variable for holding the ticker symbol
            Dim Ticker_Symbol As String
            
            ' Set an initial variable for open, close, and volume numbers
            Dim OpeningPrice As Double
            Dim ClosingPrice As Double
            Dim Volume_Total As Variant
            Dim Yearly_Change As Double
            Dim Percent_Change As Variant
            
            OpeningPrice = 0
            ClosingPrice = 0
            Yearly_Change = 0
            Percent_Change = 0
            Volume_Total = 0
            
            ' Set an initial variable for total count per unique ticker symbol
            Dim Ticker_Count As Long
            
            Ticker_Count = 0
            
            ' Keep track of the location for each ticker symbol in the summary table
            Dim Summary_Table_Row As Integer
            
            Summary_Table_Row = 2
              
            ' Loop through all rows
            For i = 2 To LastRow
            
            ' Check if we are still within the same ticker symbol
            ' If the following cell does not have the same ticker symbol as the previous cell, then:
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                  ' Set the Ticker Symbol
                  Ticker_Symbol = ws.Cells(i, 1).Value
        
                  ' Get closing price of the year and total volume for the year
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                  Ticker_Count = Ticker_Count + 1
                  
                  ClosingPrice = ws.Cells(i, 6).Value
                  OpeningPrice = ws.Cells(i - Ticker_Count + 1, 3).Value
                  Yearly_Change = ClosingPrice - OpeningPrice
                  
                  If (OpeningPrice = 0) Then
                    Percent_Change = ""
                  Else
                    Percent_Change = Yearly_Change / OpeningPrice
                  End If
        
                  ' Print the ticker symbol and values in the Summary Table
                  ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                  ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                  ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                  ws.Range("L" & Summary_Table_Row).Value = Volume_Total
        
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
            
                  ' Reset all the values back to 0
                  OpeningPrice = 0
                  ClosingPrice = 0
                  Yearly_Change = 0
                  Percent_Change = 0
                  Volume_Total = 0
                  Ticker_Count = 0
            
                ' If the cell immediately following a row is the same brand...
                Else
        
                  ' Add to the Volume Total and Ticker Count
                  
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                  Ticker_Count = Ticker_Count + 1
            
                End If
            
              Next i
              
              '----------------------------------------------------------------------------------------------------------------------------
              ' Adding formatting to the summary table
              
              ' Define last row in summary table
              Dim LastRow_Table As Long
              
              LastRow_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
              
              ' Loop through all ticker symbols in the summary table
              For j = 2 To LastRow_Table
              
                ' Change columns format
                ws.Cells(j, 10).NumberFormat = "$#,##0.00"
                ws.Cells(j, 11).NumberFormat = "0.00%"
                ws.Cells(j, 12).NumberFormat = "#,###,###"
                ws.Columns("I:L").AutoFit
                
                ' Add color formatting to yearly change column
                If ws.Cells(j, 10).Value > 0 Then
        
                    ws.Cells(j, 10).Interior.ColorIndex = 4
        
                ' Otherwise color it red
                ElseIf ws.Cells(j, 10).Value < 0 Then
        
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    ws.Cells(j, 10).Font.ColorIndex = 2
        
                End If
                
              Next j
              
              '---------------------------------------------------------------------------------------------------------------------------
              ' Challenge 1
                    
              ' Creating the table
              
              ws.Range("O2").Value = "Greatest % Increase"
              ws.Range("O3").Value = "Greatest % Decrease"
              ws.Range("O4").Value = "Greatest Total Volume"
              ws.Range("P1").Value = "Ticker"
              ws.Range("Q1").Value = "Value"
              
              ws.Range("Q2:Q3").NumberFormat = "0.00%"
              ws.Range("Q4").NumberFormat = "#,###,###"
              
              ' Getting greatest % increase, decrease and total volume values
                    
              Dim Min_Change As Double
              Dim Max_Change As Double
              Dim Max_Volume As Variant
              Dim Count_MinChange As Long
              Dim Count_MaxChange As Long
              Dim Count_MaxVolume As Long
              Dim rng_change As Variant
              Dim rng_volume As Variant
              
              ' Setting the range for the min and maximum values of percent change and total volume
              rng_change = ws.Range("K:K")
              rng_volume = ws.Range("L:L")
              
              ' Finding the minimum and maximum values and enter the values and labels to the summary table
              
              ' Greatest Decrease
              Min_Change = WorksheetFunction.Min(rng_change) 'finding the greatest decrease
              Count_MinChange = WorksheetFunction.Match(Min_Change, rng_change, 0) 'finding the row index of the greatest decrease value
        
              ws.Range("P3").Value = ws.Cells(Count_MinChange, 9).Value 'entering the ticker to the table
              ws.Range("Q3").Value = Min_Change 'entering the value to the table
              
              ' Greatest Increase
              Max_Change = WorksheetFunction.Max(rng_change) 'finding the greatest increase
              Count_MaxChange = WorksheetFunction.Match(Max_Change, rng_change, 0) 'finding the row index of the greatest increase value
              
              ' Greatest Increase
              ws.Range("P2").Value = ws.Cells(Count_MaxChange, 9).Value 'entering the ticker to the table
              ws.Range("Q2").Value = Max_Change 'entering the value to the table
              
              ' Greatest Volume
              Max_Volume = WorksheetFunction.Max(rng_volume) 'finding the greatest volume
              Count_MaxVolume = WorksheetFunction.Match(Max_Volume, rng_volume, 0) 'finding the row index of the greatest increase value
              
              ws.Range("P4").Value = ws.Cells(Count_MaxVolume, 9).Value 'entering the ticker to the table
              ws.Range("Q4").Value = Max_Volume 'entering the value to the table
              
              ws.Columns("O:Q").AutoFit
              
    Next ws
                  
                  
End Sub


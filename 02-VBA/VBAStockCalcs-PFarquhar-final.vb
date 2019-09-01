Sub stock_calcs()

  ' Set variables
  Dim Ticker As String
  Dim VolTotal As Currency
  Dim LastRow As Long
  Dim YrOpenPrice As Currency
  Dim YrClosePrice As Currency
  Dim YearlyChange As Currency
  Dim PctChange As Double
  Dim GreatestPctIncrease As Double
  Dim GreatestPctDecrease As Double
  Dim GreatestTotVolume As Currency
  Dim GreatestIncTicker As String
  Dim GreatestDecTicker As String
  Dim GreatestVolTicker As String
  Dim SummaryTableRow As Long
  Dim SummaryTableLength As Long

 'Create an instance of the Worksheet object called "ws"
    Dim ws As Worksheet

'Loop through each Worksheet object in the Worksheets collection
    For Each ws In ActiveWorkbook.Worksheets

      ' Find # of last row in sheet
        ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

      'Reset Tracking Variables for This Year
        VolTotal = 0
        YrOpenPrice = Cells(2, 3).Value
        YrClosePrice = 0
        
      ' Write table titles
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"

      ' Keep track of the location for each ticker in the summary table
        SummaryTableRow = 2

          ' Loop through tickers on this sheet
          For i = 2 To LastRow

            ' If this is last row of ticker, then....
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

              ' Set Year-End Stats for this stock
              Ticker = Cells(i, 1).Value
              VolTotal = VolTotal + Cells(i, 7).Value
              YrClosePrice = Cells(i, 6).Value
              YearlyChange = YrClosePrice - YrOpenPrice
              
              If YrOpenPrice > 0 Then
                PctChange = YearlyChange / YrOpenPrice
                
              Else
                PctChange = 0
                
              End If

              ' Print to Summary Table
              Range("I" & SummaryTableRow).Value = Ticker
              Range("J" & SummaryTableRow).Value = YearlyChange
              Range("L" & SummaryTableRow).Value = VolTotal
              
              If YrOpenPrice > 0 Then
                Range("K" & SummaryTableRow).Value = Round((PctChange * 100), 4)
                
              Else
                Range("K" & SummaryTableRow).Value = "n/a"
                
              End If
              
              'Set Percent Change Cells to Red or Green
              
              If PctChange > 0 Then
                Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                
              Else
                Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                
              End If
                
                  
              ' Add one to the summary table row
              SummaryTableRow = SummaryTableRow + 1
              
              ' Reset Total, Set Open for next stock
              VolTotal = 0
              YrOpenPrice = Cells(i + 1, 3)

            ' If the cell immediately following a row is same ticker, just add to total
            Else
               VolTotal = VolTotal + Cells(i, 7).Value

            End If

          Next i

      'We are done going through stocks- now test of annual leaders
      
      'Reset leaders variables
        
          GreatestIncTicker = 0
          GreatestDecTicker = 0
          GreatestVolTicker = 0
          GreatestPctIncrease = 0
          GreatestPctDecrease = 0
          GreatestTotVolume = 0
        
       ' Scan summary table for leaders
       
          SummaryTableLength = Range("I" & Rows.Count).End(xlUp).Row
      
            For j = 2 To SummaryTableLength
            
              Ticker = Cells(j, 9).Value
              VolTotal = Cells(j, 12).Value
              
              
              If Cells(j, 11).Value = "n/a" Then
                PctChange = 0
              
              Else
                PctChange = Cells(j, 11).Value
                    
              End If
                  
                  If PctChange > GreatestPctIncrease Then
                    GreatestPctIncrease = PctChange
                    GreatestIncTicker = Ticker
                    
                  End If
    
                  If PctChange < GreatestPctDecrease Then
                    GreatestPctDecrease = PctChange
                    GreatestDecTicker = Ticker
    
                  End If
    
                  If VolTotal > GreatestTotVolume Then
                    GreatestTotVolume = VolTotal
                    GreatestVolTicker = Ticker
                    
                  End If
                  
              Next j
            
      'We are done with sheet - Print Annual Leaders
          Cells(2, 16).Value = GreatestIncTicker
          Cells(3, 16).Value = GreatestDecTicker
          Cells(4, 16).Value = GreatestVolTicker
          Cells(2, 17).Value = GreatestPctIncrease
          Cells(3, 17).Value = GreatestPctDecrease
          Cells(4, 17).Value = GreatestTotVolume
          
      Columns("A:Q").EntireColumn.AutoFit

    Next ws
    
End Sub

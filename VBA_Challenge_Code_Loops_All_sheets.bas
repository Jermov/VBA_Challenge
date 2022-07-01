Attribute VB_Name = "Module1"
'This is an all-inclusive loop that runs through all worksheets- needs run 1 time, will produce output on each sheet

Sub WSLoop()

Dim ws As Worksheet
Dim LR As Long
Dim ClosePrice As Double
Dim OpenPrice As Double
Dim YearlyChg As Double
                   
    'iterate through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        'within each worksheet do the following
        With ws
            LR = .Cells(.Rows.Count, 1).End(xlUp).Row
     
            ' Set an initial variable for holding the total volume per ticker
            Dim volamt As Double
            volamt = 0

            ' create a counter to determine the last row needed for the unique tickers
            Dim Ticker_Row As Integer
            Ticker_Row = 2

                ' Loop through all ticker data
                For i = 2 To LR

                    ' Check if we are still within the Same ticker, if it is not...
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                        ' Set the tickername
                        Ticker = ws.Cells(i, 1).Value

                        ' Add to the volume total
                        volamt = volamt + ws.Cells(i, 7).Value

                        ' Print the unique ticker
                        ws.Range("I" & Ticker_Row).Value = Ticker

                        ' Print the volume
                        ws.Range("N" & Ticker_Row).Value = volamt
                        'adding end of year price
                        ClosePrice = ws.Cells(i, 6).Value
                         ws.Range("J" & Ticker_Row).Value = ClosePrice
                        ' Add one to the counter
                        Ticker_Row = Ticker_Row + 1
      
                        ' Reset the total volume
                        volamt = 0
            
        
                    ' If the cell immediately following a row is the same ticker...
                    Else

                        ' Add to the ticker volume
      
                        volamt = volamt + Cells(i, 7).Value
                       
    
                    End If
                    'Since the code iterates through each worksheet automatically- This was my workaround
                    'to obtain the earliest date for each ticker across the worksheets
                    'this was not covered in class- and google was not helpful
                    If ws.Cells(i, 2).Value = 20180102 Or ws.Cells(i, 2).Value = 20190102 Or ws.Cells(i, 2).Value = 20200102 Then
                            OpenPrice = ws.Cells(i, 3).Value
                            ws.Range("K" & Ticker_Row).Value = OpenPrice
 
                      End If
                Next i
'Outside the loop! All the hard work is done!
'Fill in the titles requested
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "End Price"
ws.Range("K1").Value = "Start Price"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
ws.Range("N1").Value = "Total Volume"
ws.Range("O1").Value = "Step to Pct Change"
ws.Range("Q1").Value = "Steps to Pct Change"

'Lies. Need to do the change in price and pct change
'Not covered in class- google unhelpful- adding to loop breaks the loop. added after to give static results
LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

ws.Range("L2:L" & LastRow) = ws.Evaluate("J2:J" & LastRow & "-K2:K" & LastRow)
ws.Range("O2:O" & LastRow) = ws.Evaluate("J2:J" & LastRow & "/K2:K" & LastRow)
ws.Range("Q2:Q" & LastRow) = 1
ws.Range("M2:M" & LastRow) = ws.Evaluate("O2:O" & LastRow & "-Q2:Q" & LastRow)
'Yay! Calculations done! Hard work is over!

'More Lies. Forgot to color code and because I used evaluate outside loop, now I gotta add another loop. *sigh*

For j = 2 To Ticker_Row
If ws.Cells(j, 12).Value < 0 Then
    ws.Cells(j, 12).Interior.ColorIndex = 3
Else
    ws.Cells(j, 12).Interior.ColorIndex = 4
End If
Next j
'YAY! Now I just need to clean up column width and format anything I need to look better
ws.Range("J1").ColumnWidth = 12
ws.Range("K1").ColumnWidth = 12
ws.Range("L1").ColumnWidth = 12
ws.Range("M1").ColumnWidth = 15
ws.Range("N1").ColumnWidth = 16
ws.Range("M2:M" & Ticker_Row).NumberFormat = "0.00%"
'End the section of 'to do within each worksheet
        End With
'End loop through each worksheet
    Next ws
'End Subroutine- since I have to google it, I'm adding definition to this code!
End Sub




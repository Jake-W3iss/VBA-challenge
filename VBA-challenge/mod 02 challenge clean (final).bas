Attribute VB_Name = "Module1"
Sub stockAnalyze():

    Dim stockVol As Double 'total stock volume
    Dim row As Long      'loop control
    Dim rowCount As Long 'holds the number of rows looked at in the sheet
    Dim qChange As Double 'holds quarterly change for each stock
    Dim perChange As Double 'holds percent change for each stock
    Dim sumTableRow As Long 'holds the row for the sum. table
    Dim stockStart As Long  'holds the start of stock's row in the sheet
    Dim stockOpen As Long   'holds the opening value for a stock
    Dim lastStock As String 'holds the last stock in sheet
    Dim stock As String     'holds current stock name
    
    
' loop through all worksheets in excel workbook
For Each ws In Worksheets
        'Set summary table and Aggregate table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        
        ' set up the values
        sumTableRow = 0 'sumtable row starts at 0 in the sheet (add 2) in relation to header
        stockVol = 0    'total stock vol. for a stock starts at 0
        qChange = 0     'quarterly change starts at 0
        stockStart = 2  'first stock in sheet is on row 2
        stockOpen = 2   'loaction of first open value is on row 2
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row 'last row in current sheet
        lastStock = ws.Cells(rowCount, 1).Value 'finds the last stock so we can leave loop
            
            'loop through allll the stocks
            For row = 2 To rowCount
            stock = ws.Cells(row, 1).Value 'lets me refrence what stock i'm looking at
            
                'see if stock has changed name
                If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1) Then
                
                    'add last stock volume update
                    stockVol = stockVol + ws.Range("G" & row).Value
                        
                        'make sure total stock vol. isn't 0
                        If stockVol = 0 Then
                            'print results in sum. table
                            ws.Range("I" & 2 + sumTableRow).Value = stock
                            ws.Range("J" & 2 + sumTableRow).Value = 0
                            ws.Range("K" & 2 + sumTableRow).Value = 0
                            ws.Range("L" & 2 + sumTableRow).Value = 0
                        Else
                            'find first open value for the stock
                            If ws.Cells(stockOpen, 3).Value = 0 Then
                            
                                For findValue = stockOpen To row
                                
                                ' check to see fi the next or rows afterwards  open value does not equal 0
                                    If ws.Cells(findValue, 3).Value <> 0 Then
                                    'once we find have a non zero first open value, that value becomes the row where we track first open
                                        stockOpen = findValue
                                        'break out of loop
                                        Exit For
                                    End If
                                Next findValue
                            End If
                        
                            'if the starting open value isn't 0 and and we are moving to a new stock
                            'calculate the quarterly change
                            qChange = ws.Range("F" & row).Value - ws.Range("C" & stockOpen).Value
                            
                            'calculate the percent change
                            perChange = qChange / ws.Cells(stockOpen, 3).Value
                            
                            ' print results
                            ws.Range("I" & 2 + sumTableRow).Value = stock
                            ws.Range("J" & 2 + sumTableRow).Value = qChange
                            ws.Range("K" & 2 + sumTableRow).Value = perChange
                            ws.Range("L" & 2 + sumTableRow).Value = stockVol
                            
                            'color the quaterly change
                            If qChange > 0 Then
                                ws.Range("J" & 2 + sumTableRow).Interior.ColorIndex = 4 'positive values are green
                            
                            ElseIf qChange < 0 Then
                                 ws.Range("J" & 2 + sumTableRow).Interior.ColorIndex = 3 'negative values are red
                                 
                            End If
                            
                            'reset the values for the next stock
                            stockVol = 0
                            qChange = 0
                            stockOpen = row + 1 'moves the start row to next row in the sheet
                            sumTableRow = sumTableRow + 1 'allows us to print in the next row of the sum. table
                        End If
                
                Else 'stock name hasn't changed, keep adding up stock volume
                
                    stockVol = stockVol + ws.Range("G" & row).Value 'stock vol. becomes itself plus whatever is on the current row in column G
                
                End If
                
              
                
            Next row
              'clean up (if needed)
                'find the last row of data in sum. table by finding the last ticket in sum. section
                
                'update sum. table row
                sumTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
                
                'find the last data in the extra rows from columns J-L
            Dim lastExtraRow As Long
            lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
             'loop that clears extra data from columns I-L
            For e = sumTableRow To lastExtraRow
                'for loop that goes through columns I-L (9-12)
                For Column = 9 To 12
                
                    ws.Cells(e, Column).Value = ""
                    ws.Cells(e, Column).Interior.ColorIndex = 0
                
                Next Column
            
            Next e
            
            'print the summary aggregates
            ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & sumTableRow + 2))
            ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & sumTableRow + 2))
            ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & sumTableRow + 2))
            
            'match to find the row numbers of the ticker names of our values in aggregate summary
            Dim greatIncRow As Double
            Dim greatDecRow As Double
            Dim greatVolRow As Double
            greatIncRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & sumTableRow + 2)), ws.Range("K2:K" & sumTableRow + 2), 0)
            greatDecRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & sumTableRow + 2)), ws.Range("K2:K" & sumTableRow + 2), 0)
            greatVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & sumTableRow + 2)), ws.Range("L2:L" & sumTableRow + 2), 0)
            
            ws.Range("P2").Value = ws.Cells(greatIncRow + 1, 9).Value
            ws.Range("P3").Value = ws.Cells(greatDecRow + 1, 9).Value
            ws.Range("p4").Value = ws.Cells(greatVolRow + 1, 9).Value
            
       'formatting the sum. table and aggregate table
       
            ws.Range("J:J").NumberFormat = "0.00"
            ws.Range("K:K").NumberFormat = "0.00%"
            ws.Range("L:L").NumberFormat = "#,###"
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###"
                
    
    
        ws.Columns("A:Q").AutoFit
Next ws
    
End Sub

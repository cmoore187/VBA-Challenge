Attribute VB_Name = "Module1"
Sub TickerSymbol()

'use for each loop to loop through all of the worksheets
   For Each ws In Worksheets

'Make the worksheet active
        ws.Activate

'get the count of the rows
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
' Set an initial variable for holding the brand name
        Dim tickerName As String
        
' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
'insert column names
        'add the word Ticker to I1
        ws.Range("I1").Value = "Ticker"
        
        'add the words Yearly Change to J1
        ws.Range("J1").Value = "Yearly Change"
        
        'add the words PercentChange to K1
        ws.Range("K1").Value = "Percent Change"
        
        'add the words Total stock volume to L1
        ws.Range("L1").Value = "Total Stock Volume"
        
        'loop through entire ticker column
        For i = 2 To LastRow
            
            ' Check if we are still within the same ticker symbol, if not then...
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                  ' Set the ticker symbol
                  tickerName = Cells(i, 1).Value
                
            
                 ' Print the ticker symbol in the Summary Table
                  Range("I" & Summary_Table_Row).Value = tickerName
                
                
            
                ' If the cell immediately following a row is the same ticker symbol
                'do nothing for now
                Else
                        
                   
                     
            
                End If
                
                 
        
            Next i
            
         
        Exit For
    Next ws
    









End Sub


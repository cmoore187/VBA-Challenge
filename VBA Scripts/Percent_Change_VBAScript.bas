Attribute VB_Name = "Module1"
Sub PercentChange()


' variable to keep track of opening price for specific year
Dim openingPrice As Double

' variable to keep track of closing price for specific year
Dim closingPrice As Double

' variable to keep track of yearly change
Dim yearlyChange As Double

' variable to keep track of percent change
Dim PercentChange As Double

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
                  
                  'calculate closing price
                  closingPrice = Cells(i, 6)
                  
                  'calculate yearly change
                  yearlyChange = closingPrice - openingPrice
                  
                  'calculate percent change
                  If openingPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = (yearlyChange / openingPrice)
                End If
                
            
                 ' Print the ticker symbol in the Summary Table
                  Range("I" & Summary_Table_Row).Value = tickerName
                  
                  'Add yearly change to summary table
                  Range("J" & Summary_Table_Row).Value = yearlyChange
            
                 'Add percent change to summary table
                 Range("K" & Summary_Table_Row).Value = PercentChange
                    
                
                    
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  'set openingPrice to zero
                  openingPrice = 0
                  
                
                
            
                ' If the cell immediately following a row is the same ticker symbol
                Else
                        
                      'Get opening price
                      If (openingPrice = 0) Then
                        openingPrice = Cells(i, 3).Value
                    End If
                     
                    
                   
                     
            
                End If
                
                 
        
            Next i
            
            'Change the color of the yearly change cells based on positive or negative
        For j = 2 To 290
            If (Cells(j, 10) > 0) Then
                    Cells(j, 10).Interior.ColorIndex = 4
                ElseIf (Cells(j, 10) < 0) Then
                    Cells(j, 10).Interior.ColorIndex = 3
                Else
            End If
            
         Next j
         
        Exit For
    Next ws
    









End Sub

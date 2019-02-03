Attribute VB_Name = "Module1"
Sub Easy()

    For Each ws In Worksheets
    
        Dim Stock_Ticker As String
        Dim Stock_Volume As Double
        Dim lastrow As Double

        Stock_Volume = 0

        Dim Ticker_Vol_TableRow As Long
        Ticker_Vol_TableRow = 2
    
        ' Determine last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' I Column needs to be called Ticker
        ws.Cells(1, 9).Value = "Ticker"
    
        ' J Column needs to be called Total Stock Volume
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
            For i = 2 To lastrow
                
                'Check for changes in value of cell vs. previous cell in column
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ' Identify the ticker
                    Stock_Ticker = ws.Cells(i, 1).Value
                
                    ' Add volumes
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                    ' Record ticker to list
                    ws.Range("I" & Ticker_Vol_TableRow).Value = Stock_Ticker

                    ' Record the volume total
                    ws.Range("J" & Ticker_Vol_TableRow).Value = Stock_Volume
                
                    ' Add one to the Ticker_Vol_TableRow to setup for next i
                    Ticker_Vol_TableRow = Ticker_Vol_TableRow + 1
                    
                    ' Reset the Stock_Volume to setup for next i
                    Stock_Volume = 0
                
                'if there is no change in value between cel and previous cell in colum the just keep adding to Total Stock Volume
                Else
                    ' Add to the volume Total
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                End If

            Next i
            
    Next ws

End Sub



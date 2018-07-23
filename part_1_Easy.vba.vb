Sub Total_vol()

' Loop through all of the worksheets in the active workbook

    For Each ws In Worksheets
           'lable
            ws.Cells(1, 10).Value = "Ticker"
                
            ' stock vol lable
                ws.Cells(1, 11).Value = "Total Volume"

            ' Start a stock
            Dim Stock As String
                                                                                                    
            ' Start stock total vol as zero
            Dim Total_vol As Double
            Total_vol = 0
                                                                                                    
            ' place holder for the summary table - to keep track of distinct uniqe stock rows
            Dim Stock_Row As Integer
            Stock_Row = 2
                                                                                                    
            ' Determine the Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ' Loop over all stocks
            For I = 2 To LastRow ' start of the for loop
                    
            
                    ' if it is not with the same stock .
                    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
                        ' Set the stock name - if we we are not in the same stock the sub is renewing the stock name - this will work for the first stock as well at cell 2,1
                        Stock = ws.Cells(I, 1).Value
                
                        ' sum the stock vol - if we are not in the same stock - remember this will work for the first stock as well in cell 2,1
                        Total_vol = Total_vol + ws.Cells(I, 7).Value
                
                        ' type the stock name in the stock Vol table summary
                        ws.Cells(Stock_Row, 10).Value = Stock
                
                        ' type stock vol summary in the table
                        ws.Cells(Stock_Row, 11).Value = Total_vol
                
                
                        ' when it is not true - count the stock raw by 1 to go to the next raw in table summary for the next distinct stock
                        Stock_Row = Stock_Row + 1
                        
                        ' reset the vol when the raw1 does not equal raw1 +1
                        Total_vol = 0
                
                    ' if stock at raw1 = raw1 +1 do nothing to stock name but add the vol only
                    Else
                        ' add the stock vol for this stock
                        Total_vol = Total_vol + ws.Cells(I, 7).Value
                    End If
            Next I ' next item in the for loob
    Next ws

    MsgBox("Total Vol Calculated, Done the Easy Part!")

End Sub

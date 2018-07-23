Sub Change()


    '''Part 2
    For Each ws In Worksheets

            ' Add the Yearly Change to the Column
            ws.Range("K1").EntireColumn.Insert
            
            ' Add the word Yearly Change to the First Column Header
            ws.Cells(1, 11).Value = "Yearly Change"

            ' Add the Percent Change to the Column
            ws.Range("L1").EntireColumn.Insert
            
            ' Add the word Percent Change to the First Column Header
            ws.Cells(1, 12).Value = "Percent Change"

    ' close1 is the early close of the stock, close2 is the late close of the stock
            
            Dim close1 As Double
            Dim close2 As Double
            ' counter
            Dim Stock_Row As Integer
            Stock_Row = 2

            close1 = ws.Cells(2, 6).Value
            
            ' Determine the Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ' Loop over all stocks

            For I = 2 To LastRow
            If ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
            close1 = close1 + 0

            Else
            close2 = ws.Cells(I, 6).Value

            ws.Cells(Stock_Row, 11).Value = (close2 - close1)

                If (ws.Cells(Stock_Row, 11).Value > 0) Then
                    ws.Cells(Stock_Row, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(Stock_Row, 11).Interior.ColorIndex = 3
                End If


                'it seems that some close1 values are 0(zero) - so need to trick this


                    'If close1 = 0 Then
                        'close1 = 0.000000000001
                    'End If

                ws.Cells(Stock_Row, 12).Value = ((close2 - close1) / close1)

                ws.Cells(Stock_Row, 12).Style = "Percent"
                ws.Cells(Stock_Row, 12).NumberFormat = "0.00%"


            Stock_Row = Stock_Row + 1
            close1 = ws.Cells(I + 1, 6).Value

            
            End If
            Next I

    Next ws

    MSGBOX("Chaange Calculated!, Moderate Part is Done!")
End Sub


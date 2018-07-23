Sub Up_down()

    '''Part 3
    For Each ws In Worksheets


            Dim Greatest_increase As Double
            Dim Greatest_Decrease As Double
            Dim Greatest_total_volume As Double

            Greatest_increase = 0
            Greatest_Decrease = 0
            Greatest_total_volume = 0

            Dim Ticker1 As String
            Dim Ticker2 As String
            Dim Ticker3 As String

            
            ' Fill Headers
            ws.Cells(2, 15).Value = "Greatest % increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest total volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"


            
            ' Determine the Last Row
            LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
            ' Loop over all stocks ALL BOOK ALL SHEETS

            For I = 2 To LastRow

                If ws.Cells(I, 12).Value > Greatest_increase Then

                    Greatest_increase = ws.Cells(I, 12).Value
                    Ticker1 = ws.Cells(I, 10).Value

                End If

                If ws.Cells(I, 12).Value < Greatest_Decrease Then

                    Greatest_Decrease = ws.Cells(I, 12).Value
                    Ticker2 = ws.Cells(I, 10).Value

                End If


                If ws.Cells(I, 13).Value > Greatest_total_volume Then

                    Greatest_total_volume = ws.Cells(I, 13).Value
                    Ticker3 = ws.Cells(I, 10).Value

                End If



            Next I

            ' Fill values
                
            ws.Cells(2, 16).Value = Ticker1
            ws.Cells(3, 16).Value = Ticker2
            ws.Cells(4, 16).Value = Ticker3

            ws.Cells(2, 17).Value = Greatest_increase

                ws.Cells(2, 17).Style = "Percent"
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ws.Cells(3, 17).Value = Greatest_Decrease

                ws.Cells(3, 17).Style = "Percent"
                ws.Cells(3, 17).NumberFormat = "0.00%"


            ws.Cells(4, 17).Value = Greatest_total_volume

    Next ws

    MsgBox("The Hard Part is Done!, Summary is Calculated")
End Sub



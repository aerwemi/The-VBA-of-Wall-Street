Sub Clean()

        For Each ws In Worksheets
                ' find Last Row
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
                ' loop over all rows starting from last row 
                For i = LastRow To 2 Step -1
                
                        If ((ws.Cells(i, 7).Value = 0) And (ws.Cells(i, 6).Value = 0) And (ws.Cells(i, 5).Value = 0)) Then
                            
                            Rows(i).Delete
                            
                        End If
                Next i
        Next ws

    MsgBox("Done Cleaning")

End Sub
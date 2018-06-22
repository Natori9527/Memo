Sub test2()
Dim st1, st2 As Worksheet
Set st1 = Sheets.item("Sheet2")
Set st2 = Sheets.item("Sheet3")
Dim row1S, row2S, c11, c12, c21, c22, tc
row1S = 1
c11 = 5
c12 = 8
c21 = 5
c22 = 11
tc = 17



Do While st1.Cells(row1S, c11) <> ""
    row2S = 1
    Do While st2.Cells(row2S, c21) <> ""
    
        If (st1.Cells(row1S, c11).Value = st2.Cells(row2S, c21).Value) And (st1.Cells(row1S, c12).Value = st2.Cells(row2S, c22).Value) Then
            If st1.Cells(row1S, tc).Value <> "" Then
                st1.Cells(row1S, tc).Value = st1.Cells(row1S, tc).Value & "-" & row2S
            Else
                st1.Cells(row1S, tc).Value = row2S
            End If
        End If
        
        row2S = row2S + 1
    Loop
    
    row1S = row1S + 1
Loop

End Sub

Sub multiCellsCompare()
    On Error GoTo ErrHandler

Dim sourceSht, targetSht As Worksheet
'Source sheet
Set sourceSht = Sheets.item("Sheet2")
'target sheet
Set targetSht = Sheets.item("Sheet3")
Dim sourceShtSR, targetShtSR, sourceC1, sourceC2, targetC1, targetC2, recultC
'Source Sheet start row number
sourceShtSR = 1
'Source Sheet search column 1
sourceC1 = 5
'Source Sheet search column 2
sourceC2 = 8
'target Sheet search column 1
targetC1 = 5
'target Sheet search column 2
targetC2 = 11
'output result column 2
resultC = 17

'source sheet activate
ThisWorkbook.Activate
sourceSht.Activate

Do While sourceSht.Cells(sourceShtSR, sourceC1) <> ""
   'result column clear
   If sourceSht.Cells(sourceShtSR, resultC).Value <> "" Then
       sourceSht.Cells(sourceShtSR, resultC).Value = ""
   End If
   
   'Target Sheet start row number
    targetShtSR = 1
    Do While targetSht.Cells(targetShtSR, targetC1) <> ""
    
        If (sourceSht.Cells(sourceShtSR, sourceC1).Value = targetSht.Cells(targetShtSR, targetC1).Value) And _
           (sourceSht.Cells(sourceShtSR, sourceC2).Value = targetSht.Cells(targetShtSR, targetC2).Value) Then
            If sourceSht.Cells(sourceShtSR, resultC).Value <> "" Then
                sourceSht.Cells(sourceShtSR, resultC).Value = sourceSht.Cells(sourceShtSR, resultC).Value & vbLf & targetShtSR
            Else
                sourceSht.Cells(sourceShtSR, resultC).Value = targetShtSR
            End If
        End If
        
        targetShtSR = targetShtSR + 1
    Loop
    
    sourceShtSR = sourceShtSR + 1
Loop

MsgBox ("Completed!")
Exit Sub

ErrHandler:
MsgBox (Err.Description)

End Sub

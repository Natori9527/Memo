Sub multiCellsCompare()
    On Error GoTo ErrHandler

'compared data dic
Dim dic As Object
Set dic = CreateObject("Scripting.Dictionary")

Dim sourceSht, targetSht As Worksheet
'Source sheet
Set sourceSht = Sheets.Item("Sheet1")
'target sheet
Set targetSht = Sheets.Item("Sheet2")
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
   
   'if has compared,continue the next row
   If dic.exists(Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value)) Then
       sourceSht.Cells(sourceShtSR, resultC).Value = "Same as row:" & dic.Item(Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value))
       GoTo nextRow
   End If
   
   ' add to dic
   dic.Add Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value), sourceShtSR
   
   'Target Sheet start row number
    targetShtSR = 1
    Do While targetSht.Cells(targetShtSR, targetC1) <> ""
    
        If (Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) = Val(targetSht.Cells(targetShtSR, targetC1).Value)) And _
           (Val(sourceSht.Cells(sourceShtSR, sourceC2).Value) = Val(targetSht.Cells(targetShtSR, targetC2).Value)) Then
           
            If sourceSht.Cells(sourceShtSR, resultC).Value <> "" Then
                sourceSht.Cells(sourceShtSR, resultC).Value = sourceSht.Cells(sourceShtSR, resultC).Value & vbLf & "Row:" & targetShtSR
            Else
                sourceSht.Cells(sourceShtSR, resultC).Value = "Row:" & targetShtSR
            End If
        End If
        
        targetShtSR = targetShtSR + 1
    Loop
    
    If Trim(sourceSht.Cells(sourceShtSR, resultC).Value) = "" Then
        sourceSht.Cells(sourceShtSR, resultC).Value = "Not Found"
    End If
    
nextRow:
    sourceShtSR = sourceShtSR + 1
Loop

MsgBox ("Completed!")
Exit Sub

ErrHandler:
MsgBox (Err.Description)

End Sub

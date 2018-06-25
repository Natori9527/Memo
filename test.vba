Sub multiCellsCompare()
    On Error GoTo ErrHandler

'compared data copDic
Dim copDic As Object
Set copDic = CreateObject("Scripting.Dictionary")
'uncompared data copDic
Dim ucopDic As Object
Set ucopDic = CreateObject("Scripting.Dictionary")

Dim sourceSht, targetSht As Worksheet
'Source sheet
Set sourceSht = Sheets.item("Sheet1")
'target sheet
Set targetSht = Sheets.item("Sheet2")
Dim commentR, sourceShtSR, targetShtSR, sourceC1, sourceC2, targetC1, _
    targetC2, recultC, unMatchCnt, targetRCount, idx, unMatchR
'comment row number
commentR = 1
'Source Sheet start row number
sourceShtSR = 1
'Source Sheet search column 1(Column E)
sourceC1 = 5
'Source Sheet search column 2(Column H)
sourceC2 = 8
'target Sheet search column 1(Column E)
targetC1 = 5
'target Sheet search column 2(Column K)
targetC2 = 11
'output result column(Column K)
resultC = 17
'Target Sheet start row number
targetShtSR = 1
'result unmatch row
unMatchR = 1
'Target Sheet row count
targetRCount = targetSht.Cells.Find("*", targetSht.Range("A1"), -4163, 1, 1, 2, False, False, False).Row + 1 - targetShtSR
'target Sheet unmatch record count ini
unMatchCnt = targetRCount
For idx = targetShtSR To unMatchCnt
    ucopDic.Add idx, idx
Next



'source sheet activate
ThisWorkbook.Activate
sourceSht.Activate

Do While sourceSht.Cells(sourceShtSR, sourceC1) <> ""

   'result column clear
   If sourceSht.Cells(sourceShtSR, resultC).Value <> "" Then
       sourceSht.Cells(sourceShtSR, resultC).Value = ""
   End If
   
   'if has compared,continue the next row
   If copDic.Exists(Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value)) Then
       sourceSht.Cells(sourceShtSR, resultC).Value = "Same as row:" & copDic.item(Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value))
       GoTo nextRow
   End If
   
    ' add to copDic
    copDic.Add Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) & "-" & Val(sourceSht.Cells(sourceShtSR, sourceC2).Value), sourceShtSR
   
    Do While targetSht.Cells(targetShtSR, targetC1) <> ""
    
        If (Val(sourceSht.Cells(sourceShtSR, sourceC1).Value) * 10 = Val(targetSht.Cells(targetShtSR, targetC1).Value)) And _
           (Val(sourceSht.Cells(sourceShtSR, sourceC2).Value) = Val(targetSht.Cells(targetShtSR, targetC2).Value)) Then
           ' unmatch data -1
           unMatchCnt = unMatchCnt - 1
           ucopDic.Remove targetShtSR

           
            If sourceSht.Cells(sourceShtSR, resultC).Value <> "" Then
                sourceSht.Cells(sourceShtSR, resultC).Value = sourceSht.Cells(sourceShtSR, resultC).Value & vbLf & "Row:" & targetShtSR
            Else
                sourceSht.Cells(sourceShtSR, resultC).Value = "Row:" & targetShtSR
            End If
        End If
        
        targetShtSR = targetShtSR + 1
    Loop
   
    'Target Sheet start row number reset
    targetShtSR = 1
    
    If Trim(sourceSht.Cells(sourceShtSR, resultC).Value) = "" Then
        sourceSht.Cells(sourceShtSR, resultC).Value = "Not Found"
    End If
    
nextRow:
    sourceShtSR = sourceShtSR + 1
Loop

sourceSht.Cells(1, resultC + 1).Value = "Unmatched count:" & unMatchCnt


sourceSht.Range("R" & unMatchR + 1 & ":R" & targetRCount).Clear

'Dim comment
idx = unMatchR + 1
For Each strKey In ucopDic.Keys()
'    comment = comment & "Row:" & ucopDic(strKey) & vbLf
    sourceSht.Cells(idx, resultC + 1).Value = "Row:" & ucopDic(strKey)
    idx = idx + 1
Next

'If Not sourceSht.Cells(commentR, resultC + 1).comment Is Nothing Then
'    sourceSht.Cells(commentR, resultC + 1).comment.Delete
'End If

'sourceSht.Cells(commentR, resultC + 1).AddComment comment
'sourceSht.Cells(commentR, resultC + 1).comment.Shape.Height = 10 + 11 * unMatchCnt

MsgBox ("Completed!")
Exit Sub

ErrHandler:
MsgBox (Err.Description)

End Sub


Sub getLog()
    Dim sCmd As String
    Dim logStr As String
    
    Dim sht As Worksheet
    Set sht = Application.Worksheets("svnLog")
    sht.Range(sht.Cells(10, 3), sht.Cells(65536, 3)).Clear
    
    ' ArrayList
    Dim result As Object
    Set result = CreateObject("System.Collections.ArrayList")
    sCmd = getCmd()
    Debug.Print sCmd
    Set result = ShellRun(sCmd)
    
    Dim endRow As Long
    endRow = result.Count + 9
    
    sht.Range(sht.Cells(10, 3), sht.Cells(endRow, 3)).Value = WorksheetFunction.Transpose(result.toarray)
End Sub

' svn log -v url --username XXX --password XXX -r 6995:HEAD
Public Function getCmd() As String
    Dim sht As Worksheet
    ' 配置情報取得
    Set sht = Application.Worksheets("svnLog")
    Dim repUrl, userNm, password, reviSt, reviEnd As String
    repUrl = sht.Cells(3, 3)
    userNm = sht.Cells(4, 3)
    password = sht.Cells(5, 3)
    reviSt = sht.Cells(6, 3)
    reviEnd = sht.Cells(7, 3)
    If reviEnd = "" Then
        reviEnd = "HEAD"
    End If
    
    getCmd = "svn log -v " & repUrl & " --username " & userNm & " --password " & password & " -r " & reviSt & ":" & reviEnd
End Function

Public Function ShellRun(sCmd As String) As Object
    'Run a shell command, returning the output as a string'
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    ' ArrayList
    Dim result As Object
    Set result = CreateObject("System.Collections.ArrayList")
    
    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        
        If sLine <> "" Then
            result.Add getFormatRow(sLine)
            's = s & sLine & vbCrLf
        End If
    Wend

    Set ShellRun = result

End Function

Function getFormatRow(row) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    '正規表現の指定
    With reg
        .Pattern = "^[ ]*[A|M|D]?[ ]*/develop/wh"      'パターンを指定
        .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
        .Global = True          '文字列全体を検索するか(True)、しないか(False)
    End With
    
    
    getFormatRow = reg.Replace(row, "")
End Function

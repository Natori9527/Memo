Sub main()
    Dim rootDir As String
    rootDir = "C:\Tmp"

    ' check file path,if not exists create
    chkPath rootDir
    
    Dim kinouNms As Variant
    kinouNms = Array("aaa", "bbb")
    
    For idx = 0 To UBound(kinouNms)
       Dim kinouPath As String
       kinouPath = rootDir & "\" & kinouNms(idx) & "\"
       
       MkDir kinouPath
       
       copyFileToTmp kinouNms(idx), kinouPath
    Next idx

End Sub

Sub chkPath(path)
    Dim fso As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Dir(path, vbDirectory)) > 0 Then
        fso.deleteFolder path
    End If
    MkDir path
End Sub

Sub copyFileToTmp(kinouNm, kinouPath)

    Dim MyPath As String, MyFolderName As String, MyFileName As String
    Dim i As Integer, F As Boolean
    Dim objShell As Object, objFolder As Object, AllFolders As Object, AllFiles As Object
    Dim MySheet As Worksheet
     
    On Error Resume Next

    'List all folders
    Dim basePath As String
    Dim javaPath As String
    Dim jsPath As String
    Dim ftlhPath As String
    Dim sqlPath As String
    
    basePath = "C:\Users\...\src\main"
    javaPath = "\java\...\"
    sqlPath = "\sql\META-INF\...\"
    jsPath = "\resources\static\..."
    ftlhPath = "\resources\templates\views\"
     
    Set AllFolders = CreateObject("Scripting.Dictionary")
    Set AllFiles = CreateObject("Scripting.Dictionary")
    AllFolders.Add (basePath & javaPath & kinouNm & "\"), ""
    AllFolders.Add (basePath & sqlPath & kinouNm & "\"), ""
    AllFolders.Add (basePath & jsPath & kinouNm & "\"), ""
    AllFolders.Add (basePath & jsPath & kinouNm & "List\"), ""
    AllFolders.Add (basePath & ftlhPath & kinouNm & "\"), ""
    i = 0
    Do While i < AllFolders.Count
        key = AllFolders.keys
        MyFolderName = Dir(key(i), vbDirectory)
        Do While MyFolderName <> ""
            If MyFolderName <> "." And MyFolderName <> ".." Then
                If (GetAttr(key(i) & MyFolderName) And vbDirectory) = vbDirectory Then
                    AllFolders.Add (key(i) & MyFolderName & "\"), ""
                End If
            End If
            MyFolderName = Dir
        Loop
        i = i + 1
    Loop
     
    'List all files
    For Each key In AllFolders.keys
        MyFileName = Dir(key & "*.*")
        'MyFileName = Dir(Key & "*.PDF")    'only PDF files
        Do While MyFileName <> ""
            Debug.Print key
            FileCopy key & MyFileName, kinouPath & MyFileName
            MyFileName = Dir
        Loop
    Next
     
    Set AllFolders = Nothing
    Set AllFiles = Nothing
End Sub

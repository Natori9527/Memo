Function RegKeyRead(ByVal i_RegKey As String) As String
 On Error Resume Next
    Dim myWS As Object
    Dim res As String
    
    Set myWS = CreateObject("WScript.Shell")
    RegKeyRead = myWS.RegRead(i_RegKey)

End Function

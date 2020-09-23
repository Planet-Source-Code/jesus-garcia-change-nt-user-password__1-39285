Attribute VB_Name = "modGeneral"
Declare Function WNetGetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public stServidorGlobal As String

Public Function Get_ComputerName()
    Dim lpBuff As String * 25
    Dim ret As Long, ComputerName As String
    
    ret = GetComputerName(lpBuff, 25)
    ComputerName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

    Get_ComputerName = ComputerName
End Function

Public Function Get_User_Name()
    Dim s$, cnt&, dl&
    cnt& = 199
    s$ = String$(200, 0)
    dl& = WNetGetUserName(s$, cnt)
    Get_User_Name = Left$(s$, cnt)
End Function


Sub LoadServers()
Dim Ff As Integer
Dim Buf As String

    Ff = FreeFile
    On Error GoTo ErrorHdlr
    Open WinDir & "\Servers.txt" For Input As Ff
        While Not EOF(Ff)
            Input #Ff, Buf
            frmMain.lstServidores.AddItem Buf
        Wend
    Close Ff
    Exit Sub
ErrorHdlr:
    If Err.Number = 53 Then
        Open WinDir & "\Servers.txt" For Output As Ff
        Close Ff
    End If
End Sub

Sub SaveServers()
Dim Ff As Integer
Dim Buf As String
Dim i As Integer

    Ff = FreeFile
    Open WinDir & "\Servers.txt" For Output As Ff
        For i = 0 To frmMain.lstServidores.ListCount - 1
            If frmMain.lstServidores.List(i) <> Get_ComputerName Then
                Buf = frmMain.lstServidores.List(i)
                Print #Ff, Buf
            End If
        Next
    Close Ff
End Sub
Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)


    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If
End Function

Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetSystemDirectory(t, Len(t))
    SysDir = Left(t, i)


    If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
        SysDir = SysDir & "\"
    ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
        SysDir = Left(SysDir, Len(SysDir) - 1)
    End If
End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClock 
      Interval        =   1000
      Left            =   5640
      Top             =   1440
   End
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3450
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8328
            MinWidth        =   8328
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3297
            MinWidth        =   3297
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraPasswords 
      Caption         =   "Passwords"
      Height          =   2055
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   2295
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNuevo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtActual 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "New password:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Oldl password:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   "Options"
      Height          =   1095
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optSeleccion 
         Caption         =   "Change selected"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Change all"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraServidores 
      Caption         =   "&Servers"
      Height          =   3255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   855
      End
      Begin VB.ListBox lstServidores 
         Height          =   2400
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Modify"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function NetUserChangePassword Lib "Netapi32.dll" ( _
    ByVal domainname As String, ByVal Username As String, _
    ByVal OldPassword As String, ByVal NewPassword As String) As Long




Private Sub DefaultServer()
    lstServidores.AddItem Get_ComputerName
End Sub

Private Sub cmdAgregar_Click()
    frmAgregar.Show vbModal, Me
    If stServidorGlobal <> "" Then
        lstServidores.AddItem stServidorGlobal
    End If
End Sub

Private Sub cmdCambiar_Click()
On Error GoTo ErrorHdlr
    Dim r As Long
    Dim sServer As String
    Dim sUser As String
    Dim sOldPass As String
    Dim sNewPass As String
    Dim i As Integer
    Dim HuboError As Boolean
    Dim stServidor As String
    Dim iRes As Integer
        
    If optSeleccion.Value Then
        If lstServidores.SelCount = 1 Then
            stServidor = "\\" & lstServidores.List(lstServidores.ListIndex)
            iRes = MsgBox("Do you want to change the password for the user " & txtUsuario.Text & " on " & UCase(stServidor), vbYesNo + vbQuestion, "Confirmation")
            If iRes = vbYes Then
                staBar.Panels(1).Text = "Modifying " & UCase(stServidor) & "..."
                staBar.Refresh
                sServer = StrConv(stServidor, vbUnicode)
                sUser = StrConv(txtUsuario.Text, vbUnicode)
                sOldPass = StrConv(txtActual.Text, vbUnicode)
                sNewPass = StrConv(txtNuevo.Text, vbUnicode)
                r = NetUserChangePassword(sServer, sUser, sOldPass, sNewPass)
                If r <> 0 Then
                    MsgBox "Error! The password could not be changed, due to: " & vbCrLf & vbCrLf & _
                    "The old password is incorrect (Error 86)" & vbCrLf & _
                    "or the user does not exists (Error 2221)" & vbCrLf & _
                    "or the server could not be found (Error 1351)" & vbCrLf & vbCrLf & _
                    "Server: \\" & UCase(lstServidores.List(lstServidores.ListIndex)), vbCritical, "Error: " & r
                Else
                    MsgBox "Password changed sucecsfully!", vbExclamation, "Change Password"
                End If
            Else
                MsgBox "Operation canceled...", vbOKOnly, "Advise"
            End If
        Else
            MsgBox "You must select a server from the list...", vbOKOnly + vbInformation, "Aviso"
        End If
    Else
        iRes = MsgBox("Do you want to change the password for the user " & txtUsuario.Text & " on all the server in the list?", vbYesNo + vbQuestion, "Confirmation")
        If iRes = vbYes Then
            HuboError = False
            For i = 0 To lstServidores.ListCount - 1
                stServidor = "\\" & lstServidores.List(i)
                staBar.Panels(1).Text = "Modifying " & UCase(stServidor) & "..."
                staBar.Refresh
                sServer = StrConv(stServidor, vbUnicode)
                sUser = StrConv(txtUsuario.Text, vbUnicode)
                sOldPass = StrConv(txtActual.Text, vbUnicode)
                sNewPass = StrConv(txtNuevo.Text, vbUnicode)
                r = NetUserChangePassword(sServer, sUser, sOldPass, sNewPass)
                If r <> 0 Then
                    MsgBox "Error! The password could not be changed, due to: " & vbCrLf & vbCrLf & _
                    "The old password is incorrect (Error 86)" & vbCrLf & _
                    "or the user does not exists (Error 2221)" & vbCrLf & _
                    "or the server could not be found (Error 1351)" & vbCrLf & vbCrLf & _
                    "Server: \\" & UCase(lstServidores.List(i)), vbCritical, "Error: " & r
                    HuboError = True
                End If
            Next
            If Not HuboError Then
                MsgBox "Passwords changed sucecsfully!", vbExclamation, "Change Password"
            End If
        Else
            MsgBox "Operation canceled...", vbOKOnly, "Advise"
        End If
    End If
    staBar.Panels(1).Text = "Ready"
    staBar.Refresh
    Exit Sub
    
ErrorHdlr:
    MsgBox "Internal error while changing the password: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Sub cmdQuitar_Click()
    If lstServidores.ListCount > 0 Then
        If lstServidores.SelCount = 1 Then
            lstServidores.RemoveItem lstServidores.ListIndex
        Else
            MsgBox "Debe seleccionar un servidor de la lista...", vbOKOnly + vbInformation, "Aviso"
        End If
    Else
        MsgBox "No hay servidores para eliminar...", vbOKOnly + vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    staBar.Panels(1).Text = "Ready"
    staBar.Refresh
    Call LoadServers
    Call DefaultServer
    lstServidores.Selected(0) = True
    optSeleccion.Value = True
    txtUsuario.Text = Get_User_Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveServers
End Sub

Private Sub tmrClock_Timer()
    staBar.Panels(2).Text = Now()
End Sub

Private Sub txtActual_GotFocus()
    txtActual.SelStart = 0
    txtActual.SelLength = Len(txtActual.Text)
End Sub

Private Sub txtNuevo_GotFocus()
    txtNuevo.SelStart = 0
    txtNuevo.SelLength = Len(txtNuevo.Text)
End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(txtUsuario.Text)
End Sub

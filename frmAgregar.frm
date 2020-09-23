VERSION 5.00
Begin VB.Form frmAgregar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add server to the list"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtServidor 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmAgregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    stServidorGlobal = txtServidor.Text
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtServidor_Change()
    stServidorGlobal = ""
End Sub

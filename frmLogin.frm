VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1620
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   957.15
   ScaleMode       =   0  'User
   ScaleWidth      =   4549.192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   2040
      Top             =   1080
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblani 
      BackStyle       =   0  'Transparent
      Caption         =   "! ! !"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblattend 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS DENIED !!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Dim db As Database
Dim Rec As Recordset
Dim icount As Integer
Dim k As Integer

Private Sub Form_Activate()
lblattend.Visible = False
lblani.Visible = False
Timer1.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Unload Me
Me.Show
End If
'--------------
If KeyAscii = &H1B Then
Unload Me
End If

End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\Security.Mdb", 1)
Set Rec = db.OpenRecordset("Login", 2)
End Sub
Private Sub Timer1_Timer()
If icount = 0 Then
lblani.Visible = True
icount = 1

ElseIf icount = 1 Then
lblani.Visible = False
icount = 0
End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
'----Start--------------
If KeyAscii = &H1B Then
Unload Me
End If
'-----End---------------
If KeyAscii = 13 Then
With Rec
.FindFirst "Security='" & UCase(txtPassword) & "'"
If .NoMatch = False Then
        FrmMain.Show
        LoginSucceeded = True
        Unload Me
    Else
    LoginSucceeded = False
    lblattend.Visible = True
    lbl1.Visible = False
    txtPassword.Visible = False
'    lblani.Visible = True
Timer1.Enabled = True
    End If
End With
End If
'------Test End-------------------------

End Sub

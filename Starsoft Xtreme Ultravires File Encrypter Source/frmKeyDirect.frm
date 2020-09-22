VERSION 5.00
Begin VB.Form frmKeyDirect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Enter Key"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   HelpContextID   =   340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "&Hide Typing"
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   945
      Width           =   3060
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1260
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   2
      Top             =   1260
      Width           =   1170
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   525
      Width           =   4320
   End
   Begin VB.Label lblCode 
      Caption         =   "Enter the Key"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   315
      Width           =   3795
   End
End
Attribute VB_Name = "frmKeyDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.txtCode.Text = ""
Me.txtCode.PasswordChar = "*"
Me.Check1.Value = 1
Me.txtCode.SetFocus
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.txtCode.PasswordChar = "*"
    Else
    Me.txtCode.PasswordChar = ""
    End If
End Sub

Private Sub cmdOK_Click()
If Me.txtCode.Text = "" Then Exit Sub
gstrActiveKey = Me.txtCode.Text
Me.Hide
Call CheckMenus
End Sub

Private Sub cmdCancel_Click()
Me.txtCode.Text = ""
Me.Hide
End Sub

Private Sub txtCode_Change()
If Len(Me.txtCode.Text) > 0 Then
    Me.cmdOK.Enabled = True
    Else
    Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtCode <> "" And Me.cmdOK.Enabled = True Then cmdOK_Click
    End If
End Sub


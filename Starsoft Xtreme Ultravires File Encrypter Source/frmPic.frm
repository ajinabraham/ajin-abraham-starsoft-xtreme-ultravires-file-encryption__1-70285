VERSION 5.00
Begin VB.Form frmPic 
   BackColor       =   &H00000000&
   Caption         =   " Image"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   Icon            =   "frmPic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3675
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If PicFactor = 0 Then Exit Sub
With Me
.Image1.Visible = False
x = .Width
Y = .Height - 250
If Int(x / PicFactor) <= .ScaleHeight Then
    .Image1.Width = x
    .Image1.Height = Int(x / PicFactor)
    Else
    .Image1.Height = Y
    .Image1.Width = Int(Y * PicFactor)
    End If
.Image1.Left = (.ScaleWidth - .Image1.Width) / 2
.Image1.Top = (.ScaleHeight - .Image1.Height) / 2
.Image1.Visible = True
End With
End Sub


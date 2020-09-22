VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "Starsoft Xtreme Ultravires file Encrypter Info"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4050
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8040
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
TxtFile = App.path & "\ReadMe.txt"
On Error Resume Next
FileO = FreeFile
Open TxtFile For Input As #FileO
Ftxt = Input(LOF(FileO), 1)
Close #FileO
frmText.txtInfo.Text = Ftxt
Err.Clear
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    Me.txtInfo.Width = Me.Width - 100
    Me.txtInfo.Height = Me.Height - 400
    End If
End Sub


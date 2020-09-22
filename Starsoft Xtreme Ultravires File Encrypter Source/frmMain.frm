VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starsoft Xtreme Ultravires File Encrypter"
   ClientHeight    =   3690
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":2A8B2
   ScaleHeight     =   3690
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton run 
      Caption         =   "Run"
      Height          =   975
      Left            =   5400
      Picture         =   "frmMain.frx":3EA8F4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton view 
      Caption         =   "View Picture"
      Height          =   975
      Left            =   4440
      Picture         =   "frmMain.frx":3EB6B6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton pro 
      Caption         =   "File Properties"
      Height          =   975
      Left            =   3480
      Picture         =   "frmMain.frx":3EC5BC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton crypto 
      Caption         =   "Encrypt\Decrypt"
      Height          =   975
      Left            =   2160
      Picture         =   "frmMain.frx":3ED0C6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton key 
      Caption         =   "Key"
      Height          =   975
      Left            =   1080
      Picture         =   "frmMain.frx":3EDBD0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Select 
      Caption         =   "Select"
      Height          =   975
      Left            =   0
      Picture         =   "frmMain.frx":3EE6DA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6080
      Begin VB.PictureBox picTextWidth 
         Height          =   330
         Left            =   4680
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame FrameFile 
         Height          =   50
         Left            =   840
         TabIndex        =   2
         Top             =   1365
         Width           =   5055
      End
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   5040
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label lblFileName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label lblSize 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   4995
      End
      Begin VB.Label lblVersion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   4995
      End
      Begin VB.Label lblTime 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1075
         Width           =   4995
      End
      Begin VB.Label lblSource 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   840
         TabIndex        =   1
         Top             =   1500
         Width           =   5055
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   210
         MouseIcon       =   "frmMain.frx":3EF5BC
         Top             =   375
         Width           =   480
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select..."
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuShell 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&File Properties"
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuKey 
      Caption         =   "&Key"
   End
   Begin VB.Menu mnuCrypto 
      Caption         =   "&Encode"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelpa 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help topics..."
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About File Encrypter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkPCM_Click()

End Sub

Private Sub Command1_Click()
Call GetSourceFile
End Sub

Private Sub crypto_Click()
  If KeyIsPresent = False Then Exit Sub
    Call FileCrypto
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAbout
Unload frmKey
Unload frmKeyDirect
Unload frmOptions
Unload frmPic
Unload frmProgress
Unload frmText
End Sub

Private Sub imgIcon_Click()
Call ShowPicture
End Sub

Private Sub imgIcon_DblClick()
Call ShowPicture
End Sub

Private Sub key_Click()
If gstrActiveKey <> "" Then
        RetVal = MsgBox("There is already an active key." & vbCrLf & vbCrLf & "Do you want to enter a new key ?", vbYesNo + vbQuestion, "Ultravires")
        If RetVal = vbNo Then Exit Sub
        End If
    frmKey.Show (vbModal)
End Sub

Private Sub lblSource_Change()
SetIcon
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (vbModal)
End Sub

Private Sub mnuCrypto_Click()
If KeyIsPresent = False Then Exit Sub
Call FileCrypto
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelp_Click()
frmText.Show
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show (vbModal)
End Sub

Private Sub mnuPCM_Click()

End Sub

Private Sub mnuDeletePCM_Click()

End Sub

Private Sub mnuKey_Click()
If gstrActiveKey <> "" Then
    RetVal = MsgBox("There is already an active key." & vbCrLf & vbCrLf & "Do you want to enter a new key ?", vbYesNo + vbQuestion, "Ultravires")
    If RetVal = vbNo Then Exit Sub
    End If
frmKey.Show (vbModal)
End Sub

Private Sub mnuProperties_Click()
Call ShowPropWindow(gstrSourceFile, frmMain.hwnd)
End Sub

Private Sub mnuSelect_Click()
Call GetSourceFile
End Sub

Private Sub mnuShell_Click()
Call StartFile(gstrSourceFile)
End Sub

Private Sub mnuView_Click()
Call ShowPicture
End Sub

Private Sub pro_Click()
 Call ShowPropWindow(gstrSourceFile, frmMain.hwnd)
End Sub

Private Sub run_Click()
 Call StartFile(gstrSourceFile)
End Sub

Private Sub Select_Click()
 Call GetSourceFile
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
Case "file"
    Call GetSourceFile
Case "key"
    If gstrActiveKey <> "" Then
        RetVal = MsgBox("There is already an active key." & vbCrLf & vbCrLf & "Do you want to enter a new key ?", vbYesNo + vbQuestion, "Ultravires")
        If RetVal = vbNo Then Exit Sub
        End If
    frmKey.Show (vbModal)
Case "crypto"
    If KeyIsPresent = False Then Exit Sub
    Call FileCrypto
Case "properties"
    Call ShowPropWindow(gstrSourceFile, frmMain.hwnd)
Case "view"
    Call ShowPicture
Case "run"
    Call StartFile(gstrSourceFile)
End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub lblSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub lblFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RandomFeed(X, Y)
End Sub

Private Sub view_Click()
Call ShowPicture
End Sub

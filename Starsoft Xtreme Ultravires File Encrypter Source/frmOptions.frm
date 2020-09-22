VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Starsoft Xtreme Ultra File Encrypter Options"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Files"
      Height          =   960
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   5985
      Begin VB.CommandButton cmdRegFile 
         Caption         =   "&Register FileType ucc"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Keys"
      Height          =   750
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   6000
      Begin VB.CheckBox chkKeys 
         Caption         =   "&Enter new key for each file"
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   315
         Width           =   5580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default folder"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6000
      Begin VB.CommandButton cmdDefaultOpen 
         Caption         =   "..."
         Height          =   275
         Left            =   210
         TabIndex        =   0
         Top             =   375
         Width           =   330
      End
      Begin VB.Label lblFolderPath 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   630
         TabIndex        =   5
         Top             =   375
         Width           =   5150
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4830
      TabIndex        =   4
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3465
      TabIndex        =   3
      Top             =   3000
      Width           =   1275
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tmpFolder As String
Private tmpLng As String

Private Sub Form_Activate()
Me.lblFolderPath.Caption = TrimPath(gstrDefaultFolder, Me.lblFolderPath.Width)
tmpFolder = ""
If gblnAllwaysNewKey = True Then
    Me.chkKeys.Value = 1
    Else
    Me.chkKeys.Value = 0
    End If
If gstrFileReg = "" Then
    Me.cmdRegFile.Caption = "&Register ucc filetype"
    Else
    Me.cmdRegFile.Caption = "&Unregister ucc filetype"
    End If
End Sub

Private Sub cmdRegFile_Click()
If gstrFileReg = "" Then
    'register filetype
    Call MakeFileAssociation("ucc", App.path, App.EXEName, "ULTRAVIRES Crypto Code", App.path & "\" & "UCC.ico")
    MsgBox "The filetype ucc will be recognized after restarting the computer.", vbInformation, "Ultravires"
    Else
    RetVal = MsgBox("Are you sure you want to unregister the ucc filetype?", vbQuestion + vbYesNo, "Ultravires")
    If RetVal = vbNo Then Exit Sub
    'delete filetype
    Call DeleteFileAssociation("ucc")
    MsgBox "The filetype ucc is unregistred after restarting the computer.", vbInformation, "Ultravires"
    End If
gstrFileReg = ReadKey(HKEY_CLASSES_ROOT, ".ucc", "", "")
If gstrFileReg = "" Then
    Me.cmdRegFile.Caption = "&Register ucc filetype"
    Else
    Me.cmdRegFile.Caption = "&Unregister ucc filetype"
    End If
End Sub

Private Sub cmdDefaultOpen_Click()
Dim tmp As String
tmp = Browse("Select default directorie...")
If tmp <> "" Then
    tmpFolder = tmp
    Me.lblFolderPath.Caption = TrimPath(tmp, Me.lblFolderPath.Width)
    End If
End Sub

Private Sub cmdOK_Click()
If tmpFolder <> "" Then
    SaveSetting App.EXEName, "Config", "Folder", tmpFolder
    gstrDefaultFolder = tmpFolder
    End If
gblnAllwaysNewKey = Me.chkKeys.Value
SaveSetting App.EXEName, "Config", "NewKeys", gblnAllwaysNewKey
Me.Hide
Call CheckMenus
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About"
   ClientHeight    =   3705
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1078.603
   ScaleMode       =   0  'User
   ScaleWidth      =   1123.685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2220
      Left            =   105
      ScaleHeight     =   2160
      ScaleWidth      =   900
      TabIndex        =   5
      Top             =   105
      Width           =   960
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmAbout.frx":0000
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   105
      TabIndex        =   4
      Top             =   2520
      Width           =   5790
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   362
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   1365
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   362
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   1350
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   1380
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   4635
   End
   Begin VB.Label lblTitle 
      Caption         =   "Starsoft Xtreme                             Ultra File Encrypter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1260
      TabIndex        =   2
      Top             =   0
      Width           =   4725
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 2.00.03"
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   720
      Width           =   4725
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.lblDescription.Caption = "Encryption program to encrypt files and create Self-decrypting files  using the ULTRAVIRES v1.0.0 Cryptographic algorithm." & vbCrLf & vbCrLf & "Programming and design by Ajin Abraham"
End Sub

Private Sub cmdSysInfo_Click()
On Error Resume Next
Dim SysInfoPath As String
SysInfoPath = ReadKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Shared Tools\MSINFO", "PATH", "")
If SysInfoPath = "" Then
    SysInfoPath = ReadKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Shared Tools Location", "MSINFO", "")
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        End If
    End If
Call Shell(SysInfoPath, vbNormalFocus)
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub



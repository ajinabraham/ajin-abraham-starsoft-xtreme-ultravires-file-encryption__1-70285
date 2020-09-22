VERSION 5.00
Begin VB.Form frmProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Encrypting file..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   3570
      TabIndex        =   1
      Top             =   1155
      Width           =   1170
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00008000&
      Height          =   200
      Left            =   105
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   840
      Width           =   4635
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "frmProgress.frx":0000
      Top             =   80
      Width           =   450
   End
   Begin VB.Label lblName 
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   3690
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
AbortUltraRun = True
End Sub

Private Sub Form_Activate()
With Me
Me.Refresh
.Caption = ""
.lblName.Caption = ""
.lblName.Refresh
If CurCryptoType = TypeFile Then
    If gstrCryptoVersion <> "" Then
        .picProgress.ForeColor = &HFF& 'red
        .Caption = " Decrypt file..."
        .lblName = "Busy decrypting " & CutFilePath(gstrSourceFile) & " ..."
        Else
        .picProgress.ForeColor = &H8000&     'green
        .Caption = " Encrypting file..."
        .lblName = "Busy encrypting " & CutFilePath(gstrSourceFile) & " ..."
        End If
        gstrReturnName = UltraFile(gstrSourceFile, gstrTargetFile, gstrActiveKey, gstrActivePCM)
    End If
End With
Select Case UltraReturnValue
    Case 0
        'no errors
        If gstrSourceFile = gstrTargetFile Then
            'overwrite source
            gstrTargetFile = gstrReturnName
            gstrSourceFile = gstrReturnName
            frmMain.lblFileName.Caption = CutFilePath(gstrSourceFile)
            frmMain.lblSource.Caption = TrimPath(gstrSourceFile, frmMain.lblSource.Width)
            frmMain.Caption = "ULTRAVIRES - " & CutFilePath(gstrSourceFile)
        Else
            gstrTargetFile = gstrReturnName
        End If
        Call SetProps
    Case 4
        MsgBox "File not found.", vbCritical, "Ultravires"
    Case 6
        MsgBox "The file does not contain data.", vbCritical, "Ultravires"
    Case 11
        MsgBox "Encrypting files aborted.", vbInformation, "Ultravires"
    Case 12
        MsgBox "Failed encrypting the file:" & vbCrLfLf & UltraReturnString, vbCritical, "Ultravires"
    Case 21
        MsgBox "Decrypting file aborted.", vbInformation, "Ultravires"
    Case 22
        MsgBox "Failed decrypting the file:" & vbCrLfLf & UltraReturnString, vbCritical, "Ultravires"
    Case 23
        MsgBox "Failed decrypting the file. This may be caused by the following:" & vbCrLf & vbCrLf & "- Wrong key or key contains errors." & vbCrLf & "- The Private Crypto Code is not compatible." & vbCrLf & "- The file contains errors or has been damaged", vbCritical, "Ultravires"
    Case Else
        MsgBox UltraReturnString, vbCritical
End Select
Me.Refresh
frmProgress.lblName.Refresh
Call CheckMenus
Me.Hide
End Sub


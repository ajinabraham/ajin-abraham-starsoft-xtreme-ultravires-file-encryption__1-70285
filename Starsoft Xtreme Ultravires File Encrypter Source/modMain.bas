Attribute VB_Name = "modMain"
'Extract icon from files
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id((123)) As Byte
End Type
Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private bMoveFrom As Boolean
'Get computername
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Run or open file
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Show File Properties window
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Type SHELLEXECUTEINFO
       cbSize As Long
       fMask As Long
       hwnd As Long
       lpVerb As String
       lpFile As String
       lpParameters As String
       lpDirectory As String
       nShow As Long
       hInstApp As Long
       lpIDList As Long 'Optional
       lpClass As String 'Optional
       hkeyClass As Long 'Optional
       dwHotKey As Long 'Optional
       hIcon As Long 'Optional
       hProcess As Long 'Optional
End Type
'Browsing folders
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public gstrCryptoVersion    As String
Public gstrActiveKey        As String
Public gstrActivePCM        As String
Public gblnAllwaysNewKey    As Boolean
Public gstrCurrentFolder    As String
Public gstrDefaultFolder    As String
Public gstrFileReg          As String
Public gstrSourceFile       As String
Public gstrTargetFile       As String
Public gstrReturnName       As String
Public vbCrLfLf             As String
Public CurrentLang          As String
Public PCMenable            As Boolean
Public CommandStart         As Boolean
Public CryptoType           As String
Public CurCryptoType        As CryptoType
Public CurSourceType        As SourceType
Public RetVal               As Integer
Public g_intCurrProgress    As Integer
Public PicFactor

Public Enum CryptoType
    TypeFile
    TypeUSD
End Enum

Public Enum SourceType
    SourceText
    SourcePicture
    SourceOther
    SourceCrypto
    SourceNon
End Enum

Public Const BnAbout = 4
Public Const BnSource = 5
Public Const BnInfo = 6
Public Const BnStart = 7
Public Const BnTarget = 8
Public Const BnMini = 9
Public Const BnExit = 10
Public Const BnOff = 11
Public Const BnText = 12

Public Const FILE_VERSION = "= ULTRAVIRES  ="

Public Const RATE_ENCODE = 318000
Public Const RATE_DECODE = 318000

Public Const PROGRESS_CALCFREQUENCY = 3
Public Const PROGRESS_CALCCRC = 3
Public Const PROGRESS_ENCODEHUFF = 44
Public Const PROGRESS_DECODEHUFF = 45
Public Const PROGRESS_CHECKCRC = 5
Public Const PROGRESS_ENCRYPT = 50
Public Const PROGRESS_DECRYPT = 50

Public Const CMDLG_NOCHECK = &H4
Public Const CMDLG_NOOVERWRITE = &H2
Public Const CMDLG_PATHMUSTEXIST = &H800
Public Const CMDLG_FILEMUSTEXIST = &H1000

Sub Main()
vbCrLfLf = vbCrLf & vbCrLf
Load frmMain
Load frmPic
Load frmKey
Load frmProgress
Load frmOptions
'check for ucc file type registration
gstrFileReg = ReadKey(HKEY_CLASSES_ROOT, ".ucc", "", "")
'get newkey
gblnAllwaysNewKey = GetSetting(App.EXEName, "Config", "NewKeys", 0)
'get default folder
gstrDefaultFolder = GetSetting(App.EXEName, "Config", "Folder", "c:\")
'get pcm
gstrActivePCM = LoadPCM

frmMain.Show
frmProgress.picProgress.ScaleWidth = 100
SetProps
'command
With frmMain
If Command <> "" Then
    gstrSourceFile = Command
    .lblSource.Caption = TrimPath(gstrSourceFile, frmMain.lblSource.Width)
    .lblFileName.Caption = CutFilePath(gstrSourceFile)
    .Caption = CutFilePath(gstrSourceFile)
    gstrCurrentFolder = GetFilePath(gstrSourceFile)
    CommandStart = True
    SetProps
    If UCase(Right(gstrSourceFile, 4)) = ".UCC" And gstrCryptoVersion = "" Then
        'unknown version
        Else
        Call FileCrypto
        End If
    Else
    ChDir gstrDefaultFolder
    gstrCurrentFolder = gstrDefaultFolder
    End If
End With
CommandStart = False
End Sub

Public Sub SetProps()
Dim convMin As Integer
Dim convSec As Integer
Dim tmp As String
Dim fAttrib As Integer
Dim AttMsg As String
Dim VersionBuffer() As Byte
Dim RetVersion As Integer
On Error Resume Next
Dim FileO As Integer
If FileExist(gstrSourceFile) Then
    frmMain.mnuCrypto.Enabled = True
    tmp = UCase(Right(gstrSourceFile, 4))
    If tmp = ".JPG" Or tmp = "JPEG" Or tmp = ".BMP" Or tmp = ".WMF" Or tmp = ".GIF" Then
        frmMain.imgIcon.ToolTipText = " Click here for a quick view "
        frmMain.imgIcon.MousePointer = 99
        CurSourceType = SourcePicture
        If frmPic.Visible = True Then Call ShowPicture
        Else
        frmMain.imgIcon.ToolTipText = ""
        frmMain.imgIcon.MousePointer = 0
        If tmp = ".TXT" Then
            CurSourceType = SourceText
            Else
            CurSourceType = SourceOther
            End If
        End If
    frmMain.lblSize = Format(FileLen(gstrSourceFile), "###,###,###,##0") & " Bytes"
    fAttrib = GetAttr(gstrSourceFile)
    AttMsg = ""
    If fAttrib And 4 Then AttMsg = "Systemfile  "
    If fAttrib And 1 Then AttMsg = AttMsg & "Read-only  "
    If fAttrib And 2 Then AttMsg = AttMsg & "Hidden"
    If AttMsg <> "" Then frmMain.lblSize = frmMain.lblSize & "  (" & Trim(AttMsg) & ")"
    'get crypto info from file
    RetVersion = CheckUltraFile(gstrSourceFile)
    Select Case RetVersion
        Case 0, 2 'no crypto/unknown
            If RetVersion = 0 Then
                frmMain.lblVersion = "Unprotected"
                Else
                frmMain.lblVersion = "Unknown crypto format"
                End If
            gstrCryptoVersion = ""
            convRate = RATE_ENCODE
            frmMain.mnuCrypto.Caption = "&Encrypt"
            frmMain.lblVersion.ForeColor = &HFF&
           
        Case 1 'ultra
            frmMain.lblVersion = "File Protected"
            gstrCryptoVersion = FILE_VERSION
            convRate = RATE_DECODE
            frmMain.mnuCrypto.Caption = "&Decrypt"
            frmMain.lblVersion.ForeColor = &H8000& '&H80000012
           
    End Select
    convTotSec = FileLen(gstrSourceFile) / convRate
    convMin = Int(convTotSec / 60)
    convSec = Int(convTotSec - (convMin * 60))
    If convSec = 0 Then convSec = 1
    frmMain.lblTime.Caption = "Estimated conversion time: " & Trim(Str(convMin)) & " min " & Trim(Str(convSec)) & " sec "
    Else
    frmMain.mnuCrypto.Enabled = False
    frmMain.lblSource = ""
    frmMain.lblSize = ""
    frmMain.imgIcon.Picture = Nothing
    frmMain.imgIcon.ToolTipText = ""
    CurSourceType = SourceNon
    End If
Call CheckMenus
End Sub

Public Sub GetSourceFile()
'get source
On Error Resume Next
With frmMain.ComDlg
.FileName = ""
.DialogTitle = "Select file..."
.Filter = "All files (*.*)|*.*|Ultravires Crypto Code files (*.ucc)|*.ucc"
.InitDir = gstrCurrentFolder
.FilterIndex = 1
.Flags = CMDLG_NOCHECK Or CMDLG_FILEMUSTEXIST
.ShowOpen
If Err <> 32755 Then ' cancel
    If gblnAllwaysNewKey = True Then
        gstrActiveKey = Space(Len(gstrActiveKey))
        gstrActiveKey = ""
        End If
    gstrSourceFile = .FileName
    gstrCurrentFolder = CurDir$
    frmMain.lblSource.Caption = TrimPath(.FileName, frmMain.lblSource.Width)
    frmMain.lblFileName.Caption = CutFilePath(gstrSourceFile)
    frmMain.Caption = "Ultravires - " & CutFilePath(gstrSourceFile)
    Call SetProps
    If gstrCryptoVersion <> "" And gstrActiveKey = "" Then
        CommandStart = True
        Call FileCrypto
        End If
    End If
End With
End Sub

Public Function GetTargetFile() As String
'get source
On Error Resume Next
With frmMain.ComDlg
.FileName = ""
.DialogTitle = "Select target file..."
.Filter = "All files (*.*)|*.*"
.InitDir = gstrCurrentFolder
.FilterIndex = 1
.Flags = CMDLG_NOCHECK Or CMDLG_NOOVERWRITE  'Or &H2
.FileName = ""
.ShowSave
If Err = 32755 Then ' cancel
    GetTargetFile = ""
    Else
    GetTargetFile = .FileName
    gstrCurrentFolder = CurDir$
    End If
End With
End Function

Public Sub ShowPicture()
Dim tmp As String
If Len(gstrSourceFile) < 5 Then Exit Sub
tmp = UCase(Right(gstrSourceFile, 4))
If tmp <> ".JPG" And tmp <> ".ICO" And tmp <> ".BMP" And tmp <> ".WMF" And tmp <> ".GIF" Then Exit Sub
On Error Resume Next
With frmPic
If .Visible = False Then
    .Height = 2800
    .Width = 3800
    Else
    If .Height > .Width Then
        .Width = .Height
        Else
        .Height = .Width
        End If
    End If
.Image1.Visible = False
.Refresh
.Image1.Stretch = False
frmMain.MousePointer = 11
.Image1.Picture = LoadPicture(gstrSourceFile)
frmMain.MousePointer = 0
PicFactor = .Image1.Width / .Image1.Height
.Image1.Stretch = True
If PicFactor = 0 Then Exit Sub
'.Image1.Visible = False
X = .Width
Y = .Height - 250
If Int(X / PicFactor) <= .ScaleHeight Then
    .Image1.Width = X
    .Image1.Height = Int(X / PicFactor)
    .Height = .Image1.Height + 400
    Else
    .Image1.Height = Y
    .Image1.Width = Int(Y * PicFactor)
    .Width = .Image1.Width
    End If
.Image1.Left = (.ScaleWidth - .Image1.Width) / 2
.Image1.Visible = True
.Caption = CutFilePath(gstrSourceFile)
End With
If frmPic.Visible = False Then frmPic.Show
End Sub

Public Function TrimPath(ByVal Text As String, ByVal Size As Long)
Dim TW
Dim Part1 As String
Dim Part2 As String
Dim pos As Integer
Size = Size - 200
TW = frmMain.picTextWidth.TextWidth(Text)
If TW < Size Then
    TrimPath = Text
    Exit Function
    End If
Part1 = Left(Text, 3) & "...\"
Part2 = Mid(Text, 4)
Text = Part1 & Part2
Do
TW = frmMain.picTextWidth.TextWidth(Text)
If TW >= (Size) Then
    pos = InStr(1, Part2, "\")
    If pos <> 0 And pos < Len(Part2) Then
        Part2 = Mid(Part2, pos + 1)
        Else
        Part2 = Mid(Part2, 2)
        End If
    Text = Part1 & Part2
    End If
Loop While TW > (Size)
TrimPath = Text
End Function

Public Sub EndProgram()
Unload frmMain
End Sub

Public Function KeyIsPresent() As Boolean
If gstrActiveKey = "" Then
    frmKey.Show (vbModal)
    If gstrActiveKey <> "" Then KeyIsPresent = True
    Else
    KeyIsPresent = True
    End If
End Function

Public Sub CheckMenus()
With frmMain
If CurSourceType = SourceNon Then
    .mnuCrypto.Enabled = False
    .mnuProperties.Enabled = False
    .mnuShell.Enabled = False
   
    .pro.Visible = False
    .run.Visible = False
    .crypto.Visible = False
    .view.Visible = False
    .FrameFile.Visible = False
    Else
    .mnuCrypto.Enabled = True
    .mnuProperties.Enabled = True
    .mnuShell.Enabled = True
    .pro.Visible = True
    .run.Visible = True
    .crypto.Visible = True
    .view.Visible = True
    .FrameFile.Visible = True

    End If
If CurSourceType = SourcePicture Then
    .mnuView.Enabled = True
    .view.Enabled = True
    Else
    .mnuView.Enabled = False
    .view.Enabled = False
    End If

If gstrCryptoVersion <> "" Then
    'known crypto version
    CurSourceType = SourceCrypto
    .mnuShell.Enabled = False
    .run.Enabled = False
    .crypto.Enabled = True
    .mnuCrypto.Enabled = True
    Else
    If UCase(Right(gstrSourceFile, 4)) = ".UCC" Then
        'unknown version
        .run.Enabled = False
        .mnuShell.Enabled = False
        .crypto.Enabled = False
        .mnuCrypto.Enabled = False
        Else
        If CurSourceType <> SourceNon Then .mnuShell.Enabled = True
        .run.Enabled = True
        .mnuShell.Enabled = False
        .crypto.Enabled = True
        .mnuCrypto.Enabled = True
        End If
    End If
End With
End Sub

Public Sub StartFile(ByVal FileName As String)
Dim RunCmd As String
Dim fExt As String
Dim X As Long
Dim RetS
On Error Resume Next
fExt = UCase(Right(FileName, 4))
Select Case fExt
Case ".WAV", ".MP2", ".MP3", ".MID", ".AVI"
    RunCmd = "play"
Case Else
    RunCmd = "open"
End Select
If FileExist(FileName) = False Or FileName = "" Then Exit Sub
'open file
RetS = ShellExecute(hwnd, RunCmd, FileName, "", App.path, 1)
If RetS = 31 Then
    'if open fails, show the 'open with...'
    RetS = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & gstrSourceFile)
    If RetS = 31 Then
        MsgBox "Unkwown filetype.", vbInformation, "Ultravires"
        End If
    End If
End Sub

Public Function GetPCname() As String
Dim sBuffer As String
Dim lBufSize As Long
Dim lStatus As Long
lBufSize = 255
sBuffer = String$(lBufSize, " ")
lStatus = GetComputerName(sBuffer, lBufSize)
GetPCname = ""
If lStatus <> 0 Then
    GetPCname = Left(sBuffer, lBufSize)
End If
End Function

Public Sub SetIcon()
Dim fName As String
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown
fName = Trim(gstrSourceFile)
If fName = "" Then frmMain.imgIcon.Picture = Nothing: Exit Sub
SHGetFileInfo fName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
hIcon = sh_info.hIcon
With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set icon_pic = lpUnk
frmMain.imgIcon = icon_pic
End Sub

Public Function Browse(ByVal aTitle As String) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, pos%
Dim BrowsePath As String
bInfo.hOwner = frmOptions.hwnd
bInfo.lpszTitle = aTitle
'the type of folder(s) to return
bInfo.ulFlags = &H1
'show the dialog box
pidl& = SHBrowseForFolder(bInfo)
'set the maximum characters
path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
'set the extracted path to SpecIn
Browse = Left(path$, pos - 1)
'make sure that "\" is at the end of the path
If Right$(Browse, 1) = "\" Then
    Browse = Browse
    Else
    Browse = Browse + "\"
End If
If Browse = "\" Then Browse = ""
End Function


Public Function ShowPropWindow(FileName As String, OwnerhWnd As Long) As Long
'open a file properties property page for specified file if return value
'<=32 an error occurred
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
'Fill in the SHELLEXECUTEINFO structure
If FileExist(FileName) = False Or FileName = "" Then
    ShowPropWindow = 0
    Exit Function
    End If
With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = OwnerhWnd
    .lpVerb = "properties"
    .lpFile = FileName
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With
'call the API
r = ShellExecuteEX(SEI)
'return the instance handle as a sign of success
ShowPropWindow = SEI.hInstApp
End Function

Public Sub FileCrypto()
If CommandStart = True Then
    CommandStart = False
    frmKeyDirect.Show (vbModal)
    If gstrActiveKey = "" Then Exit Sub
    End If
RetVal = MsgBox("Do You want to overwrite the original file ?", vbYesNoCancel + vbQuestion, "Ultravires")
If RetVal = vbCancel Then
    Exit Sub
ElseIf RetVal = vbNo Then
    gstrTargetFile = GetTargetFile
    If gstrTargetFile = "" Then Exit Sub
ElseIf RetVal = vbYes Then
    gstrTargetFile = gstrSourceFile
    End If
CurCryptoType = TypeFile
frmProgress.Show (vbModal)
End Sub

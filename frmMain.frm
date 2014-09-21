VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Directory Update"
   ClientHeight    =   6396
   ClientLeft      =   1596
   ClientTop       =   840
   ClientWidth     =   8388
   Icon            =   "frmMain.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6396
   ScaleWidth      =   8388
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   6360
      Top             =   1920
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "&Include subfolders"
      Height          =   192
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1692
   End
   Begin VB.Frame frameOptions 
      Caption         =   "O&ptions"
      Height          =   1152
      Left            =   120
      TabIndex        =   16
      Top             =   1380
      Width           =   5232
      Begin VB.TextBox txtDateModified 
         Height          =   312
         Left            =   2100
         TabIndex        =   19
         Top             =   780
         Width           =   2832
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "&Full Update (includes comparison of entire directory trees)..."
         Height          =   252
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   4512
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "&Quick Update"
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Value           =   -1  'True
         Width           =   1272
      End
      Begin VB.Label lblDateModified 
         AutoSize        =   -1  'True
         Caption         =   "Files modified on or after:"
         Height          =   192
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1788
      End
   End
   Begin VB.TextBox txtTarget 
      Height          =   288
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5472
   End
   Begin VB.CommandButton cmdBrowseTarget 
      Caption         =   "&Browse"
      Height          =   312
      Left            =   5820
      TabIndex        =   3
      Top             =   960
      Width           =   912
   End
   Begin VB.CommandButton cmdBrowseSource 
      Caption         =   "&Browse"
      Height          =   312
      Left            =   5820
      TabIndex        =   1
      Top             =   300
      Width           =   912
   End
   Begin VB.TextBox txtSource 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   5472
   End
   Begin ComctlLib.ListView lvwSource 
      Height          =   3132
      Left            =   180
      TabIndex        =   8
      Top             =   2880
      Width           =   8112
      _ExtentX        =   14309
      _ExtentY        =   5525
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Enabled         =   0   'False
      Height          =   372
      Left            =   7020
      TabIndex        =   10
      Top             =   720
      Width           =   1092
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   372
      Left            =   7020
      TabIndex        =   9
      Top             =   300
      Width           =   1092
   End
   Begin ComctlLib.StatusBar sbBottom 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   5
      Top             =   6084
      Width           =   8388
      _ExtentX        =   14796
      _ExtentY        =   550
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11896
            Key             =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Size"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pboxIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   7620
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pboxSmallIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7260
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame frameSelected 
      Caption         =   "Selected Files"
      Height          =   3432
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   8232
      Begin VB.Label lblLogFile 
         AutoSize        =   -1  'True
         Caption         =   "Log File: <filename>"
         Height          =   192
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1428
      End
   End
   Begin VB.Frame frameTarget 
      Caption         =   "Target Directory"
      Height          =   612
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   6672
   End
   Begin VB.Frame frameSource 
      Caption         =   "Source Directory"
      Height          =   612
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   6672
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   5880
      Top             =   1860
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   5460
      Top             =   1860
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Index           =   1
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All Files"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About DirectoryUpdate..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const lvwSubItem_Path = 0
Const lvwSubItem_RawSize = 1
Const lvwSubItem_Size = 2
Const lvwSubItem_Type = 3
Const lvwSubItem_RawDateModified = 4
Const lvwSubItem_DateModified = 5
Const KB& = 1024
Const MB& = KB * 1024
Const GB& = MB * 1024

Public AppName As String
Public ComputerName As String
Public UserName As String
Private mItem As ListItem
Private strLogFile As String
Dim LogFileUnit As Integer
Dim LogFileUnit2 As Integer
Dim ButtonOffset As Integer
Dim BrowseOffset As Integer
Dim lblLogFileTop As Integer
Dim lvwSourceOffset As Integer
Dim lvwSourceHeight As Integer
Dim frameSelectedOffset As Integer
Dim frameSelectedHeight As Integer
Dim BoxesOffset As Integer
Dim FramesOffset As Integer
Dim OriginalFormWidth As Integer
Dim OriginalFormHeight As Integer
Dim gfWebServer As Boolean
Dim strFileName As String
Dim mnuHelpAboutHit As Boolean
Dim cmdSourceBrowseHit As Boolean
Dim cmdTargetBrowseHit As Boolean
Dim fNeedSave As Boolean
Dim fSearching As Boolean
Dim fQuickCopy As Boolean

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub CheckSaveState()
    If fNeedSave Then
        If MsgBox("Configuration has changed, save before exit?", vbYesNo, App.Title) = vbYes Then mnuFileSave_Click
    End If
End Sub
Private Sub CreateDirPath(SourceFileName As String, TargetFileName As String)
    Dim i As Integer
    Dim DirSegment As String
    Dim SourceRoot As String
    Dim TargetRoot As String
    
    On Error GoTo 0
    SourceRoot = txtSource.Text
    TargetRoot = txtTarget.Text
    
    For i = Len(SourceRoot) + 1 To Len(SourceFileName)
        If Mid(SourceFileName, i, 1) = "\" Then
            DirSegment = Mid(SourceFileName, i, InStr(Right(SourceFileName, Len(SourceFileName) - i), "\"))
            If DirSegment = "" Then Exit For
         
            If (GetAttr(SourceRoot & DirSegment) And vbDirectory) = vbDirectory Then
                If Dir(TargetRoot & DirSegment, vbDirectory) = "" Then
                    On Error Resume Next
                    MkDir TargetRoot & DirSegment
                    If Err.Number <> 0 Then
                        Print #LogFileUnit, Now() & Chr(9) & "MkDir failed on """ & TargetRoot & DirSegment & """, Error #" & Err.Number & "; " & Err.Description
                        Print #LogFileUnit2, Now() & Chr(9) & "MkDir failed on """ & TargetRoot & DirSegment & """, Error #" & Err.Number & "; " & Err.Description
                        Err.Clear
                    Else
                        Print #LogFileUnit, Now() & Chr(9) & "Created " & TargetRoot & DirSegment & "..."
                        Print #LogFileUnit2, Now() & Chr(9) & "Created " & TargetRoot & DirSegment & "..."
                    End If
                    On Error GoTo 0
                End If
            End If
            SourceRoot = SourceRoot & DirSegment
            TargetRoot = TargetRoot & DirSegment
        End If
    Next i
End Sub
Private Function FileExists(strPath As String) As Boolean
    Dim vTargetTime As Variant
    
    On Error Resume Next
    'Don't use Dir() to check file existence, 'cause we might be in the midst of a search loop...
    vTargetTime = FileDateTime(strTargetPath & strDir)
    If Err.Number = 53 Then
        FileExists = False
    Else
        FileExists = True
    End If
    Err.Clear
    On Error GoTo 0
End Function
Private Function GetFiles(strSourcePath As String, strTargetPath As String, ByRef intTotalBytes As Long) As Integer
    Dim strDir As String
    Dim Icon As Integer
    Dim strIcon As String
    Dim intTotalFiles As Integer
    Dim fContinue As Boolean
   
    If Right(strSourcePath, 1) <> "\" Then strSourcePath = strSourcePath & "\"
    If Right(strTargetPath, 1) <> "\" Then strTargetPath = strTargetPath & "\"
   
    'Debug.Print "On entry to GetFiles(): strSourcePath is """ & strSourcePath & """..."
   
    strDir = Dir(strSourcePath, vbDirectory)
    Do While strDir <> ""   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        fContinue = True
        Select Case Left(strDir, 1)
            Case "."
            Case "_"
            Case Else
                If gfWebServer Then
                    If UCase(Right(strSourcePath, 5)) = "\LOG\" Or UCase(Right(strSourcePath, 6)) = "\LOGS\" Then ' Ignore log files in web servers...
                        sbBottom.Panels("Status").Text = "Skipping web site's " & strSourcePath & strDir & "..."
                        fContinue = False
                    End If
                End If
            
                If fContinue Then
                    'Debug.Print strDir & "'s Date Modified: " & FileDateTime(strSourcePath & strDir) & "(" & DateDiff("s", txtDateModified.Text, FileDateTime(strSourcePath & strDir)) & " seconds)..."
                    If (GetAttr(strSourcePath & strDir) And vbDirectory) = vbDirectory Then
                        If chkSubFolders = 1 Then
                            sbBottom.Panels("Status").Text = "Searching " & strSourcePath & strDir & "..."
                     
                            intTotalFiles = intTotalFiles + _
                                GetFiles(strSourcePath & strDir, strTargetPath & strDir, intTotalBytes)
                                
                            fContinue = True
                            If Not FileExists(strTargetPath & strDir) Then
                                fContinue = True
                            ElseIf DateDiff("s", txtDateModified.Text, FileDateTime(strSourcePath & strDir)) <= 0 Then
                                fContinue = False
                            End If
                            
                            If fContinue Then
                                LoadFileEntry strSourcePath & strDir
                                intTotalBytes = intTotalBytes + FileLen(strSourcePath & strDir)
                                intTotalFiles = intTotalFiles + 1
                            End If
                            RePosition strSourcePath, strDir
                            DoEvents
                        End If
                    Else
                        fContinue = True
                        If Not fQuickCopy And Not FileExists(strTargetPath & strDir) Then
                            fContinue = True
                        ElseIf DateDiff("s", txtDateModified.Text, FileDateTime(strSourcePath & strDir)) <= 0 Then
                            fContinue = False
                        End If
                        
                        If fContinue Then
                            LoadFileEntry strSourcePath & strDir
                            intTotalBytes = intTotalBytes + FileLen(strSourcePath & strDir)
                            intTotalFiles = intTotalFiles + 1
                        End If
                    End If
                End If
        End Select
        strDir = Dir   ' Get next entry.
        'DoEvents
        If Not fSearching Then Exit Do
    Loop
    GetFiles = intTotalFiles
End Function
Function GetFileType(pFileName As String) As String
    Dim Temp As Long
    Dim i As Integer
   
    Temp = SHGetFileInfo(pFileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS)
    GetFileType = Left(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr(0)) - 1)
End Function
Private Sub LoadFileEntry(PathName As String)
    Dim Icon As Integer
    Dim strIcon As String
    Dim FileSizeKB As Long
   
    'Debug.Print "Loading Entry: " & PathName
   
    Icon = genRetrieveIcon(PathName, pboxIcon, pboxSmallIcon, imlIcons, imlSmallIcons)
    strIcon = genstrRetrieveIcon(PathName, pboxIcon, pboxSmallIcon, imlIcons, imlSmallIcons)
    lvwSource.Icons = imlIcons
    lvwSource.SmallIcons = imlSmallIcons
    Set mItem = lvwSource.ListItems.Add()
    mItem.Text = PathName
    mItem.Key = PathName
    mItem.SmallIcon = strIcon
    mItem.Icon = strIcon
    mItem.SubItems(lvwSubItem_RawSize) = Format(FileLen(PathName), "000000000000")
    FileSizeKB = FileLen(PathName) \ 1024
    If FileSizeKB = 0 Then FileSizeKB = 1
    mItem.SubItems(lvwSubItem_Size) = Format(FileSizeKB, "##,##0 KB")
    mItem.SubItems(lvwSubItem_Type) = GetFileType(PathName)
    mItem.SubItems(lvwSubItem_RawDateModified) = DateDiff("s", FileDateTime(PathName), txtDateModified.Text)
    mItem.SubItems(lvwSubItem_DateModified) = FileDateTime(PathName)
    mItem.Selected = True
    
    If Len(PathName) > frmMain.ScaleX(lvwSource.ColumnHeaders("Path").Width, vbTwips, vbCharacters) Then
        If frmMain.ScaleX(Len(PathName), vbCharacters, vbTwips) <= (lvwSource.Width \ 2) Then lvwSource.ColumnHeaders("Path").Width = frmMain.ScaleX(Len(PathName), vbCharacters, vbTwips)
    End If
    If Len(mItem.SubItems(lvwSubItem_Size)) > frmMain.ScaleX(lvwSource.ColumnHeaders("Size").Width, vbTwips, vbCharacters) Then lvwSource.ColumnHeaders("Size").Width = frmMain.ScaleX(Len(mItem.SubItems(lvwSubItem_Size)), vbCharacters, vbTwips)
    If Len(mItem.SubItems(lvwSubItem_Type)) > frmMain.ScaleX(lvwSource.ColumnHeaders("Type").Width, vbTwips, vbCharacters) Then lvwSource.ColumnHeaders("Type").Width = frmMain.ScaleX(Len(mItem.SubItems(lvwSubItem_Type)), vbCharacters, vbTwips)
    If Len(mItem.SubItems(lvwSubItem_DateModified)) > frmMain.ScaleX(lvwSource.ColumnHeaders("Date Modified").Width, vbTwips, vbCharacters) Then lvwSource.ColumnHeaders("Date Modified").Width = frmMain.ScaleX(Len(mItem.SubItems(lvwSubItem_DateModified)), vbCharacters, vbTwips)
End Sub
Private Sub OpenConfiguration()
    If strFileName <> "" Then
        Me.MousePointer = vbHourglass
        FileUnit = FreeFile
        Open strFileName For Input Access Read As #FileUnit
        Do While Not EOF(FileUnit)
            On Error GoTo ReadError
            Line Input #FileUnit, txtLine
            On Error GoTo 0
            
            Select Case Left(txtLine, 1)
                Case "'", "!", "#", ";"     'Comment Characters...
                Case Else
                    If Len(Trim(txtLine)) = 0 Then
                    ElseIf UCase(txtLine) = "[DIRECTORYUPDATE]" Then
                        FoundKey = True
                    ElseIf FoundKey And UCase(Left(txtLine, 4)) = "TOP=" Then
                    '    Me.Top = Mid(txtLine, 5)
                    ElseIf FoundKey And UCase(Left(txtLine, 5)) = "LEFT=" Then
                    '    Me.Left = Mid(txtLine, 6)
                    ElseIf FoundKey And UCase(Left(txtLine, 7)) = "HEIGHT=" Then
                    '    Me.Height = Mid(txtLine, 8)
                    ElseIf FoundKey And UCase(Left(txtLine, 6)) = "WIDTH=" Then
                    '    Me.Width = Mid(txtLine, 7)
                    ElseIf FoundKey And UCase(Left(txtLine, 7)) = "SOURCE=" Then
                        txtSource.Text = Mid(txtLine, 8)
                    ElseIf FoundKey And UCase(Left(txtLine, 7)) = "TARGET=" Then
                        txtTarget.Text = Mid(txtLine, 8)
                    ElseIf FoundKey And UCase(Left(txtLine, 13)) = "DATEMODIFIED=" Then
                        txtDateModified.Text = Mid(txtLine, 14)
                    ElseIf FoundKey And UCase(Left(txtLine, 4)) = "LOG=" Then
                        strLogFile = Mid(txtLine, 5)
                    Else
                        MsgBox "Superfluous line in " & strFileName & ":" & vbCr & vbCr & """" & txtLine & """", vbExclamation, "Warning"
                    End If
            End Select
        Loop
        Close #FileUnit
        Me.MousePointer = vbDefault
        frmMain.Caption = AppName & " - " & strFileName
        fNeedSave = False
    End If
    Exit Sub
    
ReadError:
    On Error Resume Next
    MsgBox "Error reading " & strFileName & vbCr & "System Error: " & Err.Description & "(" & Err.Number & ")", vbCritical, "Application Error"
    Close #FileUnit
    Exit Sub
End Sub
Private Sub RePosition(strPath As String, strDirectory As String)
    Dim strDir As String
   
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strDir = Dir(strPath, vbDirectory)
    Do While strDir <> ""
        If strDir = strDirectory Then Exit Sub
        strDir = Dir
    Loop
End Sub
Private Sub SaveConfiguration()
    Dim FileUnit As Integer
    
    Me.MousePointer = vbHourglass
    FileUnit = FreeFile
    Open strFileName For Output As #FileUnit
    Print #FileUnit, "[DirectoryUpdate]"
    'Print #FileUnit, "Top=" & Me.Top
    'Print #FileUnit, "Left=" & Me.Left
    'Print #FileUnit, "Height=" & Me.Height
    'Print #FileUnit, "Width=" & Me.Width
    Print #FileUnit, "Source=" & txtSource.Text
    Print #FileUnit, "Target=" & txtTarget.Text
    Print #FileUnit, "DateModified=" & txtDateModified.Text
    Print #FileUnit, "Log=" & strLogFile
    Close #FileUnit
    Me.MousePointer = vbDefault
    fNeedSave = False
End Sub
Private Sub ToggleLock(LockSource As Boolean)
    If LockSource Then
        txtSource.Enabled = False
        txtTarget.Enabled = False
        txtDateModified.Enabled = False
        optOptions(0).Enabled = False
        optOptions(1).Enabled = False
        chkSubFolders.Enabled = False
        cmdSearch.Enabled = False
        
        cmdCopy.Enabled = True
        mnuClear.Enabled = True
        mnuSelectAll.Enabled = True
    Else
        txtSource.Enabled = True
        txtTarget.Enabled = True
        txtDateModified.Enabled = True
        optOptions(0).Enabled = True
        optOptions(1).Enabled = True
        chkSubFolders.Enabled = True
        cmdSearch.Enabled = True
        
        cmdCopy.Enabled = False
        mnuClear.Enabled = False
        mnuSelectAll.Enabled = False
    End If
End Sub
Private Sub cmdBrowseSource_Click()
    cmdSourceBrowseHit = True
    frmChooseFolder.txtPath.Text = txtSource.Text
    frmChooseFolder.Show vbModal
    If frmChooseFolder.txtPath.Text <> "" Then
        txtSource.Text = frmChooseFolder.txtPath.Text
        fNeedSave = True
    End If

    If Dir(txtSource.Text & "\global.asa", vbNormal) <> "" Then
        gfWebServer = True
    Else
        gfWebServer = False
    End If
    cmdSourceBrowseHit = False
End Sub
Private Sub cmdBrowseTarget_Click()
    cmdTargetBrowseHit = True
    frmChooseFolder.txtPath.Text = txtTarget.Text
    frmChooseFolder.Show vbModal
    If frmChooseFolder.txtPath.Text <> "" Then
        txtTarget.Text = frmChooseFolder.txtPath.Text
        fNeedSave = True
    End If
    cmdTargetBrowseHit = False
End Sub
Private Sub cmdCopy_Click()
    Dim Response As Integer
    Dim i As Integer
    Dim FileName As String
    Dim TempFileName As String
    Dim Seconds As Long
    Dim DoCopy As Boolean
    Dim strIcon As String
    Dim FilesCopied As Integer
    Dim FilesNotCopied As Integer
    Dim Errors As Integer
    Dim SourceFileAttr As VbFileAttribute
    Dim TargetFileAttr As VbFileAttribute
    Dim LogFileName As String
    
    LogFileName = txtTarget.Text & "\DirectoryUpdate.log"
    lblLogFile.Caption = "Log File: " & LogFileName
    LogFileUnit = FreeFile
    Open LogFileName For Append As #LogFileUnit
    Print #LogFileUnit, String(132, "-")
    Print #LogFileUnit, Now() & Chr(9) & "User: " & UserName & " (from " & ComputerName & " machine) starting " & AppName & "..."
    
    Errors = 0
    FilesCopied = 0
    FilesNotCopied = 0
    For i = 1 To lvwSource.ListItems.Count
        If lvwSource.ListItems(i).Selected Then
            sbBottom.Panels("Status").Text = "Copying " & lvwSource.ListItems(i).Text & " to " & txtTarget.Text
            DoCopy = True
            FileName = txtTarget.Text & Right(lvwSource.ListItems(i).Text, Len(lvwSource.ListItems(i)) - Len(txtSource.Text))
    
            If Dir(FileName) <> "" Then
                Seconds = DateDiff("s", lvwSource.ListItems(i).SubItems(lvwSubItem_DateModified), FileDateTime(FileName))
            
                If Seconds = 0 Then
                    Print #LogFileUnit, Now() & Chr(9) & "Files have same date/time stamp, """ & FileName & """ not copied."
                    Print #LogFileUnit2, Now() & Chr(9) & "Files have same date/time stamp, """ & FileName & """ not copied."
                    DoCopy = False
                End If
            
                If Seconds > 0 Then
                    ' Bring up the Confirm File Replace Dialog...
               
                    If frmConfirmFileReplace.gfYesToAll Then
                        DoCopy = True
                    Else
                        If frmConfirmFileReplace.gfNoToAll Or frmConfirmFileReplace.gfCancel Then
                            Print #LogFileUnit, Now() & Chr(9) & "Newer " & FileName & " not overwritten at user's request."
                            Print #LogFileUnit2, Now() & Chr(9) & "Newer " & FileName & " not overwritten at user's request."
                            DoCopy = False
                        Else
                            hImgLarge = SHGetFileInfo(FileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
                     
                            ' Draw the associated icons into the picture boxes
                            r& = ImageList_Draw(hImgLarge, shinfo.iIcon, frmConfirmFileReplace.picSource.hDC, 0, 0, ILD_TRANSPARENT)
                            r& = ImageList_Draw(hImgLarge, shinfo.iIcon, frmConfirmFileReplace.picTarget.hDC, 0, 0, ILD_TRANSPARENT)
                            
                            frmConfirmFileReplace.lblFileName = "This folder already contains a file called '" & FileName & "'."
                            frmConfirmFileReplace.lblSourceDate = "Modified on " & Format(lvwSource.ListItems(i).SubItems(lvwSubItem_DateModified), "dddd, mmmm d yyyy, hh:nn:ss AMPM")
                            frmConfirmFileReplace.lblSourceSize = lvwSource.ListItems(i).SubItems(lvwSubItem_Size)
                            frmConfirmFileReplace.lblTargetDate = "Modified on " & Format(FileDateTime(FileName), "dddd, mmmm d yyyy, hh:nn:ss AMPM")
                            frmConfirmFileReplace.lblTargetSize = Format(FileLen(FileName) \ 1024, "##,##0 KB")
                            frmConfirmFileReplace.Show vbModal
                            
                            If frmConfirmFileReplace.gfCancel Then
                                Print #LogFileUnit, Now() & Chr(9) & "Copy operation aborted (at user's request). Remaining files not copied to target directory."
                                Print #LogFileUnit2, Now() & Chr(9) & "Copy operation aborted (at user's request). Remaining files not copied to target directory."
                                MsgBox "Copy operation aborted. Remaining files not copied to target directory.", vbExclamation + vbOKOnly
                                DoCopy = False
                            Else
                                If frmConfirmFileReplace.gfYes Or frmConfirmFileReplace.gfYesToAll Then
                                    DoCopy = True
                                Else
                                    Print #LogFileUnit, Now() & Chr(9) & FileName & " not overwritten at user's request."
                                    Print #LogFileUnit2, Now() & Chr(9) & FileName & " not overwritten at user's request."
                                    DoCopy = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
         
            If (GetAttr(lvwSource.ListItems(i).Text) And vbDirectory) = vbDirectory Then
                Print #LogFileUnit, Now() & Chr(9) & "Directory " & FileName & " not copied (as a rule, we don't do directories)."
                Print #LogFileUnit2, Now() & Chr(9) & "Directory " & FileName & " not copied (as a rule, we don't do directories)."
                DoCopy = False
            End If
         
            If DoCopy Then
                SourceFileAttr = GetAttr(lvwSource.ListItems(i).Text)
                SetAttr lvwSource.ListItems(i).Text, vbNormal
                If Dir(FileName, vbNormal) <> "" Then
                    TargetFileAttr = GetAttr(FileName)
                    SetAttr FileName, vbNormal
                End If
            
                On Error Resume Next
                FileCopy lvwSource.ListItems(i).Text, FileName
                If Err.Number <> 0 Then
                    If Err.Number = 76 Then
                        Err.Clear
                
                        ' Make sure target directory path exists...
                        On Error GoTo 0
                        CreateDirPath lvwSource.ListItems(i).Text, FileName
                        On Error Resume Next
                        FileCopy lvwSource.ListItems(i).Text, FileName
                    End If
                End If
            
                If Err.Number <> 0 Then
                    Print #LogFileUnit, Now() & Chr(9) & "Copy failed on """ & FileName & """, Error #" & Err.Number & "; " & Err.Description
                    Print #LogFileUnit2, Now() & Chr(9) & "Copy failed on """ & FileName & """, Error #" & Err.Number & "; " & Err.Description
                    lvwSource.ListItems(i).Selected = False
                    Err.Clear
                    
                    FilesNotCopied = FilesNotCopied + 1
                    Errors = Errors + 1
                Else
                    Print #LogFileUnit, Now() & Chr(9) & "Copied " & lvwSource.ListItems(i).Text & " to " & txtTarget.Text
                    Print #LogFileUnit2, Now() & Chr(9) & "Copied " & lvwSource.ListItems(i).Text & " to " & txtTarget.Text
                    If LCase(Right(FileName, 3)) <> "cnt" Then SetAttr FileName, vbReadOnly
                End If
                On Error GoTo 0
                
                SetAttr lvwSource.ListItems(i).Text, SourceFileAttr
                SetAttr FileName, TargetFileAttr
                
                FilesCopied = FilesCopied + 1
            Else
                lvwSource.ListItems(i).Selected = False
                FilesNotCopied = FilesNotCopied + 1
            End If
        End If
        DoEvents
    Next i

    frmConfirmFileReplace.InitFlags
    sbBottom.Panels("Status").Text = Format(FilesCopied, "###,##0") & " file(s) copied."
    
    Print #LogFileUnit, Now() & Chr(9) & "Copy Complete; " & FilesCopied & " file(s) copied, " & Errors & " error(s) reported."
    Print #LogFileUnit2, Now() & Chr(9) & "Copy Complete; " & FilesCopied & " file(s) copied, " & Errors & " error(s) reported."
    Close #LogFileUnit
    
    If Errors > 0 Then
        Response = MsgBox(Errors & " error(s) reported. See " & LogFileName & " for details. View log file now...?", vbExclamation + vbYesNo)
        If Response = vbYes Then
            Debug.Print Shell("notepad.exe " & LogFileName, vbNormalFocus)
        End If
    End If
End Sub
Private Sub cmdSearch_Click()
    Dim intTotalFiles As Integer
    Dim TotalBytes As Long
   
    If Not fSearching Then
        If txtSource.Text = "" Then
            MsgBox "Source Directory must be specified.", vbExclamation + vbOKOnly
            txtSource.SetFocus
            Exit Sub
        End If
        If txtTarget.Text = "" Then
            MsgBox "Target Directory must be specified.", vbExclamation + vbOKOnly
            txtTarget.SetFocus
            Exit Sub
        End If
        If txtDateModified.Text = "" Then
            MsgBox "Date Modified must be specified.", vbExclamation + vbOKOnly
            txtDateModified.SetFocus
            Exit Sub
        End If
        
        On Error GoTo 0
        
        fSearching = True
        cmdSearch.Caption = "&Stop"
        SaveSetting AppName, "Preferences", "Last Run", Now()
        intTotalFiles = GetFiles(txtSource.Text, txtTarget.Text, TotalBytes)
        sbBottom.Panels("Status").Text = Format(intTotalFiles, "###,##0") & " file(s) found."
        fSearching = False
        
        If TotalBytes > GB Then
            sbBottom.Panels("Size").Text = Format(TotalBytes / GB, "#,##0.00 GB")
        Else
            If TotalBytes > MB Then
                sbBottom.Panels("Size").Text = Format(TotalBytes / MB, "#,##0.00 MB")
            Else
                If TotalBytes > KB Then
                    sbBottom.Panels("Size").Text = Format(TotalBytes / KB, "#,##0.00 KB")
                Else
                    If TotalBytes > 0 Then sbBottom.Panels("Size").Text = Format(TotalBytes, "#,##0 Bytes")
                End If
            End If
        End If
        
        If intTotalFiles > 0 Then
            ToggleLock True
            lvwSource.SetFocus
        Else
            ToggleLock False
        End If
        cmdSearch.Caption = "&Search"
    Else
        fSearching = False
    End If
End Sub
Private Sub Form_Load()
    Const MAX_COMPUTERNAME_LENGTH = 15
    Dim mNode As Node
    Dim strDir As String
    Dim strDisk As String
    Dim i As Integer
    Dim s$
    Dim dl&
    Dim sz&
    
    Load frmSplash
    frmSplash.Timer1.Enabled = True
    frmSplash.Timer1.Interval = 2500
    frmSplash.Show vbModal  'vbModeless
    
    s$ = String$(MAX_COMPUTERNAME_LENGTH + 1, Chr(0))
    sz& = MAX_COMPUTERNAME_LENGTH + 1
    dl& = GetComputerName(s$, sz)
    ComputerName = Left(s$, InStr(s$, Chr(0)) - 1)
    
    cnt& = 199
    s$ = String$(cnt&, Chr(0))
    dl& = GetUserName(s$, cnt)
    UserName = Left(s$, InStr(s$, Chr(0)) - 1)
    
    AppName = "Directory Update"
    
    ' Remember original field positions for use when resizing the window...
    OriginalFormWidth = Me.Width
    OriginalFormHeight = 5244  'Me.Height
    BoxesOffset = Me.Width - txtSource.Width
    BrowseOffset = Me.Width - cmdBrowseSource.Left
    ButtonOffset = Me.Width - cmdSearch.Left
    lblLogFileTop = Me.Height - lblLogFile.Top
    FramesOffset = Me.Width - frameSource.Width
    frameSelectedOffset = Me.Width - frameSelected.Width
    frameSelectedHeight = Me.Height - frameSelected.Height
    lvwSourceOffset = Me.Width - lvwSource.Width
    lvwSourceHeight = Me.Height - lvwSource.Height
    
    frmMain.Caption = AppName
    frmMain.Height = GetSetting(AppName, "Preferences", "Height", "7116")
    frmMain.Top = GetSetting(AppName, "Preferences", "Top", "84")
    frmMain.Width = GetSetting(AppName, "Preferences", "Width", "8532")
    frmMain.Left = GetSetting(AppName, "Preferences", "Left", "84")
    
    frmMain.txtDateModified.Text = GetSetting(AppName, "Preferences", "DateModified", Now())
    frmMain.txtDateModified.Text = GetSetting(AppName, "Preferences", "Last Run", Now())
    frmMain.txtSource.Text = GetSetting(AppName, "Preferences", "Source", "")
    frmMain.txtTarget.Text = GetSetting(AppName, "Preferences", "Target", "")
    strLogFile = GetSetting(AppName, "Preferences", "Log", "C:\" & AppName & ".log")
    strFileName = GetSetting(AppName, "Preferences", "FileName", "")
    If strFileName <> "" Then frmMain.Caption = AppName & " - " & strFileName
    If strFileName <> "Untitled" Then OpenConfiguration
    
    If Dir(txtSource.Text & "\global.asa", vbNormal) <> "" Then
        gfWebServer = True
    Else
        gfWebServer = False
    End If
   
    imlIcons.ListImages.Clear
    imlSmallIcons.ListImages.Clear
    
    lvwSource.ColumnHeaders.Clear
    lvwSource.ColumnHeaders.Add , "Path", "Path", lvwSource.Width \ 8
    lvwSource.ColumnHeaders.Add , "RawSize", "RawSize", 0
    lvwSource.ColumnHeaders.Add , "Size", "Size", 0, lvwColumnRight
    lvwSource.ColumnHeaders.Add , "Type", "Type", 0
    lvwSource.ColumnHeaders.Add , "RawDateModified", "RawModified", 0
    lvwSource.ColumnHeaders.Add , "Date Modified", "Date Modified", 0
    
    lvwSource.ListItems.Clear
    
    lblLogFile.Caption = "Log File: " & strLogFile
    LogFileUnit2 = FreeFile
    Open strLogFile For Append As #LogFileUnit2
    Print #LogFileUnit2, String(132, "-")
    Print #LogFileUnit2, Now() & Chr(9) & "User: " & UserName & " (from " & ComputerName & " machine) starting " & AppName & "..."
    
    ToggleLock False
    fQuickCopy = True
    Load frmConfirmFileReplace
    Load frmChooseFolder
    Load frmPreferences
    Unload frmSplash
End Sub
Private Sub Form_Resize()
    If Me.Width < OriginalFormWidth Or Me.Height < OriginalFormHeight Then Exit Sub
   
    'frameSelected.Width = Me.Width - (3 * frameSelected.Left)
    frameSelected.Width = Me.Width - frameSelectedOffset
    frameSelected.Height = Me.Height - frameSelectedHeight
    
    frameSource.Width = Me.Width - FramesOffset
    frameTarget.Width = Me.Width - FramesOffset
    frameOptions.Width = Me.Width - FramesOffset
    
    txtSource.Width = Me.Width - BoxesOffset
    txtTarget.Width = Me.Width - BoxesOffset
    
    'lvwSource.Width = frameSelected.Width - (2 * (lvwSource.Left - frameSelected.Left))
    lvwSource.Width = Me.Width - lvwSourceOffset
    lvwSource.Height = Me.Height - lvwSourceHeight
    If lvwSource.ColumnHeaders.Count > 0 Then
        'lvwSource.ColumnHeaders("Path").Width = (lvwSource.Width \ 2)
        'lvwSource.ColumnHeaders("RawSize").Width = 0
        'lvwSource.ColumnHeaders("Size").Width = (lvwSource.Width \ 8)
        'lvwSource.ColumnHeaders("RawDateModified").Width = 0
        'lvwSource.ColumnHeaders("Date Modified").Width = (lvwSource.Width \ 4)
    End If
   
    cmdSearch.Left = Me.Width - ButtonOffset
    cmdCopy.Left = Me.Width - ButtonOffset
    cmdBrowseSource.Left = Me.Width - BrowseOffset
    cmdBrowseTarget.Left = Me.Width - BrowseOffset
    
    lblLogFile.Top = Me.Height - lblLogFileTop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Close #LogFileUnit
    
    CheckSaveState
    SaveSetting AppName, "Preferences", "Top", Me.Top
    SaveSetting AppName, "Preferences", "Left", Me.Left
    SaveSetting AppName, "Preferences", "Height", Me.Height
    SaveSetting AppName, "Preferences", "Width", Me.Width
    'SaveSetting AppName, "Preferences", "Source", txtSource.Text
    'SaveSetting AppName, "Preferences", "Target", txtTarget.Text
    'SaveSetting AppName, "Preferences", "DateModified", txtDateModified.Text
    'SaveSetting AppName, "Preferences", "Log", strLogFile
    SaveSetting AppName, "Preferences", "FileName", strFileName
    
    End
End Sub
Private Sub lvwSource_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Static LastColumnClicked As Integer
    Static LastSortOrder As Integer
    Dim ColumnClicked As Integer
    
    ColumnClicked = ColumnHeader.Index - 1
    Select Case ColumnClicked
        Case lvwSubItem_Size
            lvwSource.SortKey = lvwSubItem_RawSize
        Case lvwSubItem_DateModified
            lvwSource.SortKey = lvwSubItem_RawDateModified
        Case Else
            lvwSource.SortKey = ColumnClicked
    End Select
   
    If LastColumnClicked = ColumnClicked Then
        If LastSortOrder <> lvwDescending Then
            lvwSource.SortOrder = lvwDescending
        Else
            lvwSource.SortOrder = lvwAscending
        End If
    Else
        lvwSource.SortOrder = lvwAscending
    End If
    LastColumnClicked = ColumnClicked
    LastSortOrder = lvwSource.SortOrder
    lvwSource.Sorted = True
End Sub
Private Sub mnuClear_Click()
    ToggleLock False
    lvwSource.ListItems.Clear
    sbBottom.Panels("Status").Text = ""
    sbBottom.Panels("Size").Text = ""
    lblLogFile.Caption = "Log File: " & strLogFile
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileNew_Click()
    CheckSaveState

    mnuClear_Click
    txtSource.Text = ""
    txtTarget.Text = ""
    txtDateModified.Text = ""
    strFileName = "Untitled"
    frmMain.Caption = AppName & " - " & strFileName
    fNeedSave = False
End Sub
Private Sub mnuFileOpen_Click()
    Dim FileUnit As Integer
    Dim txtLine As String
    Dim FoundKey As Boolean
    
    CheckSaveState
    
    On Error GoTo ErrorHandler
    With cdgMain
        .CancelError = True
        .DefaultExt = ".duc"
        .Filter = "All Files (*.*)|*.*|DirectoryUpdate Config Files (*.duc)|*.duc"
        .FilterIndex = 2
        .ShowOpen
        strFileName = .FileName
    End With
    OpenConfiguration
    Exit Sub

ErrorHandler:
    If Err.Number = 32755 Then Exit Sub
        
    MsgBox "Error encountered while attempting to open file." & vbCr & vbCr & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub
Private Sub mnuFileSave_Click()
    Dim FileUnit As Integer
    
    If strFileName = "" Or strFileName = "Untitled" Then
        mnuFileSaveAs_Click
    Else
        SaveConfiguration
    End If
End Sub
Private Sub mnuFileSaveAs_Click()
    On Error GoTo ErrorHandler
    With cdgMain
        .CancelError = True
        .DefaultExt = ".duc"
        .Filter = "All Files (*.*)|*.*|DirectoryUpdate Config Files (*.duc)|*.duc"
        .FilterIndex = 2
        .ShowSave
        strFileName = .FileName
    End With
    frmMain.Caption = AppName & " - " & strFileName
    SaveConfiguration
    Exit Sub
    
ErrorHandler:
    If Err.Number = 32755 Then Exit Sub
        
    MsgBox "Error encountered while attempting to save file." & vbCr & vbCr & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub
Private Sub mnuHelpAbout_Click()
    mnuHelpAboutHit = True
    Load frmSplash
    frmSplash.Timer1.Enabled = False
    frmSplash.Show vbModal
    mnuHelpAboutHit = False
End Sub
Private Sub mnuPreferences_Click()
    frmPreferences.txtSource.Text = txtSource.Text
    frmPreferences.txtTarget.Text = txtTarget.Text
    frmPreferences.txtDateModified.Text = txtDateModified.Text
    frmPreferences.txtLogFile.Text = strLogFile
    frmPreferences.Show vbModal
    txtSource.Text = frmPreferences.txtSource.Text
    txtTarget.Text = frmPreferences.txtTarget.Text
    txtDateModified.Text = frmPreferences.txtDateModified.Text
    
    If strLogFile <> frmPreferences.txtLogFile.Text Then
        Print #LogFileUnit2, Now() & Chr(9) & "Switching to new log file: " & frmPreferences.txtLogFile.Text
        Close #LogFileUnit2
        strLogFile = frmPreferences.txtLogFile.Text
        lblLogFile.Caption = "Log File: " & strLogFile
        LogFileUnit2 = FreeFile
        Open strLogFile For Append As #LogFileUnit2
        Print #LogFileUnit2, String(132, "-")
    End If
End Sub
Private Sub mnuSelectAll_Click()
    Dim i As Integer
   
    For i = 1 To lvwSource.ListItems.Count
        lvwSource.ListItems(i).Selected = True
    Next i
End Sub
Private Sub optOptions_Click(Index As Integer)
    ToggleLock False
    Select Case Index
        Case 0
            fQuickCopy = True
        Case 1
            'ToggleLock True
            fQuickCopy = False
    End Select
End Sub
Private Sub txtDateModified_Change()
    fNeedSave = True
End Sub
Private Sub txtDateModified_LostFocus()
    txtDateModified.Text = Format(txtDateModified.Text, "ddddd ttttt")
End Sub
Private Sub txtSource_Change()
    fNeedSave = True
End Sub
Private Sub txtSource_LostFocus()
    If mnuHelpAboutHit Then Exit Sub
    If cmdTargetBrowseHit Then Exit Sub
    
    If txtSource.Text = "" Then
        cmdBrowseSource_Click
        Exit Sub
    End If
   
    If Right(txtSource.Text, 1) = "\" Then txtSource.Text = Left(txtSource.Text, Len(txtSource.Text) - 1)
    On Error Resume Next
    If Dir(txtSource.Text, vbDirectory) <> "" Then
        If (GetAttr(txtSource.Text) And vbDirectory) <> vbDirectory Then
            MsgBox "Path specified does not represent a directory.", vbExclamation + vbOKOnly
            txtSource.SetFocus
        End If
      
        If Dir(txtSource.Text & "\global.asa", vbNormal) <> "" Then
            gfWebServer = True
        Else
            gfWebServer = False
        End If
    Else
        MsgBox "Path specified does not exist.", vbExclamation + vbOKOnly
        txtSource.SetFocus
    End If
    On Error GoTo 0
End Sub
Private Sub txtTarget_Change()
    fNeedSave = True
End Sub
Private Sub txtTarget_LostFocus()
    If mnuHelpAboutHit Then Exit Sub
    If cmdSourceBrowseHit Then Exit Sub
    
    If txtTarget.Text = "" Then
        cmdBrowseTarget_Click
        Exit Sub
    End If
   
    If Right(txtTarget.Text, 1) = "\" Then txtTarget.Text = Left(txtTarget.Text, Len(txtTarget.Text) - 1)
    On Error Resume Next
    If Dir(txtTarget.Text, vbDirectory) <> "" Then
        If (GetAttr(txtTarget.Text) And vbDirectory) <> vbDirectory Then
            MsgBox "Path specified does not represent a directory.", vbExclamation + vbOKOnly
            txtTarget.SetFocus
        End If
    Else
        MsgBox "Path specified does not exist.", vbExclamation + vbOKOnly
        txtTarget.SetFocus
    End If
    On Error GoTo 0
End Sub
Private Sub sbBody_PanelClick(ByVal Panel As ComctlLib.Panel)
    MsgBox "Don't click here..."
End Sub


Attribute VB_Name = "IconHandler"
Option Explicit

Public hImgSmall As Long      ' The handle to the system image list
Public hImgLarge As Long
Dim FileName As String     ' The file name to get icon fro
Public r As Long
   
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Public Const SHGFI_LARGEICON = &H0        ' Large icon
Public Const SHGFI_SMALLICON = &H1        ' Small icon
Public Const SHGFI_SELECTED = &H10000
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const ILD_TRANSPARENT = &H1        ' Display transparent

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Public shinfo As SHFILEINFO

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long
Function genRetrieveIcon(ByVal pFileName As String, ByRef pboxIcon As PictureBox, ByRef pboxSmallIcon As PictureBox, ByRef imlIcons As ImageList, ByRef imlSmallIcons As ImageList) As Integer
   Dim strIcon As String
   Dim intRetrieveIcon As Integer
   Dim TempVal As Integer
   
    ' Get the system icons associated with the file
    hImgLarge = SHGetFileInfo(pFileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    hImgSmall = SHGetFileInfo(pFileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON Or SHGFI_SELECTED)
    strIcon = "A" & Trim(Str(shinfo.iIcon))

    TempVal = genRetrieveIconNumber(strIcon, imlIcons)
    If TempVal = 0 Then
        pboxIcon.Picture = LoadPicture()
        pboxSmallIcon.Picture = LoadPicture()
    
        ' Draw the associated icons into the picture boxes
        r& = ImageList_Draw(hImgLarge, shinfo.iIcon, pboxIcon.hDC, 0, 0, ILD_TRANSPARENT)
        r& = ImageList_Draw(hImgSmall, shinfo.iIcon, pboxSmallIcon.hDC, 0, 0, ILD_TRANSPARENT)
        imlIcons.ListImages.Add , strIcon, pboxIcon.Image
        imlSmallIcons.ListImages.Add , strIcon, pboxSmallIcon.Image
    End If
    If TempVal = 0 Then
        genRetrieveIcon = imlIcons.ListImages.Count
    Else
        genRetrieveIcon = TempVal
    End If
End Function
Function genstrRetrieveIcon(ByVal pFileName As String, ByRef pboxIcon As PictureBox, ByRef pboxSmallIcon As PictureBox, ByRef imlIcons As ImageList, ByRef imlSmallIcons As ImageList) As String
   genstrRetrieveIcon = imlIcons.ListImages(genRetrieveIcon(pFileName, pboxIcon, pboxSmallIcon, imlIcons, imlSmallIcons)).Key
End Function

Function genRetrieveIconNumber(pTypeName As String, ByRef imlIcons As ImageList) As Integer
   Dim retIcon As Integer
   For retIcon = 1 To imlIcons.ListImages.Count
      'Debug.Print "imlIcons.ListImages(" & retIcon & "): " & frmMain.imlIcons.ListImages(retIcon).Key
      If imlIcons.ListImages(retIcon).Key = pTypeName Then genRetrieveIconNumber = retIcon: Exit Function
   Next retIcon
   genRetrieveIconNumber = 0
End Function

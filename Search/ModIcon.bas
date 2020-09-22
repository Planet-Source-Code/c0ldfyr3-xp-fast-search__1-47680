Attribute VB_Name = "ModIcon"
Option Explicit
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal X&, ByVal y&, ByVal flags&) As Long
Private Const SHGFI_SYSICONINDEX As Long = &H4000
Private Const SHGFI_SHELLICONSIZE As Long = &H4
Private Const SHGFI_DISPLAYNAME As Long = &H200
Private Const SHGFI_EXETYPE     As Long = &H2000
Private Const SHGFI_LARGEICON   As Long = &H0
Private Const SHGFI_SMALLICON   As Long = &H1
Private Const SHGFI_TYPENAME    As Long = &H400
Private Const ILD_TRANSPARENT   As Long = &H1
Private Const BASIC_SHGFI_FLAGS As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Type PicBmp
    Size                        As Long
    tType                       As Long
    hBmp                        As Long
    hPal                        As Long
    Reserved                    As Long
End Type
Private Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(7)                    As Byte
End Type
Private Type SHFILEINFO
    hIcon                       As Long
    iIcon                       As Long
    dwAttributes                As Long
    szDisplayName               As String * MAX_PATH
    szTypeName                  As String * 80
End Type
Private ShInfo                  As SHFILEINFO
Private Type tImages
    PicLargeIcon                As Picture
    PicSmallIcon                As Picture
End Type
Public Function GetIcon(FileName As String, Index As Long, sShell32 As String) As Long
    Dim hLIcon                  As Long
    Dim hSIcon                  As Long
    Dim imgObj                  As ListImage
    Dim r                       As Long
    Dim lIcons                  As tImages
    hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    'Get the file info
    If Not hLIcon = 0 Then
        'Locked onto an icon...
        With FrmMain.pic32
            Set .Picture = LoadPicture("") 'Clear the existing pic.
            .AutoRedraw = True
            r = ImageList_Draw(hLIcon, ShInfo.iIcon, FrmMain.pic32.hDc, 0, 0, ILD_TRANSPARENT)
            'Draw the image into the Pic Box
            .Refresh
        End With
        With FrmMain.pic16
            Set .Picture = LoadPicture("") 'Clear the existing pic.
            .AutoRedraw = True
            r = ImageList_Draw(hSIcon, ShInfo.iIcon, FrmMain.pic16.hDc, 0, 0, ILD_TRANSPARENT)
            'Draw the image into the Pic Box
            .Refresh
        End With
        'Add it to the image list
        Set imgObj = FrmMain.iml32.ListImages.Add(Index, , FrmMain.pic32.Image)
        Set imgObj = FrmMain.iml16.ListImages.Add(Index, , FrmMain.pic16.Image)
    Else
        lIcons = GetIconFromFile(sShell32, 0, True)
        With FrmMain.pic32
            .AutoRedraw = True
            Set .Picture = lIcons.PicLargeIcon
            .Refresh
        End With
        With FrmMain.pic16
            .AutoRedraw = True
            Set .Picture = lIcons.PicSmallIcon  'Get the default "No Icon" icon from shell32.dll in %WinDir%\System32
            .Refresh
        End With
        'Add it to the image list.
        Set imgObj = FrmMain.iml32.ListImages.Add(Index, , FrmMain.pic32.Image)
        Set imgObj = FrmMain.iml16.ListImages.Add(Index, , FrmMain.pic16.Image)
    End If
End Function
Public Function GetIconFromFile(FileName As String, IconIndex As Long, UseLargeIcon As Boolean) As tImages
    Dim hLargeIcon                  As Long
    Dim hSmallIcon                  As Long
    Dim Pic(1)                      As PicBmp
    Dim IPic(1)                     As IPicture
    Dim IID_IDispatch(1)            As GUID
    If ExtractIconEx(FileName, IconIndex, hLargeIcon, hSmallIcon, 1) > 0 Then
        With IID_IDispatch(0)
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        With IID_IDispatch(1)
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        With Pic(0)
            .Size = Len(Pic(0))
            .tType = vbPicTypeIcon
            .hBmp = hLargeIcon
        End With
        With Pic(1)
            .Size = Len(Pic(1))
            .tType = vbPicTypeIcon
            .hBmp = hSmallIcon
        End With
        Call OleCreatePictureIndirect(Pic(0), IID_IDispatch(0), 1, IPic(0))
        Call OleCreatePictureIndirect(Pic(1), IID_IDispatch(1), 1, IPic(1))
        Set GetIconFromFile.PicLargeIcon = IPic(0)
        Set GetIconFromFile.PicSmallIcon = IPic(1)
        DestroyIcon hSmallIcon
        DestroyIcon hLargeIcon
    End If
End Function

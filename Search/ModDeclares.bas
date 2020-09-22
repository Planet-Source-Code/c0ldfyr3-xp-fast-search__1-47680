Attribute VB_Name = "ModDeclares"
Option Explicit
Public VB6 As Boolean

'Alot of this wasn't written originally by me, but I'll comment parts of it.
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal I As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function PathGetDriveNumber Lib "SHLWAPI.DLL" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long
Private Declare Function PathStripToRoot Lib "SHLWAPI.DLL" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Public Declare Function PathIsNetworkPath Lib "SHLWAPI.DLL" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Boolean
Private Declare Function PathIsUNCServerShare Lib "SHLWAPI.DLL" Alias "PathIsUNCServerShareA" (ByVal pszPath As String) As Boolean
Public Declare Function PathIsUNCServer Lib "SHLWAPI.DLL" Alias "PathIsUNCServerA" (ByVal pszPath As String) As Boolean
Public Declare Function PathIsUNC Lib "SHLWAPI.DLL" Alias "PathIsUNCA" (ByVal pszPath As String) As Boolean
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As enumFileAttributes
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDc As Long, qrc As RECT, ByVal edge As enumBorderEdges, ByVal grfFlags As enumBorderFlags) As Long
Private Declare Function OpenFileHandle Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const MAX_PATH           As Integer = 400
Private Const SH_TYPENAME       As Long = &H400
Private Const SH_DISPLAYNAME    As Long = &H200
Private Const SH_EXETYPE        As Long = &H2000
Private Const SH_SYSICONINDEX   As Long = &H4000
Private Const SH_LARGEICON      As Long = &H0
Private Const SH_SMALLICON      As Long = &H1
Private Const SH_SHELLICONSIZE  As Long = &H4
Private Const ILD_TRANSPARENT   As Long = &H1
Private Const OFS_MAXPATHNAME   As Long = 128
Private Const GOOD_RETURN_CODE  As Long = 0
Private Const MAX_PATH_LENGTH   As Long = 260
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2
Private Const flags             As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const BASIC_SH_FLAGS    As Long = SH_TYPENAME Or SH_SHELLICONSIZE Or SH_SYSICONINDEX Or SH_DISPLAYNAME Or SH_EXETYPE
Private Const SH_USEFILEATTRIBUTES As Long = &H10
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Enum HKEYS
    vHKEY_CLASSES_ROOT = &H80000000
    vHKEY_CURRENT_USER = &H80000001
    vHKEY_LOCAL_MACHINE = &H80000002
    vHKEY_USERS = &H80000003
    vHKEY_PERFORMcANCE_DATA = &H80000004
    vHKEY_CURRENT_CONFIG = &H80000005
    vHKEY_DYN_DATA = &H80000006
End Enum
Private Type SHFILEINFO
    hIcon                       As Long
    iIcon                       As Long
    dwAttributes                As Long
    szDisplayName               As String * 260
    szTypeName                  As String * 80
End Type
Private Type sArrType
    Name                        As String
    Picture                     As IPictureDisp
End Type
Public Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long
    ftCreationTime              As FILETIME
    ftLastAccessTime            As FILETIME
    ftLastWriteTime             As FILETIME
    nFileSizeHigh               As Long
    nFileSizeLow                As Long
    dwReserved0                 As Long
    dwReserved1                 As Long
    cFileName                   As String * 260
    cShortFileName              As String * 14
End Type
Public Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type
Public Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type
Public Type OFSTRUCT
    cBytes                      As Byte
    fFixedDisk                  As Byte
    nErrCode                    As Integer
    Reserved1                   As Integer
    Reserved2                   As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes            As Long
    ftCreationTime              As FILETIME
    ftLastAccessTime            As FILETIME
    ftLastWriteTime             As FILETIME
    dwVolumeSerialNumber        As Long
    nFileSizeHigh               As Long
    nFileSizeLow                As Long
    nNumberOfLinks              As Long
    nFileIndexHigh              As Long
    nFileIndexLow               As Long
End Type
Public Type typeVolumeInformation
    sRootPathName               As String
    sVolumeName                 As String
    lVolumeSerialNo             As Long
    lMaximumComponentLength     As Long
    lFileSystemFlags            As Long
    sFileSystemName             As String
End Type
Public Type tCompareDate
    tType                       As eCompareDate
    tTime                       As Integer
    tCompare                    As Boolean
    tWhich                      As eCompareType
End Type
Public Enum enumFileAttributes
    efaARCHIVE = &H20
    efaCOMPRESSED = &H800
    efaDIRECTORY = &H10
    efaHIDDEN = &H2
    efaNORMAL = &H80
    efaREADONLY = &H1
    efaSYSTEM = &H4
    efaTEMPORARY = &H100
End Enum
Public Enum enumDriveTypes
    DRIVE_CDROM = 5
    DRIVE_FIXED = 3
    DRIVE_RAMDISK = 6
    DRIVE_REMOTE = 4
    DRIVE_REMOVABLE = 2
End Enum
Public Enum enumBorderFlags
    BF_ADJUST = &H2000
    BF_BOTTOM = &H8
    BF_DIAGONAL = &H10
    BF_FLAT = &H4000
    BF_LEFT = &H1
    BF_MIDDLE = &H800
    BF_MONO = &H8000
    BF_RIGHT = &H4
    BF_SOFT = &H1000
    BF_TOP = &H2
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
End Enum
Public Enum enumBorderEdges
    BDR_RAISEDINNER = &H4
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENINNER = &H8
    BDR_SUNKENOUTER = &H2
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum
Public Enum enumFileNameParts
    efpFileName = 2 ^ 0
    efpFileExt = 2 ^ 1
    efpFilePath = 2 ^ 2
    efpFileNameAndExt = efpFileName + efpFileExt
    efpFileNameAndPath = efpFilePath + efpFileName
End Enum
Public Enum eCompareDate
    eDay = 0
    eMonth = 1
End Enum
Public Enum eCompareType
    eCreated = 0
    eModified = 1
End Enum
Public Sub Main()
    On Error GoTo ErrClear
    VB6 = True
    Dim sTmp                    As String
    Dim sArr()                  As String
    Dim X                       As Integer
    Dim bShowSplash             As Boolean
    Load FrmMain
    bShowSplash = True
    If Len(Command) > 0 Then
        sArr = Split(Replace(Command, """", ""), " /")
        If J_UBound(sArr) = 0 Then
            frmSplash.Show
            FrmMain.txtFolder.Text = Command
        Else
            For X = 0 To J_UBound(sArr)
                If Len(Dir(FixPath(sArr(X)))) > 0 Then
                    FrmMain.txtFolder.Text = sArr(X)
                End If
                Select Case LCase(sArr(X))
                    Case Is = "noload"
                        bShowSplash = False
                End Select
            Next
        End If
    End If
    If bShowSplash = True Then frmSplash.Show
    'Load settings saved on previous shut down.
    sTmp = GetSetting("Search Sub Folders")
    FrmMain.chkSubFolders.Value = IIf(sTmp = "1", vbChecked, vbUnchecked)
    sTmp = GetSetting("Search Hidden Folders")
    FrmMain.chkHidden.Value = IIf(sTmp = "1", vbChecked, vbUnchecked)
    sTmp = GetSetting("Search System Folders")
    FrmMain.chkSystem.Value = IIf(sTmp = "1", vbChecked, vbUnchecked)
    sTmp = GetSetting("Search Zip Folders")
    FrmMain.chkZips.Value = IIf(sTmp = "1", vbChecked, vbUnchecked)
    sTmp = GetSetting("Last Location")
    If Len(sTmp) = 0 Then sTmp = sPathRoot(SpecialFolder(Windows))
    If Len(Command) = 0 Then FrmMain.txtFolder.Text = sTmp
    sTmp = GetSetting("View Type")
    If Len(sTmp) = 0 Then sTmp = "3"
    FrmMain.lstFind.View = IIf(IsNumeric(sTmp), CInt(sTmp), 3)
    If bShowSplash = False Then FrmMain.Show
    Call EnumTypes(FrmMain.ImageCombo, FrmMain.imgCombo)
    Exit Sub
ErrClear:
    MsgBox "#" & Err.Number & " : " & Err.Description & " : " & Err.Source
    End
End Sub
Private Function GetSetting(Setting As String, Optional Default As String = "") As String
    'Time saving function.
    GetSetting = REGGetSetting(vHKEY_CURRENT_USER, "Software\EliteProdigy\Search\Settings", Setting, Default)
End Function
Public Function sFilename(ByVal sFile As String, ByVal ePortions As enumFileNameParts) As String
    'This gets parts of the file name, depending on what you asked for.
    Dim lFirstPeriod            As Long
    Dim lFirstBackSlash         As Long
    Dim sPath                   As String
    Dim sname                   As String
    Dim sExt                    As String
    Dim sRet                    As String
    lFirstPeriod = InStrRev(sFile, ".")
    lFirstBackSlash = InStrRev(sFile, "\")
    If lFirstBackSlash > 0 Then
        sPath = Left(sFile, lFirstBackSlash)
    End If
    If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
        sExt = Mid(sFile, lFirstPeriod + 1)
        sname = Mid(sFile, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
    Else
        sname = Mid(sFile, lFirstBackSlash + 1)
    End If
    If ePortions And efpFilePath Then
        sRet = sRet & sPath
    End If
    If ePortions And efpFileName Then
        sRet = sRet & sname
    End If
    If ePortions And efpFileExt Then
        If sRet <> "" Then
            sRet = sRet & "." & sExt
        Else
            sRet = sRet & sExt
        End If
    End If
    sFilename = sRet
End Function
Public Sub MakeOntop(FormName As Form)
    On Error GoTo Error
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags) 'Sets the Form on top of all other windows.
    Exit Sub
Error:
    Err.Clear
End Sub
Public Function FileInformation(ByVal sFilename As String) As BY_HANDLE_FILE_INFORMATION
    'Get all file Info.
    Dim tInfo                   As BY_HANDLE_FILE_INFORMATION
    Dim tOF                     As OFSTRUCT
    Dim lHandle                 As Long
    lHandle = OpenFileHandle(sFilename, tOF, 0)
    If lHandle > 0 Then GetFileInformationByHandle lHandle, tInfo
    FileInformation = tInfo
    CloseHandle lHandle
End Function
Public Function timeSysToDate(st As SYSTEMTIME) As Date
    'Converts System Time to an actual readable date.
    timeSysToDate = CDate(Format(st.wMonth, "00") & "/" & Format(st.wDay, "00") & "/" & Format(st.wYear, "0000") & " " & Format(st.wHour, "00") & ":" & Format(st.wMinute, "00") & ":" & Format(st.wSecond, "00"))
End Function
Public Function timeFileToDate(ft As FILETIME) As Date
    'Converts file time to a legible date.
    Dim tSysTime                As SYSTEMTIME
    FileTimeToSystemTime ft, tSysTime
    timeFileToDate = timeSysToDate(tSysTime)
End Function
Public Function fileLength(ByVal sFilename As String) As Long
    'File Size
    Dim FileHandle              As Integer
    FileHandle = FreeFile
    On Error Resume Next
    Open sFilename For Input As #FileHandle
    fileLength = LOF(FileHandle)
    Close #FileHandle
End Function
Public Function lSize(sFile As String) As Long
    'Public size calling routine.
    On Error Resume Next
    lSize = FileLen(sFile)
    If Err.Number <> 0 Then 'If we couldnt get it using our own function, use an api call.
        lSize = fileLength(sFile)
    End If
End Function
Public Function sNT(ByVal sString As String) As String
    Dim iNullLoc                As Integer
    iNullLoc = InStr(sString, Chr(0))
    If iNullLoc > 0 Then
        sNT = Left(sString, iNullLoc - 1)
    Else
        sNT = sString
    End If
End Function
Public Function FileShortName(ByVal sFilename As String) As String
    Dim sBuffer                 As String
    'Gets the Windows Short File Name
    'Ex C:\Progra~1\Counte~1\cstrike.exe
    sBuffer = Space(1024)
    GetShortPathName sFilename, sBuffer, Len(sBuffer)
    FileShortName = sNT(sBuffer)
End Function
Public Function FileRoot(ByVal sFilename As String) As String
    Dim lngResult               As Long
    'Get the Root of a file path.
    'Ex: C:\Program Files\Program.exe will return C:\
    lngResult = PathStripToRoot(sFilename)
    If lngResult <> 0 Then
        If InStr(sFilename, vbNullChar) > 0 Then
            FileRoot = Left$(sFilename, InStr(sFilename, vbNullChar) - 1)
        Else
            FileRoot = sFilename
        End If
    End If
End Function
Public Function sPathRoot(sFile As String) As String
    Dim sRet                    As String
    sRet = FileRoot(sFile)
    If Right(sRet, 1) <> "\" And Trim(sRet) <> "" Then
        sRet = sRet & "\"
    End If
    sPathRoot = sRet
End Function
Public Function eAttributes(sFilename) As enumFileAttributes
    eAttributes = GetFileAttributes(sFilename)
End Function
Public Function volumeInformation(ByVal sDrive As String) As typeVolumeInformation
    Dim Ret                     As typeVolumeInformation
    Ret.sRootPathName = sDrive
    Ret.sFileSystemName = Space(1024)
    Ret.sVolumeName = Space(1024)
    GetVolumeInformation Ret.sRootPathName, Ret.sVolumeName, Len(Ret.sVolumeName), Ret.lVolumeSerialNo, Ret.lMaximumComponentLength, Ret.lFileSystemFlags, Ret.sFileSystemName, Len(Ret.sFileSystemName)
    Ret.sFileSystemName = sNT(Ret.sFileSystemName)
    Ret.sVolumeName = sNT(Ret.sVolumeName)
    volumeInformation = Ret
End Function
Public Function fileOpenStructure(ByVal sFilename As String) As OFSTRUCT
    Dim tOF                     As OFSTRUCT
    Dim lHandle                 As Long
    lHandle = OpenFileHandle(sFilename, tOF, 0)
    CloseHandle lHandle
    fileOpenStructure = tOF
End Function
Public Sub EnumTypes(imgCombo As ImageCombo, imgList As ImageList)
    'This is a little routine I wrote to get all the file types which are installed on a system.
    Screen.MousePointer = 11 'Hour Glass
    FrmMain.StatusBar1.Panels(1).Text = "Loading File Formats"
    FrmMain.StatusBar1.Panels(2).Text = "Status: Working ."
    FrmMain.tmrWorking.Enabled = True
    DoEvents
    Call LockControl(imgCombo, False) 'Stop the Image Combo from refreshing on every new item, saves alot of time and processing power.
    Dim sFileTypeName           As String
    Dim sFileExtension          As String
    Dim lIcon                   As Long
    Dim lIco2                   As Long
    Dim ShInfo                  As SHFILEINFO
    Dim sTmp()                  As String
    Dim sArr()                  As String
    Dim lRegKeyIndex            As Long
    Dim sRegSubkey              As String * MAX_PATH_LENGTH
    Dim sRegKeyClass            As String * MAX_PATH_LENGTH
    Dim FTime                   As FILETIME
    Dim bLoaded                 As Boolean
    Dim aCnt                    As Long
    Dim cTmp                    As ComboItem
    aCnt = -1
    Do While RegEnumKeyEx(vHKEY_CLASSES_ROOT, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, FTime) = GOOD_RETURN_CODE 'Enumerate ALL keys in that folder.
        If Asc(sRegSubkey) = 46 Then 'If it starts with a dot.
            lIco2 = SHGetFileInfo(sRegSubkey, FILE_ATTRIBUTE_NORMAL, ShInfo, Len(ShInfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_LARGEICON)
            lIcon = SHGetFileInfo(sRegSubkey, FILE_ATTRIBUTE_NORMAL, ShInfo, Len(ShInfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
                'Get the icon.
            sFileTypeName = TrimNull(ShInfo.szTypeName)
                'Type Name
            sFileExtension = TrimNull(sRegSubkey)
                'Extension associated with it.
            sFileExtension = Right(sFileExtension, Len(sFileExtension) - 1)
            If InArray(sArr, FirstUcase(sFileTypeName), aCnt) = False Then
                DoEvents
                FrmMain.pSmall.Picture = LoadPicture()
                Call ImageList_Draw(lIcon, ShInfo.iIcon, FrmMain.pSmall.hDc, 0, 0, ILD_TRANSPARENT)
                FrmMain.pSmall.Picture = FrmMain.pSmall.Image
                imgList.ListImages.Add , "#" & sFileExtension & "#", FrmMain.pSmall.Picture
                'Add the image to the image list for the combo to use.
                'Check if it exists allready, because multiple extensions for the same type.
                'Ex Log and Txt are both Text Document.
                'So we only want to add each one once.
                aCnt = aCnt + 1
                ReDim Preserve sArr(aCnt)
                sArr(aCnt) = FirstUcase(sFileTypeName) & "|^|" & sFileExtension
                'I could easily have added them here and saved on a little processing, but they would be in a stupid order.
            End If
        End If
        lRegKeyIndex = lRegKeyIndex + 1
    Loop
    Call TriQuickSortString(sArr, SortAscending) 'Sort the array.
    imgCombo.ImageList = imgList
    For aCnt = 0 To aCnt
        'Loop through and add each one.
        sTmp = Split(sArr(aCnt), "|^|")
        Set cTmp = imgCombo.ComboItems.Add(, , sTmp(0), imgList.ListImages("#" & sTmp(1) & "#").Key, imgList.ListImages("#" & sTmp(1) & "#").Key, 0)
        cTmp.Tag = sTmp(1)
    Next
    'Stop
    Set imgCombo.SelectedItem = imgCombo.ComboItems(1)
    Screen.MousePointer = 0
    FrmMain.tmrWorking.Enabled = False
    FrmMain.StatusBar1.Panels(1).Text = "Idle"
    FrmMain.StatusBar1.Panels(2).Text = ""
    Call LockControl(imgCombo, True)
End Sub
Private Function InArray(sArr() As String, Text As String, U_Bound As Long) As Boolean
    Dim X                       As Long
    Dim tArr()                  As String
    'Check for the existance of a string in an array.
    For X = 0 To U_Bound
        tArr = Split(sArr(X), "|^|")
        If StrComp(tArr(0), Text, vbTextCompare) = 0 Then
            InArray = True
            Exit Function
        End If
    Next
End Function
Private Function FirstUcase(Text As String) As String
    Dim sTmp                    As String * 1
    Dim sArr()                  As String
    Dim X                       As Integer
    'This takes a string lets say, "It is a raInY daY"
    'And returns "It Is A Rainy Day"
    'In home buying, its location, location, location.
    'In programming, its presentation, presentation, presentation !!! =)
    If Len(Text) = 0 Then Exit Function
    sArr = Split(Text, " ")
    For X = 0 To UBound(sArr)
        If Len(sArr(X)) < 2 Then 'If the word is only one char long, return that upper cased.
            sArr(X) = UCase(sArr(X))
        Else
            sTmp = sArr(X)
            'First letter upper case, rest lower.
            sArr(X) = UCase(sTmp) & LCase(Mid(sArr(X), 2))
        End If
    Next
    FirstUcase = Join(sArr, " ")
End Function
Public Sub LockControl(Cntrl As Control, Optional UnLockIt As Boolean = False)
    On Error Resume Next
    'This stops a control from updating while you work with it.
    'Saves ALOT of time when adding lists etc.
    'When I say alot, I really mean alot lol.
    Select Case UnLockIt
        Case True
            Call LockWindowUpdate(0)
            Cntrl.Refresh
        Case False
            Call LockWindowUpdate(Cntrl.hwnd)
    End Select
End Sub
Public Function GetType(Extension As String) As String
    'Get a file type name using its extension
    Dim sTmp                    As String
    sTmp = REGGetSetting(vHKEY_CLASSES_ROOT, "." & Extension, "")
    If Len(sTmp) > 0 Then
        sTmp = REGGetSetting(vHKEY_CLASSES_ROOT, sTmp, "")
    End If
    GetType = sTmp
End Function
Public Function StringSize(Size As Long) As String
    'Return a size as a string in B, KB, MB or GB depending on its size.
'    Stop
    Select Case Size
        Case 0 To 1023
            StringSize = CStr(Format(Size, "##0") & " Bytes")
        Case 1024 To 1048575
            StringSize = CStr(Format(Size / 1024#, "#,##0") & " KB")
        Case 1024# ^ 2 To 1073741823
            StringSize = CStr(Format(Size / (1024# ^ 2), "#,##0.00") & " MB")
        Case Is > 1073741823#
            StringSize = CStr(Format(Size / (1024# ^ 3), "#,###,##0.00") & " GB")
    End Select
End Function
Public Function FixPath(sPath As String) As String
    FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "") 'Add a \ if it doesn't exist.
End Function
Public Function J_UBound(ByRef sArr() As String) As Long
    On Error GoTo ErrClear
    'My own little upper bound function.
    'Won't give me mass errors when I try and check array demensions.
    'Instead just returns -1 if the array is not initialised =)
    J_UBound = UBound(sArr)
    Exit Function
ErrClear:
    J_UBound = -1
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MaxLFNPath        As Integer = 260
Private Const vbBackslash       As String = "\"
Private Const vbAllFiles        As String = "*.*"
Private Const vbKeyDot          As Integer = 46
Public sShell32                 As String
Public Enum ResType
    TEXT_RESOURCE_WORKSPACE = 4162
    TEXT_RESOURCE_MYCOMPUTER = 9216
    TEXT_RESOURCE_CONTROLPANEL = 4161
    
    TEXT_RESOURCE_COL_NAME = 8976
    TEXT_RESOURCE_COL_SIZE = 8978
    TEXT_RESOURCE_COL_TYPE = 8979
    TEXT_RESOURCE_COL_MODIFIED = 8980
    TEXT_RESOURCE_COL_CREATED = 8996
    
    ICON_RESOURCE_WORKSPACE = 34
    ICON_RESOURCE_MYCOMPUTER = 16
    ICON_RESOURCE_MYDOCUMENTS = 20
    ICON_RESOURCE_NETWOORK = 17
    ICON_RESOURCE_CONTROLPANEL = 35
End Enum
Private Enum FileAtribs
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum
Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type
Private Type FILETIME
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
    cFileName                   As String * MaxLFNPath
    cShortFileName              As String * 14
End Type
Private Type TIME_ZONE_INFORMATION
    bias                        As Long
    StandardName(0 To 63)       As Byte
    StandardDate                As SYSTEMTIME
    StandardBias                As Long
    DaylightName(0 To 63)       As Byte
    DaylightDate                As SYSTEMTIME
    DaylightBias                As Long
End Type
Private WFD                     As WIN32_FIND_DATA
Private sZipFolders()           As String
Private hItem                   As Long
Private hFile                   As Long
Private msFileSpec              As String
Private msPath                  As String
Private mbSearching             As Boolean
Private mbSubDirs               As Boolean
Private mbSystemDirs            As Boolean
Private mbHiddenDirs            As Boolean
Private mbUseCase               As Boolean
Private mbSearchString          As String
Private sExtension              As String
Private mbExtension             As Boolean
Private mbUseString             As Boolean
Private mbZipFiles              As Boolean
Private mbCompareDate           As tCompareDate
Private mbSize                  As Double
Private mbSizeType              As Integer
Private mlTotalDirs             As Long
Private mlTotalFiles            As Long
Public Event BeginFindFiles()
Public Event EndFindFiles(FileCount As Long)
Public Event FoundFile(FileName As String, FilePath As String, Size As Long, InZip As Boolean, sFileType As String, sTmpLocation As String, sDate As String, Cancel As Boolean)
Public Event FolderChange(sFolder As String)
Private Function ConvertTime(fFiletime As FILETIME, Optional bFull As Boolean = False) As String
    Dim SysTime                 As SYSTEMTIME
    Dim Ret                     As Long
    Dim sDateSetial             As Single
    'Convert file time to the system time
    Ret = FileTimeToSystemTime(fFiletime, SysTime)
    If Ret Then
        sDateSetial = DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay) 'Date serial to convert to string
        If sDateSetial > 0 Then
            'Convert the time to a string we can read
            If bFull = True Then
                ConvertTime = Format$(sDateSetial, "dd/mm/yy " & String(2 - Len(CStr(SysTime.wHour)), "0") & SysTime.wHour & ":" & String(2 - Len(CStr(SysTime.wMinute)), "0") & SysTime.wMinute)
                'Stop
            Else
                ConvertTime = Format$(sDateSetial, "dd/mm/yy")
            End If
        Else
            ConvertTime = "(No Date)"
        End If
    End If
End Function
Public Sub SetCompareProps(tType As eCompareDate, tTime As Integer, tCompare As Boolean, tWhich As eCompareType)
    'Fast easy way to set multiple properties at once =)
    With mbCompareDate
        .tCompare = tCompare
        .tType = tType
        .tTime = tTime
        .tWhich = tWhich
    End With
End Sub
Public Property Let sType(ByVal vData As String)
    'Only certain type ?
    sExtension = vData
    If Len(vData) = 0 Then
        mbExtension = 0
    Else
        mbExtension = 1
    End If
End Property
Public Property Get sType() As String
    sType = sExtension
End Property
Public Property Let SearchZips(ByVal vData As Boolean)
    'Search Zip files ?
    mbZipFiles = vData
End Property
Public Property Get SearchZips() As Boolean
    SearchZips = mbZipFiles
End Property
Public Property Let CompareSizeType(ByVal vData As Integer)
    'Use Size matching ? If so which type.
    mbSizeType = vData
End Property
Public Property Get CompareSizeType() As Integer
    CompareSizeType = mbSizeType
End Property
Public Property Let CompareSize(ByVal vData As Double)
    'The size limit we might be searching for
    mbSize = vData
End Property
Public Property Get CompareSize() As Double
    CompareSize = mbSize
End Property
Public Property Let SearchWords(ByVal vData As String)
    'Containing text....
    mbUseString = (Len(vData) > 0)
    mbSearchString = vData
End Property
Public Property Get SearchWords() As String
    SearchWords = mbSearchString
End Property
Public Property Let CaseSensitive(ByVal vData As Boolean)
    'Ignore case ?
    mbUseCase = vData
End Property
Public Property Get CaseSensitive() As Boolean
    CaseSensitive = mbUseCase
End Property
Public Property Let SearchHiddenFolders(ByVal vData As Boolean)
    'Search hidden folders ?
    mbHiddenDirs = vData
End Property
Public Property Get SearchHiddenFolders() As Boolean
    SearchHiddenFolders = mbHiddenDirs
End Property
Public Property Let SearchSystemFolders(ByVal vData As Boolean)
    'Search system folders ?
    mbSystemDirs = vData
End Property
Public Property Get SearchSystemFolders() As Boolean
    SearchSystemFolders = mbSystemDirs
End Property
Public Property Let SearchSubFolders(ByVal vData As Boolean)
    'Search sub folders ?
    mbSubDirs = vData
End Property
Public Property Get SearchSubFolders() As Boolean
    SearchSubFolders = mbSubDirs
End Property
Public Property Let Path(ByVal vData As String)
    'This is always required, the path to search.
    msPath = vData
End Property
Public Property Get Path() As String
    Path = msPath
End Property
Public Property Let FileSpec(ByVal vData As String)
    'The regular expression to match.
    msFileSpec = vData
End Property
Public Property Get FileSpec() As String
    FileSpec = msFileSpec
End Property
Public Property Let Searching(ByVal vData As Boolean)
    'Our property to cancel the search
    mbSearching = vData
End Property
Public Property Get Searching() As Boolean
    Searching = mbSearching
End Property
Public Function FindAll(ByRef FileListArray() As String) As Boolean
    Dim asfiles()               As String
    'Our find all files function, this was the original, edited of course.
    'I really should just remove it for efficiency but im a lazy git to be honest.
    Erase sZipFolders
    If FindFiles(asfiles()) Then
'    Stop
        FileListArray = asfiles()
        FindAll = True
    Else
        FindAll = False
    End If
eop_error:
    Select Case Err.Number
        Case Is > 0
            FindAll = False
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
eop_exit:
End Function
Private Function FindFiles(asFoundFiles() As String) As Boolean
    FindFiles = False
    If Searching Then
        Searching = False
        GoTo eop_exit
    End If
    'On Error Resume Next
    ' A parsing routine could be implemented here for
    ' multiple file spec searches, i.e. "*.bmp,*.wmf", etc.
    ' See the MS KB article Q130860 for information on how
    ' FindFirstFile() does not handle the "?" wildcard char correctly !!
    If Len(FileSpec) = 0 Then GoTo eop_exit
    If Len(Path) = 0 Then GoTo eop_exit
    mbSearching = True
    RaiseEvent BeginFindFiles
    Call SearchDirs(Path, asFoundFiles())
    Searching = False
    mlTotalFiles = J_UBound(asFoundFiles)
'    Stop
    RaiseEvent EndFindFiles(mlTotalFiles)
    FindFiles = True
eop_exit:
End Function
Private Sub SearchDirs(CurPath As String, asFoundFiles() As String)
    Dim dirs                      As Integer
    Dim dirbuf()                  As String
    Dim I                         As Integer
    On Error GoTo ErrClear
    DoEvents
    'Tell the program we are searching a new folder.
    RaiseEvent FolderChange(CurPath)
    'If user cancelled then exit.
    If Not Searching Then GoTo eop_exit
    'Find the first file.
    hItem& = FindFirstFile(CurPath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        Do
'        Stop
            If mbSubDirs And (WFD.dwFileAttributes And vbDirectory) Then
                'If its a folder.
                'Check the following if we need to.
                If Not mbHiddenDirs And (WFD.dwFileAttributes And FileAtribs.FILE_ATTRIBUTE_HIDDEN) = FileAtribs.FILE_ATTRIBUTE_HIDDEN Then GoTo Skipper
                If Not mbSystemDirs And (WFD.dwFileAttributes And FileAtribs.FILE_ATTRIBUTE_SYSTEM) = FileAtribs.FILE_ATTRIBUTE_SYSTEM Then GoTo Skipper
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    'If its the good old '.', aka folder up.
'                    Stop
                    mlTotalDirs = mlTotalDirs + 1
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    'Saves procesing time to only add every ten ticks.
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            End If
Skipper:
        'Continue looping while we are finding more files.
        Loop While FindNextFile(hItem&, WFD)
        'Stop the searching.
        Call FindClose(hItem&)
    End If
    'Our function to search indiviual files.
    Call SearchFileSpec(CurPath$, asFoundFiles())
    For I = 1 To dirs
        'Search sub folders.
        SearchDirs CurPath$ & dirbuf$(I%) & vbBackslash, asFoundFiles()
    Next
    'My dirty little hack for searching inside zip files.
    Call SearchZipFiles(asFoundFiles())
eop_exit:
Exit Sub
ErrClear:
    MsgBox ("SearchDirs() : Error: #" & Err.Number & " : " & Err.Description & " : " & Err.Source)
End Sub
Private Sub SearchFileSpec(CurPath As String, asFoundFiles() As String, Optional ZipFolder As String = "")
    Dim Cancel                  As Boolean
    Dim sTempFile               As String
    Dim Pos                     As Integer
    Dim sFileTime               As String
    Dim wftTime                 As FILETIME
    Dim sTmp                    As String
    Dim sExt                    As String
    Dim sTime                   As String
    On Error GoTo ErrClear
    hFile& = FindFirstFile(CurPath$ & FileSpec, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If Not mbSearching Then GoTo eop_exit
            sTempFile = CurPath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            'Full path
            Pos = InStrRev(sTempFile, "\")
            'Check for end of folder name.
            If Not Right$(sTempFile, Len(sTempFile) - Pos) = "." And Not Right$(sTempFile, Len(sTempFile) - Pos) = ".." Then
                'If its not a REAL folder...
                If Not mbHiddenDirs And (WFD.dwFileAttributes And FileAtribs.FILE_ATTRIBUTE_HIDDEN) = FileAtribs.FILE_ATTRIBUTE_HIDDEN Then GoTo Skipperd
                If mbZipFiles = True Then
                    'If we want to search inside these later...
                    If Right$(sTempFile, 4) = ".zip" Then
                        ReDim Preserve sZipFolders(J_UBound(sZipFolders) + 1)
                        sZipFolders(J_UBound(sZipFolders)) = sTempFile
                        'Add it to our array of zip files.
                    End If
                End If
                If mbUseString Then If CheckForString(sTempFile, mbSearchString) = False Then GoTo Skipperd 'Check containing string
                Pos = InStrRev(sTempFile, ".")
                If mbExtension = True And Pos = 0 Then GoTo Skipperd
                sExt = Right(sTempFile, Len(sTempFile) - Pos)
                sTmp = GetType(sExt)
                If mbExtension = True Then If Not StrComp(sTmp, sExtension, vbBinaryCompare) = 0 Then GoTo Skipperd
                If DirExists(sTempFile) = True Then
                    sTmp = "File Folder"
                Else
                    If Len(sTmp) = 0 And Len(sExt) > 0 Then sTmp = UCase(sExt) & " File"
                End If
                sTime = ConvertTime(WFD.ftLastWriteTime, True)
                If mbCompareDate.tCompare = True Then
                    'If we need to compare file dates.
                    If (mbCompareDate.tWhich = eCreated) Then
                        'Set to check Creation time.
                        wftTime = WFD.ftCreationTime
                        sFileTime = ConvertTime(wftTime)
                    Else
                        'Set to check Modification time.
                        wftTime = WFD.ftLastWriteTime
                        sFileTime = ConvertTime(wftTime)
                    End If
                    If sFileTime = "(No Date)" Then GoTo Skipperd 'Has no date
                    If CompareDates(sFileTime) = False Then GoTo Skipperd 'Isnt within our search parametres.
                End If
                If mbSize > 0 Then
                    'If we need to check file size
                    If mbSizeType = 1 Then
                        'Check is the file size greater than allowed
                        If WFD.nFileSizeLow > CDbl(mbSize * 1024) Then GoTo Skipperd
                    Else
                        'Check is the file size less than allowed
                        If WFD.nFileSizeLow < CDbl(mbSize * 1024) Then GoTo Skipperd
                    End If
                End If
                ReDim Preserve asFoundFiles(J_UBound(asFoundFiles) + 1)
                Pos = InStrRev(sTempFile, "\")
                asFoundFiles(J_UBound(asFoundFiles)) = sTempFile
                'Add the file to the found files array.
                Cancel = False
'                Stop
                RaiseEvent FoundFile(Right(sTempFile, Len(sTempFile) - Pos), Mid(sTempFile, 1, Pos), WFD.nFileSizeLow, False, sTmp, vbNullString, sTime, Cancel)
                mbSearching = Not Cancel
Skipperd:
            End If
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
    End If
eop_exit:
Exit Sub
ErrClear:
    MsgBox ("SearchFileSpec() : Error: #" & Err.Number & " : " & Err.Description & " : " & Err.Source)
End Sub
Private Function CompareDates(sFileDate As String) As Boolean
    Dim sNow                    As String
    Dim iInterval               As Integer
    sNow = Format$(Now, "dd/mm/yy")
    'Format now to be the way I want it to.
    'Otherwise on cross continent box's, the date would be mm/dd/yy which is opposite to what the file time is converted to.
    If mbCompareDate.tType = eDay Then
        iInterval = DateDiff("d", sFileDate, sNow)
        'Check the date differance in days.
        If iInterval <= mbCompareDate.tTime Then
            CompareDates = True
        Else
            CompareDates = False
        End If
    Else
        iInterval = DateDiff("m", sFileDate, sNow)
        'Check the date difference in months.
        If iInterval <= mbCompareDate.tTime Then
            CompareDates = True
        Else
            CompareDates = False
        End If
    End If
End Function
Private Function CheckForString(sFile As String, sWords As String) As Boolean
    Dim sContents               As String
    Dim cCompare                As VbCompareMethod
    cCompare = IIf(mbUseCase, vbBinaryCompare, vbTextCompare)
    sContents = OpenFile(sFile)
    'Search the contents of the file for our search key words.
    'Note : I couldnt find a decent Regular Expression searching example so I didnt use it =)
    CheckForString = (InStr(1, sContents, sWords, cCompare) > 0)
End Function
Private Sub CreateFile(Path As String, Size As Long)
    Dim sTmp                    As String
    Dim Free                    As Integer
    Free = FreeFile
    If Right(Path, 1) = "\" Then Exit Sub
    'Create the temporary file for zip searching.
    Open Path For Binary Access Write As #Free
        Put Free, , String(Size, 0)
    Close #Free
End Sub
Private Function RemDirs(ByVal sFile As String) As String
    Dim Pos                     As Long
    'Remove any folders which may be included in the zip
    'I probably should use them, but its hard when Im doing it like this
    'And im a very lazy person !
    sFile = Replace(sFile, "/", "\")
    'Seems zip files use / instead of \ to seperate folders.
    'I didnt know this untill today !
    If InStr(1, sFile, "\") > 0 Then
        Pos = InStr(1, sFile, "\")
        If Pos > 0 Then
            RemDirs = Mid(sFile, Pos + 1, Len(sFile) - Pos + 1)
        Else
            RemDirs = sFile
        End If
    Else
        RemDirs = sFile
    End If
End Function
Private Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    'My little function to check folder existance.
    'Best working one I ever used.
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    Err.Clear
End Function
Private Sub SearchZipFiles(asFoundFiles() As String)
    On Error Resume Next
    'The dirtiest hack I have ever ever used.
    'This is really disgracefull but I tried and tried and could not get Regular Expression searching to work
    'So I used the windows FindFirst file again which uses Regular Expressions as we have seen before.
    'What I do is create files the same size as the ones inside the zip files we are searching
    'in a temporary folder and search that folder, deleting all the contents again afterwards.
    'As I said, very very very dirty hack and should NOT be used as an example of how to search zips !
    'I only added this here as an added extra and so you could see how to read Zip Headers to create your own method.
    'I am sure someone can find a much better, cleaner, more efficient and worthwhile method so someone please do !
    'If you do come up with a better method or even some way to use Regular Expression searching so I can
    'just search the filenames in the zip files, then email me at c0ldfyr3@eliteprodigy.com
    Dim sZips()                 As String
    Dim X                       As Long
    Dim sTmp                    As String
    Dim y                       As Long
    Dim aTmp()                  As String
    y = J_UBound(sZipFolders)
    For y = 0 To y
        sZips = ModZip.SearchZip(sZipFolders(y))
        'Get array of file names contained in the zip file
        X = J_UBound(sZips)
        If X > -1 Then
            sTmp = SpecialFolder(Temp)
            'Get the current users Temp folder
            If Not Right(sTmp, 1) = "\" Then sTmp = sTmp & "\"
            'Add a \ if not present.
            sTmp = sTmp & "cSearch\"
            'Use our sub own sub folder.
            If Not DirExists(sTmp) Then MkDir (sTmp)
            'Make the sub folder if it isnt there.
            For X = 0 To J_UBound(sZips)
                'Loop through the files creating each one.
                aTmp = Split(sZips(X), "|")
                'Split the Name|Size array.
                Do While True
                    If J_UBound(aTmp) >= 0 Then
                        aTmp(0) = RemDirs(aTmp(0))
                        'Remove all sub folders in the file name/zip archive.
                        If InStr(1, aTmp(0), "/") = 0 And InStr(1, aTmp(0), "\") = 0 Then Exit Do
                    End If
                Loop
                Call CreateFile(sTmp & aTmp(0), CLng(aTmp(1)))
                'Create the temp file in the folder to search.
            Next
            Call SearchZipTemp(sTmp$, sZipFolders(y), asFoundFiles())
            'Search this folder using only name and size comparison.
            For X = 0 To J_UBound(sZips)
                aTmp = Split(sZips(X), "|")
                Do While True
                    If J_UBound(sZips) >= 0 Then
                        aTmp(0) = RemDirs(aTmp(0))
                        'I done this twice, im too lazy to change it =)
                        If InStr(1, aTmp(0), "/") = 0 And InStr(1, aTmp(0), "\") = 0 Then Exit Do
                    End If
                Loop
                'Remove all the files we created so they aren't found in the next zip we "search".
                If Len(aTmp(0)) > 0 Then Kill sTmp & aTmp(0)
            Next
        End If
    Next
End Sub
Private Sub SearchZipTemp(Location As String, ZipLocation As String, asFoundFiles() As String)
    Dim Cancel                  As Boolean
    Dim sTempFile               As String
    Dim Pos                     As Integer
    Dim sFileTime               As String
    Dim wftTime                 As FILETIME
    Dim sTmp                    As String
    Dim sExt                    As String
    'Do I really need to comment this function ?
    'Come on; its the same as the last one with some of the comparisons removed for gods sake !
    'Ok ok, I will for laughs =)
    RaiseEvent FolderChange(ZipLocation)
    'We changed folders again, this time to a zip file WoWy !
    hFile& = FindFirstFile(Location & FileSpec, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If Not mbSearching Then GoTo eop_exit
            'If we cancelled, then exit.
            sTempFile = Location & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            'Full location of file.
            Pos = InStrRev(sTempFile, "\")
            'Get folder and file seperation.
            If Not Right$(sTempFile, Len(sTempFile) - Pos) = "." And Not Right$(sTempFile, Len(sTempFile) - Pos) = ".." Then
                'If its not a "real" folder ....
                If mbSize > 0 Then
                    'Size comparison again.
                    If mbSizeType = 1 Then
                        If WFD.nFileSizeLow > CDbl(mbSize * 1024) Then GoTo Skipperd
                    Else
                        If WFD.nFileSizeLow < CDbl(mbSize * 1024) Then GoTo Skipperd
                    End If
                End If
                Pos = InStrRev(sTempFile, ".")
                If mbExtension = True And Pos = 0 Then GoTo Skipperd
                sExt = Right(sTempFile, Len(sTempFile) - Pos)
                sTmp = GetType(sExt)
                If mbExtension = True Then If Not StrComp(sTmp, sExtension, vbBinaryCompare) = 0 Then GoTo Skipperd
                If DirExists(sTempFile) = True Then
                    sTmp = "File Folder"
                Else
                    If Len(sTmp) = 0 And Len(sExtension) > 0 Then sTmp = UCase(sExtension) & " File"
                End If
                ReDim Preserve asFoundFiles(J_UBound(asFoundFiles) + 1)
                Pos = InStrRev(sTempFile, "\")
                'Add the filename.
                asFoundFiles(J_UBound(asFoundFiles)) = sTempFile
                Cancel = False
                'Tell the rest of the program we found a file !
                'Stop
                RaiseEvent FoundFile(Right(sTempFile, Len(sTempFile) - Pos), ZipLocation, WFD.nFileSizeLow, True, sTmp, Location, "(No Date)", Cancel)
                mbSearching = Not Cancel
Skipperd:
            End If
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
  End If
eop_exit:
End Sub
Public Function GetResourceString(ResourceType As ResType) As String
    'Gets strings from Shell32.dll
    GetResourceString = GetResourceStringFromFile(sShell32, ResourceType)
End Function
Private Function GetResourceStringFromFile(sModule As String, idString As Long) As String
    Dim hModule                 As Long
    Dim nChars                  As Long
    Dim Buffer                  As String * 260
    hModule = LoadLibrary(sModule) 'Load the library (dll)
    If hModule Then
        nChars = LoadString(hModule, idString, Buffer, 260) 'Load the string from the position specified.
        If nChars Then GetResourceStringFromFile = Left(Buffer, nChars) 'Trim it.
        FreeLibrary hModule 'Release the library from memory.
    End If
End Function
Private Sub Class_Initialize()
    sShell32 = FixPath(SpecialFolder(WinSystem)) & "shell32.dll" 'Get the shell32.dll location into memory for later use.
End Sub
Private Sub Class_Terminate()
    sShell32 = ""
End Sub

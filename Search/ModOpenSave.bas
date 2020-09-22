Attribute VB_Name = "ModOpenSave"
Option Explicit
'I didn't write any of this so I won't comment it.
'I don't know where I found it either, I have it about three years and have been using it in every project since.
'If you are decent with Visual Basic, I would suggest looking over the api calls and routines.
'Sorry =(
Private OpeningAFile            As Boolean
Private Xhwnd                   As Long
Private PXhwnd                  As Long
Private Xpos                    As Long
Private Ypos                    As Long
Private prevID                  As Long
Private WWdt                    As Long
Private OldProcedura            As Long
Private CLL                     As Long
Private DIALOGHWND              As Long
Private TMR                     As Long
Private TMRHWND                 As Long
Private Const GWL_ID            As Long = (-12)
Private Type PointAPI
     X                          As Long
     y                          As Long
End Type
Private Type BROWSEINFOTYPE
    hOwner                      As Long
    pidlRoot                    As Long
    pszDisplayName              As String
    lpszTitle                   As String
    ulFlags                     As Long
    lpfn                        As Long
    lParam                      As Long
    iImage                      As Long
End Type
Private Const SM_CXBORDER       As Long = 5
Private Const SM_CXDLGFRAME     As Long = 7
Private Const SM_CYCAPTION      As Long = 4
Private Const SM_CYDLGFRAME     As Long = 8
Private Const SM_CYBORDER       As Long = 6
Private Const SM_CYMENU         As Long = 15
Private Const SM_CYMENUSIZE     As Long = 55
Private Const SM_CXMENUSIZE     As Long = 54
Private Const WM_USER           As Long = &H400
Private Const WM_NOTIFY         As Long = &H4E
Private Const WM_PAINT          As Long = &HF
Private Const WM_DRAWITEM       As Long = &H2B
Private Const WM_SETTEXT        As Long = &HC
Private Const WM_SETREDRAW      As Long = &HB
Private Const WM_CLOSE          As Long = &H10
Private Const WM_INITDIALOG     As Long = &H110
Private Const WM_CREATE         As Long = &H1
Private Const WM_PARENTNOTIFY   As Long = &H210
Private Const WM_NCCREATE       As Long = &H81
Private Const WM_DESTROY        As Long = &H2
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOACTIVATE    As Long = &H10
Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type
Private Const MAX_PATH = 260
Private Type OPENFILENAME
    lStructSize                 As Long
    hWndOwner                   As Long
    hInstance                   As Long
    lpstrFilter                 As String
    lpstrCustomFilter           As String
    nMaxCustFilter              As Long
    nFilterIndex                As Long
    lpstrFile                   As String
    nMaxFile                    As Long
    lpstrFileTitle              As String
    nMaxFileTitle               As Long
    lpstrInitialDir             As String
    lpstrTitle                  As String
    flags                       As OFN_Flags
    nFileOffset                 As Integer
    nFileExtension              As Integer
    lpstrDefExt                 As String
    lCustData                   As Long
    lpfnHook                    As Long
    lpTemplateName              As String
End Type
Public Enum OFN_Flags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
    OFN_ENABLEINCLUDENOTIFY = &H400000
    OFN_ENABLESIZING = &H800000
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Property Let SetWidth(ByVal newwidth As Long)
    WWdt = newwidth
End Property
Public Sub InsertCtrl(ByVal Ctrlhwnd As Long, ByVal Parhwnd As Long, ByVal X As Long, ByVal y As Long)
    Xhwnd = Ctrlhwnd
    PXhwnd = Parhwnd
    Xpos = X + 4
    Ypos = y + 23
End Sub
Public Function GetOpenFilePath(hwnd As Long, sFilter As String, iFilter As Integer, sFile As String, sInitDir As String, sTitle As String, sRtnPath As String, lFlags As Long) As Boolean
    Dim ofn As OPENFILENAME
    OpeningAFile = True
    With ofn
        .lStructSize = Len(ofn)
        .hWndOwner = hwnd
        .lpstrFilter = sFilter & vbNullChar & vbNullChar
        .nFilterIndex = iFilter
        .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
        .nMaxFile = MAX_PATH
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle & vbNullChar
        .flags = lFlags
        .lpfnHook = GetAddress(AddressOf HookX)
    End With
    If GetOpenFileName(ofn) Then
        iFilter = ofn.nFilterIndex
        sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
        sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
        GetOpenFilePath = True
    End If
End Function
Public Function GetSaveFilePath(hwnd As Long, sFilter As String, iFilter As Integer, sDefExt As String, sFile As String, sInitDir As String, sTitle As String, sRtnPath As String, lFlags As Long) As Boolean
    Dim ofn As OPENFILENAME
    OpeningAFile = False
    With ofn
        .lStructSize = Len(ofn)
        .hWndOwner = hwnd
        .lpstrFilter = sFilter & vbNullChar & vbNullChar
        .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
        .lpstrDefExt = sDefExt
        .nMaxFile = MAX_PATH
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle & vbNullChar
        .flags = lFlags
        .lpfnHook = GetAddress(AddressOf HookX)
    End With
    If GetSaveFileName(ofn) Then
        iFilter = ofn.nFilterIndex
        sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
        sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
        GetSaveFilePath = True
    End If
End Function
Public Function GetStrFromBufferA(szA As String) As String
    If InStr(szA, vbNullChar) Then
        GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
    Else
        GetStrFromBufferA = szA
    End If
End Function
Public Function GetAddress(ByVal address As Long) As Long
    GetAddress = address
End Function
Public Function HookX(ByVal hDlg As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim przRECT                 As RECT
    Dim XY                      As PointAPI
    Dim RC1                     As RECT
    Dim Parhwnd                 As Long
    Dim ctrlX                   As Long
    Dim MTY                     As Long
    Dim MTX                     As Long
    Dim X                       As Long
    Dim y                       As Long
    On Error GoTo Ender
    Select Case uiMsg
        Case WM_INITDIALOG
            Parhwnd = GetParent(hDlg)
            DIALOGHWND = Parhwnd
            ctrlX = GetDlgItem(Parhwnd, &H1)
            CLL = ctrlX
            OldProcedura = SetWindowLong(ctrlX, -4, AddressOf Provjera2)
            SetWindowText CLL, OkButtonText
            SetParent Xhwnd, Parhwnd
            prevID = GetDlgCtrlID(Xhwnd)
            SetWindowLong Xhwnd, GWL_ID, &H6000
            GetWindowRect Parhwnd, RC1
            XY.X = RC1.Left
            XY.y = RC1.Top
            ScreenToClient Parhwnd, XY
            GetWindowRect Xhwnd, RC1
            MTY = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER)
            MTX = GetSystemMetrics(SM_CXBORDER)
            MoveWindow Xhwnd, XY.X + Xpos + MTX, XY.y + Ypos + MTY, RC1.Right - RC1.Left, RC1.Bottom - RC1.Top, 1
            ShowWindow Xhwnd, 1
            SetDlgItemText Parhwnd, &H2, CancelButtonText
            SetDlgItemText Parhwnd, &H443, "There:"
            SetDlgItemText Parhwnd, &H442, "File Name:"
            SetDlgItemText Parhwnd, &H441, IIf(OpeningAFile, "Files Of Type:", "Save As Type:")
            GetWindowRect Parhwnd, RC1
            MoveWindow Parhwnd, 0, 0, RC1.Right - RC1.Left, WWdt, 1
            GetWindowRect Parhwnd, przRECT
            X = (Screen.Width / 15 - (przRECT.Right - przRECT.Left)) / 2
            y = (Screen.Height / 15 - (przRECT.Bottom - przRECT.Top)) / 2
            SetWindowPos Parhwnd, 0, X, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        Case WM_DESTROY
            ShowWindow Xhwnd, 0
            SetParent Xhwnd, Parhwnd
            SetWindowLong Xhwnd, GWL_ID, prevID
            Call SetWindowLong(CLL, -4, OldProcedura)
    End Select
    Exit Function
Ender:
    Err.Clear
End Function
Public Function OkButtonText(Optional TheText As String) As String
    Static strText              As String
    If Len(TheText) = 0 Then
        If strText = "" Then strText = "Ok"
        OkButtonText = strText
    Else
        strText = TheText
        OkButtonText = TheText
    End If
End Function
Public Function CancelButtonText(Optional TheText As String) As String
    Static strText              As String
    If Len(TheText) = 0 Then
        If strText = "" Then strText = "Cancel"
        CancelButtonText = strText
    Else
        strText = TheText
        CancelButtonText = TheText
    End If
End Function
Public Function Provjera2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim TTX                     As String
    Dim TX()                    As Byte
    If uMsg = WM_SETTEXT Then
        TTX = OkButtonText() & Chr(CByte(0))
        TX = StrConv(TTX, vbFromUnicode)
        lParam = VarPtr(TX(0))
    End If
    Provjera2 = CallWindowProc(OldProcedura, hwnd, uMsg, wParam, lParam)
End Function
Public Sub Provjera(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
    Dim TXT1                    As String
    Dim ltxt1                   As Long
    TXT1 = Space(20)
    ltxt1 = GetWindowText(hwnd, TXT1, Len(TXT1))
    TXT1 = Left(TXT1, ltxt1)
    If TXT1 = "&Open" Then
        SetWindowText hwnd, "Open Folder"
    ElseIf TXT1 = "&Save" Then
        SetWindowText hwnd, "Save"
    End If
End Sub
Public Sub CloseDLG()
    PostMessage DIALOGHWND, WM_CLOSE, 0, 0
End Sub
Private Function FunctionPointer(FunctionAddress As Long) As Long
    FunctionPointer = FunctionAddress
End Function
Public Function BrowseForFolder(startPath As String, Optional ByVal strTitle As String) As String
    Dim Browse_for_folder       As BROWSEINFOTYPE
    Dim itemID                  As Long
    Dim selectedPathPointer     As Long
    Dim tmpPath                 As String * 256
    Dim selectedPath            As String
    Dim LPTR                    As Long
    selectedPath = startPath
    If Len(selectedPath) > 0 Then
        If Not Right$(selectedPath, 1) <> "\" Then selectedPath = Left$(selectedPath, Len(selectedPath) - 1)
    End If
    With Browse_for_folder
        .hOwner = 0
        If Len(strTitle) > 0 Then
            .lpszTitle = strTitle
        Else
            .lpszTitle = "Please, Select a folder."
        End If
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
        selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1)
        CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1
        .lParam = selectedPathPointer
    End With
    itemID = SHBrowseForFolder(Browse_for_folder)
    If itemID Then
        If SHGetPathFromIDList(itemID, tmpPath) Then
            BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(itemID)
    End If
    Call LocalFree(selectedPathPointer)
End Function
Private Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = 1 Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If
End Function
Public Function CreateFilterType(Name As String, Ext As String) As String
    CreateFilterType = Name & " (*." & Ext & ")" + Chr$(0) + "*." & Ext
End Function
Public Function CreateFilters(StrFilters() As String) As String
    CreateFilters = Join(StrFilters, Chr(0))
End Function

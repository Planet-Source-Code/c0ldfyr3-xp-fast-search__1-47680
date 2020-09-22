Attribute VB_Name = "ModArray"
Option Explicit
'None of this is mine, check the Credits for where to get an example.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum
Public Sub TriQuickSortString(ByRef sArray() As String, Optional ByVal SortOrder As SortOrder = SortAscending)
    Dim iLBound                 As Long
    Dim iUBound                 As Long
    Dim I                       As Long
    Dim j                       As Long
    Dim sTemp                   As String
    If J_UBound(sArray) = -1 Then Exit Sub
    iLBound = LBound(sArray)
    iUBound = UBound(sArray)
    TriQuickSortString2 sArray, 4, iLBound, iUBound
    InsertionSortString sArray, iLBound, iUBound
    If SortOrder = SortDescending Then ReverseStringArray sArray
End Sub
Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
    Dim I                       As Long
    Dim j                       As Long
    Dim sTemp                   As String
    If (iMax - iMin) > iSplit Then
        I = (iMax + iMin) / 2
        If sArray(iMin) > sArray(I) Then SwapStrings sArray(iMin), sArray(I)
        If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
        If sArray(I) > sArray(iMax) Then SwapStrings sArray(I), sArray(iMax)
        j = iMax - 1
        SwapStrings sArray(I), sArray(j)
        I = iMin
        CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)
        Do
            Do
                I = I + 1
            Loop While sArray(I) < sTemp
            Do
                j = j - 1
            Loop While sArray(j) > sTemp
            If j < I Then Exit Do
                SwapStrings sArray(I), sArray(j)
        Loop
        SwapStrings sArray(I), sArray(iMax - 1)
        TriQuickSortString2 sArray, iSplit, iMin, j
        TriQuickSortString2 sArray, iSplit, I + 1, iMax
    End If
    I = 0
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(I), 4
End Sub
Private Sub InsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
    Dim I                       As Long
    Dim j                       As Long
    Dim sTemp                   As String
    For I = iMin + 1 To iMax
        CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(I)), 4 ' sTemp = sArray(i)
        j = I
        Do While j > iMin
            If sArray(j - 1) <= sTemp Then Exit Do
            CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4 ' sArray(j) = sArray(j - 1)
            j = j - 1
        Loop
        CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
    Next I
    I = 0
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(I), 4
End Sub
Public Sub ReverseStringArray(ByRef sArray() As String)
    Dim iLBound                 As Long
    Dim iUBound                 As Long
    iLBound = LBound(sArray)
    iUBound = UBound(sArray)
    While iLBound < iUBound
        SwapStrings sArray(iLBound), sArray(iUBound)
        iLBound = iLBound + 1
        iUBound = iUBound - 1
    Wend
End Sub
Private Sub SwapStrings(ByRef s1 As String, ByRef s2 As String)
    Dim I                       As Long
    I = StrPtr(s1)
    If I = 0 Then CopyMemory ByVal VarPtr(I), ByVal VarPtr(s1), 4
    CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
    CopyMemory ByVal VarPtr(s2), I, 4
End Sub

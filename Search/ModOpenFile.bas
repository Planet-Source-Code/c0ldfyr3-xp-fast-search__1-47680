Attribute VB_Name = "ModOpenFile"
Option Explicit
Public Function OpenFile(Location As String) As String
    'Open a file in its entirity.
    Dim StrFinal                    As String
    Dim FileL                       As Long
    Dim Free                        As Integer
    Free = FreeFile
    If Len(Dir(Location)) > 0 Then
        Open Location For Binary Access Read As #Free
            FileL = LOF(Free)
            StrFinal = Space(FileL)
            Get #Free, , StrFinal
        Close #Free
    End If
    OpenFile = StrFinal
End Function
Public Function SaveFile(Location As String, sData As String)
    'Save a file in its entirity.
    Dim Free                    As Integer
    Free = FreeFile
    If Len(Dir(Location)) > 0 Then Kill (Location)
    Open Location For Binary Access Write As #Free
        Put #Free, , sData
    Close #Free
End Function

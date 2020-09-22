Attribute VB_Name = "ModZip"
Option Explicit
'This code was ripped from someone, no idea who.
'I left in the parts for getting the file information even though I don't need it.
'I may use it in a later release.
Private Type ZipFile
    Version                     As Integer
    Flag                        As Integer
    CompressionMethod           As Integer
    Time                        As Integer
    Date                        As Integer
    CRC32                       As Long
    CompressedSize              As Long
    UncompressedSize            As Long
    FileNameLength              As Integer
    ExtraFieldLength            As Integer
    FileName                    As String
End Type
Private Const LocalFileSig      As Long = &H4034B50
Private Const CentralFileSig    As Long = &H2014B50
Private Const EndCentralDirSig  As Long = &H6054B50
Public Function SearchZip(ZipPath As String) As String()
    Dim sArr()                  As String
    Dim Sig                     As Long
    Dim ZipStream               As Integer
    Dim Res                     As Long
    Dim Name                    As String
    Dim zFile                   As ZipFile
    Dim I                       As Integer
    If ZipPath = "" Then Exit Function
    ZipStream = FreeFile
    Open ZipPath For Binary As ZipStream
        Do While True
            Get ZipStream, , Sig
            If Sig = LocalFileSig Then
                Get ZipStream, , zFile.Version
                Get ZipStream, , zFile.Flag
                Get ZipStream, , zFile.CompressionMethod
                Get ZipStream, , zFile.Time
                Get ZipStream, , zFile.Date
                Get ZipStream, , zFile.CRC32
                Get ZipStream, , zFile.CompressedSize
                Get ZipStream, , zFile.UncompressedSize
                Get ZipStream, , zFile.FileNameLength
                Get ZipStream, , zFile.ExtraFieldLength
                Name = String$(zFile.FileNameLength, " ")
                Get ZipStream, , Name
                ReDim Preserve sArr(J_UBound(sArr) + 1)
                'Add the file to the array we will be returning.
                sArr(J_UBound(sArr)) = Mid$(Name, 1, zFile.FileNameLength) & "|" & zFile.UncompressedSize
                Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
            Else
                If Sig = CentralFileSig Or Sig = 0 Then
                    Exit Do
                Else
                    If Sig = EndCentralDirSig Then
                        Exit Do
                    End If
                End If
            End If
        Loop
    Close ZipStream
    SearchZip = sArr()
End Function

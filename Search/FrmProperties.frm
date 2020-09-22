VERSION 5.00
Begin VB.Form FrmProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmVolumeInfo 
      Caption         =   " Volume Info "
      Height          =   1305
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   6225
      Begin VB.CheckBox chkERemote 
         Appearance      =   0  'Flat
         Caption         =   "Remote?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4080
         TabIndex        =   31
         Top             =   570
         Width           =   1365
      End
      Begin VB.CheckBox chkEFixedDisk 
         Appearance      =   0  'Flat
         Caption         =   "Fixed Disk?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   34
         Top             =   570
         Width           =   1365
      End
      Begin VB.CheckBox chkUNCServer 
         Appearance      =   0  'Flat
         Caption         =   "UNC Server?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4080
         TabIndex        =   30
         Top             =   930
         Width           =   1785
      End
      Begin VB.CheckBox chkUNC 
         Appearance      =   0  'Flat
         Caption         =   "UNC?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   930
         Width           =   1785
      End
      Begin VB.CheckBox chkNetworkPath 
         Appearance      =   0  'Flat
         Caption         =   "Network Path?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         Top             =   930
         Width           =   1785
      End
      Begin VB.CheckBox chkECDRom 
         Appearance      =   0  'Flat
         Caption         =   "CD Rom?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   570
         Width           =   1365
      End
      Begin VB.CheckBox chkERemovable 
         Appearance      =   0  'Flat
         Caption         =   "Removable?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   930
         Width           =   1365
      End
      Begin VB.CheckBox chkERamDisk 
         Appearance      =   0  'Flat
         Caption         =   "Ram Disk?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2880
         TabIndex        =   32
         Top             =   570
         Width           =   1365
      End
      Begin VB.TextBox txtVolSerialNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   27
         Text            =   "txtVolSerialNumber"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtVolLabel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   690
         TabIndex        =   26
         Text            =   "txtVolLabel"
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label lblVolSerialNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial No:"
         Height          =   195
         Left            =   2970
         TabIndex        =   37
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lblVolLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Label:"
         Height          =   195
         Left            =   225
         TabIndex        =   36
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame frmFileDates 
      Caption         =   " Dates "
      Height          =   1215
      Left            =   3120
      TabIndex        =   21
      Top             =   1920
      Width           =   3225
      Begin VB.TextBox meModified 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox meAccessed 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox meCreated 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCreated 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         Height          =   195
         Left            =   615
         TabIndex        =   24
         Top             =   270
         Width           =   600
      End
      Begin VB.Label lblAccessed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Last Accessed:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lblModified 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         Height          =   195
         Left            =   570
         TabIndex        =   22
         Top             =   840
         Width           =   645
      End
   End
   Begin VB.Frame frmAttributes 
      Caption         =   " Attributes "
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
      Begin VB.CheckBox chkNormal 
         Appearance      =   0  'Flat
         Caption         =   "Normal?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chkReadonly 
         Appearance      =   0  'Flat
         Caption         =   "Read Only?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   1275
      End
      Begin VB.CheckBox chkSystem 
         Appearance      =   0  'Flat
         Caption         =   "System?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   690
         Width           =   1275
      End
      Begin VB.CheckBox chkTemporary 
         Appearance      =   0  'Flat
         Caption         =   "Temporary?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   900
         Width           =   1275
      End
      Begin VB.CheckBox chkArchive 
         Appearance      =   0  'Flat
         Caption         =   "Archive?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1425
      End
      Begin VB.CheckBox chkCompressed 
         Appearance      =   0  'Flat
         Caption         =   "Compressed?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   480
         Width           =   1425
      End
      Begin VB.CheckBox chkDirectory 
         Appearance      =   0  'Flat
         Caption         =   "Directory?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   690
         Width           =   1425
      End
      Begin VB.CheckBox chkHidden 
         Appearance      =   0  'Flat
         Caption         =   "Hidden?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   900
         Width           =   1425
      End
   End
   Begin VB.TextBox txtFullFileName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Text            =   "txtFullFileName"
      Top             =   120
      Width           =   5205
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Text            =   "txtFileName"
      Top             =   1110
      Width           =   5205
   End
   Begin VB.TextBox txtFileExtension 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Text            =   "txtFileExtension"
      Top             =   1440
      Width           =   585
   End
   Begin VB.TextBox txtFileSize 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Text            =   "txtFileSize"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtShortFullFilename 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Text            =   "txtShortFullFilename"
      Top             =   450
      Width           =   5205
   End
   Begin VB.TextBox txtPathRoot 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Text            =   "txtPathRoot"
      Top             =   780
      Width           =   5205
   End
   Begin VB.Label lblFullFilename 
      Alignment       =   1  'Right Justify
      Caption         =   "Filename:"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label lblExtension 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ext:"
      Height          =   195
      Left            =   825
      TabIndex        =   9
      Top             =   1530
      Width           =   270
   End
   Begin VB.Label lblFileSize 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   1500
      Width           =   345
   End
   Begin VB.Label lblShortFullFilename 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Short Filename:"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblPathRoot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Path Root:"
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   840
      Width           =   765
   End
End
Attribute VB_Name = "FrmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sMyFileName                As String
Public sMyFilePath                As String
Private sMyFullFile               As String
Public Sub LoadData()
    Dim TmpAtrib                    As enumFileAttributes
    Dim TmpVol                      As typeVolumeInformation
    Dim TmpType                     As enumDriveTypes
    Dim lTmp                        As Long
    Dim sTmp                        As String
    Me.Caption = "Properties for " & sMyFileName 'Caption
    sMyFullFile = sMyFilePath & sMyFileName 'Full path of file.
    txtFileExtension = sFilename(sMyFullFile, efpFileExt) 'Extension
    txtFileName = sMyFilePath 'The path to the file
    txtFullFileName = sMyFileName 'File Name only
    meAccessed.Text = timeFileToDate(FileInformation(sMyFullFile).ftLastAccessTime)
         'Last access time. (Note: Will always be the time this function is called)
         'When they created windows, the great authors made a large booboo, and to get the accessed time, the file has to be accessed, see what Im saying ?
    meCreated.Text = timeFileToDate(FileInformation(sMyFullFile).ftCreationTime) 'Creation time.
    meModified.Text = timeFileToDate(FileInformation(sMyFullFile).ftLastWriteTime) 'Last Modified.
    lTmp = lSize(sMyFullFile) 'The Size in bytes.
    sTmp = StringSize(lTmp) 'Convert to stirng
    txtFileSize.Text = sTmp
    txtShortFullFilename.Text = FileShortName(sMyFullFile) 'Windows Short File name.
    txtPathRoot.Text = sPathRoot(sMyFilePath) 'Root path, Ex: C:\,D:\ etc
    
    TmpAtrib = eAttributes(sMyFullFile)
    chkArchive = Abs(CBool(TmpAtrib And efaARCHIVE)) 'Archive ?
    chkCompressed = Abs(CBool(TmpAtrib And efaCOMPRESSED)) 'Compressed ?
    chkDirectory = Abs(CBool(TmpAtrib And efaDIRECTORY)) 'Folder ?
    chkHidden = Abs(CBool(TmpAtrib And efaHIDDEN)) 'Hidden ?
    chkNormal = Abs(CBool(TmpAtrib And efaNORMAL))
    chkReadonly = Abs(CBool(TmpAtrib And efaREADONLY)) 'Read only ?
    chkSystem = Abs(CBool(TmpAtrib And efaSYSTEM)) 'System ?
    chkTemporary = Abs(CBool(TmpAtrib And efaTEMPORARY)) 'Temp File ?
    
    TmpVol = volumeInformation(txtPathRoot)
    txtVolLabel.Text = TmpVol.sVolumeName 'Drive Label
    txtVolSerialNumber.Text = TmpVol.lVolumeSerialNo 'Serial Number
    TmpType = GetDriveType(txtPathRoot.Text)
    chkECDRom = Abs(TmpType = DRIVE_CDROM) 'CD Drive  ?
    chkEFixedDisk = Abs(TmpType = DRIVE_FIXED) 'Fixed Disk ?
    chkERamDisk = Abs(TmpType = DRIVE_RAMDISK) 'Ram Drive ?
    chkERemote = Abs(TmpType = DRIVE_REMOTE) 'Remove Disk ?
    chkERemovable = Abs(TmpType = DRIVE_REMOVABLE) 'Removable Disk (Floppy, Zip Disk)?
    
    chkNetworkPath = Abs(PathIsNetworkPath(sMyFullFile))
    chkUNC = Abs(PathIsUNC(sMyFullFile))
    chkUNCServer = Abs(PathIsUNCServer(sMyFullFile))
End Sub

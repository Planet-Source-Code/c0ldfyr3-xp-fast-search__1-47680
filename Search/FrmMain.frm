VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Search"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13815
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWorking 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12840
      Top             =   8040
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   12480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   12840
      Picture         =   "FrmMain.frx":6DC2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   12960
      Picture         =   "FrmMain.frx":7A04
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12600
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   12960
      Top             =   6360
   End
   Begin VB.Timer tmrFrames 
      Interval        =   1
      Left            =   12600
      Top             =   8520
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E19272&
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   120
      ScaleHeight     =   9735
      ScaleWidth      =   4215
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   36
         Top             =   3960
         Width           =   3975
         Begin VB.TextBox txtSize 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   40
            Text            =   "1"
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cmbSize 
            Height          =   315
            ItemData        =   "FrmMain.frx":7D46
            Left            =   120
            List            =   "FrmMain.frx":7D50
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   480
            Width           =   1335
         End
         Begin MSComCtl2.UpDown UpDownSize 
            Height          =   255
            Left            =   2205
            TabIndex        =   41
            Top             =   480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            Max             =   1048576
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "KB"
            Height          =   255
            Left            =   2520
            TabIndex        =   42
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   60
            Width           =   375
         End
         Begin VB.Label lblBack 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   2535
         End
         Begin VB.Image MenuHeader 
            Height          =   375
            Index           =   3
            Left            =   1200
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   43
         Top             =   4440
         Width           =   3975
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Options"
            Height          =   255
            Left            =   840
            TabIndex        =   51
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox chkCase 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Case Sensitive"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1920
            Width           =   2055
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search System Folders"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkHidden 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search Hidden Files && Folders"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkSubFolders 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search Sub Folders"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkZips 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search Zip Files"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Advanced Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lblBack 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   2535
         End
         Begin VB.Image MenuHeader 
            Height          =   375
            Index           =   4
            Left            =   1200
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   32
         Top             =   3480
         Width           =   3975
         Begin MSComctlLib.ImageCombo ImageCombo 
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "File Types"
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   60
            Width           =   435
         End
         Begin VB.Label lblBack 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   2535
         End
         Begin VB.Image MenuHeader 
            Height          =   375
            Index           =   2
            Left            =   1200
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   19
         Top             =   3000
         Width           =   3975
         Begin VB.TextBox txtMonths 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Text            =   "1"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtDays 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   26
            Text            =   "1"
            Top             =   1080
            Width           =   495
         End
         Begin VB.OptionButton OpDate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Created"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OpDate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Modified"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OpLast 
            BackColor       =   &H00FFFFFF&
            Caption         =   "In the last"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OpLast 
            BackColor       =   &H00FFFFFF&
            Caption         =   "In the last"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   1095
         End
         Begin MSComCtl2.UpDown UpDownDays 
            Height          =   255
            Left            =   1980
            TabIndex        =   28
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            Max             =   32
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownMonths 
            Height          =   255
            Left            =   1980
            TabIndex        =   29
            Top             =   1440
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            Max             =   120
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   60
            Width           =   420
         End
         Begin VB.Label lblBack 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Months"
            Height          =   255
            Left            =   2280
            TabIndex        =   31
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Days"
            Height          =   255
            Left            =   2280
            TabIndex        =   30
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image MenuHeader 
            Height          =   375
            Index           =   1
            Left            =   1200
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   0
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   3975
         TabIndex        =   3
         Top             =   120
         Width           =   3975
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search Now"
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdFolder 
            Caption         =   "..."
            Height          =   255
            Left            =   3480
            TabIndex        =   14
            Top             =   1920
            Width           =   375
         End
         Begin VB.TextBox txtFolder 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtContaining 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Containing text:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Look in:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Search for files or folders named:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search For Files And Folders"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   60
            Width           =   2460
         End
         Begin VB.Label lblBack 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   2535
         End
         Begin VB.Image MenuHeader 
            Height          =   375
            Index           =   0
            Left            =   1200
            Top             =   0
            Width           =   2655
         End
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   10005
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4762
            MinWidth        =   4762
            Text            =   "Status: Idle"
            TextSave        =   "Status: Idle"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstFind 
      Height          =   9735
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   17171
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   12600
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   12480
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgNull 
      Left            =   12600
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgCombo 
      Left            =   12480
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Image imgUp 
      Height          =   375
      Left            =   10440
      Picture         =   "FrmMain.frx":7D67
      Top             =   7680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgDown 
      Height          =   375
      Left            =   10440
      Picture         =   "FrmMain.frx":B3F5
      Top             =   8040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Menu mnRClick 
      Caption         =   "RightClickMenu"
      Visible         =   0   'False
      Begin VB.Menu mnMainView 
         Caption         =   "View"
         Begin VB.Menu mnView 
            Caption         =   "Icon"
            Index           =   0
         End
         Begin VB.Menu mnView 
            Caption         =   "Small Icon"
            Index           =   1
         End
         Begin VB.Menu mnView 
            Caption         =   "List"
            Index           =   2
         End
         Begin VB.Menu mnView 
            Caption         =   "Report"
            Index           =   3
         End
      End
      Begin VB.Menu mnSort 
         Caption         =   "Sort By ..."
         Begin VB.Menu mnSortName 
            Caption         =   "Name"
         End
         Begin VB.Menu mnSortPath 
            Caption         =   "Path"
         End
         Begin VB.Menu mnSortSize 
            Caption         =   "Size"
         End
         Begin VB.Menu mnSortType 
            Caption         =   "Type"
         End
      End
      Begin VB.Menu mnOpenFolder 
         Caption         =   "Open Containing Folder"
      End
      Begin VB.Menu mnProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const LARGE_ICON        As Integer = 32
Private Const SMALL_ICON        As Integer = 16
Private Const Speed             As Integer = 60
Private Expand                  As Boolean
Private Frame                   As Integer
Private DoResize                As Boolean
Private WithEvents cFind        As clsSearch
Attribute cFind.VB_VarHelpID = -1
Private mbDate                  As Boolean
Private mbType                  As Boolean
Private mbSize                  As Boolean
Private Sub cFind_BeginFindFiles()
    'Set the Status Panels to work when the Searching routine has begun.
    StatusBar1.Panels(2).Text = "Status: Working .."
    tmrWorking.Enabled = True
End Sub
Private Sub cFind_EndFindFiles(FileCount As Long)
    'Finished, so tell the user.
    tmrWorking.Enabled = False
    cmdSearch.Enabled = True
    cmdSearch.Caption = "Search Now"
    StatusBar1.Panels(1).Text = "Found : " & lstFind.ListItems.Count & " Items"
    StatusBar1.Panels(2).Text = "Status: Idle"
End Sub
Private Sub cFind_FolderChange(sFolder As String)
    'When the Search routine moves out of one folder and into another, tell the user.
    StatusBar1.Panels(1).Text = "Searching: " & sFolder
End Sub
Private Sub cFind_FoundFile(FileName As String, FilePath As String, Size As Long, InZip As Boolean, sFileType As String, sTmpLocation As String, sFileTime As String, Cancel As Boolean)
    'Ok, found a file, so perform the following....
    On Error GoTo ErrClear
    Dim Lst                     As ListItem
    Dim lTmp                    As Long
    Dim sTmp                    As String
    Set Lst = lstFind.ListItems.Add(, , FileName)
    'Easiest method of working with List Items is set a ListItem object.
    Lst.ListSubItems.Add , , FilePath 'Add FilePath
    sTmp = StringSize(Size)
    Lst.ListSubItems.Add , , sTmp 'Add the size.
    Lst.ListSubItems.Add , , String(12 - Len(CStr(Size)), "0") & Size 'This is a hidden column we use for sorting, it is the raw byte size with added 0's at the front.
    Lst.ListSubItems.Add , , sFileType 'Add the Type
    Lst.Tag = sFileTime 'Set the file time as the tag so we can show it in the Status Bar when the user selects a list item.
    If InZip = True Then
        Call GetIcon(sTmpLocation & FileName, Lst.Index, cFind.sShell32) 'If its in a Zip file, we need the Temporary Location of the files we created to get the file icon
    Else
        Call GetIcon(FilePath & FileName, Lst.Index, cFind.sShell32) 'If not, use the file location.
    End If
    With lstFind
        .Icons = iml32
        .SmallIcons = iml16
        Lst.Icon = Lst.Index
        Lst.SmallIcon = Lst.Index
    End With
    Exit Sub
ErrClear:
    MsgBox ("cFind_FoundFile() : Error: #" & Err.Number & " : " & Err.Description & " : " & Err.Source)
    Stop
End Sub
Private Sub cmdFolder_Click()
    Dim sTmp                    As String
    sTmp = IIf(Len(Dir(txtFolder.Text, vbDirectory)) > 0, txtFolder.Text, "") 'If the text in the Folder text box is not a directory, we cannot use it as the start up for the Folder Open API or it will crash.
    sTmp = BrowseForFolder(sTmp, "Select Folder To Search...") 'Show the Open Folder dialog.
    If Len(sTmp) > 0 Then txtFolder.Text = sTmp
End Sub
Private Sub cmdOptions_Click()
    FrmOptions.Show vbModal, Me
End Sub
Private Sub cmdSearch_Click()
    Dim sTmp()                  As String
    Select Case cmdSearch.Caption 'Easy method of Starting and Stopping, detect the button caption =)
        Case Is = "Search Now" 'Start Searching
            cmdSearch.Caption = "Stop"
            lstFind.ListItems.Clear
            imgNull.ListImages.Add , , Me.Icon 'This is a small hack to remove all icons in the lstFind cache.
            lstFind.SmallIcons = imgNull
            lstFind.Icons = imgNull
            iml32.ListImages.Clear 'Now that they aren't in use, clear all the Icons from both Image Lists
            iml16.ListImages.Clear
            If Len(Dir(txtFolder.Text, vbDirectory)) > 0 And Len(txtFolder.Text) > 0 Then 'If the txtFolder is valid, then perform the search.
                If Not Right(txtFolder.Text, 1) = "\" Then txtFolder.Text = txtFolder.Text & "\" 'If we dont have a trailing slash, add one
                If Not IsNumeric(txtDays.Text) Then Call MsgBox("The Days Text Box must be a numerical value !", vbExclamation, "Error"): txtDays.SetFocus: Exit Sub 'We must have numeric day count to compare.
                If Not IsNumeric(txtMonths.Text) Then Call MsgBox("The Months Text Box must be a numerical value !", vbExclamation, "Error"): txtMonths.SetFocus: Exit Sub 'Must also have numeric month count to compare.
                If Not IsNumeric(txtSize.Text) Then Call MsgBox("The size of the file must be a numerical value !", vbExclamation, "Error"): txtSize.SetFocus: Exit Sub 'And the same for comparing size.
                With cFind
'                Stop
                    .Path = txtFolder.Text 'Set path
                    .CaseSensitive = chkCase.Value 'Case Sensitive ?
                    .CompareSize = IIf(mbSize, IIf(Label6.Caption = "MB", CDbl(txtSize.Text * 1024), CDbl(txtSize.Text)), 0)
                        'Set the Size Comparation
                        'If the label is MB then multiply the size by 1024 because there is 1024 KB's in a MB =)
                    .CompareSizeType = cmbSize.ListIndex 'More or Less than given size ?
                    .SearchHiddenFolders = chkHidden.Value
                    .SearchSubFolders = chkSubFolders.Value
                    .SearchSystemFolders = chkSystem.Value
                    .SearchZips = chkZips.Value
                    .SearchWords = txtContaining.Text
                    If mbType = True Then 'If we need to search File Types.
                        .sType = ImageCombo.SelectedItem.Text
                    Else
                        .sType = ""
                    End If
                    Call .SetCompareProps(IIf(OpLast(0).Value, eDay, eMonth), IIf(OpLast(0).Value, CInt(txtDays.Text), CInt(txtMonths.Text)), mbDate, IIf(OpDate(0).Value, eModified, eCreated)) 'This is a sub I used to set a couple of options at once. Made things easier.
                    .FileSpec = IIf(Len(txtSearch.Text) = 0, "*.*", txtSearch.Text) 'What we need to search for.
                    Call .FindAll(sTmp()) 'Start the search.
'                    Stop
                End With
            Else
                Call MsgBox("You must enter a valid Folder to search !", vbExclamation, "Error") 'No valid folder message.
            End If
'            Stop
        Case Is = "Stop"
            cmdSearch.Enabled = False 'Disable the button to let the search routine finish before someone starts it again.
            cFind.Searching = False
    End Select
End Sub
Private Sub Form_Load()
    Set cFind = New clsSearch
    Dim I                       As Integer
    For I = 0 To MenuHeader.Count - 1 'Setup the XP Menu frames.
        If FrameXPMenu(I).Height = imgUp.Height Then
            MenuHeader(I).Picture = imgDown.Picture
        Else
            MenuHeader(I).Picture = imgUp.Picture
        End If
        MenuHeader(I).Height = 375
        MenuHeader(I).Width = FrameXPMenu(I).Width
    Next
    DoResize = False 'To stop it resizing the form while we make adjustments.
    pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX 'Set the Temp Picture Box properties.
    pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY 'Set the Temp Picture Box properties.
    pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX 'Set the Temp Picture Box properties.
    pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY 'Set the Temp Picture Box properties.
    OpDate(0).Enabled = False 'Disable all the check box's while they are not visible.
    OpDate(1).Enabled = False
    OpLast(0).Enabled = False
    OpLast(1).Enabled = False
    txtDays.Enabled = False
    txtMonths.Enabled = False
    UpDownDays.Enabled = False
    UpDownMonths.Enabled = False
    cmbSize.Enabled = False
    txtSize.Enabled = False
    UpDownSize.Enabled = False
    cmbSize.ListIndex = 0
    With lstFind 'Add column headers.
        .ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_NAME), 1600 'File Name Column
        .ColumnHeaders.Add , , "Path", 2000 'File Path Column
        .ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_SIZE), 1000 'File Size Column
        .ColumnHeaders.Add , , "Byte Size", 0 'Raw Size Column for sorting only !
        .ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_TYPE), 2000 'Type Column
        .ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_MODIFIED), 3000 'Type Column
        .ColumnHeaders.Add , , cFind.GetResourceString(TEXT_RESOURCE_COL_CREATED), 3000  'Type Column
    End With
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub 'It will error if its minimized !
    If Me.Height <= 7305 Then Me.Height = 7305 'Don't let it go too small !
    If Me.Width <= 10185 Then Me.Width = 10185
    ' Now calculate the size's for resizing.
    Picture3.Height = Me.Height - 1000
    lstFind.Height = Picture3.Top + Picture3.Height - lstFind.Top
    lstFind.Width = Me.Width - 195 - lstFind.Left
    StatusBar1.Panels(1).Width = Me.Width - StatusBar1.Panels(2).Width - 400
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim vValue                  As Integer
    'Save some data to the registry for next load time.
    vValue = chkSubFolders.Value
    Call SaveSetting("Search Sub Folders", CStr(vValue))
    vValue = chkHidden.Value
    Call SaveSetting("Search Hidden Folders", CStr(vValue))
    vValue = chkSystem.Value
    Call SaveSetting("Search System Folders", CStr(vValue))
    vValue = chkZips.Value
    Call SaveSetting("Search Zip Folders", CStr(vValue))
    Call SaveSetting("Last Location", txtFolder.Text)
    vValue = lstFind.View
    Call SaveSetting("View Type", CStr(vValue))
    Set cFind = Nothing 'Unset the Search Object
    End
    'Alot of people say don't use End, this is BULLSHIT.
    'End stops the process in its tracks, and unloads all variables from memory.
    'One thing you must do before using this, is to unset all the objects you may have loaded, otherwise memory leaks will occur.
End Sub
Private Sub SaveSetting(Setting As String, Value As String)
    'Small time saving function, it just saved me typing the whole line multiple times =P
    Call REGSaveSetting(vHKEY_CURRENT_USER, "Software\EliteProdigy\Search\Settings", Setting, Value)
End Sub
Private Sub Label6_Click()
    'Switch between MB and KB size comparision.
    'I should probably call this object something meaningfull, but oh well.
    Label6.Caption = IIf(Label6.Caption = "KB", "MB", "KB")
End Sub
Private Sub lstFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'This event is fired when a user clicks on one of the column headers.
    If ColumnHeader.Index = 3 Then
        'If they click on the Size column, we need to set the SortKey to the raw size column.
        'Example:
        
        '1      MB
        '100    KB
        '2      MB
        '200    Bytes
        
        'I think thats right anyway, but thats how it would be sorted,
        'But our hidden column contains...
        
        '000002097152 ( 2 MB )
        '000001048576 ( 1 MB )
        '000000102400 ( 100 KB )
        '000000000200 ( 200 Bytes )
        'Now do you see the method to my madness ? =D
        
        lstFind.SortKey = 3
    Else
        'Otherwise set the SortKey to the column clicked.
        lstFind.SortKey = ColumnHeader.Index - 1
    End If
    lstFind.Sorted = True 'Sort it.
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending) 'SortOrder = Not SortOrder
End Sub
Private Sub lstFind_DblClick()
    Call ShellExecute(&O0, "", lstFind.SelectedItem.ListSubItems(1).Text & lstFind.SelectedItem.Text, vbNullString, vbNullString, 1) 'Shell Execute the selected Item.
End Sub
Private Sub lstFind_ItemClick(ByVal item As MSComctlLib.ListItem)
    StatusBar1.Panels(1).Text = "In Folder " & item.ListSubItems(1) & "; Type " & item.ListSubItems(4) & "; Date Modified " & item.Tag & "; Size " & item.ListSubItems(2)
    'Set the Status Bar Panel to hold the info of the current clicked item.
End Sub
Private Sub lstFind_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next 'We need this to know if an item is selected.
    If Button = 2 Then
        mnProperties.Visible = False 'Set both of the Item independant menu items to not visibile.
        mnOpenFolder.Visible = False
        If Not Len(lstFind.SelectedItem.Text) = 0 Then 'If no item is selected, this will error, but we are resuming so this should be False
            'But if by some miracle, the above statement ends up true, with no item selected, it will have errored anyway.
            'So, if no error, then we have an item.
            If Err.Number = 0 Then
                'So show the menu items needed !
                mnProperties.Visible = True
                mnOpenFolder.Visible = True
            End If
        Else
            'Otherwise reset them to false anyway.
            mnProperties.Visible = False
            mnOpenFolder.Visible = False
        End If
        PopupMenu mnRClick 'Show the popupmenu.
    End If
End Sub
Private Sub mnOpenFolder_Click()
    Dim Lst                     As MSComctlLib.ListItem
    Set Lst = lstFind.SelectedItem
    'Open the containing folder, will either be a zip file or a directory, so shell execute the first sub item.
    Call ShellExecute(&O0, "", Lst.ListSubItems(1), vbNullString, vbNullString, 1)
End Sub
Private Sub mnProperties_Click()
    Dim Frm                     As New FrmProperties
    Load Frm
    'Show the File Properties dialog window, our own of course.
    Frm.sMyFileName = lstFind.SelectedItem.Text
    Frm.sMyFilePath = lstFind.SelectedItem.ListSubItems(1).Text
    Frm.LoadData
    Frm.Show
End Sub
'The following four functions are in the menu and set the sort method.
Private Sub mnSortName_Click()
    lstFind.SortKey = 0
    lstFind.Sorted = True
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub
Private Sub mnSortPath_Click()
    lstFind.SortKey = 1
    lstFind.Sorted = True
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub
Private Sub mnSortSize_Click()
    lstFind.SortKey = 3
    lstFind.Sorted = True
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub
Private Sub mnSortType_Click()
    lstFind.SortKey = 4
    lstFind.Sorted = True
    lstFind.SortOrder = IIf(lstFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub
Private Sub mnView_Click(Index As Integer)
    'Set the View Type from the menu.
    lstFind.View = Index
End Sub
Private Sub OpLast_Click(Index As Integer)
    'Date comaparision
    Select Case True
        Case OpLast(1).Value
            'If we are comparing by months.
            txtDays.Enabled = False
            txtMonths.Enabled = True
            UpDownDays.Enabled = False
            UpDownMonths.Enabled = True
        Case OpLast(0).Value
            'If the comparision is based on days.
            txtDays.Enabled = True
            txtMonths.Enabled = False
            UpDownDays.Enabled = True
            UpDownMonths.Enabled = False
    End Select
End Sub
Private Sub tmrFrames_Timer()
    On Error Resume Next
    'Produce the moving effect on the XP Style frame arrangement.
    'I didn't create this, check the Credits for where to get a basic example of this.
    Dim FrameExpandHeight       As Integer: FrameExpandHeight = 0
    Dim I                       As Integer
    For I = 1 To FrameXPMenu.Count - 1
        FrameXPMenu(I).Top = FrameXPMenu(I - 1).Top + FrameXPMenu(I - 1).Height + 120
    Next
    If DoResize = True Then
        If Expand = False Then
            MenuHeader(Frame).Picture = imgDown.Picture
            FrameXPMenu(Frame).Height = FrameXPMenu(Frame).Height - Speed
            If FrameXPMenu(Frame).Height <= MenuHeader(Frame).Height Then DoResize = False: FrameXPMenu(Frame).Height = MenuHeader(Frame).Height
        Else
            MenuHeader(Frame).Picture = imgUp.Picture
            FrameXPMenu(Frame).Height = FrameXPMenu(Frame).Height + Speed
            For I = 0 To Me.Count - 1
                If Controls(I).Container.Name = FrameXPMenu(Frame).Name Then
                    If Controls(I).Container.Index = FrameXPMenu(Frame).Index Then
                        If Controls(I).Top + Controls(I).Height > FrameExpandHeight Then
                            FrameExpandHeight = Controls(I).Top + Controls(I).Height
                        End If
                    End If
                End If
            Next
            If FrameXPMenu(Frame).Height >= FrameExpandHeight + 120 Then DoResize = False: FrameXPMenu(Frame).Height = FrameExpandHeight + 120 'Stop the frame from resizing
        End If
    End If
End Sub
Private Sub tmrWorking_Timer()
    'This just keeps rotating the text in the second status bar panel while its working.
    'Nice effect, but not needed really, just for show.
    'Rotates between
    'Working ..
    'Working ...
    'etc etc
    Static sExtra               As String
    If Len(sExtra) = 0 Then sExtra = "."
    sExtra = sExtra & "."
    If sExtra = "............" Then sExtra = ".."
    StatusBar1.Panels(2).Text = "Status: Working " & sExtra
End Sub
Private Sub UpDownDays_DownClick()
    'Set the text box value for the day comparision
    If txtDays.Text = "1" Then Exit Sub
    txtDays.Text = CLng(txtDays.Text) - 1
    UpDownDays.Value = CLng(txtDays.Text)
End Sub
Private Sub UpDownDays_UpClick()
    'Set the text box value for the day comparision
    txtDays.Text = CLng(txtDays.Text) + 1
    UpDownDays.Value = CLng(txtDays.Text)
End Sub
Private Sub UpDownMonths_DownClick()
    'Set the text box value for the month comparision
    If txtMonths.Text = "1" Then Exit Sub
    txtMonths.Text = CLng(txtMonths.Text) - 1
    UpDownMonths.Value = CLng(txtMonths.Text)
End Sub
Private Sub UpDownMonths_UpClick()
    'Set the text box value for the month comparision
    txtMonths.Text = CLng(txtMonths.Text) + 1
    UpDownMonths.Value = CLng(txtMonths.Text)
End Sub
Private Sub lblHeader_Click(Index As Integer)
    MenuHeader_Click (Index)
End Sub
Private Sub MenuHeader_Click(Index As Integer)
    'Start the moving of the panels.
    If Index = 0 Then Exit Sub
    Dim I                       As Integer
    If DoResize = False Then
        If FrameXPMenu(Index).Height = MenuHeader(I).Height Then 'minimised
            Expand = True
        Else
            Expand = False
        End If
        'Enable and disable items on a panel depending on their state so that they cannot tab to it.
        Select Case Index
            Case 1
                mbDate = Expand
                OpDate(0).Enabled = Expand
                OpDate(1).Enabled = Expand
                OpLast(1).Enabled = Expand
                OpLast(0).Enabled = Expand
                txtDays.Enabled = OpLast(0).Value
                txtMonths.Enabled = OpLast(1).Value
                UpDownDays.Enabled = OpLast(0).Value
                UpDownMonths.Enabled = OpLast(1).Value
            Case 2
                mbType = Expand
                ImageCombo.Enabled = Expand
            Case 3
                mbSize = Expand
                cmbSize.Enabled = Expand
                txtSize.Enabled = Expand
                UpDownSize.Enabled = Expand
        End Select
        DoResize = True
        Frame = Index
    End If
End Sub
Private Sub UpDownSize_DownClick()
    'Set the Size comparison text box.
    If txtSize.Text = "1" Then Exit Sub
    txtSize.Text = CLng(txtSize.Text) - 1
    UpDownSize.Value = CLng(txtSize.Text)
End Sub
Private Sub UpDownSize_UpClick()
    txtSize.Text = CLng(txtSize.Text) + 1
    UpDownSize.Value = CLng(txtSize.Text)
End Sub

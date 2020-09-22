VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Options"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Menu Extensions"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdShortCuts 
         Caption         =   "Install ShortCuts"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmOptions.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdShortCuts_Click()
    Select Case cmdShortCuts.Tag
        Case Is = "Add"
            Call REGSaveSetting(vHKEY_CLASSES_ROOT, "Directory\shell\Start Searching &Here\command", "", """" & FixPath(App.Path) & App.EXEName & ".exe"" ""%1 /noload""")
            Call REGSaveSetting(vHKEY_CLASSES_ROOT, "Drive\shell\Start Searching &Here\command", "", """" & FixPath(App.Path) & App.EXEName & ".exe"" ""%1 /noload""")
            cmdShortCuts.Tag = "Remove"
            cmdShortCuts.Caption = "Remove ShortCuts"
        Case Is = "Remove"
            Call REGDeleteSetting(vHKEY_CLASSES_ROOT, "Directory\shell\Start Searching &Here\command", "")
            Call REGDeleteSetting(vHKEY_CLASSES_ROOT, "Drive\shell\Start Searching &Here\command", "")
            Call REGDeleteSetting(vHKEY_CLASSES_ROOT, "Directory\shell\Start Searching &Here", "")
            Call REGDeleteSetting(vHKEY_CLASSES_ROOT, "Drive\shell\Start Searching &Here", "")
            cmdShortCuts.Caption = "Install ShortCuts"
            cmdShortCuts.Tag = "Add"
    End Select
End Sub
Private Sub Form_Load()
    If Len(REGGetSetting(vHKEY_CLASSES_ROOT, "Directory\shell\Start Searching &Here\command", "")) > 0 Then
        cmdShortCuts.Caption = "Remove ShortCuts"
        cmdShortCuts.Tag = "Remove"
    Else
        cmdShortCuts.Caption = "Install ShortCuts"
        cmdShortCuts.Tag = "Add"
    End If
    Icon = FrmMain.Icon
End Sub

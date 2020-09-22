VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "EZ Open"
   ClientHeight    =   2325
   ClientLeft      =   1710
   ClientTop       =   1725
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   960
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   495
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblRaph 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "by c0ldfyr3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XP Fast Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   -360
      TabIndex        =   1
      Top             =   795
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Line L2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      X1              =   480
      X2              =   840
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      X1              =   5880
      X2              =   6240
      Y1              =   240
      Y2              =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XX1                         As Integer
Dim XX2                         As Integer
Dim YY1                         As Integer
Dim YY2                         As Integer
Dim XXX1                        As Integer
Dim XXX2                        As Integer
Dim YYY1                        As Integer
Dim YYY2                        As Integer
Dim When                        As Integer
Dim Start                       As Boolean
Dim I                           As Integer
Const sname = "XP Name Scanner"
Private Sub Form_DblClick()
    FrmBlank.Show
    Unload Me
End Sub
Private Sub Form_Load()
    MakeOntop Me
    YY1 = L.Y1
    YY2 = L.Y2
    XX1 = L.X1
    XX2 = L.X2
    YYY1 = L2.Y1
    YYY2 = L2.Y2
    XXX1 = L2.X1
    XXX2 = L2.X2
    Start = False
    I = 1
    lblAppName = ""
    lblRaph.Font = 29
End Sub
Private Sub lblAppName_DblClick()
    FrmBlank.Show
    Unload frmSplash
End Sub
Private Sub Timer2_Timer()
    On Error GoTo ErrClear
    YY2 = YY2 - 100: If YY2 = 600 Then YY2 = 0
    YY1 = YY1 + 100: If YY1 = 600 Then YY1 = 0
    XX2 = XX2 - 100: If XX2 = 0 Then XX2 = 600
    XX1 = XX1 - 100: If XX1 = 0 Then XX1 = 600
    L.X1 = XX1
    L.X2 = XX2
    L.Y1 = YY1
    L.Y2 = YY2
    Exit Sub
ErrClear:
    Err.Clear
    'Timer2.Enabled = False
    'FrmMain.Show
    'Unload Me
End Sub
Private Sub Timer3_Timer()
    On Error GoTo ErrClear
    Dim S                       As Integer
    YYY2 = YYY2 - 100: If YY2 = 0 Then YY2 = 600
    YYY1 = YYY1 + 100: If YY1 = 600 Then YY1 = 0
    XXX2 = XXX2 + 100: If XX2 = 600 Then XX2 = 0
    XXX1 = XXX1 + 100: If XX1 = 600 Then XX1 = 0
    L2.X1 = XXX1
    L2.X2 = XXX2
    L2.Y1 = YYY1
    L2.Y2 = YYY2
    If L.X1 = 3180 Then
        lblAppName.Visible = True
        Start = True
    End If
    If Start = True Then
        If L2.X1 = 6480 And L2.Y1 = 6120 Then
            FinishSplash
        ElseIf I = Len(sname) + 1 Then
            Exit Sub
        Else
            a = lblAppName
            b = Mid(sname, I, 1)
            a = a & b
            lblAppName = a
            I = I + 1
        End If
    End If
ErrClear:
    Err.Clear
    'Timer3.Enabled = False
    'FrmMain.Show
    'Unload Me
End Sub
Sub FinishSplash()
    Dim X                       As Currency
    Dim J_Wait                  As New clsWaitableTimer
    lblRaph.Visible = True
    X = 30
    Do
        DoEvents
        X = X - 0.5
        lblRaph.FontSize = X
        J_Wait.Wait (5)
        If X = 12.5 Then Shape1.Visible = True
    Loop Until X = 10.5
    J_Wait.Wait (1000)
    FrmMain.Show
    Unload Me
End Sub

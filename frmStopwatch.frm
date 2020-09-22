VERSION 5.00
Begin VB.Form frmStopwatch 
   Caption         =   "Stopwatch"
   ClientHeight    =   4440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStopwatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer0 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label lblStopWatch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1215
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   5055
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu Seperator 
         Caption         =   "-"
      End
      Begin VB.Menu mPopAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TSeconds, Seconds, Minutes As Integer
Dim Stime As String
Dim Result As Long
Private Sub cmdClear_Click()
100     TSeconds = 0
110     Seconds = 0
120     Minutes = 0
130     lblStopWatch.Caption = "00:00.0"
End Sub
Private Sub cmdStart_Click()
100     Timer0.Enabled = True
End Sub
Private Sub cmdStop_Click()
100     Timer0.Enabled = False
End Sub
Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
100 Me.Show
110 Me.Refresh
120 With nid
130 .cbSize = Len(nid)
140 .hwnd = Me.hwnd
150 .uId = vbNull
160 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
170 .uCallBackMessage = WM_MOUSEMOVE
180 .hIcon = Me.Icon
190 .szTip = "Your ToolTip" & vbNullChar
200 End With
210 Shell_NotifyIcon NIM_ADD, nid
220     Timer0.Enabled = False
230     TSeconds = 0
240     Seconds = 0
250     Minutes = 0
260     lblStopWatch.Caption = "00:00"
End Sub
Private Sub mnuAbout_Click()
100     frmAbout.Show
End Sub
Private Sub mnuExit_Click()
100     End
End Sub
Private Sub Timer0_Timer()
100     TSeconds = TSeconds + 1
110     If (TSeconds = 10) Then
120     Seconds = Seconds + 1
130     TSeconds = 0
140         If (Seconds = 60) Then
150         Minutes = Minutes + 1
160         Seconds = 0
170         TSeconds = 0
180             If (Minutes = 60) Then
190             Minutes = 0
200             Seconds = 0
210             TSeconds = 0
220             End If
230         End If
240     End If
250     lblStopWatch.Caption = Format(Minutes, "00") & ":" & Format(Seconds, "00") & "." & Format(TSeconds, "0")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim msg As Long
'the value of X will vary depending upon the scalemode setting
100 If Me.ScaleMode = vbPixels Then
110 msg = X
120 Else
130 msg = X / Screen.TwipsPerPixelX
140 End If
150 Select Case msg
    Case WM_LBUTTONUP        '514 restore form window
160 Me.WindowState = vbNormal
170 Result = SetForegroundWindow(Me.hwnd)
180 Me.Show
    Case WM_LBUTTONDBLCLK    '515 restore form window
190 Me.WindowState = vbNormal
200 Result = SetForegroundWindow(Me.hwnd)
210 Me.Show
    Case WM_RBUTTONUP        '517 display popup menu
220 Result = SetForegroundWindow(Me.hwnd)
230 Me.PopupMenu Me.mPopupSys
240 End Select
End Sub
Private Sub Form_Resize()
'this is necessary to assure that the minimized window is hidden
100 If Me.WindowState = vbMinimized Then Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
'this removes the icon from the system tray
100 Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
100 Unload Me
End Sub
Private Sub mPopRestore_Click()
'called when the user clicks the popup menu Restore command
100 Me.WindowState = vbNormal
110 Result = SetForegroundWindow(Me.hwnd)
120 Me.Show
End Sub


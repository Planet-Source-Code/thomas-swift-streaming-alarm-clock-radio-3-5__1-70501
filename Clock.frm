VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Clock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Streaming Alarm Clock Radio 3.5"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "M5"
      Height          =   330
      Left            =   8850
      TabIndex        =   11
      Top             =   1020
      Width           =   390
   End
   Begin VB.CommandButton Command6 
      Caption         =   "M4"
      Height          =   330
      Left            =   8430
      TabIndex        =   10
      Top             =   1020
      Width           =   390
   End
   Begin VB.CommandButton Command5 
      Caption         =   "M3"
      Height          =   330
      Left            =   8010
      TabIndex        =   9
      Top             =   1020
      Width           =   390
   End
   Begin VB.CommandButton Command4 
      Caption         =   "M2"
      Height          =   330
      Left            =   7590
      TabIndex        =   8
      Top             =   1020
      Width           =   390
   End
   Begin VB.CommandButton Command3 
      Caption         =   "M1"
      Height          =   330
      Left            =   7170
      TabIndex        =   7
      Top             =   1020
      Width           =   390
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   780
      Top             =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9360
      TabIndex        =   6
      ToolTipText     =   "Minimize to tray."
      Top             =   1020
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Settings"
      Height          =   330
      Left            =   6090
      TabIndex        =   5
      ToolTipText     =   "Settings"
      Top             =   1020
      Width           =   885
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Alarm Set"
      Height          =   255
      Left            =   5025
      TabIndex        =   4
      Top             =   1058
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   300
      Top             =   2040
   End
   Begin VB.Label AlarmTimeDsp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8145
      TabIndex        =   3
      ToolTipText     =   "Alarm Time"
      Top             =   525
      Width           =   1560
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   930
      Left            =   5025
      TabIndex        =   2
      Top             =   60
      Width           =   4740
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8361
      _cy             =   1640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "TAS Independent Programming"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   4905
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   930
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Time. You can change between 12 or 24 hour display in settings."
      Top             =   60
      Width           =   4905
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "mnuSysTray"
      Visible         =   0   'False
      Begin VB.Menu MnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu MnuVisualize 
         Caption         =   "Visualize"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Military As Boolean
Public AlarmTime As Date
Public AlarmStream As String
Public AlarmFile As String
Public OgVol As Integer
Private MyCap As String
Private OldMem As String
Private Sub Command2_Click()
    Me.Hide
    Unload Settings
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    WindowsMediaPlayer1.SetFocus
Preselect 1, True
End Sub
Public Function Preselect(Memory As Integer, Optional Playing As Boolean)
Dim MemStr As String
'Dim Playing As Boolean
If WindowsMediaPlayer1.playState = wmppsPlaying Then Playing = True
SaveSetting "Streaming Radio Alarm Clock", "Settings", "M" & OldMem & "V", WindowsMediaPlayer1.Settings.volume
Me.Caption = MyCap & " - M" & Memory
Select Case Memory
         Case 1
            MemStr = GetSetting("Streaming Radio Alarm Clock", "Settings", "M1", vbNullString)
         Case 2:
            MemStr = GetSetting("Streaming Radio Alarm Clock", "Settings", "M2", vbNullString)
         Case 3:
            MemStr = GetSetting("Streaming Radio Alarm Clock", "Settings", "M3", vbNullString)
         Case 4:
            MemStr = GetSetting("Streaming Radio Alarm Clock", "Settings", "M4", vbNullString)
         Case 5:
            MemStr = GetSetting("Streaming Radio Alarm Clock", "Settings", "M5", vbNullString)
End Select
SaveSetting "Streaming Radio Alarm Clock", "Settings", "Memory", Memory
WindowsMediaPlayer1.Settings.volume = GetSetting("Streaming Radio Alarm Clock", "Settings", "M" & Memory & "V", "100")
If MemStr <> vbNullString Then
AlarmStream = MemStr
WindowsMediaPlayer1.URL = AlarmStream
If Playing = True Then WindowsMediaPlayer1.Controls.play
End If
OldMem = Memory
End Function
Private Sub Command4_Click()
On Error Resume Next
    WindowsMediaPlayer1.SetFocus
Preselect 2, True
End Sub
Private Sub Command5_Click()
On Error Resume Next
    WindowsMediaPlayer1.SetFocus
Preselect 3, True
End Sub
Private Sub Command6_Click()
On Error Resume Next
    WindowsMediaPlayer1.SetFocus
Preselect 4, True
End Sub
Private Sub Command7_Click()
On Error Resume Next
    WindowsMediaPlayer1.SetFocus
Preselect 5, True
End Sub
Private Sub MnuVisualize_Click()
    If WindowsMediaPlayer1.playState = wmppsPlaying Then
        WindowsMediaPlayer1.fullScreen = True
    Else
        MsgBox "Media must be playing before you can use this feature !"
    End If
End Sub
Private Sub AlarmTimeDsp_Click()
    Command1_Click
End Sub
Private Sub Check1_Click()
    On Error Resume Next
    WindowsMediaPlayer1.SetFocus
    If Check1.Value = 1 And AlarmFile = vbNullString Then MsgBox "Please set the backup alarm file.": Settings.Show: Timer2.Enabled = True: Check1.Value = 0
    If Check1.Value = 1 And AlarmStream = vbNullString Then MsgBox "Please set at least one stream.": Settings.Show: Timer2.Enabled = True: Check1.Value = 0
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "AlarmSet", Check1.Value
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    WindowsMediaPlayer1.SetFocus
    Settings.Show
    Timer2.Enabled = True
    If (Me.Top + Me.Height) + Settings.Height > Screen.Height Then Me.Top = Screen.Height - ((Me.Height + Settings.Height) + 1000)
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        Unload Settings
        Unload Me
    End If
    Me.Hide
    MyCap = Me.Caption
    AlarmFile = GetSetting("Streaming Radio Alarm Clock", "Settings", "AlarmFile", vbNullString)
    WindowsMediaPlayer1.Settings.autoStart = False 'This here is a very important ingrediant
    WindowsMediaPlayer1.Settings.volume = 100
    Preselect GetSetting("Streaming Radio Alarm Clock", "Settings", "Memory", "1")
    AlarmTime = CDate(GetSetting("Streaming Radio Alarm Clock", "Settings", "AlarmTime", "6:30:00 AM"))
    Check1.Value = GetSetting("Streaming Radio Alarm Clock", "Settings", "AlarmSet", Check1.Value)
    AddToTray Me, "Streaming Radio Alarm Clock", Me.Icon
    btnFlat Command1
    btnFlat Command2
    btnFlat Command3
    btnFlat Command4
    btnFlat Command5
    btnFlat Command6
    btnFlat Command7
    OgVol = GetMasterVolume_Value
    AlarmTimeDsp.Caption = Format(AlarmTime, "h:mm AMPM")
    Clock
    If AlarmFile = vbNullString Then MsgBox "Please set the backup alarm file.": Me.Show: Command1_Click: Exit Sub
    If AlarmStream = vbNullString Then MsgBox "Please set at least one stream.": Me.Show: Command1_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RemoveFromTray
    Unload Settings
    Unload Me
End Sub
Private Sub MnuExit_Click()
    Unload Me
End Sub
Private Sub MnuShow_Click()
    Me.Show
End Sub
Private Sub Timer1_Timer()
    If Hour(Now) = Hour(AlarmTime) And Minute(Now) = Minute(DateAdd("n", -1, AlarmTime)) And Second(Now) = "0" And Check1.Value = 1 And WindowsMediaPlayer1.playState = wmppsPlaying Then WindowsMediaPlayer1.Controls.stop 'Makes sure mediaplayer is ready to alarm. Just in case someone goes to sleep with the radio going.
    If Hour(Now) = Hour(AlarmTime) And Minute(Now) = Minute(AlarmTime) And Second(Now) = "0" And Check1.Value = 1 Then
        Me.Show
        PlayItLoud
    End If
    Clock
End Sub
Private Sub Clock()
    If Military = True Then
        Label1.Caption = Format(Now, "h:mm:ss")
    Else
        Label1.Caption = Format(Now, "h:mm:ss AMPM")
    End If
    Label2.Caption = Format(Now, "dddd, mmm d yyyy")
End Sub
Private Sub Timer2_Timer() 'Settings docking timer
    Settings.Top = Me.Top + Me.Height
    Settings.Left = (Me.Left + Me.Width) - Settings.Width
End Sub
Private Sub WindowsMediaPlayer1_MediaError(ByVal pMediaObject As Object)
    WindowsMediaPlayer1.URL = AlarmFile
    WindowsMediaPlayer1.Controls.play
    MsgBox "Couldn't connect to internet !"
End Sub
Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)
    Select Case NewState
    Case wmppsMediaEnded
        WindowsMediaPlayer1.Controls.stop
    Case wmppsReady
        WindowsMediaPlayer1.Controls.stop
    Case wmppsStopped
        SetMasterVolume_Value OgVol
        WindowsMediaPlayer1.URL = AlarmStream
    End Select
End Sub
Public Sub PlayItLoud()
    Dim eType As eConnectionType
    Dim sName As String
    If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsUndefined Or WindowsMediaPlayer1.playState = wmppsReady Then
        If GetSetting("Streaming Radio Alarm Clock", "Settings", "RadioStream", "1") = 0 Or InternetConnected(sName, eType) = False Or AlarmStream = vbNullString Then
            If AlarmFile = vbNullString Then MsgBox "Please choose a alarm sound file !": Exit Sub
            OgVol = GetMasterVolume_Value
            MasterVolume_Mute True, False
            SetMasterVolume_Value "90"
            WindowsMediaPlayer1.URL = AlarmFile
        Else
            OgVol = GetMasterVolume_Value
            MasterVolume_Mute True, False
            SetMasterVolume_Value "90"
            WindowsMediaPlayer1.URL = AlarmStream
        End If
        WindowsMediaPlayer1.Controls.play
    Else
        WindowsMediaPlayer1.Controls.stop
        SetMasterVolume_Value OgVol
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP
        SetForegroundWindow Me.hwnd
        PopupMenu mnuSysTray
    Case WM_LBUTTONUP
        Me.Show
    End Select
End Sub

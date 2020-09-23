VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   4995
      TabIndex        =   20
      Top             =   435
      Width           =   3825
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Test"
      Height          =   270
      Left            =   8835
      TabIndex        =   19
      Top             =   1950
      Width           =   525
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Test"
      Height          =   270
      Left            =   8835
      TabIndex        =   18
      Top             =   1575
      Width           =   525
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Test"
      Height          =   270
      Left            =   8835
      TabIndex        =   17
      Top             =   1200
      Width           =   525
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test"
      Height          =   270
      Left            =   8835
      TabIndex        =   16
      Top             =   825
      Width           =   525
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   4995
      TabIndex        =   15
      Top             =   1935
      Width           =   3825
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   4995
      TabIndex        =   14
      Top             =   1560
      Width           =   3825
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   4995
      TabIndex        =   13
      Top             =   1185
      Width           =   3825
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   4995
      TabIndex        =   12
      Top             =   810
      Width           =   3825
   End
   Begin VB.CheckBox Check1 
      Caption         =   "24 Hour Display"
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose Alarm File"
      Height          =   330
      Left            =   158
      TabIndex        =   5
      Top             =   525
      Width           =   1785
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test"
      Height          =   315
      Left            =   2048
      TabIndex        =   4
      ToolTipText     =   "You don't need to set your volume high for alarm. It will automatically be set to 90% when it goes off."
      Top             =   540
      Width           =   1020
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Radio Stream"
      Height          =   225
      Left            =   3158
      TabIndex        =   3
      Top             =   585
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test"
      Height          =   270
      Left            =   8835
      TabIndex        =   2
      ToolTipText     =   "You don't need to set your volume high for alarm. It will automatically be set to 90% when it goes off."
      Top             =   450
      Width           =   525
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   300
      Left            =   8970
      TabIndex        =   1
      Top             =   90
      Width           =   255
   End
   Begin MSComDlg.CommonDialog ComD1 
      Left            =   2070
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   360
      Left            =   60
      TabIndex        =   7
      Top             =   1170
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   635
      _Version        =   393216
      Max             =   23
      SelStart        =   6
      Value           =   6
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   360
      Left            =   60
      TabIndex        =   8
      Top             =   1905
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   635
      _Version        =   393216
      Max             =   59
      SelStart        =   30
      Value           =   30
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M5"
      Height          =   210
      Index           =   4
      Left            =   4710
      TabIndex        =   25
      Top             =   1980
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M4"
      Height          =   210
      Index           =   3
      Left            =   4710
      TabIndex        =   24
      Top             =   1605
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M3"
      Height          =   210
      Index           =   2
      Left            =   4710
      TabIndex        =   23
      Top             =   1230
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M2"
      Height          =   210
      Index           =   1
      Left            =   4710
      TabIndex        =   22
      Top             =   855
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M1"
      Height          =   210
      Index           =   0
      Left            =   4710
      TabIndex        =   21
      Top             =   480
      Width           =   255
   End
   Begin VB.Label AlarmTimeDsp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   375
      Left            =   2445
      TabIndex        =   11
      ToolTipText     =   "Alarm Time"
      Top             =   30
      Width           =   1740
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hours"
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
      Left            =   1710
      TabIndex        =   10
      Top             =   990
      Width           =   1290
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Minutes"
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
      Left            =   1710
      TabIndex        =   9
      Top             =   1710
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Stream URL's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6465
      TabIndex        =   0
      Top             =   165
      Width           =   1290
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
    On Error Resume Next
    Text1.SetFocus
    Clock.Military = Check1.Value
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "Military", Check1.Value
End Sub
Private Sub Check2_Click()
    On Error Resume Next
    Text1.SetFocus
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "RadioStream", Check2.Value
End Sub
Private Sub Command1_Click()
    On Error GoTo Error
    With ComD1
        .InitDir = GetSetting("Streaming Radio Alarm Clock", "Settings", "InitDir", "C:\")
        .Flags = &H1
        .Flags = &H2
        '.DefaultExt = "wav"
        .Filter = "Audio Files (*.wav)(*.mp3)|*.wav;*.mp3"
        .DialogTitle = "Open Alarm Sound"
        .Flags = &H4
        .Flags = &H1000
        .ShowOpen
    End With
    MousePointer = vbHourglass
    If ComD1.FileName = vbNullString Then
        MousePointer = vbDefault
        Exit Sub
    End If
    Clock.AlarmFile = ComD1.FileName
    
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "AlarmFile", Clock.AlarmFile
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "InitDir", GetFilePath(Clock.AlarmFile)
    MousePointer = vbDefault
    Debug.Print Clock.AlarmFile
    Exit Sub
Error:
    MousePointer = vbDefault
    If Err.Number = 32755 Then Exit Sub
    MsgBox "Error loading Sound File. " & vbNewLine & _
            "ERROR #" & Err.Number & " - " & Error$(Err.Number), vbCritical, "File Load Error"
    Close #1
    Exit Sub
    
End Sub
Public Property Get GetFilePath(FileNamePath As String) As String
On Error GoTo FunctionError:
Dim X As Long
Dim tString As String
GetFilePath = vbNullString
For X = Len(FileNamePath) To 0 Step -1
    tString = Mid$(FileNamePath, X, 1)
    If tString = "\" Then
        GetFilePath = Left(FileNamePath, X)
        Exit Property
    End If
Next X
FunctionError:
GetFilePath = -1
End Property
Public Property Get GetFileName(file As String) As String
Dim m As Long
Dim GetChr0 As String
Dim GetChr1 As String
GetFileName = vbNullString
For m = 1 To Len(file)
    GetChr0 = Right(file, m)
    GetChr1 = Left(GetChr0, 1)
    If GetChr1 = "\" Or GetChr1 = "/" Then
        GetFileName = Right(GetChr0, m - 1): Exit Property
    End If
Next m
End Property
Public Sub Command2_Click()
    On Error Resume Next
    Text1.SetFocus
    If Clock.AlarmFile = vbNullString Then MsgBox "Please choose a alarm sound file !": Exit Sub
    If Clock.WindowsMediaPlayer1.playState = wmppsStopped Or Clock.WindowsMediaPlayer1.playState = wmppsUndefined Or Clock.WindowsMediaPlayer1.playState = wmppsReady Then
        Clock.WindowsMediaPlayer1.URL = Clock.AlarmFile
        Clock.WindowsMediaPlayer1.Controls.play
    Else
        Clock.WindowsMediaPlayer1.Controls.stop
    End If
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    Text1.SetFocus
    ShellExecute 0&, vbNullString, App.Path & "\Complimentary Stream Links.txt", vbNullString, vbNullString, 1
    'MsgBox "You will find a list of streams in this programs shortcut folder on your Start menu."
End Sub
Private Sub Command4_Click()
    On Error Resume Next
    Text1.SetFocus
'Clock.AlarmStream = Text1.Text
TestAlarm 1
End Sub
Private Function TestAlarm(Mem As Integer)
    If Clock.AlarmStream = vbNullString Then MsgBox "Please write or paste a stream URL in the box to the left !": Exit Function
    If Clock.WindowsMediaPlayer1.playState = wmppsStopped Or Clock.WindowsMediaPlayer1.playState = wmppsUndefined Or Clock.WindowsMediaPlayer1.playState = wmppsReady Then
        SavMems
        Call Clock.Preselect(Mem, True)
        'Clock.WindowsMediaPlayer1.Controls.play
    Else
        Clock.WindowsMediaPlayer1.Controls.stop
    End If
End Function
Private Sub Command5_Click()
    On Error Resume Next
    Text2.SetFocus
'Clock.AlarmStream = Text2.Text
TestAlarm 2
End Sub
Private Sub Command6_Click()
    On Error Resume Next
    Text3.SetFocus
'Clock.AlarmStream = Text3.Text
TestAlarm 3
End Sub
Private Sub Command7_Click()
    On Error Resume Next
    Text4.SetFocus
'Clock.AlarmStream = Text4.Text
TestAlarm 4
End Sub
Private Sub Command8_Click()
    On Error Resume Next
    Text5.SetFocus
'Clock.AlarmStream = Text5.Text
TestAlarm 5
End Sub
Private Sub Form_Load()
    Dim ATime As Date
    btnFlat Command1
    btnFlat Command2
    btnFlat Command3
    btnFlat Command4
    btnFlat Command5
    btnFlat Command6
    btnFlat Command7
    btnFlat Command8
    Check1.Value = GetSetting("Streaming Radio Alarm Clock", "Settings", "Military", Check1.Value)
    Check2.Value = GetSetting("Streaming Radio Alarm Clock", "Settings", "RadioStream", Check2.Value)
    Text1.Text = GetSetting("Streaming Radio Alarm Clock", "Settings", "M1", vbNullString)
    Text2.Text = GetSetting("Streaming Radio Alarm Clock", "Settings", "M2", vbNullString)
    Text3.Text = GetSetting("Streaming Radio Alarm Clock", "Settings", "M3", vbNullString)
    Text4.Text = GetSetting("Streaming Radio Alarm Clock", "Settings", "M4", vbNullString)
    Text5.Text = GetSetting("Streaming Radio Alarm Clock", "Settings", "M5", vbNullString)
    Slider1.Value = Hour(Clock.AlarmTime)
    Slider2.Value = Minute(Clock.AlarmTime)
    ShowATime
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "AlarmTime", CStr(Clock.AlarmTime)
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "AlarmStream", Clock.AlarmStream
    SavMems
    Clock.Timer2.Enabled = False
End Sub
Private Sub SavMems()
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "M1", Text1.Text
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "M2", Text2.Text
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "M3", Text3.Text
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "M4", Text4.Text
    SaveSetting "Streaming Radio Alarm Clock", "Settings", "M5", Text5.Text
End Sub
Public Sub ShowATime()
    AlarmTimeDsp.Caption = Format(Clock.AlarmTime, "h:mm AMPM")
    Clock.AlarmTimeDsp.Caption = Format(Clock.AlarmTime, "h:mm AMPM")
End Sub
Public Sub SetATime()
    Clock.AlarmTime = DateAdd("h", Slider1.Value, 0)
    Clock.AlarmTime = DateAdd("n", Slider2.Value, Clock.AlarmTime)
    ShowATime
End Sub
Private Sub Slider1_Scroll()
    SetATime
End Sub
Private Sub Slider2_Scroll()
    SetATime
End Sub


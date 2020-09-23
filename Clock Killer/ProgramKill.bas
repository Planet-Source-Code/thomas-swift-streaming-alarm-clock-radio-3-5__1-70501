Attribute VB_Name = "KillProgram"
Option Explicit

'***************************************************************************************
'   API Declares
'***************************************************************************************
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
'***************************************************************************************
'   Types Used to Retrieve Information From Windows
'***************************************************************************************
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&
'Used to determine what OS Version
Private Const WINNT As Integer = 2
Private Const WIN98 As Integer = 1
'Used to return error code
Public KillAppReturn As Boolean
Public Function getVersion() As Integer
  Dim udtOSInfo As OSVERSIONINFO
  Dim intRetVal As Integer
         
  'Initialize the type's buffer sizes
    With udtOSInfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
    End With
    
  'Make an API Call to Retrieve the OSVersion info
    intRetVal = GetVersionExA(udtOSInfo)
  
  'Set the return value
    getVersion = udtOSInfo.dwPlatformId
End Function
Public Function Killapp(myName As String)
Select Case getVersion()
Case WIN98 'Windows 95/98
Killapp9X (myName)
Case WINNT 'Windows NT
KillappNT (myName)
End Select
End Function
Private Function KillappNT(myName As String)
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim I As Integer
    On Local Error GoTo Finish
    appCount = 0
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        I = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillAppReturn = True
            appCount = appCount + 1
            myProcess = OpenProcess(1&, -1&, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, 0&)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
KillAppReturn = False
End Function
Private Function Killapp9X(myName As String)
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim I As Integer
    On Local Error GoTo Finish
    appCount = 0
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        I = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillAppReturn = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
KillAppReturn = False
End Function

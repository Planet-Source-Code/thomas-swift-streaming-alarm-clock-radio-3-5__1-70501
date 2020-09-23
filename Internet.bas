Attribute VB_Name = "Internet"
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Boolean
Public Enum eConnectionType
    INTERNET_CONNECTION_MODEM = &H1&
    INTERNET_CONNECTION_LAN = &H2&
    INTERNET_CONNECTION_PROXY = &H4&
    INTERNET_RAS_INSTALLED = &H10&
    INTERNET_CONNECTION_OFFLINE = &H20&
    INTERNET_CONNECTION_CONFIGURED = &H40&
End Enum
'Purpose   :    Determines basic information regarding the local machine's internet connection state.
'Inputs    :    sConnectionName                 See outputs.
'               eType                           See outputs.
'Inputs    :    Returns True if connected +
'               sConnectionName                 The name of the internet connection
'               eType                           The type of internet connection (see eConnectionType)
'Author    :    Andrew Baker
'Date      :    25/03/2001
'Notes     :
Public Function InternetConnected(Optional ByRef sConnectionName As String, Optional ByRef eType As eConnectionType) As Boolean
    Dim sConNameBuffer As String * 513
    Dim lPos As Long
    'Clear output values
    sConnectionName = vbNullString
    eType = 0
    'Call API
    InternetConnected = InternetGetConnectedStateEx(eType, sConNameBuffer, 512, 0&)
    'Clean up output
    lPos = InStr(sConNameBuffer, vbNullChar)
    If lPos > 0 Then
        sConnectionName = Left$(sConNameBuffer, lPos - 1)
    End If
End Function


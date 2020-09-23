Attribute VB_Name = "Mod_SoundMixer"
Option Explicit
''
''  Function GetMasterVolume_Value() As Integer
''  Function SetMasterVolume_Value(GetPercentVal As Integer)
''      (Function MasterVolume_Mute(SetFlag As Boolean, Optional Mute As Boolean) As Boolean)
''      (Function MasterVolume_Value(SetFlag As Boolean, Optional NewVol As Long) As MasterVolumeType)
''
Public Type MasterVolumeType
    Min As Long
    Max As Long
    CurVal As Long
    Mute As Boolean
End Type
''
''============================ WINMM.DLL
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
''
''
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_OBJECTF_MIXER = &H0&
Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
''
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
''
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
''
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
''
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
''
''
''
Private Type MIXERCAPS
    wMid As Integer '  manufacturer id
    wPid As Integer '  product id
    vDriverVersion As Long '  version of the driver
    szPname As String * MAXPNAMELEN '  product name
    fdwSupport As Long '  misc. support bits
    cDestinations As Long '  count of destinations
End Type
''
Private Type MIXERCONTROL
    cbStruct As Long '  size in Byte of MIXERCONTROL
    dwControlID As Long '  unique control id for mixer device
    dwControlType As Long '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte ' short name of Control
    szName(1 To MIXER_LONG_NAME_CHARS) As Byte ' long name of control
    lMinimum As Long '  Minimum value
    lMaximum As Long '  Maximum value
    reserved(10) As Long '  reserved structure space
End Type
''
Private Type MIXERCONTROLDETAILS
    cbStruct As Long '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long '  control id to get/set details on
    cChannels As Long '  number of channels in paDetails array
    item As Long '  hwndOwner or cMultipleItems
    cbDetails As Long '  size of _one_ details_XX struct
    paDetails As Long '  pointer to array of details_XX structs
End Type
''
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long '  value of the control
End Type
Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type
''
Private Type MIXERLINE
    cbStruct As Long '  size of MIXERLINE structure
    dwDestination As Long '  zero based destination index
    dwSource As Long '  zero based source index (if source)
    dwLineID As Long '  unique line id for mixer device
    fdwLine As Long '  state/information about line
    dwUser As Long '  driver specific information
    dwComponentType As Long '  component type line connects to
    cChannels As Long '  number of channels line supports
    cConnections As Long '  number of connections (possible)
    cControls As Long '  number of controls at this line
    szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
    szName(1 To MIXER_LONG_NAME_CHARS) As Byte
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
''
Private Type MIXERLINECONTROLS
    cbStruct As Long '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long '  line id (from MIXERLINE.dwLineID)
    dwControl As Long '  MIXER_GETLINECONTROLSF_ONEBYID or MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long '  count of controls pmxctrl points to
    cbmxctrl As Long '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long '  pointer to first MIXERCONTROL array
End Type
''
''
''
Public Function GetMasterVolume_Value() As Integer
    Dim MVolType As MasterVolumeType
    MVolType = MasterVolume_Value2(False)
    GetMasterVolume_Value = (MVolType.CurVal * 100) / MVolType.Max '.Min
End Function
''
Public Function SetMasterVolume_Value(GetPercentVal As Integer)
    Dim MVolType As MasterVolumeType
    Dim CurVal As Long
    If GetPercentVal > 100 Then GetPercentVal = 100
    If GetPercentVal < 0 Then GetPercentVal = 0
    ''
    MVolType = MasterVolume_Value2(False)
    CurVal = (MVolType.Max * GetPercentVal) / 100
    Call MasterVolume_Value2(True, CurVal)
End Function

'
''##############################################################################################
Private Function MasterVolume_Value2(SetVolume As Boolean, Optional NewVol As Long) As MasterVolumeType
    '' SetFlag is used to set Master volume or get current Master volume
    Dim hMixer As Long
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim mxc As MIXERCONTROL
    Dim mxcd As MIXERCONTROLDETAILS
    Dim mxcdu As MIXERCONTROLDETAILS_UNSIGNED
    ''
    If mixerOpen(hMixer, 0, 0, 0, MIXER_OBJECTF_MIXER) <> MMSYSERR_NOERROR Then MsgBox "mixerOpen ERROR"
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
    ''
    If mixerGetLineInfo(hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE) <> MMSYSERR_NOERROR Then MsgBox "mixerGetLineInfo ERROR"
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    mxlc.cControls = 1
    mxlc.cbmxctrl = Len(mxc)
    mxlc.pamxctrl = VarPtr(mxc)
    mxc.cbStruct = Len(mxc)
    ''
    If mixerGetLineControls(hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE) <> MMSYSERR_NOERROR Then MsgBox "mixerGetLineControls ERROR"
    '' Get:
    MasterVolume_Value2.Max = mxc.lMaximum '<###############
    MasterVolume_Value2.Min = mxc.lMinimum '<###############
    ''
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.item = 0
    mxcd.cbDetails = Len(mxcdu)
    mxcd.paDetails = VarPtr(mxcdu)
    mxcd.cChannels = 1
    '' Set/Get:
    If SetVolume = True Then
        mxcdu.dwValue = NewVol '<###############
        If mixerSetControlDetails(hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE) <> MMSYSERR_NOERROR Then MsgBox "0 mixerSetControlDetails ERROR"
        MasterVolume_Value2.CurVal = NewVol
    Else
        If mixerGetControlDetails(hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE) <> MMSYSERR_NOERROR Then MsgBox "1 mixerSetControlDetails ERROR"
        MasterVolume_Value2.CurVal = mxcdu.dwValue '<###############
    End If
    ''
    If mixerClose(hMixer) <> MMSYSERR_NOERROR Then MsgBox "mixerClose ERROR"
End Function
''
''
''
''##############################################################################################
Public Function MasterVolume_Mute(SetMute As Boolean, Optional MuteFlag As Boolean) As Boolean
    '' SetFlag is used to set mute on/off Master volume on get current mute status
    Dim hMixer As Long
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim mxc As MIXERCONTROL
    Dim mxcd As MIXERCONTROLDETAILS
    Dim mxcdb As MIXERCONTROLDETAILS_BOOLEAN
    ''
    If mixerOpen(hMixer, 0, 0, 0, MIXER_OBJECTF_MIXER) <> MMSYSERR_NOERROR Then MsgBox "mixerOpen ERROR"
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
    ''
    If mixerGetLineInfo(hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE) <> MMSYSERR_NOERROR Then MsgBox "mixerGetLineInfo ERROR"
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.dwControl = MIXERCONTROL_CONTROLTYPE_MUTE
    mxlc.cControls = 1
    mxlc.cbmxctrl = Len(mxc)
    mxlc.pamxctrl = VarPtr(mxc)
    mxc.cbStruct = Len(mxc)
    ''
    If mixerGetLineControls(hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE) <> MMSYSERR_NOERROR Then MsgBox "mixerGetLineControls ERROR"
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.item = 0
    mxcd.cbDetails = Len(mxcdb)
    mxcd.paDetails = VarPtr(mxcdb)
    mxcd.cChannels = 1
    ''
    '' Set/Get:
    If SetMute = True Then
        mxcdb.fValue = MuteFlag '<###############
        If mixerSetControlDetails(hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE) <> MMSYSERR_NOERROR Then MsgBox "0 mixerGetControlDetails ERROR"
        MasterVolume_Mute = MuteFlag
    Else
        If mixerGetControlDetails(hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE) <> MMSYSERR_NOERROR Then MsgBox "1 mixerGetControlDetails ERROR"
        MasterVolume_Mute = mxcdb.fValue '<###############
    End If
    ''
    If mixerClose(hMixer) <> MMSYSERR_NOERROR Then MsgBox "mixerClose ERROR"
End Function


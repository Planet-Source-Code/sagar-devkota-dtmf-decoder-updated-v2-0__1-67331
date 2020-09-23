VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveInRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetModuleHandleA Lib "kernel32" ( _
    ByVal lpModuleName As String _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Private Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" ( _
    ByVal lpfn As Long _
) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As Long, _
    lpdwProcessId As Long _
) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" ( _
) As Long

Private Declare Function IsWindow Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    ByVal lpParam As Long _
) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Private Declare Sub ZeroMem Lib "kernel32.dll" _
Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal cBytes As Long _
)

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long _
) As Long

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As GMEMFlags, _
    ByVal dwBytes As Long _
) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal cBytes As Long _
)

Private Declare Function waveInAddBuffer Lib "winmm" ( _
    ByVal hwi As Long, _
    pwh As Any, _
    ByVal cbwh As Long _
) As Long

Private Declare Function waveInClose Lib "winmm" ( _
    ByVal hwi As Long _
) As Long

Private Declare Function waveInGetDevCaps Lib "winmm" _
Alias "waveInGetDevCapsA" ( _
    ByVal hwi As Long, _
    pwic As Any, _
    ByVal cbwic As Long _
) As Long

Private Declare Function waveInGetErrorText Lib "winmm" _
Alias "waveInGetErrorTextA" ( _
    ByVal mmrError As Long, _
    ByVal pszText As String, _
    ByVal cchText As Long _
) As Long

Private Declare Function waveInGetID Lib "winmm" ( _
    ByVal hwi As Long, _
    puDeviceID As Long _
) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function waveInGetPosition Lib "winmm" ( _
    ByVal hwi As Long, _
    pmmt As Any, _
    ByVal cbmmt As Long _
) As Long

Private Declare Function waveInMessage Lib "winmm" ( _
    ByVal deviceid As Long, _
    ByVal uMsg As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long _
) As Long

Private Declare Function waveInOpen Lib "winmm" ( _
    phwi As Long, _
    ByVal uDeviceID As Long, _
    pwfx As Any, _
    ByVal dwCallback As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function waveInPrepareHeader Lib "winmm" ( _
    ByVal hwi As Long, _
    pwh As Any, _
    ByVal cbwh As Long _
) As Long

Private Declare Function waveInUnprepareHeader Lib "winmm.dll" ( _
    ByVal hWaveIn As Long, _
    lpWaveInHdr As Any, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveInReset Lib "winmm" ( _
    ByVal hwi As Long _
) As Long

Private Declare Function waveInStart Lib "winmm" ( _
    ByVal hwi As Long _
) As Long

Private Declare Function waveInStop Lib "winmm" ( _
    ByVal hwi As Long _
) As Long

Private Declare Function mixerClose Lib "winmm" ( _
    ByVal hmx As Long _
) As Long

Private Declare Function mixerOpen Lib "winmm" ( _
    phmx As Long, _
    ByVal uMxId As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function mixerGetControlDetails Lib "winmm" _
Alias "mixerGetControlDetailsA" ( _
    ByVal hmxobj As Long, _
    pmxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long _
) As Long

Private Declare Function mixerGetLineControls Lib "winmm" _
Alias "mixerGetLineControlsA" ( _
    ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long _
) As Long

Private Declare Function mixerGetLineInfo Lib "winmm" _
Alias "mixerGetLineInfoA" ( _
    ByVal hmxobj As Long, _
    pmxl As MIXERLINE, _
    ByVal fdwInfo As Long _
) As Long

Private Declare Function mixerSetControlDetails Lib "winmm" ( _
    ByVal hmxobj As Long, _
    pmxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long _
) As Long

Private Declare Function mixerGetDevCaps Lib "winmm" _
Alias "mixerGetDevCapsA" ( _
    ByVal uMxId As Long, _
    pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long _
) As Long

' of course there are a lot more line ids
Public Enum MIXER_RECORDING_LINES
    MIXERLINE_ANALOG = &H100A&
    MIXERLINE_AUXILIARY = &H1009&
    MIXERLINE_COMPACTDISC = &H1005&
    MIXERLINE_DIGITAL = &H1001&
    MIXERLINE_LINE = &H1002&
    MIXERLINE_MICROPHONE = &H1003&
    MIXERLINE_PCSPEAKER = &H1007&
    MIXERLINE_SYNTHESIZER = &H1004&
    MIXERLINE_TELEPHONE = &H1006&
    MIXERLINE_UNDEFINED = &H1000&
    MIXERLINE_WAVEOUT = &H1008&
End Enum

Private Enum GMEMFlags
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_ZEROINIT = &H40
End Enum

' cSelfSub stuff
Private Enum eMsgWhen                               ' When to callback
    MSG_BEFORE = 1                                  ' Callback before the original WndProc
    MSG_AFTER = 2                                   ' Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER      ' Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1            ' All messages callback
Private Const MSG_ENTRIES   As Long = 32            ' Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38          ' Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4            ' SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1             ' Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2             ' Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9             ' Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11            ' Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12            ' Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13            ' Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                 ' Thunk base address
Private z_Sc(64)            As Long                 ' Thunk machine-code initialised here
Private z_Funk              As Collection           ' hWnd/thunk-address collection
' end cSelfSub stuff

Private Const WAVE_FORMAT_PCM                           As Long = 1&

Private Const MAXPNAMELEN                               As Long = 32&

Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT         As Long = &H1&
Private Const MIXER_GETCONTROLDETAILSF_VALUE            As Long = &H0&

Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE          As Long = &H2&
    
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE          As Long = &H3&
Private Const MIXER_GETLINEINFOF_LINEID                 As Long = &H2&
Private Const MIXER_GETLINEINFOF_SOURCE                 As Long = &H1&
Private Const MIXER_GETLINEINFOF_DESTINATION            As Long = &H0&

Private Const MIXER_LONG_NAME_CHARS                     As Long = 64
Private Const MIXER_SHORT_NAME_CHARS                    As Long = 16

Private Const MIXERCONTROL_CONTROLTYPE_VOLUME           As Long = &H50030001
Private Const MIXERCONTROL_CONTROLTYPE_MUTE             As Long = &H20010002

Private Const MIXER_SETCONTROLDETAILSF_VALUE            As Long = &H0&

Private Const MIXER_OBJECTF_WAVEIN                      As Long = &H20000000

Private Const MIXERCONTROL_CT_UNITS_BOOLEAN             As Long = &H10000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE          As Long = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE            As Long = 0&
Private Const MIXERCONTROL_CT_CLASS_LIST                As Long = &H70000000

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT   As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT     As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MIXER            As Long = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MUX              As Long = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)

Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN        As Long = &H7&

Private Const CALLBACK_WINDOW                           As Long = &H10000

Private Const MM_WIM_OPEN                               As Long = &H3BE
Private Const MM_WIM_CLOSE                              As Long = &H3BF
Private Const MM_WIM_DATA                               As Long = &H3C0

Private Const WM_DESTROY                                As Long = &H2&
Private Const MMSYSERR_NOERROR                          As Long = 0&

Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1                                    As Long
    dwParam2                                    As Long
    szName                                      As Long
End Type

Private Type MIXERCONTROL
    cbStruct                                    As Long
    dwControlID                                 As Long
    dwControlType                               As Long
    fdwControl                                  As Long
    cMultipleItems                              As Long
    szShortName(MIXER_SHORT_NAME_CHARS / 2 - 1) As Integer
    szName(MIXER_LONG_NAME_CHARS / 2 - 1)       As Integer
    Bounds(5)                                   As Long
    Metrics(5)                                  As Long
End Type

Private Type MIXERCONTROLDETAILS
    cbStruct                                    As Long
    dwControlID                                 As Long
    cChannels                                   As Long
    item                                        As Long
    cbDetails                                   As Long
    paDetails                                   As Long
End Type

Private Type Target
    dwType                                      As Long
    dwDeviceID                                  As Long
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVersion                              As Long
    szPname                                     As String * MAXPNAMELEN
End Type

Private Type MIXERLINE
    cbStruct                                    As Long
    dwDestination                               As Long
    dwSource                                    As Long
    dwLineID                                    As Long
    fdwLine                                     As Long
    dwUser                                      As Long
    dwComponentType                             As Long
    cChannels                                   As Long
    cConnections                                As Long
    cControls                                   As Long
    szShortName                                 As String * MIXER_SHORT_NAME_CHARS
    szName                                      As String * MIXER_LONG_NAME_CHARS
    tTarget                                     As Target
End Type

Private Type MIXERLINECONTROLS
    cbStruct                                    As Long
    dwLineID                                    As Long
    dwControl                                   As Long
    cControls                                   As Long
    cbmxctrl                                    As Long
    pamxctrl                                    As Long
End Type

Private Type MIXERCAPS
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVersion                              As Long
    szPname                                     As String * 32
    fdwSupport                                  As Long
    cDestinations                               As Long
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue                                     As Long
End Type

Private Type volume_stereo
    L                                           As Long
    r                                           As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag                                  As Integer
    nChannels                                   As Integer
    nSamplesPerSec                              As Long
    nAvgBytesPerSec                             As Long
    nBlockAlign                                 As Integer
    wBitsPerSample                              As Integer
    cbSize                                      As Integer
End Type

Private Type WAVEHDR
    lpData                                      As Long
    dwBufferLen                                 As Long
    dwBytesRec                                  As Long
    dwUser                                      As Long
    dwFlags                                     As Long
    dwLoops                                     As Long
    lpNext                                      As Long
    reserved                                    As Long
End Type

Private Type WAVEFORMAT
    wFormatTag                                  As Integer
    nChannels                                   As Integer
    nSamplesPerSec                              As Long
    nAvgBytesPerSec                             As Long
    nBlockAlign                                 As Integer
End Type

Private Type WAVEINCAPS
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVer                                  As Long
    szPname                                     As String * MAXPNAMELEN
    dwFormats                                   As Long
    wChannels                                   As Integer
    wReserved1                                  As Integer
End Type

Private Type MMTIME
    wType                                       As Long
    U                                           As Long
End Type

Private Type WaveInBuffer
    hdr                                         As WAVEHDR
    intBuffer()                                 As Integer
    pMem                                        As Long
End Type

Private lngCurDev                               As Long
Private lngBufSize                              As Long
Private lngBufCnt                               As Long

Private udtBuffers()                            As WaveInBuffer

Private lngPrevProc                             As Long
Private lngHwnd                                 As Long

Private hWaveIn                                 As Long
Private hMixer                                  As Long

Public Event GotData( _
    intBuffer() As Integer, _
    lngLen As Long _
)

Public Property Get IsRecording( _
) As Boolean

    IsRecording = hWaveIn <> 0
End Property

' http://www.ureader.com/message/1359669.aspx
Private Function SetMixerLine2( _
    ByVal devid As Long, _
    ByVal Index As Long _
) As Boolean

    Dim retval      As Long
    Dim terr        As Long
    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control     As MIXERCONTROL
    Dim hControl    As Long
    Dim pControl    As Long
    Dim cd          As MIXERCONTROLDETAILS
    Dim hr          As Long
    Dim i           As Long
    Dim c           As Long
    Dim j           As Long
    Dim lv          As MIXERCONTROLDETAILS_BOOLEAN
    Dim hLV         As Long
    Dim pLV         As Long

    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_DESTINATION)

    If (hr <> 0) Then Exit Function

    For i = 0 To mxl.dwDestination - 1
        controls.cbStruct = Len(controls)
        controls.dwLineID = mxl.dwLineID
        controls.cControls = mxl.cControls
        controls.cbmxctrl = Len(control)

        hControl = GlobalAlloc(GMEM_ZEROINIT, Len(control) * controls.cControls)
        pControl = GlobalLock(hControl)

        controls.pamxctrl = pControl
        controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

        terr = mixerGetLineControls(devid, controls, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If terr <> 0 Then Exit Function
        CpyMem control, ByVal pControl, Len(control)
        If Index > control.cMultipleItems Then
            GoTo SkipItem
        End If

        For c = 0 To controls.cControls - 1
            CpyMem control, ByVal pControl + Len(control) * c, Len(control)
            If (MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control.dwControlType)) Then
                cd.cbStruct = Len(cd)
                cd.dwControlID = control.dwControlID
                cd.cChannels = 1
                cd.item = control.cMultipleItems
                cd.cbDetails = Len(lv)

                hLV = GlobalAlloc(GMEM_ZEROINIT, cd.cChannels * cd.item * cd.cbDetails)
                pLV = GlobalLock(hLV)
                
                cd.paDetails = pLV

                terr = mixerGetControlDetails(devid, cd, MIXER_GETCONTROLDETAILSF_VALUE)

                For j = 0 To cd.item - 1
                    CpyMem lv, ByVal pLV + Len(lv) * j, Len(lv)
                    If lv.fValue Then retval = i
                    lv.fValue = Abs(CBool(j = Index))
                    CpyMem ByVal pLV + Len(lv) * j, lv, Len(lv)
                Next

                terr = mixerSetControlDetails(devid, cd, MIXER_SETCONTROLDETAILSF_VALUE)

                GlobalUnlock hLV
                GlobalFree hLV
            End If
        Next

SkipItem:
        GlobalUnlock hControl
        GlobalFree hControl
    Next

    SetMixerLine2 = True
End Function

' http://www.ureader.com/message/1359669.aspx
Private Function GetMixerLine2( _
    ByVal devid As Long _
) As Long

    Dim retval      As Long
    Dim terr        As Long
    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control     As MIXERCONTROL
    Dim hControl    As Long
    Dim pControl    As Long
    Dim cd          As MIXERCONTROLDETAILS
    Dim hr          As Long
    Dim i           As Long
    Dim c           As Long
    Dim j           As Long
    Dim lv          As MIXERCONTROLDETAILS_BOOLEAN
    Dim hLV         As Long
    Dim pLV         As Long

    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_DESTINATION)

    If (hr <> 0) Then Exit Function

    For i = 0 To mxl.dwDestination - 1
        controls.cbStruct = Len(controls)
        controls.dwLineID = mxl.dwLineID
        controls.cControls = mxl.cControls
        controls.cbmxctrl = Len(control)

        hControl = GlobalAlloc(GMEM_ZEROINIT, Len(control) * controls.cControls)
        pControl = GlobalLock(hControl)

        controls.pamxctrl = pControl
        controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

        terr = mixerGetLineControls(devid, controls, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If terr <> 0 Then Exit Function

        For c = 0 To controls.cControls - 1
            CpyMem control, ByVal pControl + Len(control) * c, Len(control)
            If (MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control.dwControlType)) Then
                cd.cbStruct = Len(cd)
                cd.dwControlID = control.dwControlID
                cd.cChannels = 1
                cd.item = control.cMultipleItems
                cd.cbDetails = Len(lv)

                hLV = GlobalAlloc(GMEM_ZEROINIT, cd.cChannels * cd.item * cd.cbDetails)
                pLV = GlobalLock(hLV)
                
                cd.paDetails = pLV

                terr = mixerGetControlDetails(devid, cd, MIXER_GETCONTROLDETAILSF_VALUE)

                For j = 0 To cd.item - 1
                    CpyMem lv, ByVal pLV + Len(lv) * j, Len(lv)
                    If lv.fValue Then retval = j
                Next

                GlobalUnlock hLV
                GlobalFree hLV
            End If
        Next

SkipItem:
        GlobalUnlock hControl
        GlobalFree hControl
    Next

    GetMixerLine2 = retval
End Function

Public Function StopRecord( _
) As Boolean

    Dim i   As Long
    Dim res As Long

    If hWaveIn <> 0 Then
        ' make sure all buffers are returned to us
        res = waveInStop(hWaveIn)
        If res <> 0 Then Debug.Print "waveInStop: " & res
        res = waveInReset(hWaveIn)
        If res <> 0 Then Debug.Print "waveInReset: " & res

        ' destroy all buffers
        For i = 0 To lngBufCnt - 1
            res = waveInUnprepareHeader(hWaveIn, _
                                  udtBuffers(i).hdr, _
                                  Len(udtBuffers(i).hdr))

            If res <> 0 Then Debug.Print "waveInUnprep " & i & ": " & res
        Next

        res = waveInClose(hWaveIn)
        If res <> 0 Then Debug.Print "waveInClose: " & res

        hWaveIn = 0
    End If

    DeinitWaveInWnd

    StopRecord = True
End Function

Public Function StartRecord( _
    ByVal samplerate As Long, _
    ByVal Channels As Integer _
) As Boolean

    Dim udtWFX  As WAVEFORMATEX
    Dim res     As Long
    Dim i       As Long
    Dim j       As Long

    ReDim udtBuffers(lngBufCnt - 1) As WaveInBuffer

    InitWaveInWnd

    With udtWFX
        .wFormatTag = WAVE_FORMAT_PCM
        .nSamplesPerSec = 8000
        .nChannels = 1
        .wBitsPerSample = 16
        .nBlockAlign = Channels * (.wBitsPerSample / 8)
        .nAvgBytesPerSec = .nBlockAlign * .nSamplesPerSec
        .cbSize = 0
    End With

    res = waveInOpen(hWaveIn, lngCurDev, udtWFX, lngHwnd, 0, CALLBACK_WINDOW)
    If res <> MMSYSERR_NOERROR Then
        DeinitWaveInWnd
        Exit Function
    End If

    ' prepare headers/buffers
    For i = 0 To lngBufCnt - 1
        With udtBuffers(i)
            ReDim .intBuffer(lngBufSize / 2 - 1) As Integer
            .pMem = VarPtr(.intBuffer(0))
            .hdr.dwBufferLen = lngBufSize
            .hdr.lpData = .pMem
            .hdr.dwUser = i

            res = waveInPrepareHeader(hWaveIn, .hdr, Len(.hdr))
            If res <> MMSYSERR_NOERROR Then
                ' on error unprepare all prepared headers
                For j = (i - 1) To 0 Step -1
                    waveInUnprepareHeader hWaveIn, .hdr, Len(.hdr)
                Next

                waveInClose hWaveIn
                hWaveIn = 0

                DeinitWaveInWnd

                Exit Function
            End If
        End With
    Next

    res = waveInStart(hWaveIn)
    If res <> MMSYSERR_NOERROR Then
        For i = 0 To lngBufCnt - 1
            waveInUnprepareHeader hWaveIn, _
                                  udtBuffers(i).hdr, _
                                  Len(udtBuffers(i).hdr)
        Next

        waveInClose hWaveIn
        hWaveIn = 0

        DeinitWaveInWnd

        Exit Function
    End If

    For i = 0 To lngBufCnt - 1
        waveInAddBuffer hWaveIn, _
                        udtBuffers(i).hdr, _
                        Len(udtBuffers(i).hdr)
    Next

    StartRecord = True
End Function

' number of buffers to be used for recording
Public Property Get BufferCount( _
) As Long

    BufferCount = lngBufCnt
End Property

Public Property Let BufferCount( _
    ByVal Index As Long _
)

    lngBufCnt = Index
End Property

' size of a buffer in bytes
' (must be multiple of 2 because we use 16 bit integers)
Public Property Get BufferSize( _
) As Long

    BufferSize = lngBufSize
End Property

Public Property Let BufferSize( _
    ByVal bytes As Long _
)

    lngBufSize = bytes
End Property

Public Property Get DeviceCount( _
) As Long

    DeviceCount = waveInGetNumDevs()
End Property

Public Property Get DeviceName( _
    ByVal Index As Long _
) As String

    Dim udtInfo As WAVEINCAPS
    Dim strName As String

    waveInGetDevCaps Index, udtInfo, Len(udtInfo)
    strName = udtInfo.szPname

    If InStr(strName, Chr$(0)) > 0 Then
        strName = Left$(strName, InStr(strName, Chr$(0)) - 1)
    End If

    DeviceName = udtInfo.szPname
End Property

Public Function SelectDevice( _
    ByVal Index As Long _
) As Boolean

    If Index < 0 Or Index > DeviceCount - 1 Then
        Exit Function
    End If

    lngCurDev = -1

    If hMixer <> 0 Then
        mixerClose hMixer
        hMixer = 0
    End If

    mixerOpen hMixer, Index, 0, 0, MIXER_OBJECTF_WAVEIN
    If hMixer = 0 Then Exit Function

    lngCurDev = Index

    SelectDevice = True
End Function

Public Property Get MixerLineCount( _
) As Long

    Dim i   As Long

    If lngCurDev < 0 Then Exit Property

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineCount = MixerDestinationConnections(hMixer, i)
            Exit For
        End If
    Next
End Property

Public Property Get MixerLineName( _
    ByVal Index As Long _
) As String

    Dim i   As Long

    If lngCurDev < 0 Then Exit Property

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineName = MixerConnectionName(hMixer, i, Index)
            Exit For
        End If
    Next
End Property

Public Function SelectMixerLine( _
    ByVal Index As Long _
) As Boolean

    If lngCurDev < 0 Then Exit Function

    SelectMixerLine = SetMixerLine2(hMixer, MixerLineCount - Index - 1)
End Function

Public Property Get SelectedDevice( _
) As Long

    SelectedDevice = lngCurDev
End Property

Public Property Get SelectedMixerLine( _
) As Long

     SelectedMixerLine = MixerLineCount - GetMixerLine2(hMixer) - 1
  
     
End Property

Public Property Get MixerLineType( _
    Optional ByVal Index As Long = -1 _
) As Long

    Dim i       As Long
    Dim lngLine As Long

    If Index = -1 Then
        lngLine = SelectedMixerLine
    Else
        lngLine = Index
    End If

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineType = MixerConnectionType(hMixer, i, lngLine)
            Exit For
        End If
    Next
End Property

Public Property Get MixerLineVolume( _
) As Long

    Dim i       As Long
    Dim udtVol  As volume_stereo

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            udtVol = MixerConnectionVolume(hMixer, i, SelectedMixerLine)
            Exit For
        End If
    Next

    MixerLineVolume = udtVol.L
End Property

Public Property Let MixerLineVolume( _
    ByVal Volume As Long _
)

    Dim i       As Long
    Dim udtVol  As volume_stereo

    udtVol.L = Volume
    udtVol.r = Volume

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerSetConnectionVolume hMixer, i, SelectedMixerLine, udtVol
            Exit For
        End If
    Next
End Property

' subclass the created fake window
Private Sub Attach( _
    ByVal hWnd As Long _
)

    lngHwnd = hWnd
    sc_Subclass lngHwnd
    sc_AddMsg lngHwnd, MM_WIM_DATA
End Sub

' remove our window proc
Private Sub Detach()
    If lngHwnd Then
        sc_UnSubclass lngHwnd
        lngHwnd = 0
    End If
End Sub

' waveIn API requires either a callback or a
' window to send messages to.
' As adding a buffer to the WaveIn Queue in a
' direct callback causes a dead lock, better
' use a fake window.
Private Sub InitWaveInWnd()
    lngHwnd = CreateWindowEx(0&, _
                             "STATIC", _
                             "WAVE_IN_WINDOW", _
                             0&, 0&, 0&, 0&, _
                             0&, 0&, 0&, _
                             App.hInstance, _
                             0&)

    Attach lngHwnd
End Sub

Private Sub DeinitWaveInWnd()
    Detach
    DestroyWindow lngHwnd: lngHwnd = 0: lngPrevProc = 0
End Sub

Private Sub Class_Initialize()
    lngBufCnt = 5
    lngBufSize = 10& * 1024&
    lngCurDev = -1
End Sub

Private Sub Class_Terminate()
    If IsRecording Then
        StopRecord
    End If

    If hMixer <> 0 Then
        mixerClose hMixer
    End If
End Sub

' /////////////////////////////////////////////////
' ////////// WINDOWS MIXER FUNCTIONS
' /////////////////////////////////////////////////

Private Property Get MixerConnectionType( _
    ByVal devid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo devid, udtML, MIXER_GETLINEINFOF_SOURCE

    MixerConnectionType = udtML.dwComponentType
End Property

Private Property Get MixerDestinations( _
    ByVal deviceid As Long _
) As Long

    Dim udtCaps As MIXERCAPS

    mixerGetDevCaps deviceid, udtCaps, Len(udtCaps)

    MixerDestinations = udtCaps.cDestinations
End Property

Private Property Get MixerConnectionName( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As String

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_SOURCE

    MixerConnectionName = udtML.szName
End Property

Private Property Get MixerDestinationConnections( _
    ByVal deviceid As Long, _
    ByVal Destination As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_DESTINATION

    MixerDestinationConnections = udtML.cConnections
End Property

Private Property Get MixerDestinationType( _
    ByVal deviceid As Long, _
    ByVal Destination As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_DESTINATION

    MixerDestinationType = udtML.dwComponentType
End Property

Private Sub MixerSetConnectionVolume( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long, _
    vol As volume_stereo _
)

    Dim udtML   As MIXERLINE
    Dim udtMCL  As MIXERCONTROLDETAILS
    Dim CtrlID  As Long

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, udtML, MIXER_GETLINEINFOF_SOURCE

    CtrlID = MixerGetControlID(deviceid, _
                               udtML.dwComponentType, _
                               MIXERCONTROL_CONTROLTYPE_VOLUME, _
                               udtML.dwLineID)

    With udtMCL
        .cbDetails = 4
        .cbStruct = Len(udtMCL)
        .cChannels = udtML.cChannels
        .dwControlID = CtrlID
        .item = 0
        .paDetails = VarPtr(vol)
    End With

    mixerSetControlDetails deviceid, _
                           udtMCL, _
                           MIXER_SETCONTROLDETAILSF_VALUE
End Sub

Private Property Get MixerConnectionVolume( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As volume_stereo

    Dim udtML   As MIXERLINE
    Dim udtMCL  As MIXERCONTROLDETAILS
    Dim udtVol  As volume_stereo
    Dim CtrlID  As Long

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, udtML, MIXER_GETLINEINFOF_SOURCE

    CtrlID = MixerGetControlID(deviceid, _
                               udtML.dwComponentType, _
                               MIXERCONTROL_CONTROLTYPE_VOLUME, _
                               udtML.dwLineID)

    With udtMCL
        .cbDetails = 4
        .cbStruct = Len(udtMCL)
        .cChannels = udtML.cChannels
        .dwControlID = CtrlID
        .item = 0
        .paDetails = VarPtr(udtVol)
    End With

    mixerGetControlDetails deviceid, _
                           udtMCL, _
                           MIXER_GETCONTROLDETAILSF_VALUE

    MixerConnectionVolume = udtVol
End Property

Private Function MixerGetControlID( _
    ByVal deviceid As Long, _
    ByVal ComponentType As Long, _
    ByVal ControlType As Long, _
    ByVal LineID As Long _
) As Long

    Dim hMem     As Long
    Dim MC       As MIXERCONTROL
    Dim MxrLine  As MIXERLINE
    Dim MLC      As MIXERLINECONTROLS

    MxrLine.cbStruct = Len(MxrLine)
    MxrLine.dwComponentType = ComponentType

    If mixerGetLineInfo(deviceid, _
                        MxrLine, _
                        MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then

        MLC.cbStruct = Len(MLC)
        MLC.dwLineID = LineID
        MLC.dwControl = ControlType
        MLC.cControls = 1
        MLC.cbmxctrl = Len(MC)

        hMem = GlobalAlloc(&H40, Len(MC))
        If hMem = 0 Then Exit Function
        MLC.pamxctrl = GlobalLock(hMem)

        MC.cbStruct = Len(MC)

        If mixerGetLineControls(deviceid, _
                                MLC, _
                                MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then

            CpyMem MC, ByVal MLC.pamxctrl, Len(MC)
            MixerGetControlID = MC.dwControlID
        End If

        GlobalUnlock hMem
        GlobalFree hMem
    End If
End Function


'-SelfSub code------------------------------------------------------------------------------------
' by Paul Caton

Private Function sc_Subclass( _
    ByVal lng_hWnd As Long, _
    Optional ByVal lParamUser As Long = 0, _
    Optional ByVal nOrdinal As Long = 1, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal bIdeSafety As Boolean = True _
) As Boolean 'Subclass the specified window handle

'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
    If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    nMyID = GetCurrentProcessId                                               'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
    If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
  
    If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
        Set oCallback = Me                                                      'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
    If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
    
    If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
        Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
    End If
  
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

    If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
        On Error GoTo CatchDoubleSub                                            'Catch double subclassing
         z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
  
        If bIdeSafety Then                                                      'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
        End If
    
        z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
        z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
        nAddr = SetWindowLong(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        
        z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
        'RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
        CpyMem ByVal z_ScMem, z_Sc(0), CODE_LEN
        sc_Subclass = True                                                      'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If
  
    Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
    Dim i As Long

    If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
        With z_Funk
            For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
                z_ScMem = .item(i)                                                  'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
                End If
            Next i                                                                'Next member of the collection
        End With
        Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
    End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass( _
    ByVal lng_hWnd As Long _
)

    If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
        End If
        z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
    End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg( _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER _
)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
        End If
        If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
        End If
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg( _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER _
)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
        End If
    End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc( _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        sc_CallOrigWndProc = _
            CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser( _
    ByVal lng_hWnd As Long _
) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
    End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser( _
    ByVal lng_hWnd As Long, _
    ByVal NewValue As Long _
)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
    End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg( _
    ByVal uMsg As Long, _
    ByVal nTable As Long _
)

    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                       'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = 0 Then                                                  'If the element is free...
                zData(i) = uMsg                                                     'Use this element
                GoTo Bail                                                           'Bail
            ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry

        nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                    'Store the message in the appended table entry
    End If

    zData(0) = nCount                                                         'Store the new table entry count
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg( _
    ByVal uMsg As Long, _
    ByVal nTable As Long _
)

    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                            'Zero the table entry count
    Else
        nCount = zData(0)                                                       'Get the table entry count
    
        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = uMsg Then                                               'If the message is found...
                zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry
    
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
  
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError( _
    ByVal sRoutine As String, _
    ByVal sMsg As String _
)

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr( _
    ByVal sDLL As String, _
    ByVal sProc As String _
) As Long

    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
    Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd( _
    ByVal lng_hWnd As Long _
) As Long

    If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
        On Error GoTo Catch                                                     'Catch unsubclassed window handles
        z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
        zMap_hWnd = z_ScMem
    End If
  
    Exit Function                                                             'Exit returning the thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf( _
    ByVal oCallback As Object, _
    ByVal nOrdinal As Long _
) As Long

    Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                         'Address of the vTable
    Dim i     As Long                                                         'Loop index
    Dim j     As Long                                                         'Loop limit
  
    'RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
    CpyMem nAddr, ByVal ObjPtr(oCallback), 4

    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
                Exit Function                                                       'Bail...
            End If
        End If
    End If
  
    i = i + 4                                                                 'Bump to the next entry
    j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        'RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
        CpyMem nAddr, ByVal i, 4
    
        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            'RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            CpyMem zAddressOf, ByVal (i - (nOrdinal * 4)), 4
            Exit Do                                                               'Bad method signature, quit loop
        End If

        'RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        CpyMem bVal, ByVal nAddr, 1
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            'RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            CpyMem zAddressOf, ByVal (i - (nOrdinal * 4)), 4
            Exit Do                                                               'Bad method signature, quit loop
        End If
    
        i = i + 4                                                             'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe( _
    ByVal nStart As Long, _
    ByRef nMethod As Long, _
    ByRef bSub As Byte _
) As Boolean

    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries
    Do While nAddr < nLimit                                                   'While we've not reached our probe depth
        'RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
        CpyMem nEntry, ByVal nAddr, 4
    
        If nEntry <> 0 Then                                                     'If not an implemented interface
            'RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
            CpyMem bVal, ByVal nEntry, 1
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Function                                                       'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop
End Function

Private Property Get zData( _
    ByVal nIndex As Long _
) As Long

    'RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
    CpyMem zData, ByVal z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData( _
    ByVal nIndex As Long, _
    ByVal nValue As Long _
)

    'RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
    CpyMem ByVal z_ScMem + (nIndex * 4), nValue, 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1( _
    ByVal bBefore As Boolean, _
    ByRef bHandled As Boolean, _
    ByRef lReturn As Long, _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef lParamUser As Long _
)

'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************

    Dim udtHdr  As WAVEHDR

    If uMsg = MM_WIM_DATA Then
        If IsBadReadPtr(ByVal lParam, Len(udtHdr)) = 0 Then

            CpyMem udtHdr, _
                   ByVal lParam, _
                   Len(udtHdr)

            RaiseEvent GotData(udtBuffers(udtHdr.dwUser).intBuffer, _
                               udtHdr.dwBytesRec)

            ' place the buffer in the waveIn queue again
            waveInAddBuffer hWaveIn, _
                            udtBuffers(udtHdr.dwUser), _
                            Len(udtHdr)

        End If
    End If
End Sub



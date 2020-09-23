VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Digital Signal Processing Class for 16 bit samples
Private Const Pi                    As Single = 3.14159265358979
Private lngSamplerate               As Long
Private intChannels                 As Integer

Public Property Get samplerate( _
) As Long

    samplerate = lngSamplerate
End Property

Public Property Let samplerate( _
    ByVal lngSR As Long _
)

    If lngSR < 1 Or lngSR > 96000 Then
        Err.Raise 32000, "invalid samplerate"
    Else
        lngSamplerate = lngSR
        UpdateFX
    End If
End Property

Public Property Get Channels( _
) As Integer

    Channels = intChannels
End Property

Public Property Let Channels( _
    ByVal intCh As Integer _
)

    If intCh < 1 Or intCh > 2 Then
        Err.Raise 32000, "invalid channels"
    Else
        intChannels = intCh
        UpdateFX
    End If
End Property




Public Sub ProcessSamples( _
    intSamples() As Integer _
)

    Dim i   As Long

    If lngSamplerate = 0 Then Exit Sub
    If intChannels = 0 Then Exit Sub

    
End Sub


Private Sub UpdateFX()
    If lngSamplerate = 0 Then Exit Sub
    If intChannels = 0 Then Exit Sub

End Sub

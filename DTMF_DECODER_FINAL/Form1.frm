VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DTMF Remote "
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Number Received From Remote"
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrVis 
      Interval        =   10
      Left            =   3120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number Received From Remote Control"
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   570
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Action"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
      Begin VB.Label Label6 
         BackColor       =   &H00400000&
         Caption         =   "6.  Media Player"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00400000&
         Caption         =   "5. Calculator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00400000&
         Caption         =   "4.  Notepad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00400000&
         Caption         =   "3. Word Pad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400000&
         Caption         =   "2.  Ms- Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "1.  Ms- Word"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m As Integer
Dim dgt As String
Private WithEvents clsRecorder  As WaveInRecorder
Attribute clsRecorder.VB_VarHelpID = -1
Private clsDSP                  As clsDSP
Private intSamples()            As Integer
Private blnLoaded               As Boolean

Private Sub clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)

    ' save the current buffer for visualizing it
    intSamples = intBuffer
    If Not clsRecorder.IsRecording Then Exit Sub
End Sub



Private Sub Record_init()
    If clsRecorder.IsRecording Then
        If Not clsRecorder.StopRecord Then
            MsgBox "Could not stop recording!", vbExclamation
        End If
    Else
        clsDSP.samplerate = CLng(8000)
        clsDSP.Channels = 1 + 1

        If Not clsRecorder.StartRecord(8000, 1 + 1) Then
            MsgBox "Could not start recording!", vbExclamation
        End If
    End If
       

End Sub





Private Sub Form_Load()
exec = False
freqs(0) = 697
freqs(1) = 770
freqs(2) = 852
freqs(3) = 941
freqs(4) = 1209
freqs(5) = 1336
freqs(6) = 1477
freqs(7) = 1633
MAX_BINS = 8
GOERTZEL_N = 92
SAMPLING_RATE = 8000
'sample_count = 0
    
    
Set clsRecorder = New WaveInRecorder
Set clsDSP = New clsDSP
ReDim intSamples(FFT_SAMPLES - 1) As Integer
blnLoaded = True
clsDSP.samplerate = CLng(8000)
clsDSP.Channels = 1 ' for sterio
Me.Show
Record_init
    
End Sub





Private Sub Form_Unload(Cancel As Integer)

    If clsRecorder.IsRecording Then
        Record_init
    End If

    Set clsRecorder = Nothing
End Sub



Private Sub tmrVis_Timer()
Dim i As Long
Dim Curnum As String

For i = 0 To UBound(intSamples)
    goertzel intSamples(i), dgt
    intSamples(i) = 0
    If Len(dgt) >= 2 Then Exit For
Next i

If Len(dgt) >= 2 Then
Text1.Text = Text1.Text & Left$(dgt, 1)
Curnum = Left$(dgt, 1)
Select Case Curnum
    Case "1"
        Shell "D:\Program Files\Microsoft Office\Office\winword.exe", vbMaximizedFocus
    Case "2"
        Shell "D:\Program Files\Microsoft Office\Office\excel.exe", vbMaximizedFocus
    Case "3"
    
        Shell "D:\Program Files\Windows NT\Accessories\wordpad.exe", vbMaximizedFocus

    Case "4"
        Shell "D:\windows\system32\notepad.exe", vbMaximizedFocus

    Case "5"
        Shell "D:\windows\system32\calc.exe", vbMaximizedFocus
    Case "6"
    Shell "D:\Program Files\Windows Media Player\wmplayer.exe", vbMaximizedFocus
   



End Select

End If
dgt = ""


End Sub


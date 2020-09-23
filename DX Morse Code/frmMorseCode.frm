VERSION 5.00
Begin VB.Form frmMorseCode 
   Caption         =   "DX Text to Morse Code "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnglish 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Text            =   "SOS"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblEnglish 
      Caption         =   "Enter Valid Text to Play in Morse Code!"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmMorseCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DirectSound MorseCode Sample
' By Jim : James A Hersey II<jhersey@biggear.com>
'
'For that guy whose email I accidently deleted!
'
'If the duration of a dot is taken to be one unit
'then that of a dash is three units.
'The space between the components of one character is one unit,
'between characters is three units and between words seven units.
'To indicate that a mistake has been made and for the receiver to
'delete the last word send ........ (eight dots).



Option Explicit
Dim lpBuffer() As Integer
Dim lBufferSize As Long

Private Const MorseUnitLength = 100

Private Const MorseDash = 3
Private Const MorseDot = 1
Private Const MorseWordGap = 7
Private Const morseUnitComponentGap = 1
Private Const morseCharacterGap = 3

Private aMorseCode() As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Function createBaseSound() As Integer
    Dim two_pi As Double
    two_pi = 8 * Atn(1)
    Dim lFrequency&, lSampleRate&, lMultiplier&, I&
    lFrequency = 100
    lSampleRate = 22050
    lMultiplier = lSampleRate \ lFrequency
    lBufferSize = lMultiplier * 500
    ReDim lpBuffer(lBufferSize)
    For I = 0 To lBufferSize - 1
        '// For 16-bit integer PCM
        lpBuffer(I) = 32767 * Sin(I * two_pi * lFrequency / lSampleRate)
    Next
    lpBuffer() = lpBuffer()
    lBufferSize = lBufferSize
     
       
End Function


Private Sub cmdPlay_Click()
    Dim sSequence As String
    Dim iRet As Integer
    
    iRet = CreateMorseCodeSequence(txtEnglish.Text, sSequence)
    If iRet <> -1 Then
        Play sSequence
    Else
        MsgBox "String was not able to converted to Morse Code", vbInformation
    End If
End Sub

Private Sub Form_Load()
    Dim iRet As Integer
    ReDim aMorseCode(33 To 90) ' The numbers are the ascii value
    
    'NUMBERS
    aMorseCode(48) = "-----" '0
    aMorseCode(49) = ".----" '1
    aMorseCode(50) = "..---" '2
    aMorseCode(51) = "...--" '3
    aMorseCode(52) = "....-" '4
    aMorseCode(53) = "....." '5
    aMorseCode(54) = "-...." '6
    aMorseCode(55) = "--..." '7
    aMorseCode(56) = "---.." '8
    aMorseCode(57) = "----." '9

    'PUNCTUATION
    aMorseCode(33) = ".----." '!
    aMorseCode(34) = ".-..-." '"
    aMorseCode(40) = "-.--.-" '(
    aMorseCode(44) = "--..--" ',
    aMorseCode(45) = "-....-" '-
    aMorseCode(46) = ".-.-.-" '.
    aMorseCode(47) = "-..-. " '/
    aMorseCode(58) = "---..." ':
    aMorseCode(63) = "..--.." '?
 
    'CHARACTERS
    aMorseCode(65) = ".-" 'A
    aMorseCode(66) = "-..." 'B
    aMorseCode(67) = "-.-." 'C
    aMorseCode(68) = "-.." 'D
    aMorseCode(69) = "." 'E
    aMorseCode(70) = "..-." 'F
    aMorseCode(71) = "--." 'G
    aMorseCode(72) = "...." 'H
    aMorseCode(73) = ".." 'I
    aMorseCode(74) = ".---" 'J
    aMorseCode(75) = "-.-" 'K
    aMorseCode(76) = ".-.." 'L
    aMorseCode(77) = "--" 'M
    aMorseCode(78) = "-." 'N
    aMorseCode(79) = "---" 'O
    aMorseCode(80) = ".--." 'P
    aMorseCode(81) = "--.-" 'Q
    aMorseCode(82) = ".-." 'R
    aMorseCode(83) = "..." 'S
    aMorseCode(84) = "-" 'T
    aMorseCode(85) = "..-" 'U
    aMorseCode(86) = "...-" 'V
    aMorseCode(87) = ".--" 'W
    aMorseCode(88) = "-..-" 'X
    aMorseCode(89) = "-.--" 'Y
    aMorseCode(90) = "--.." 'Z
        
    iRet = createBaseSound
    
End Sub

Private Function CreateMorseCodeSequence(ByVal sEnglish As String, ByRef sSequence As String) As Integer
Dim x As Long
Dim sChar As String
Dim iAscii As Integer

sSequence = ""

For x = 1 To Len(sEnglish)
    sChar = Mid(sEnglish, x, 1)
    iAscii = Asc(UCase(sChar))
    Select Case iAscii
    Case 32
        sSequence = sSequence & "W"
    Case 33 To 90
        If aMorseCode(iAscii) <> "" Then
            sSequence = sSequence & aMorseCode(iAscii) & "C"
        Else
            GoTo ErrHandler
        End If
    Case Else
            GoTo ErrHandler
    End Select
Next x
Exit Function
ErrHandler:
Select Case Err.Number
Case Else
        MsgBox "Character not yet developed in MorseCode array", vbInformation
        CreateMorseCodeSequence = -1
End Select
End Function


Private Sub Play(sSequence As String)
    
    Dim x As Long
    Dim DX As DirectX8
    Dim DS As DirectSound8
    Dim DSBD1 As DSBUFFERDESC
    Dim DSB As DirectSoundSecondaryBuffer8
    Dim DSFormat As WAVEFORMATEX
     
    Set DX = New DirectX8
    Set DS = DX.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
     
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 22050
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
     
    DSBD1.fxFormat = DSFormat
    DSBD1.lBufferBytes = lBufferSize
    Set DSB = DS.CreateSoundBuffer(DSBD1)
    DSB.WriteBuffer 0, 0, lpBuffer(0), DSBLOCK_ENTIREBUFFER
    
    Dim lTime As Long
    For x = 1 To Len(sSequence)
        Select Case Mid$(sSequence, x, 1)
            Case "."
                DSB.Play DSBPLAY_LOOPING
                Sleep MorseUnitLength * MorseDot
                DSB.Stop
                Sleep MorseUnitLength * morseUnitComponentGap
            Case "-"
                DSB.Play DSBPLAY_LOOPING
                Sleep MorseUnitLength * MorseDash
                DSB.Stop
                Sleep MorseUnitLength * morseUnitComponentGap
            Case "C" ' Character Gap
                Sleep MorseUnitLength * morseCharacterGap
            Case "W" ' Word Gap
                Sleep MorseUnitLength * MorseWordGap
            Case Else
                MsgBox "Tone not developed", vbCritical
        End Select
           
    Next x
    Set DSB = Nothing
    Set DS = Nothing
    Set DX = Nothing
 

End Sub

Attribute VB_Name = "CHAT_Defs"
Option Explicit

' Application User Defined Types...
' Sound Format
Public Const WAVE_FORMAT_PCM = &H1                  ' Microsoft Windows PCM Wave Format
Public Const WAVE_FORMAT_ADPCM = &H11               ' ADPCM Wave Format
Public Const WAVE_FORMAT_IMA_ADPCM = &H11           ' IMA ADPCM Wave Format
Public Const WAVE_FORMAT_DVI_ADPCM = &H11           ' DVI ADPCM Wave Format
Public Const WAVE_FORMAT_DSPGROUP_TRUESPEECH = &H22 ' DSP Group Wave Format
Public Const WAVE_FORMAT_GSM610 = &H31              ' GSM610 Wave Format
Public Const WAVE_FORMAT_MSN_AUDIO = &H32           ' MSN Audio Wave Format

Public Const TIMESLICE = 0.2            ' Time Slicing 1/5 Second

' Application Constants...
Public Const NoOfRings = 1                  ' Number Of Times In/Out Bound Calls Ring...

Public Const phoneHungUp = 3                ' Hangup Status Icon...
Public Const phoneRingIng = 2               ' Ringing Status Icon...
Public Const phoneAnswered = 1              ' Answered Status Icon...
Public Const mikeNO = 6
Public Const mikeOFF = 7
Public Const mikeON = 8
Public Const speakNO = 9
Public Const speakOFF = 10
Public Const speakON = 11

Public Const RingInId = 101                 ' Ringing InBound Sound...
Public Const RingOutId = 102                ' Ringing OutBound Sound...

' Toolbar constants...
Public Const tbCALL = 2
Public Const tbHANGUP = 3
Public Const tbAUTOANSWER = 5

'== flag values for wFlags parameter ==================================
Public Const SND_SYNC = &H0                 '  play synchronously (default)
Public Const SND_ASYNC = &H1                '  play asynchronously
Public Const SND_NODEFAULT = &H2            '  don't use default sound
Public Const SND_MEMORY = &H4               '  lpszSoundName points to a memory file
Public Const SND_LOOP = &H8                 '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10              '  don't stop any currently playing sound

'== MCI Wave API Declarations ================================================
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SoundData As Any, ByVal uFlags As Long) As Long


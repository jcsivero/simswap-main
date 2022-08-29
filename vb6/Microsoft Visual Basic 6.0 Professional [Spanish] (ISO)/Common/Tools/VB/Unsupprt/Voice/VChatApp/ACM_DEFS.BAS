Attribute VB_Name = "ACM_Defs"
Option Explicit
'== ACM API Constants ================================================
Public Const ACMERR_BASE = 512
Public Const ACMERR_NOTPOSSIBLE = (ACMERR_BASE + 0)
Public Const ACMERR_BUSY = (ACMERR_BASE + 1)
Public Const ACMERR_UNPREPARED = (ACMERR_BASE + 2)
Public Const ACMERR_CANCELED = (ACMERR_BASE + 3)

' AcmStreamSizeFormat Constants
Public Const ACM_STREAMSIZEF_SOURCE = &H0
Public Const ACM_STREAMSIZEF_DESTINATION = &H1
Public Const ACM_STREAMSIZEF_QUERYMASK = &HF

' acmStreamConvert Formats
Public Const ACM_STREAMCONVERTF_BLOCKALIGN = &H4
Public Const ACM_STREAMCONVERTF_START = &H10
Public Const ACM_STREAMCONVERTF_END = &H20

' Done Bits For ACMSTREAMHEADER.fdwStatus
Public Const ACMSTREAMHEADER_STATUSF_DONE = &H10000
Public Const ACMSTREAMHEADER_STATUSF_PREPARED = &H20000
Public Const ACMSTREAMHEADER_STATUSF_INQUEUE = &H100000

' Done Bits For acmStreamOpen Formats
Public Const ACM_STREAMOPENF_QUERY = &H1
Public Const ACM_STREAMOPENF_ASYNC = &H2
Public Const ACM_STREAMOPENF_NONREALTIME = &H4

'== ACM API Declarations ================================================
'Declare Function acmStreamOpen Lib "MSACM32" (ByVal hAS As Integer, ByVal hADrv As Integer, wfxSrc As WAVEFORMATEX, wfxDst As WAVEFORMATEX, wFltr As Any, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Integer
'Declare Function acmStreamOpen Lib "MSACM32" (ByVal hAS As Long, ByVal hADrv As Long, wfxSrc As WAVEFORMATEX, wfxDst As WAVEFORMATEX, wFltr As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function acmStreamOpen Lib "MSACM32" (ByVal hAS As Long, ByVal hADrv As Long, wfxSrc As WAVEFORMATEX, wfxDst As WAVEFORMATEX, ByVal wFltr As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long

Declare Function acmStreamPrepareHeader Lib "MSACM32" (ByVal hAS As Integer, hASHdr As ACMSTREAMHEADER, ByVal dwPrepare As Long) As Integer
Declare Function acmStreamUnprepareHeader Lib "MSACM32" (ByVal hAS As Integer, hASHdr As ACMSTREAMHEADER, ByVal dwUnPrepare As Long) As Integer
Declare Function acmStreamConvert Lib "MSACM32" (ByVal hAS As Integer, hASHdr As ACMSTREAMHEADER, ByVal dwConvert As Long) As Integer
Declare Function acmStreamClose Lib "MSACM32" (ByVal hAS As Integer, ByVal dwClose As Long) As Integer
Declare Function acmStreamReset Lib "MSACM32" (ByVal hAS As Integer, ByVal dwReset As Long) As Integer
Declare Function acmStreamSize Lib "MSACM32" (ByVal hAS As Integer, ByVal cbInput As Long, ByVal dwOutBytes As Long, ByVal dwSize As Long) As Integer


'== ACM User Defined Datatypes ================================================
Type WAVEFILTER
    cbStruct      As Long
    dwFilterTag   As Long
    fdwFilter     As Long
    dwReserved(5) As Long
End Type

Type ACMSTREAMHEADER            ' [ACM STREAM HEADER TYPE]
    cbStruct As Long            ' Size of header in bytes
    dwStatus As Long            ' Conversion status buffer
    dwUser As Long              ' 32 bits of user data specified by application
    pbSrc As Long               ' Source data buffer pointer
    cbSrcLength As Long         ' Source data buffer size in bytes
    cbSrcLengthUsed As Long     ' Source data buffer size used in bytes
    dwSrcUser As Long           ' 32 bits of user data specified by application
    cbDst As Long               ' Dest data buffer pointer
    cbDstLength As Long         ' Dest data buffer size in bytes
    cbDstLengthUsed As Long     ' Dest data buffer size used in bytes
    dwDstUser As Long           ' 32 bits of user data specified by application
    dwReservedDriver(10) As Long ' Reserved and should not be used
End Type
'==============================================================================

'------------------------------------------------------------------
Public Function acmCompress(srcWavefmt As WAVEFORMATEX, dstWavefmt As WAVEFORMATEX) As Boolean
'------------------------------------------------------------------
    Dim rc As Long
    Dim hAS As Long
    Dim hASHdr As ACMSTREAMHEADER
'   Dim wFltr As WAVEFILTER
    Dim dwConvert As Long, dwClose As Long, dwReset As Long
    Dim cbInput As Long, dwOutBytes As Long, dwSize As Long
'------------------------------------------------------------------
    ' Open/Configure an acm Stream Handle For Compression
'   rc = acmStreamOpen(hAS, 0, srcWavefmt, dstWavefmt, wFltr, 0, 0, dwOpen)
    rc = acmStreamOpen(hAS, 0, srcWavefmt, dstWavefmt, 0, 0, 0, ACM_STREAMOPENF_ASYNC)
    Debug.Print "acmStreamOpen rc= ", rc
    
    ' Prepare acm Stream Header
    rc = acmStreamPrepareHeader(hAS, hASHdr, 0)
    Debug.Print "acmStreamPrepareHeader rc= ", rc
    
        cbInput = 255 ' must be non zero
    
        ' Calculate acm Stream Size of Output Buffer
        rc = acmStreamSize(hAS, cbInput, dwOutBytes, ACM_STREAMSIZEF_SOURCE)
        Debug.Print "acmStreamSize(input) rc= ", rc
        
        ' Calculate acm Stream Size of Output Buffer
        rc = acmStreamSize(hAS, cbInput, dwOutBytes, ACM_STREAMSIZEF_DESTINATION)
        Debug.Print "acmStreamSize(output) rc= ", rc
    
    ' Convert/Compress acm Stream Wave Buffer
    rc = acmStreamConvert(hAS, hASHdr, dwConvert)
    Debug.Print "acmStreamConvert rc= ", rc
    
    ' Wait Until Conversion Complete...
    Do                                                          ' Loop Until Conversion Is Done
        DoEvents                                                ' Post Events...
    Loop Until hASHdr.dwStatus And ACMSTREAMHEADER_STATUSF_DONE ' Check For The DONE Flag.

    ' UnPrepare acm Stream Header
    rc = acmStreamUnprepareHeader(hAS, hASHdr, 0)
    Debug.Print "acmStreamUnprepareHeader rc= ", rc
    
    ' Close acm Stream Handle
    rc = acmStreamClose(hAS, dwClose)
    Debug.Print "acmStreamClose rc= ", rc
'------------------------------------------------------------------
End Function
'------------------------------------------------------------------

'--------------------------------------------------------------
Public Sub InitAcmHDR(asHdr As ACMSTREAMHEADER, srcHdr As WAVEHDR)
' Initialize's An Input Wave Header's DataBuffer And Size Members...
'--------------------------------------------------------------
    Dim rc As Long                                      ' Function Return Code...
'--------------------------------------------------------------
    asHdr.cbStruct = Len(asHdr)                         ' Size of header in bytes
    asHdr.pbSrc = srcHdr.lpData                         ' Copy pointer To uncompressed data
    asHdr.cbSrcLength = srcHdr.lpData                   ' Copy size of uncompress data
    
    asHdr.cbDstLength = asHdr.cbSrcLength               ' Allocate Enough Memory For Compression
    asHdr.dwDstUser = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT, _
                                  asHdr.cbDstLength)    ' Allocate Global Memory
    asHdr.cbDst = GlobalLock(asHdr.dwDstUser)           ' Lock Memory handle
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------

'--------------------------------------------------------------
Public Sub WaitForACMCallBack(CallBackBit As Long, cbFlag As Long)
' Waits For Asynchronous Function Callback Bit To Be Set.
'--------------------------------------------------------------
    Do                                  ' Loop Until CallBack Bit Is Set!
        DoEvents                        ' Post Events...
    Loop Until (((CallBackBit And cbFlag) = cbFlag) Or _
                 (CallBackBit = 0))     ' Check For (CallBack Bit Or Null)...
'--------------------------------------------------------------
End Sub
'--------------------------------------------------------------




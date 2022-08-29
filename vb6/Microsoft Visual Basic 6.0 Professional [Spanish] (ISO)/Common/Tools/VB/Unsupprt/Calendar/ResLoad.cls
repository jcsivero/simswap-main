VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' ResLoader.cls
'----------------------------------------------------------------------
' Implementation file for the ResLoader class which is responsible
' for locating and loading resources such as strings and images.
'----------------------------------------------------------------------
' Copyright (c) 1996, Microsoft Corporation
'              All Rights Reserved
'
' Information Contained Herin is Proprietary and Confidential
'----------------------------------------------------------------------

Option Explicit

Private Const ERR_BASE = vbObjectError

'======================================================================
' Public Enumerations
'======================================================================
Public Enum ErrorIDs
    errPropValue = ERR_BASE + 1
    errPropValueRange = ERR_BASE + 2
    errCantChange = ERR_BASE + 3
End Enum 'ErrorIDs

'======================================================================
' Public Methods
'======================================================================

'----------------------------------------------------------------------
' LoadResString()
'----------------------------------------------------------------------
' Purpose:  Loads the desired string matching the ID passed in.
' Inputs:   ID of string to load
' Outputs:  none
'----------------------------------------------------------------------
Public Function LoadResString(StringID As Long) As String
    LoadResString = VB.LoadResString(StringID)
End Function 'LoadResString

'----------------------------------------------------------------------
' RaiseUserError()
'----------------------------------------------------------------------
' Purpose:  Loads an error string matching the ID passed in and matches
'           parameter values passed in with markers in the string
' Inputs:   ID of string to load, parameters to stuff
' Outputs:  none
'----------------------------------------------------------------------
Public Sub RaiseUserError(ErrorID As ErrorIDs, Params As Variant)
    Dim sErrText As String      'raw error string
    Dim nLBound As Long         'lbound of the param array
    Dim nUBound As Long         'ubound of the param array
    Dim nPos As Long            'position marker in the string
    Dim ct As Long              'loop counter
    
    'get the raw string
    sErrText = LoadResString(ErrorID - ERR_BASE)
    
    'The Params() array is a one-dimentional, zero-based array of
    'variant values created by using the Array function in VBA.
    'the values in this must be coerceable to strings.
    
    'Param markers are signalled by a "%<num>" format where the <num>
    'is the number of the param.  This routine will match the first
    'element of the array to "%1", the second to "%2" and so on.
    
    'This routine replaces these param markers with the supplied
    'in the param array and then displays the resulting error message.
    'Extra params are ignored and if less params are passed in than
    'param markers, the unreplaced markers will stay as is.
    
    'get UBound and LBound of the array
    nLBound = LBound(Params)
    nUBound = UBound(Params)
    
    'loop over all the parameters
    For ct = nLBound To nUBound
        'find the param marker
        nPos = InStr(sErrText, "%" & (ct + 1))
        
        If nPos > 0 Then
            'replace the param marker with the param value
            sErrText = Left$(sErrText, nPos - 1) & _
                        Params(ct) & Mid$(sErrText, nPos + 2)
        End If 'found the param
    Next ct
    
    'finally, raise the error
    Err.Raise ErrorID, "MSVBCalendar", sErrText
End Sub 'RaiseUserError()

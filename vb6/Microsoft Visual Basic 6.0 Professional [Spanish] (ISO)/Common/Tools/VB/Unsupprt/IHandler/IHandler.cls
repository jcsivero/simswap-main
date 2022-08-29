VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExtractIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------
'- Public Enums...
'----------------------------------------------------------------
Public Enum INPUTFLAGS
    FOR_SHELL = GIL_FORSHELL
    OPEN_ICON = GIL_OPENICON
End Enum

Public Enum RETURNFLAGS
    DONTCACHE = GIL_DONTCACHE
    NOTFILENAME = GIL_NOTFILENAME
    PERCLASS = GIL_PERCLASS
    PERINSTANCE = GIL_PERINSTANCE
    SIMULATEDOC = GIL_SIMULATEDOC
End Enum

'----------------------------------------------------------------
Public Sub GetIconLocation(clsid As String, iFlag As INPUTFLAGS, Idx As Long, IconFile As String, rFlags As RETURNFLAGS)
'----------------------------------------------------------------
    Dim ExIcon As IExtractIcon                          ' Object --> IExtractIcon Interface
    Dim pUnk As IUnknown                                ' Object --> IUnknown Interface
    Dim szIconFile As String                            ' Icon file path...
    Dim cchMax As Long                                  ' Char count of icon file path
'----------------------------------------------------------------
    Set pUnk = CreateObjectLocal(clsid)                 ' Get IUnknown pointer to clsid object
    Set ExIcon = pUnk                                   ' Implement Known Interface (IEctractIcon) from IUnknown...
    
    szIconFile = String(255, 0)                         ' Preallocate 255 null chars for string
    cchMax = Len(szIconFile)                            ' Count length of string...
    
    ' Call GetIconLocation from clsid's IExtractIcon interface...
    ExIcon.GetIconLocation iFlag, StrPtr(szIconFile), cchMax, Idx, rFlags
    
    IconFile = StrConv(szIconFile, vbUnicode)           ' Convert string to Unicode...
    
    Set ExIcon = Nothing                                ' Destroy IExtractIcon Interface reference
    Set pUnk = Nothing                                  ' Destroy IUnknown Interface reference...
'----------------------------------------------------------------
End Sub
'----------------------------------------------------------------

'----------------------------------------------------------------
Public Function CreateObjectLocal(strCLS As String) As IUnknown
'----------------------------------------------------------------
    Dim rclsid As GUID                              ' Class identifier (CLSID) of object
    Dim IID_IUnknown As GUID                        ' Reference to identifier of IUnknown interface
    Dim pvObj As IUnknown                           ' Indirect pointer to requested interface
    Dim hr As Long                                  ' HRESULT
'----------------------------------------------------------------
    hr = CLSIDFromString(ByVal StrPtr(strCLS), rclsid) ' Convert classid to guid
    
    If (hr = 0) Then                                ' If Success
        With IID_IUnknown                           ' Build IUnknown Guid
           .Data4(0) = &HC0
           .Data4(7) = &H46
        End With
        
        hr = CoCreateInstance(rclsid, ByVal 0&, CLSCTX_INPROC_SERVER, IID_IUnknown, pvObj) ' Get instance of object from classid
        
        If (hr = 0) Then                            ' If Success
            Set CreateObjectLocal = pvObj           ' Return Created object
            Exit Function
        End If
    End If
    
    If hr Then Err.Raise hr                         ' Validate HRESULT
'----------------------------------------------------------------
End Function
'----------------------------------------------------------------



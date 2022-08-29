Attribute VB_Name = "mExtractIcon"
Option Explicit

'----------------------------------------------------------------
'- Public type used in Ole32 api calls...
'----------------------------------------------------------------
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'----------------------------------------------------------------
'- Public API Declares...
'----------------------------------------------------------------
Public Declare Function CLSIDFromString Lib "ole32.dll" (strCLS As Long, clsid As GUID) As Long
Public Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As GUID, pUnkOuter As Any, ByVal dwClsContext As Long, riid As GUID, ppvObj As IUnknown) As Long

'----------------------------------------------------------------
'- Public Constants...
'----------------------------------------------------------------
Public Const CLSCTX_INPROC_SERVER = 1
Public Const CLSCTX_INPROC_HANDLER = 2
Public Const CLSCTX_LOCAL_SERVER = 4

Attribute VB_Name = "servermain"
Option Explicit

Public Type UserType
    DisplayName As String   ' user's name
    EntryID As String       ' entryid
    ReportIndex As Integer  ' index of the corresponding entry in aReport
End Type

Public Type UserListType
    cUsers As Integer       'number of elements in aUsers
    aUsers() As UserType
End Type

Public Type CategoryListType
    cCats As Integer        'number of elements in cCats
    aCats() As String
End Type


Global objSession As Object 'session object

Global UserList As UserListType  'list of all the users
Global CategoryList As CategoryListType 'for sending request
Global PayPeriod As Date 'for sending

Global Const UserListFile As String = "Users.dat"   'file to save users to
Global Const CatsListFile As String = "categs.dat"  'file to save categories to
Global Const ClientExePath As String = "d:\mapisamp\timecard.cli\client\tmcli.exe" 'path to the client executable
Global Const ClientExeName As String = "tmcli.exe"


Global Const mapiFileData As Integer = 1
Global Const E_NOT_FOUND As Integer = -1


Public Sub GetReceivIPCFolder(objFolder As Object)
'Finds the receiving folder for IPC messages, which is the
'top folder of the default message store.
'This is the only folder that is its own parent.

On Error GoTo error_olemsg

Dim objReceivFolder As Object
Dim objRecFolParent As Object
Dim parentid As String

If objSession Is Nothing Then
    MsgBox "Not logged on"
    Exit Sub
End If

Set objRecFolParent = objSession.inbox
If objRecFolParent Is Nothing Then
    Exit Sub
End If

Do
    Set objReceivFolder = objRecFolParent
    parentid = objReceivFolder.folderid 'get parent's id
    
    Set objRecFolParent = objSession.getfolder(parentid)
    
    If objRecFolParent Is Nothing Then Exit Sub 'error

Loop While Not objReceivFolder.id = parentid


Set objFolder = objReceivFolder

Exit Sub

error_olemsg:
    MsgBox "Error " & Str(err) & ": " & Error$(err)
    Resume Next

End Sub


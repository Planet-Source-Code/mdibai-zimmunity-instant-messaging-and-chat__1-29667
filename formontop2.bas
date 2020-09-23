Attribute VB_Name = "Module2"
Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

' =========================
Public RTChatRemoteIP As String
Public RTChatRemoteNick As String
Public RTCListen As Boolean
Public FileSendRemoteIP As String
Public FileSendRemoteNick As String
Public KeySection As String
Public KeyKey As String
Public KeyValue As String
Public RemoteNick As String
Public gFileNum As Long
Public RTChatTemp As String

Public MyPersonalInfo As MyPersonalData

Public Type MyPersonalData
    Sex As String * 7
    Country As String * 21
    BirthDay As String * 11
    Age As String * 4
    Webpage As String * 101
    About As String * 451
End Type

Declare Function WritePrivateProfileString _
Lib "KERNEL32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lsString As Any, _
ByVal lplFilename As String) As Long

Declare Function GetPrivateProfileString Lib _
"KERNEL32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As _
String) As Long

Public Sub LoadINI()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\Settings.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
'Call MsgBox("An error has occurred while calling the API function(The INI file probably doesn't exist)", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub

Public Sub SaveINI()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\Settings.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
'Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub


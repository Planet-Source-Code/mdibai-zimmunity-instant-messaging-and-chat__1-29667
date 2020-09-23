Attribute VB_Name = "Module3"
Declare Sub ReleaseCapture Lib "user32" ()


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long


Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Sub

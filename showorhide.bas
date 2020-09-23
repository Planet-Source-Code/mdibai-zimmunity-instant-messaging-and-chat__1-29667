Attribute VB_Name = "Module7"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Global showorhide As Integer
Global channel As String
Global zmess, zmessOLD, zmessb2, zmessOLDb2, zmessb3, zmessOLDb3, zmessb4, zmessOLDb4, zmessb5, zmessOLDb5, internetTEST, zmWELCOME As String
Global zmhighlight As BorderStyleConstants
Global zmcolor2 As BackStyleConstants
Global searchPoint As Integer
Global adtemp, adtext, linktext, link As String

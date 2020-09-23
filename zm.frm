VERSION 5.00
Begin VB.Form zm 
   BackColor       =   &H00A56E3A&
   BorderStyle     =   0  'None
   Caption         =   "NullStr"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   Icon            =   "zm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   103
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   0
      ScaleHeight     =   1543.237
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   0
      Top             =   0
      Width           =   4005
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   -300
         ScaleHeight     =   960
         ScaleWidth      =   4395
         TabIndex        =   1
         Top             =   540
         Width           =   4425
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   330
            TabIndex        =   2
            Top             =   -30
            Width           =   3885
            Begin VB.TextBox message 
               Appearance      =   0  'Flat
               BackColor       =   &H00A56E3A&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   180
               MaxLength       =   331
               TabIndex        =   3
               Top             =   240
               Width           =   3525
            End
            Begin VB.Label lblmess 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Show Messages  >>"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   1020
               MousePointer    =   99  'Custom
               TabIndex        =   6
               Top             =   600
               UseMnemonic     =   0   'False
               Width           =   1845
            End
            Begin VB.Label lblhide 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Hide"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   2940
               MousePointer    =   99  'Custom
               TabIndex        =   5
               Top             =   600
               UseMnemonic     =   0   'False
               Width           =   795
            End
            Begin VB.Label lblsend 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Send"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   120
               MousePointer    =   99  'Custom
               TabIndex        =   4
               Top             =   600
               UseMnemonic     =   0   'False
               Width           =   825
            End
            Begin VB.Shape shpsend 
               BackColor       =   &H00A56E3A&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H00A56E3A&
               Height          =   255
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   570
               Width           =   825
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00A56E3A&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               Height          =   285
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   210
               Width           =   3615
            End
            Begin VB.Shape shphide 
               BackColor       =   &H00A56E3A&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H00A56E3A&
               Height          =   255
               Left            =   2940
               Shape           =   4  'Rounded Rectangle
               Top             =   570
               Width           =   795
            End
            Begin VB.Shape shpmess 
               BackColor       =   &H00A56E3A&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H00A56E3A&
               Height          =   255
               Left            =   1020
               Shape           =   4  'Rounded Rectangle
               Top             =   570
               Width           =   1845
            End
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3600
         Top             =   1110
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -180
         ScaleHeight     =   735
         ScaleWidth      =   5055
         TabIndex        =   7
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox picTray 
            Height          =   225
            Left            =   450
            ScaleHeight     =   165
            ScaleWidth      =   1875
            TabIndex        =   12
            Top             =   510
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00A56E3A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3690
            ScaleHeight     =   195
            ScaleWidth      =   225
            TabIndex        =   9
            Top             =   30
            Width           =   225
            Begin VB.Label lblmin 
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   60
               TabIndex        =   11
               Top             =   -30
               Width           =   165
            End
            Begin VB.Shape shpmin 
               BorderColor     =   &H00000000&
               Height          =   285
               Left            =   30
               Shape           =   5  'Rounded Square
               Top             =   -30
               Width           =   165
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00A56E3A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3900
            ScaleHeight     =   195
            ScaleWidth      =   225
            TabIndex        =   8
            Top             =   30
            Width           =   225
            Begin VB.Label lblx 
               BackStyle       =   0  'Transparent
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   10
               Top             =   -60
               Width           =   165
            End
            Begin VB.Shape shpx 
               BorderColor     =   &H00000000&
               Height          =   285
               Left            =   30
               Shape           =   5  'Rounded Square
               Top             =   -30
               Width           =   165
            End
         End
      End
   End
End
Attribute VB_Name = "zm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zmformtop, zmformleft, RandomNumber, RandomAds, RandomAds2, tester As Integer

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)

On Error Resume Next

    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Sub

Sub Form_Load()

On Error GoTo makenewini

If App.PrevInstance Or FindWindow("Zimmunity", App.Title) > 0 Then
    AppTitle$ = App.Title
        
    App.Title = "#$#"

    zm.Caption = "#$#"
    AppActivate AppTitle$
    End
            
End If

startzm:

zm.Caption = App.Title

channel = " Zimmunity (General)"

AddIcon picTray, "Zimmunity v1.4.1 Beta", zm

AlwaysOnTop zm, True

Open App.Path & "\settings.ini" For Input As #1
Input #1, zmformtop, zmformleft, showorhide, channel, zmcolor2, zmhighlight
Close #1

zm.Top = zmformtop
zm.Left = zmformleft

zmcolor (zmcolor2)

zmm.messages.ForeColor = &H808080
message.ForeColor = &H808080

zmm.channel.Text = channel

Call SetAllBordersBlack

If showorhide = 1 Then

    zm.Show
    zmm.Show
    lblmess.Caption = "Hide Messages  <<"
    showorhide = 1

Else

    zm.Show
    zmm.Hide
    lblmess.Caption = "Show Messages  >>"
    showorhide = 2

End If

zmm.Left = zm.Left
zmm.Top = zm.Top + zm.Height + 60
        
zmm.messages.ForeColor = &H808080
message.ForeColor = &H808080
        
lblsend.Enabled = False
zmm.lbldisconnect.Enabled = False
zmm.lblconnect.Enabled = True
zmm.channel.Enabled = True
        
zmm.lbldisconnect.Visible = False
zmm.lblconnect.Visible = True

zmm.lblstatus.ForeColor = &H80&
zmm.lblstatus.Caption = "Disconnected."

zmm.shpgreen.BackColor = &HC0C0C0
zmm.shpyellow.BackColor = &HC0C0C0
zmm.shpred.BackColor = &HC0&

makenewini:

        Select Case Err.Number
        
            Case 53
            
                Open App.Path & "\settings.ini" For Output As #1
                Write #1, 400, 400, 1, channel, &HA56E3A, &HFFFF00
                Close #1
                Resume
                
            Case Else
                
                Resume Next
            
        End Select

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub


Sub lblhide_Click()

On Error Resume Next

Call SetAllBordersBlack

zmm.Hide
zm.Hide

End Sub

Private Sub lblhide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.lblmin.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    ' end
    '*****************************

shphide.BorderColor = zmhighlight

End Sub

Private Sub lblmess_Click()

On Error Resume Next

If lblmess.Caption = "Show Messages  >>" Then

    zmm.Show
    lblmess.Caption = "Hide Messages  <<"
    zmm.messages.SelStart = Len(zmm.messages.Text)
    showorhide = 1
    
ElseIf lblmess.Caption = "Hide Messages  <<" Then

    zmm.Hide
    lblmess.Caption = "Show Messages  >>"
    showorhide = 2
    
End If

message.SetFocus

End Sub

Private Sub lblmess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.lblmin.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************
    
shpmess.BorderColor = zmhighlight

End Sub

Private Sub lblmin_Click()

On Error Resume Next

    Call lblhide_Click

End Sub

Private Sub lblmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

shpmin.BorderColor = zmhighlight
lblmin.ForeColor = zmhighlight

End Sub

Private Sub lblsend_Click()

On Error Resume Next
    
If message.Text = "" Or zmess <> zmessOLD Or zmessb2 <> zmessOLDb2 Or zmessb3 <> zmessOLDb3 Or zmessb4 <> zmessOLDb4 Or zmessb5 <> zmessOLDb5 Then
    message.Text = ""
Else
    
End If


End Sub


Private Sub lblsend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.lblmin.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

shpsend.BorderColor = zmhighlight

End Sub

Sub lblx_Click()

On Error GoTo makenewini2

zmformtop = zm.Top
zmformleft = zm.Left

Open App.Path & "\settings.ini" For Output As #1
Write #1, zmformtop, zmformleft, showorhide, channel, zmcolor2, zmhighlight
Close #1
        
Timer1.Enabled = False

DeleteIcon picTray
        
Unload zmabout
Unload zmm
Unload zm

Set zmabout = Nothing
Set zmm = Nothing
Set zm = Nothing

makenewini2:

        Select Case Err.Number
        
            Case 55
                Close #1
                Resume
                
            Case Else
                
                Resume Next
            
        End Select

End
    
End Sub

Private Sub lblx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblmin.ForeColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

shpx.BorderColor = zmhighlight
lblx.ForeColor = zmhighlight
End Sub

Private Sub message_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then

    If lblsend.Enabled = True Then
        Call lblsend_Click
        message.SetFocus
    End If
    
    KeyAscii = 0

End If
    
End Sub


Private Sub message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call SetAllBordersBlack

End Sub

Private Sub Picture4_Click()
    
On Error Resume Next

    Call lblx_Click

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblmin.ForeColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

    shpx.BorderColor = zmhighlight
    lblx.ForeColor = zmhighlight

End Sub

Private Sub Picture5_Click()

On Error Resume Next

    Call lblhide_Click

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

    shpmin.BorderColor = zmhighlight
    lblmin.ForeColor = zmhighlight

End Sub

Private Sub Timer1_Timer()

On Error Resume Next

zmm.Left = zm.Left
zmm.Top = zm.Top + zm.Height + 60
    
End Sub

Sub ConnecttoZimmunity()

End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error Resume Next
    
    Select Case X
        Case trayLBUTTONDOWN
        
            Call SetAllBordersBlack
            
            zm.Show
            
            If zm.lblmess.Caption = "Hide Messages  <<" Then
                zmm.Show
                zmm.messages.SelStart = Len(zmm.messages.Text)
            Else
                zmm.Hide
            End If

            zm.message.SetFocus
            
        Case trayRBUTTONDOWN
        
            Call SetAllBordersBlack
            
            zm.Show
            
            If zm.lblmess.Caption = "Hide Messages  <<" Then
                zmm.Show
                zmm.messages.SelStart = Len(zmm.messages.Text)
            Else
                zmm.Hide
            End If

            zm.message.SetFocus
            
        Case trayMOUSEMOVE
        
            Call SetAllBordersBlack
            
        Case Else
        
    End Select
    
End Sub



Sub SetAllBordersBlack()

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    zmm.shpabout.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.lblmin.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

End Sub

Sub zmcolor(zmcolor As BackStyleConstants)

On Error Resume Next

zmm.messages.BackColor = zmcolor
zmm.shpabout.BackColor = zmcolor
zmm.shpclear.BackColor = zmcolor
zmm.shpconnect.BackColor = zmcolor
zmm.shpexit.BackColor = zmcolor
zmm.Picture1.BackColor = zmcolor
zmm.Picture3.BackColor = zmcolor
zmm.channel.BackColor = zmcolor
zm.Picture1.BackColor = zmcolor
zm.Picture3.BackColor = zmcolor
zm.Picture4.BackColor = zmcolor
zm.Picture5.BackColor = zmcolor
zm.shphide.BackColor = zmcolor
zm.shpmess.BackColor = zmcolor
zm.shpmin.BackColor = zmcolor
zm.shpsend.BackColor = zmcolor
zm.Shape1.BackColor = zmcolor
zm.message.BackColor = zmcolor
zmabout.Picture1.BackColor = zmcolor
zmabout.shpclose.BackColor = zmcolor
zmabout.Picture3.BackColor = zmcolor
zmabout.Picture4.BackColor = zmcolor
zmabout.Picture2.BackColor = zmcolor
zmabout.picclose.BackColor = zmcolor
zmabout.picgeneral.BackColor = zmcolor
zmabout.picsupport.BackColor = zmcolor
zmabout.piccopyright.BackColor = zmcolor
zmabout.BackColor = zmcolor
zmm.BackColor = zmcolor
zm.BackColor = zmcolor


End Sub

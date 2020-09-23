VERSION 5.00
Begin VB.Form zmabout 
   Appearance      =   0  'Flat
   BackColor       =   &H00A56E3A&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   Icon            =   "zmabout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4365
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   135
         ScaleHeight     =   3165
         ScaleWidth      =   5250
         TabIndex        =   11
         Top             =   1065
         Visible         =   0   'False
         Width           =   5250
         Begin VB.Shape shpsupport2 
            Height          =   3165
            Left            =   0
            Top             =   0
            Width           =   5250
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   135
         ScaleHeight     =   3165
         ScaleWidth      =   5250
         TabIndex        =   1
         Top             =   1065
         Visible         =   0   'False
         Width           =   5250
         Begin VB.Shape shpcopyright2 
            Height          =   3165
            Left            =   0
            Top             =   0
            Width           =   5250
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   135
         ScaleHeight     =   3165
         ScaleWidth      =   5250
         TabIndex        =   10
         Top             =   1065
         Visible         =   0   'False
         Width           =   5250
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   165
            ScaleHeight     =   435
            ScaleWidth      =   4860
            TabIndex        =   25
            Top             =   1065
            Width           =   4890
            Begin VB.PictureBox Picture20 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1380
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   31
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture19 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   945
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   30
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture18 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   510
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   29
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture17 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFF00&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   75
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   28
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture16 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF00FF&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1815
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   27
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture15 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2250
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   26
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   165
            ScaleHeight     =   435
            ScaleWidth      =   4860
            TabIndex        =   13
            Top             =   300
            Width           =   4890
            Begin VB.PictureBox Picture8 
               Appearance      =   0  'Flat
               BackColor       =   &H00000040&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   4425
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   24
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               BackColor       =   &H00337878&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3990
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   23
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox picdarkblue 
               Appearance      =   0  'Flat
               BackColor       =   &H00663300&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3555
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   22
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox Picture6 
               Appearance      =   0  'Flat
               BackColor       =   &H00963333&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3120
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   21
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox picdarkpurple 
               Appearance      =   0  'Flat
               BackColor       =   &H00663737&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2685
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   20
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox picgrey 
               Appearance      =   0  'Flat
               BackColor       =   &H00333333&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2250
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   19
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox piclightred 
               Appearance      =   0  'Flat
               BackColor       =   &H0033338C&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1815
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   18
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox piclightblue 
               Appearance      =   0  'Flat
               BackColor       =   &H00A56E3A&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   75
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   17
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox pictur 
               Appearance      =   0  'Flat
               BackColor       =   &H00787833&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   510
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   16
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox picgreen 
               Appearance      =   0  'Flat
               BackColor       =   &H00335633&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   945
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   15
               Top             =   60
               Width           =   360
            End
            Begin VB.PictureBox picpurple 
               Appearance      =   0  'Flat
               BackColor       =   &H00996666&
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1380
               ScaleHeight     =   285
               ScaleWidth      =   330
               TabIndex        =   14
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.Shape shpgeneral2 
            Height          =   3165
            Left            =   0
            Top             =   0
            Width           =   5250
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Highlight Color"
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
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Program Color"
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
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   75
            Width           =   1155
         End
      End
      Begin VB.PictureBox picclose 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5280
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   8
         Top             =   30
         Width           =   225
         Begin VB.Shape shpclose 
            BorderColor     =   &H00000000&
            Height          =   285
            Left            =   30
            Shape           =   5  'Rounded Square
            Top             =   -30
            Width           =   165
         End
         Begin VB.Label lblclose 
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
            TabIndex        =   9
            Top             =   -60
            Width           =   165
         End
      End
      Begin VB.PictureBox piccopyright 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2985
         ScaleHeight     =   375
         ScaleWidth      =   1365
         TabIndex        =   2
         Top             =   705
         Width           =   1365
         Begin VB.Shape shpcopyright 
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblcopyright 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright"
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
            Height          =   195
            Left            =   -15
            TabIndex        =   5
            Top             =   90
            Width           =   1365
         End
      End
      Begin VB.PictureBox picgeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         ScaleHeight     =   375
         ScaleWidth      =   1365
         TabIndex        =   4
         Top             =   705
         Width           =   1365
         Begin VB.Shape shpgeneral 
            BackColor       =   &H00A56E3A&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblgeneral 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "General"
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
            Height          =   195
            Left            =   0
            TabIndex        =   7
            Top             =   90
            Width           =   1365
         End
      End
      Begin VB.PictureBox picsupport 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   1365
         TabIndex        =   3
         Top             =   705
         Width           =   1365
         Begin VB.Shape shpsupport 
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1365
         End
         Begin VB.Label lblsupport 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Support"
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
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   90
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "zmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Form_Load()

On Error Resume Next

zm.zmcolor (zmcolor2)

AlwaysOnTop zmabout, True

shpgeneral.BorderColor = zmhighlight
shpgeneral2.BorderColor = zmhighlight
Picture3.Visible = True
Picture4.Visible = False
Picture1.Visible = False

End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call zm.SetAllBordersBlack
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Call zm.SetAllBordersBlack
End Sub

Private Sub Label14_Click()
On Error Resume Next

OpenWebsite ("mailto:support@zetamicrosystems.com")

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub lblads_Click()

On Error Resume Next

OpenWebsite ("mailto:advertise@zetamicrosystems.com")

End Sub

Private Sub lblads_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblads.ForeColor = &HC0C0C0

End Sub

Private Sub lblbugs_Click()

On Error Resume Next

OpenWebsite ("mailto:bugs@zetamicrosystems.com")

End Sub

Private Sub lblbugs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblbugs.ForeColor = &HC0C0C0

End Sub

Private Sub lblchannels_Click()
On Error Resume Next

OpenWebsite ("mailto:channels@zetamicrosystems.com")
End Sub

Private Sub lblchannels_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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

Private Sub lblclose_Click()

On Error Resume Next

Unload zmabout

End Sub

Private Sub lblclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

shpclose.BorderColor = zmhighlight
lblclose.ForeColor = zmhighlight

End Sub

Sub lblcopyright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

shpgeneral.BorderColor = &H0&
shpgeneral2.BorderColor = &H0&
shpsupport.BorderColor = &H0&
shpsupport2.BorderColor = &H0&
shpcopyright.BorderColor = zmhighlight
shpcopyright2.BorderColor = zmhighlight


Picture1.Visible = True
Picture3.Visible = False
Picture4.Visible = False

End Sub

Private Sub lblemail1_Click()

On Error Resume Next

OpenWebsite ("mailto: questions@zetamicrosystems.com")

End Sub

Private Sub lblemail1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblemail1.ForeColor = &HC0C0C0

End Sub



Private Sub lblemail2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblemail2.ForeColor = &HC0C0C0

End Sub

Sub lblgeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

shpgeneral.BorderColor = zmhighlight
shpgeneral2.BorderColor = zmhighlight
shpsupport.BorderColor = &H0&
shpsupport2.BorderColor = &H0&
shpcopyright.BorderColor = &H0&
shpcopyright2.BorderColor = &H0&

Picture3.Visible = True
Picture1.Visible = False
Picture4.Visible = False

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
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

End Sub

Private Sub lblsup_Click()
On Error Resume Next

OpenWebsite ("mailto:support@zetamicrosystems.com")
End Sub

Private Sub lblsup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblsup.ForeColor = &HC0C0C0

End Sub

Sub lblsupport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

shpgeneral.BorderColor = &H0&
shpgeneral2.BorderColor = &H0&
shpsupport.BorderColor = zmhighlight
shpsupport2.BorderColor = zmhighlight
shpcopyright.BorderColor = &H0&
shpcopyright2.BorderColor = &H0&

Picture4.Visible = True
Picture3.Visible = False
Picture1.Visible = False

End Sub

Private Sub lblweb1_Click()

On Error Resume Next

OpenWebsite ("http://www.zetamicrosystems.com")

End Sub

Private Sub lblweb1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblweb1.ForeColor = &HC0C0C0

End Sub

Private Sub lblweb2_Click()

On Error Resume Next

OpenWebsite ("http://www.zimmunity.com")

End Sub

Private Sub lblweb2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

lblweb2.ForeColor = &HC0C0C0

End Sub

Private Sub picclose_Click()

On Error Resume Next

Unload zmabout

End Sub

Private Sub picclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

shpclose.BorderColor = zmhighlight
lblclose.ForeColor = zmhighlight

End Sub

Private Sub piccopyright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

shpgeneral.BorderColor = &H0&
shpgeneral2.BorderColor = &H0&
shpsupport.BorderColor = &H0&
shpsupport2.BorderColor = &H0&
shpcopyright.BorderColor = zmhighlight
shpcopyright2.BorderColor = zmhighlight


Picture1.Visible = True
Picture3.Visible = False
Picture4.Visible = False

End Sub

Private Sub picdarkblue_Click()

On Error Resume Next

zmcolor2 = &H663300
zm.zmcolor (&H663300)
End Sub

Private Sub picdarkpurple_Click()

On Error Resume Next

zmcolor2 = &H663737
zm.zmcolor (&H663737)
End Sub

Private Sub picgeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

shpgeneral.BorderColor = zmhighlight
shpgeneral2.BorderColor = zmhighlight
shpsupport.BorderColor = &H0&
shpsupport2.BorderColor = &H0&
shpcopyright.BorderColor = &H0&
shpcopyright2.BorderColor = &H0&

Picture3.Visible = True
Picture1.Visible = False
Picture4.Visible = False

End Sub

Private Sub picgreen_Click()

On Error Resume Next

zmcolor2 = &H335633
zm.zmcolor (&H335633)
End Sub

Private Sub picgrey_Click()

On Error Resume Next

zmcolor2 = &H333333
zm.zmcolor (&H333333)
End Sub

Private Sub piclightblue_Click()
zmcolor2 = &HA56E3A
zm.zmcolor (&HA56E3A)
End Sub

Private Sub piclightred_Click()
zmcolor2 = &H33338C
zm.zmcolor (&H33338C)
End Sub

Private Sub picpurple_Click()

On Error Resume Next

zmcolor2 = &H996666
zm.zmcolor (&H996666)
End Sub

Private Sub picsupport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

shpgeneral.BorderColor = &H0&
shpgeneral2.BorderColor = &H0&
shpsupport.BorderColor = zmhighlight
shpsupport2.BorderColor = zmhighlight
shpcopyright.BorderColor = &H0&
shpcopyright2.BorderColor = &H0&

Picture4.Visible = True
Picture3.Visible = False
Picture1.Visible = False

End Sub

Private Sub pictur_Click()

On Error Resume Next

zmcolor2 = &H787833
zm.zmcolor (&H787833)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Picture15_Click()

On Error Resume Next

shpgeneral.BorderColor = &HFF&
shpgeneral2.BorderColor = &HFF&
zmhighlight = &HFF&
End Sub

Private Sub Picture16_Click()

On Error Resume Next

shpgeneral.BorderColor = &HFF00FF
shpgeneral2.BorderColor = &HFF00FF
zmhighlight = &HFF00FF
End Sub

Private Sub Picture17_Click()

On Error Resume Next

shpgeneral.BorderColor = &HFFFF00
shpgeneral2.BorderColor = &HFFFF00
zmhighlight = &HFFFF00
End Sub

Private Sub Picture18_Click()

On Error Resume Next

shpgeneral.BorderColor = &HFF00&
shpgeneral2.BorderColor = &HFF00&
zmhighlight = &HFF00&
End Sub

Private Sub Picture19_Click()

On Error Resume Next

shpgeneral.BorderColor = &HFFFF&
shpgeneral2.BorderColor = &HFFFF&
zmhighlight = &HFFFF&
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

formdrag Me
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Picture20_Click()

On Error Resume Next

shpgeneral.BorderColor = &H80FF&
shpgeneral2.BorderColor = &H80FF&
zmhighlight = &H80FF&
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack
End Sub


Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack
End Sub

Private Sub Picture6_Click()

On Error Resume Next

zmcolor2 = &H963333
zm.zmcolor (&H963333)
End Sub

Private Sub Picture7_Click()

On Error Resume Next

zmcolor2 = &H337878
zm.zmcolor (&H337878)
End Sub

Private Sub Picture8_Click()

On Error Resume Next

zmcolor2 = &H40&
zm.zmcolor (&H40&)
End Sub

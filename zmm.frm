VERSION 5.00
Begin VB.Form zmm 
   Appearance      =   0  'Flat
   BackColor       =   &H00A56E3A&
   BorderStyle     =   0  'None
   ClientHeight    =   6045
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
   ForeColor       =   &H00000000&
   Icon            =   "zmm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   3975
      TabIndex        =   7
      Top             =   5760
      Width           =   4005
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   30
         Width           =   645
      End
      Begin VB.Shape shpgreen 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   135
         Left            =   3300
         Shape           =   3  'Circle
         Top             =   60
         Width           =   255
      End
      Begin VB.Shape shpyellow 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   135
         Left            =   3510
         Shape           =   3  'Circle
         Top             =   60
         Width           =   255
      End
      Begin VB.Shape shpred 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   135
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblstatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   30
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   3975
      TabIndex        =   2
      Top             =   4740
      Width           =   4005
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   60
         TabIndex        =   3
         Top             =   -30
         Width           =   3855
         Begin VB.ComboBox channel 
            BackColor       =   &H00A56E3A&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            ItemData        =   "zmm.frx":000C
            Left            =   1320
            List            =   "zmm.frx":0049
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   210
            Width           =   2385
         End
         Begin VB.Label lblabout 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
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
            Left            =   2040
            TabIndex        =   12
            Top             =   630
            UseMnemonic     =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpabout 
            BackColor       =   &H00A56E3A&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00A56E3A&
            Height          =   255
            Left            =   2040
            Shape           =   4  'Rounded Rectangle
            Top             =   600
            Width           =   825
         End
         Begin VB.Label lbldisconnect 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Disconnect"
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
            TabIndex        =   9
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblclear 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Clear Messages"
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
            TabIndex        =   6
            Top             =   630
            UseMnemonic     =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblexit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
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
            TabIndex        =   5
            Top             =   630
            UseMnemonic     =   0   'False
            Width           =   765
         End
         Begin VB.Label lblconnect 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Connect"
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
            TabIndex        =   4
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   1125
         End
         Begin VB.Shape shpclear 
            BackColor       =   &H00A56E3A&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00A56E3A&
            Height          =   255
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   600
            Width           =   1845
         End
         Begin VB.Shape shpexit 
            BackColor       =   &H00A56E3A&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00A56E3A&
            Height          =   255
            Left            =   2940
            Shape           =   4  'Rounded Rectangle
            Top             =   600
            Width           =   765
         End
         Begin VB.Shape shpconnect 
            BackColor       =   &H00A56E3A&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00A56E3A&
            Height          =   255
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   1125
         End
      End
   End
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
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5114.7
      ScaleMode       =   0  'User
      ScaleWidth      =   3994
      TabIndex        =   0
      Top             =   0
      Width           =   4005
      Begin VB.TextBox messages 
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
         Height          =   4725
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   0
         Width           =   3915
      End
   End
End
Attribute VB_Name = "zmm"
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

Private Sub Combo1_GotFocus()

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Form_Load()

On Error Resume Next

zm.zmcolor (zmcolor2)

AlwaysOnTop zmm, True

Left = zm.Left
Top = zm.Top + zm.Height + 60

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub lblclear_Click()

On Error Resume Next

messages.Text = ""
zm.message.SetFocus

End Sub

Private Sub lblclear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
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

shpclear.BorderColor = zmhighlight

End Sub

Private Sub lblconnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
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

shpconnect.BorderColor = zmhighlight

End Sub

Sub lbldisconnect_Click()

End Sub

Private Sub lbldisconnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
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

shpconnect.BorderColor = zmhighlight

End Sub

Sub lblexit_Click()

On Error Resume Next

    Call zm.lblx_Click

End Sub

Private Sub lblexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
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

shpexit.BorderColor = zmhighlight

End Sub

Private Sub lblabout_Click()

On Error Resume Next

zmabout.Show

End Sub

Private Sub lblabout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error Resume Next
    
    '*****************************
    ' set all borders black
    
    zmm.shpconnect.BorderColor = &H0&
    zmm.shpclear.BorderColor = &H0&
    zmm.shpexit.BorderColor = &H0&
    
    zm.lblx.ForeColor = &H0&
    zm.lblmin.ForeColor = &H0&
    zm.shpx.BorderColor = &H0&
    zm.shpmin.BorderColor = &H0&
    zm.shpsend.BorderColor = &H0&
    zm.shpmess.BorderColor = &H0&
    zm.shphide.BorderColor = &H0&
    ' end
    '*****************************

shpabout.BorderColor = zmhighlight
End Sub

Private Sub lblstatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub messages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

Call zm.SetAllBordersBlack

End Sub


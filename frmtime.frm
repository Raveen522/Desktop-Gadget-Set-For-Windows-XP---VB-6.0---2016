VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtime 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   1575
   ClientLeft      =   12930
   ClientTop       =   285
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3900
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20709378
      CurrentDate     =   36494
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Alarm"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image16 
      Height          =   495
      Left            =   4320
      Picture         =   "frmtime.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image15 
      Height          =   495
      Left            =   4320
      Picture         =   "frmtime.frx":2442
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   345
      Left            =   1320
      Picture         =   "frmtime.frx":4884
      Stretch         =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   1540
      Left            =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   350
      Left            =   1850
      TabIndex        =   16
      Top             =   80
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image14 
      Height          =   495
      Left            =   3600
      Picture         =   "frmtime.frx":6CC6
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   2760
      Picture         =   "frmtime.frx":9108
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   350
      Left            =   840
      Picture         =   "frmtime.frx":B54A
      Stretch         =   -1  'True
      Top             =   80
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image10 
      Height          =   350
      Left            =   240
      Picture         =   "frmtime.frx":D98C
      Stretch         =   -1  'True
      Top             =   80
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   3600
      Picture         =   "frmtime.frx":FDCE
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   2760
      Picture         =   "frmtime.frx":12210
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   480
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1720
      _cy             =   873
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Browse"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Alarm"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   330
      Left            =   2880
      Picture         =   "frmtime.frx":14652
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Image Image5 
      Height          =   330
      Left            =   2880
      Picture         =   "frmtime.frx":15926
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   120
      Picture         =   "frmtime.frx":16BFA
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1200
      Picture         =   "frmtime.frx":17ECE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "frmtime.frx":191A2
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Height          =   75
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   2055
      Left            =   0
      Picture         =   "frmtime.frx":1A476
      Stretch         =   -1  'True
      Top             =   1635
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   345
      Left            =   70
      TabIndex        =   7
      Top             =   80
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Height          =   350
      Left            =   70
      TabIndex        =   6
      Top             =   80
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblampm 
      BackStyle       =   0  'Transparent
      Caption         =   "am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblseconds 
      BackStyle       =   0  'Transparent
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lbl2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      Caption         =   "08:55:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgback 
      Height          =   1575
      Left            =   0
      Picture         =   "frmtime.frx":66BDA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
DTPicker1.Value = "12:00:00 AM"
End Sub

Private Sub Image10_Click()
frmtime.Height = 3740
Shape1.Height = 3740
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Picture = Image14.Picture
End Sub

Private Sub Image11_Click()
frmsettings.Show
ProjectDesktopGagets.frmsettings.SSTab1.Tab = 1
 
ProjectDesktopGagets.frmsettings.Line16.Visible = True
ProjectDesktopGagets.frmsettings.Line13.X1 = 960
ProjectDesktopGagets.frmsettings.Line13.X2 = 1680
ProjectDesktopGagets.frmsettings.Line15.Visible = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Picture = Image8.Picture
End Sub

Private Sub Image12_Click()
frmstopwatch.Show
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Picture = Image15.Picture
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF00&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Visible = False
Image11.Picture = Image9.Picture
Image11.Visible = False
Label9.ForeColor = &HFF&
Label9.Visible = False
Image10.Visible = False
Image10.Picture = Image13.Picture
Image12.Visible = False
End Sub

Private Sub Label10_Click()
Me.Height = 1545
Shape1.Height = 1540
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H8000&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Visible = True
Image11.Visible = True
Label9.Visible = True
Image10.Visible = True
Image12.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Picture = Image9.Picture
Label9.ForeColor = &HFF&
Image10.Picture = Image13.Picture
Image12.Picture = Image16.Picture
End Sub

Private Sub Label5_Click()
Timer2.Enabled = True
DTPicker1.Enabled = False
frmtime.Height = 1545
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image5.Picture
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image6.Picture
End Sub

Private Sub Label6_Click()
Timer2.Enabled = False
DTPicker1.Enabled = True
WindowsMediaPlayer1.Controls.stop
DTPicker1.Value = "12:00:00 AM"
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = Image5.Picture
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = Image6.Picture
End Sub

Private Sub Label7_Click()
CommonDialog1.Filter = "Sound files|*.mp3;*.wav"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = Image5.Picture
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = Image6.Picture
End Sub



Private Sub Label8_Click()
Label8.Visible = False
Timer2.Enabled = False
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub Label9_Click()
Unload Me
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &H80&
End Sub

Private Sub Timer1_Timer()
lbltime.Caption = FormatDateTime(Time, vbShortTime) & ":"
lbl1.Caption = Time
a = Mid(lbl1.Caption, 9, 11)
lbl2.Caption = a
lblseconds.Caption = Second(Time)

If lbl2.Caption = " AM" Then
lblampm.Caption = "AM"
End If
If lbl2.Caption = "AM" Then
lblampm.Caption = "AM"
End If
If lbl2.Caption = " PM" Then
lblampm.Caption = "PM"
End If
If lbl2.Caption = "PM" Then
lblampm.Caption = "PM"
End If
If lblseconds.Visible = True Then
lbltime.Caption = FormatDateTime(Time, vbShortTime) & ":"
Else
lbltime.Caption = FormatDateTime(Time, vbShortTime)
End If
End Sub

Private Sub Timer2_Timer()
If lbl1.Caption = DTPicker1.Value Then
WindowsMediaPlayer1.URL = Text1.Text
Label8.Visible = True

End If

End Sub


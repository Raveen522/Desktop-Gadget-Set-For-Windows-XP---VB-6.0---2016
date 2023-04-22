VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmplayer 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   1905
   ClientLeft      =   12750
   ClientTop       =   3705
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   1905
      Left            =   0
      Top             =   0
      Width           =   2610
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1320
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   2655
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
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4683
      _cy             =   2328
   End
   Begin VB.Image Image25 
      Height          =   375
      Left            =   4440
      Picture         =   "frmplayer.frx":0000
      Stretch         =   -1  'True
      Top             =   1490
      Width           =   375
   End
   Begin VB.Image Image24 
      Height          =   375
      Left            =   3840
      Picture         =   "frmplayer.frx":0342
      Stretch         =   -1  'True
      Top             =   1490
      Width           =   375
   End
   Begin VB.Image Image23 
      Height          =   465
      Left            =   3240
      Picture         =   "frmplayer.frx":0654
      Stretch         =   -1  'True
      Top             =   1400
      Width           =   495
   End
   Begin VB.Image Image22 
      Height          =   495
      Left            =   2685
      Picture         =   "frmplayer.frx":0D0E
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   495
   End
   Begin VB.Image Image21 
      Height          =   495
      Left            =   4320
      Picture         =   "frmplayer.frx":1C14
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image20 
      Height          =   495
      Left            =   4320
      Picture         =   "frmplayer.frx":1EF6
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image19 
      Height          =   495
      Left            =   4320
      Picture         =   "frmplayer.frx":21A0
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image Image18 
      Height          =   495
      Left            =   4320
      Picture         =   "frmplayer.frx":26CE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   3480
      Picture         =   "frmplayer.frx":3370
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image16 
      Height          =   495
      Left            =   3480
      Picture         =   "frmplayer.frx":4832
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image15 
      Height          =   615
      Left            =   3480
      Picture         =   "frmplayer.frx":5BD4
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image14 
      Height          =   495
      Left            =   3480
      Picture         =   "frmplayer.frx":7B06
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   2640
      Picture         =   "frmplayer.frx":BEF8
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   2640
      Picture         =   "frmplayer.frx":C23A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   2640
      Picture         =   "frmplayer.frx":C54C
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   2640
      Picture         =   "frmplayer.frx":CC06
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   615
      Left            =   2640
      Picture         =   "frmplayer.frx":DB0C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Image imgback 
      Height          =   1305
      Left            =   0
      Picture         =   "frmplayer.frx":15CAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2625
   End
   Begin VB.Image imgopen 
      Height          =   255
      Left            =   1430
      Picture         =   "frmplayer.frx":165BE
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgstop 
      Height          =   255
      Left            =   1070
      Picture         =   "frmplayer.frx":17A00
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgpuase 
      Height          =   255
      Left            =   700
      Picture         =   "frmplayer.frx":18E42
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgplay 
      Height          =   495
      Left            =   100
      Picture         =   "frmplayer.frx":1A284
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   600
      Picture         =   "frmplayer.frx":1B6C6
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   0
      Picture         =   "frmplayer.frx":1CB08
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   600
      Picture         =   "frmplayer.frx":1DF4A
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Picture         =   "frmplayer.frx":1F38C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   600
      Picture         =   "frmplayer.frx":207CE
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmplayer.frx":21C10
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   600
      Picture         =   "frmplayer.frx":23052
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmplayer.frx":24494
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "frmplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Image22_Click()
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub Image22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Picture = Image18.Picture
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Picture = Image14.Picture
End Sub

Private Sub Image22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Picture = Image10.Picture
End Sub

Private Sub Image23_Click()
WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub Image23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.Picture = Image19.Picture
End Sub

Private Sub Image23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.Picture = Image15.Picture
End Sub

Private Sub Image23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.Picture = Image11.Picture
End Sub

Private Sub Image24_Click()
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub Image24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.Picture = Image20.Picture
End Sub

Private Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.Picture = Image16.Picture
End Sub

Private Sub Image24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.Picture = Image12.Picture
End Sub

Private Sub Image25_Click()
On Error GoTo errhandler:
CommonDialog1.ShowOpen
WindowsMediaPlayer1.URL = CommonDialog1.FileName
errhandler:
Exit Sub
End Sub

Private Sub Image25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Picture = Image21.Picture
End Sub

Private Sub Image25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Picture = Image17.Picture
End Sub

Private Sub Image25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Picture = Image13.Picture
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Picture = Image10.Picture
Image23.Picture = Image11.Picture
Image24.Picture = Image12.Picture
Image25.Picture = Image13.Picture
End Sub

Private Sub imgopen_Click()
On Error GoTo errhandler:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
MMControl1.Command = "Close"
MMControl1.FileName = CommonDialog1.FileName
errhandler:
Exit Sub
End Sub

Private Sub imgopen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgopen.Picture = Image7.Picture
End Sub

Private Sub imgopen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgopen.Picture = Image8.Picture
End Sub

Private Sub imgplay_Click()
MMControl1.FileName = CommonDialog1.FileName
MMControl1.Command = "Open"

MMControl1.Command = "Prev"
MMControl1.Command = "Play"
End Sub

Private Sub imgplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgplay.Picture = Image1.Picture
End Sub

Private Sub imgplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgplay.Picture = Image2.Picture
End Sub

Private Sub imgpuase_Click()
MMControl1.Command = "Pause"
End Sub

Private Sub imgpuase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgpuase.Picture = Image3.Picture
End Sub

Private Sub imgpuase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgpuase.Picture = Image4.Picture
End Sub

Private Sub imgstop_Click()
MMControl1.Command = "Stop"
End Sub

Private Sub imgstop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgstop.Picture = Image5.Picture
End Sub

Private Sub imgstop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgstop.Picture = Image6.Picture
End Sub

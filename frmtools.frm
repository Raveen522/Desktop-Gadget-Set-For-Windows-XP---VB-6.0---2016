VERSION 5.00
Begin VB.Form frmtools 
   BorderStyle     =   0  'None
   ClientHeight    =   825
   ClientLeft      =   12945
   ClientTop       =   10050
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   6
      FillColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image7 
      Height          =   720
      Left            =   1440
      Picture         =   "frmtools.frx":0000
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   120
      Picture         =   "frmtools.frx":0D42
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   1440
      Picture         =   "frmtools.frx":1A84
      Top             =   1080
      Width           =   720
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   120
      Picture         =   "frmtools.frx":27C6
      Top             =   1080
      Width           =   720
   End
   Begin VB.Image imgrestart 
      Height          =   600
      Left            =   1560
      Picture         =   "frmtools.frx":3508
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image imgshutdown 
      Height          =   600
      Left            =   240
      Picture         =   "frmtools.frx":424A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image imgback 
      Height          =   855
      Left            =   0
      Picture         =   "frmtools.frx":4F8C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmtools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowEx Lib "User32" (ByVal dwoptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2


Private Sub imgrestart_Click()
'm = MsgBox("are you want to Restart your computer", vbYesNo + vbQuestion, "Restart")
'If m - vbYes Then
'lng Result = ExitWindowEx(EWX_REBOOT, 0&)
'Else
'End If
End Sub

Private Sub imgrestart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgrestart.Picture = Image5.Picture
End Sub
Private Sub imgrestart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgrestart.Picture = Image7.Picture
End Sub

Private Sub imgshutdown_Click()
'm = MsgBox("Are want to Shutdown your computer", vbYesNo + vbQuestion, "Shutdown")
'If m = vbYes Then
'lng Result = ExitWindowEx(EWX_SHUTDOWN, 0&)
'Else
'End If
End Sub

Private Sub imgshutdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgshutdown.Picture = Image4.Picture
End Sub
Private Sub imgshutdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgshutdown.Picture = Image6.Picture
End Sub

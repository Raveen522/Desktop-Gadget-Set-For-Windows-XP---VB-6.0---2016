VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmoffice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   12810
   ClientTop       =   7155
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   1470
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image Image29 
      Height          =   255
      Left            =   2640
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image28 
      Height          =   255
      Left            =   2280
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image27 
      Height          =   255
      Left            =   1920
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image26 
      Height          =   255
      Left            =   1560
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image25 
      Height          =   255
      Left            =   1200
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image24 
      Height          =   255
      Left            =   840
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image23 
      Height          =   255
      Left            =   480
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image22 
      Height          =   350
      Left            =   1440
      Picture         =   "frmoffice.frx":0000
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   350
   End
   Begin VB.Image Image21 
      Height          =   350
      Left            =   960
      Picture         =   "frmoffice.frx":12F5A
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   350
   End
   Begin VB.Image Image20 
      Height          =   255
      Left            =   1920
      Picture         =   "frmoffice.frx":25EB4
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image19 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image18 
      Height          =   255
      Left            =   2640
      Picture         =   "frmoffice.frx":38D76
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image17 
      Height          =   255
      Left            =   2280
      Picture         =   "frmoffice.frx":4B7D8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image16 
      Height          =   255
      Left            =   1560
      Picture         =   "frmoffice.frx":5ED86
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image15 
      Height          =   255
      Left            =   1200
      Picture         =   "frmoffice.frx":725B8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image14 
      Height          =   255
      Left            =   840
      Picture         =   "frmoffice.frx":85512
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image13 
      Height          =   255
      Left            =   480
      Picture         =   "frmoffice.frx":9846C
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image12 
      Height          =   255
      Left            =   120
      Picture         =   "frmoffice.frx":AB82E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image img16 
      Height          =   615
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image img13 
      Height          =   615
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image img7 
      Height          =   615
      Left            =   4080
      Picture         =   "frmoffice.frx":BF958
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image img10 
      Height          =   615
      Left            =   2760
      Picture         =   "frmoffice.frx":D6D4C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   350
      Left            =   1920
      Picture         =   "frmoffice.frx":2BF5F6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   350
   End
   Begin VB.Image Image6 
      Height          =   350
      Left            =   1440
      Picture         =   "frmoffice.frx":2D29B8
      Stretch         =   -1  'True
      Top             =   600
      Width           =   350
   End
   Begin VB.Image Image5 
      Height          =   350
      Left            =   960
      Picture         =   "frmoffice.frx":2E587A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   350
   End
   Begin VB.Image Image4 
      Height          =   350
      Left            =   1920
      Picture         =   "frmoffice.frx":2F99A4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   350
   End
   Begin VB.Image Image3 
      Height          =   350
      Left            =   1440
      Picture         =   "frmoffice.frx":30D1D6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   350
   End
   Begin VB.Image Image2 
      Height          =   350
      Left            =   960
      Picture         =   "frmoffice.frx":31FC38
      Stretch         =   -1  'True
      Top             =   120
      Width           =   350
   End
   Begin VB.Image imgback 
      Height          =   1455
      Left            =   0
      Picture         =   "frmoffice.frx":3331E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmoffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error GoTo errhand:

If ProjectDesktopGagets.frmGagets.txt2007b.Text = "True" Then
imgback.Picture = img7.Picture
End If

If ProjectDesktopGagets.frmGagets.txt2010b.Text = "True" Then
imgback.Picture = img10.Picture
End If

If ProjectDesktopGagets.frmGagets.txt2013b.Text = "True" Then
imgback.Picture = img13.Picture
End If

If ProjectDesktopGagets.frmGagets.txt2016b.Text = "True" Then
imgback.Picture = img16.Picture
End If

errhand:
Exit Sub

End Sub


Private Sub Image2_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\WINWORD.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\WINWORD.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Image21_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Image22_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\ONENOTE.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\ONENOTE.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Image3_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\EXCEL.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub


Private Sub Image4_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\POWERPNT.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub


Private Sub Image5_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\MSACCESS.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\MSACCESS.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Image6_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\MSPUB.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\MSPUB.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Image7_Click()
On Error GoTo errhandler:
If imgback.Picture = img7.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office12\INFOPATH.EXE"), vbMaximizedFocus
End If

If imgback.Picture = img10.Picture Then
Shell ("C:\Program Files\Microsoft Office\Office14\INFOPATH.EXE"), vbMaximizedFocus
End If
errhandler:
Exit Sub
End Sub


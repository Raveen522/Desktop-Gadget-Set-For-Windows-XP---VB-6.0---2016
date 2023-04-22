VERSION 5.00
Begin VB.Form frmslide 
   BorderStyle     =   0  'None
   ClientHeight    =   1350
   ClientLeft      =   13005
   ClientTop       =   8670
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   2400
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2640
      Pattern         =   "*.gif;*.emf;*.dib;*.bmp;*.jpeg;*.wmf;*.jpg"
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1000
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ll"
      BeginProperty Font 
         Name            =   "David"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1000
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   0
      Picture         =   "frmslide.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmslide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub File1_Click()

a = Dir1.path & "\" & File1.FileName
Image1.Picture = LoadPicture(a)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Visible = False
Label5.Visible = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = True Then
Label4.Visible = True
Label5.Visible = False
Else
Label5.Visible = True
Label4.Visible = False
End If
End Sub

Private Sub Label4_Click()
Timer1.Enabled = False
Label5.Visible = True
Label4.Visible = False
End Sub

Private Sub Label5_Click()
Label5.Visible = False
Timer1.Enabled = True
Label4.Visible = True

End Sub

Private Sub Timer1_Timer()
On Error GoTo errhandler:
File1.ListIndex = File1.ListIndex + 1
Image1.Picture = LoadPicture(a)
Label1.Caption = File1.ListCount
Label2.Caption = File1.ListIndex

If Label2.Caption = Label1.Caption - 1 Then
Image1.Picture = LoadPicture(a)
File1.ListIndex = 0
End If
errhandler:
Exit Sub
End Sub

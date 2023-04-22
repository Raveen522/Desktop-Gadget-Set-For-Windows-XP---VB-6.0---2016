VERSION 5.00
Begin VB.Form frmdate 
   BorderStyle     =   0  'None
   ClientHeight    =   1680
   ClientLeft      =   13065
   ClientTop       =   1935
   ClientWidth     =   2280
   BeginProperty Font 
      Name            =   "Narkisim"
      Size            =   20.25
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   340
      Left            =   1800
      TabIndex        =   7
      Top             =   130
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   130
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1970
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      FillColor       =   &H00C0C000&
      Height          =   1680
      Left            =   0
      Top             =   0
      Width           =   2280
   End
   Begin VB.Label lblmonth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   15.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblyear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   1290
      Width           =   2295
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   48
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1995
   End
   Begin VB.Image imgback 
      Height          =   1695
      Left            =   0
      Picture         =   "frmdate.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0&
End Sub

Private Sub Label3_Click()
frmlcal.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H808000
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Visible = False
Label3.Visible = False
Label1.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Visible = True
Label3.Visible = True
Label1.Visible = True
Label3.ForeColor = vbBlack
Label1.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
lbldate.Caption = Format(Date, "dd")
lblmonth.Caption = Format(Date, "mmmm yyyy")
lblyear.Caption = Format(Date, "dddd")
End Sub

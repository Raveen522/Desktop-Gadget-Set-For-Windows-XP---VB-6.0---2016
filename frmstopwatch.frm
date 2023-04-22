VERSION 5.00
Begin VB.Form frmstopwatch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stop Watch"
   ClientHeight    =   1335
   ClientLeft      =   7605
   ClientTop       =   555
   ClientWidth     =   4605
   Icon            =   "frmstopwatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4605
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   840
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Puase"
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   615
      Left            =   3480
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   615
      Left            =   2400
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   615
      Left            =   1320
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Napa Heavy SF"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Napa Heavy SF"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Napa Heavy SF"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Napa Heavy SF"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmstopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim milse As Integer

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = &HC0C000
Label11.BackColor = &HC0C000
End Sub


Private Sub Label10_Click()
If Label5.Caption = "0" And Label2.Caption = "0" And Label3.Caption = "0" And Label4.Caption = "0" Then
Timer1.Enabled = True
Timer2.Enabled = True
Else
Dim strmsg As String
strmsg = MsgBox("Do you want to start a new session?", vbYesNo + vbExclamation)
If strmsg = vbYes Then
Label5.Caption = "0"
Label2.Caption = "0"
Label3.Caption = "0"
milse = 0
Timer1.Enabled = True
Timer2.Enabled = True
Else
Timer1.Enabled = True
Timer2.Enabled = True
End If
End If

End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = &H808000
End Sub

Private Sub Label11_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.BackColor = &H808000
End Sub





Private Sub Timer1_Timer()
Label3.Caption = Val(Label3.Caption) + Val(1)
If Label3.Caption = 60 Then
Label3.Caption = 0
Label2.Caption = Val(Label2.Caption) + Val(1)
Else
If Label2.Caption = 60 Then
Label2.Caption = 0
Label5.Caption = Val(Label5.Caption) + Val(1)
Else
End If
End If

End Sub

Private Sub Timer2_Timer()
milse = milse + 1
If milse = 100 Then
milse = 0
Else
Label4.Caption = milse
End If

End Sub

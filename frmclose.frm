VERSION 5.00
Begin VB.Form frmclose 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   255
   ClientLeft      =   14640
   ClientTop       =   0
   ClientWidth     =   705
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "frmclose.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   400
      TabIndex        =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "frmclose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
frmsettings.Show
End Sub

Private Sub Label1_Click()
Unload frmGagets
Unload frmtime
Unload frmdate

Unload frmtools
Unload frmplayer
Unload frmvirusgards
Unload frmoffice
Unload frmslide
Unload Me
End Sub

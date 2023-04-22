VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmlcal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   9285
   ClientTop       =   3435
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   70
      TabIndex        =   5
      Top             =   480
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2630
      TabIndex        =   4
      Top             =   360
      Width           =   110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Blippo Light SF"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   70
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   70
      TabIndex        =   2
      Top             =   690
      Width           =   2655
   End
   Begin MSComCtl2.MonthView MonthView1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   2310
      Left            =   70
      TabIndex        =   0
      Top             =   360
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      StartOfWeek     =   20709377
      TitleBackColor  =   -2147483647
      TrailingForeColor=   -2147483645
      CurrentDate     =   42579
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   50
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   6
      Height          =   2720
      Left            =   30
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Calender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -240
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmlcal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Text1.Text = MonthView1.Month & "/" & MonthView1.Year
MonthView1.Value = Date

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = vbRed
End Sub


Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0&
End Sub


Private Sub MonthView1_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
Text1.Text = MonthView1.Month & "/" & MonthView1.Year
End Sub

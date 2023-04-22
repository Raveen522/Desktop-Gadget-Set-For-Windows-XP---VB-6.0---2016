VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmvirusgards 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   12795
   ClientTop       =   5670
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmvirusgards.frx":0000
      OLEDBString     =   $"frmvirusgards.frx":0091
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   1410
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image imgother 
      Height          =   840
      Left            =   3120
      Picture         =   "frmvirusgards.frx":0122
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   165
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgavira 
      Height          =   615
      Left            =   4560
      Picture         =   "frmvirusgards.frx":FF38
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image imgavg 
      Height          =   615
      Left            =   3120
      Picture         =   "frmvirusgards.frx":FA57A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image imgavast 
      Height          =   615
      Left            =   4560
      Picture         =   "frmvirusgards.frx":232DBC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgkaspaskey 
      Height          =   735
      Left            =   4560
      Picture         =   "frmvirusgards.frx":33F86E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgmcafee 
      Height          =   615
      Left            =   3120
      Picture         =   "frmvirusgards.frx":357E60
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgeset 
      Height          =   735
      Left            =   3120
      Picture         =   "frmvirusgards.frx":4424A4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   0
      Picture         =   "frmvirusgards.frx":44F695
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmvirusgards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
If ProjectDesktopGagets.frmGagets.txtavastb.Text = "True" Then
Image1.Picture = imgavast.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "Avast"
End If

If ProjectDesktopGagets.frmGagets.txtavgb.Text = "True" Then
Image1.Picture = imgavg.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "AVG"
End If

If ProjectDesktopGagets.frmGagets.txtavirab.Text = "True" Then
Image1.Picture = imgavira.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "Avira"
End If

If ProjectDesktopGagets.frmGagets.txtkasb.Text = "True" Then
Image1.Picture = imgkaspaskey.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "Kaspaskey"
End If

If ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "True" Then
Image1.Picture = imgmcafee.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "McAfee"
End If

If ProjectDesktopGagets.frmGagets.txteset.Text = "True" Then
Image1.Picture = imgeset.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "ESET"
End If

If ProjectDesktopGagets.frmGagets.txtotherb.Text = "True" Then
Image1.Picture = imgother.Picture

Set txt1.DataSource = Adodc1
txt1.DataField = "Other"
End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False
End Sub



Private Sub Label2_Click()
On Error GoTo errhandler:
Dim path As String
path = txt1.Text

If txt1.Text = "" Then
MsgBox "Unavialable path please enter the path", vbCritical, "Error"
Else
Shell (path), vbNormalFocus
End If
errhandler:
Exit Sub
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = True
End Sub

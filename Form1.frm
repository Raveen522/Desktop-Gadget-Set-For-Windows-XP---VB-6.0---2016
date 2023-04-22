VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGagets 
   Caption         =   "Gagets"
   ClientHeight    =   3405
   ClientLeft      =   195
   ClientTop       =   7860
   ClientWidth     =   7710
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7710
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Virus Gard"
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   7695
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   1320
         Top             =   1560
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         Connect         =   $"Form1.frx":179A1
         OLEDBString     =   $"Form1.frx":17A32
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Table1"
         Caption         =   "Adodc3"
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
      Begin VB.TextBox txtotherp 
         Height          =   285
         Left            =   6720
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtmcafeep 
         Height          =   285
         Left            =   5640
         TabIndex        =   29
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtkasp 
         Height          =   285
         Left            =   4320
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtesetp 
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtavgp 
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtavirap 
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtavstp 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtotherb 
         Height          =   285
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtmcafeeb 
         Height          =   285
         Left            =   5640
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtkasb 
         Height          =   285
         Left            =   4320
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtesetb 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtavgb 
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtavirab 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtavastb 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtother 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Text            =   "Text15"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtmcafee 
         Height          =   285
         Left            =   5640
         TabIndex        =   15
         Text            =   "Text14"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtkas 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Text            =   "Text13"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txteset 
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Text            =   "ESET"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtavg 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "AVG"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtavira 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   "Avira"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtavst 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Avast"
         Top             =   240
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   0
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Connect         =   $"Form1.frx":17AC3
         OLEDBString     =   $"Form1.frx":17B56
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Table1"
         Caption         =   "Adodc2"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Office"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txt2016b 
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt2013b 
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt2010b 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt2007b 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt2016 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Text            =   "2016"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt2013 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   "2013"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt2010 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "2010"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt2007 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "2007"
         Top             =   240
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   960
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         Connect         =   $"Form1.frx":17BE9
         OLEDBString     =   $"Form1.frx":17C72
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
   End
End
Attribute VB_Name = "frmGagets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo errhandler:

'**********SHOW********
frmtime.Show
frmdate.Show
frmclose.Show

frmtools.Show

'********END SHOW**********

'**********DATA BASE********
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "Table1"

Adodc2.CommandType = adCmdTable
Adodc2.RecordSource = "Table1"

Adodc3.CommandType = adCmdTable
Adodc3.RecordSource = "Table1"

'*********virus gard********
Set txtavastb.DataSource = Adodc2
txtavastb.DataField = "Avast"

Set txtavirab.DataSource = Adodc2
txtavirab.DataField = "Avira"

Set txtavgb.DataSource = Adodc2
txtavgb.DataField = "AVG"

Set txtesetb.DataSource = Adodc2
txtesetb.DataField = "ESET"

Set txtkasb.DataSource = Adodc2
txtkasb.DataField = "Kaspaskey"

Set txtmcafeeb.DataSource = Adodc2
txtmcafeeb.DataField = "McAfee"

Set txtotherb.DataSource = Adodc2
txtotherb.DataField = "Other"
'*********************************
Set txt2007b.DataSource = Adodc1
txt2007b.DataField = "2007"

Set txt2010b.DataSource = Adodc1
txt2010b.DataField = "2010"

Set txt2013b.DataSource = Adodc1
txt2013b.DataField = "2013"

Set txt2016b.DataSource = Adodc1
txt2016b.DataField = "2016"

'**********END DATA BASE********

errhandler:
Exit Sub
End Sub


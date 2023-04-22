VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsettings 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   6855
   ClientLeft      =   1665
   ClientTop       =   1245
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   5640
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      BackColor       =   -2147483639
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Gagets"
      TabPicture(0)   =   "frmsettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgclock"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgcalender"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgplayer"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgslide_show"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgtool_box"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Clock"
      TabPicture(1)   =   "frmsettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Calender"
      TabPicture(2)   =   "frmsettings.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Player"
      TabPicture(3)   =   "frmsettings.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label15"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Slide Show"
      TabPicture(4)   =   "frmsettings.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Adodc2"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).Control(2)=   "Label16"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Tool Box"
      TabPicture(5)   =   "frmsettings.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label17"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).Control(2)=   "Frame11"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Office"
      TabPicture(6)   =   "frmsettings.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "CommonDialog2"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame8"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label18"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Virus Gard"
      TabPicture(7)   =   "frmsettings.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label19"
      Tab(7).Control(1)=   "Frame6"
      Tab(7).Control(2)=   "Frame9"
      Tab(7).Control(3)=   "Adodc1"
      Tab(7).Control(4)=   "txtab"
      Tab(7).Control(5)=   "txtarb"
      Tab(7).Control(6)=   "txtavb"
      Tab(7).Control(7)=   "txtesb"
      Tab(7).Control(8)=   "txtkab"
      Tab(7).Control(9)=   "txtmcab"
      Tab(7).Control(10)=   "txtotb"
      Tab(7).Control(11)=   "Command24"
      Tab(7).Control(12)=   "Adodc3"
      Tab(7).ControlCount=   13
      TabCaption(8)   =   "About"
      TabPicture(8)   =   "frmsettings.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label29"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Image6"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Label30"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "Label42"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "Label43"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "Text4"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).ControlCount=   6
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   -72600
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Appearence"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   130
         Top             =   240
         Width           =   8175
         Begin VB.CommandButton Command26 
            Caption         =   "Browse"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1080
            TabIndex        =   131
            Top             =   840
            Width           =   4575
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C000&
            BorderWidth     =   2
            Height          =   735
            Left            =   5760
            Top             =   600
            Width           =   2295
         End
         Begin VB.Image Image12 
            Height          =   495
            Left            =   7320
            Picture         =   "frmsettings.frx":00FC
            Stretch         =   -1  'True
            Top             =   720
            Width           =   495
         End
         Begin VB.Image Image11 
            Height          =   495
            Left            =   6000
            Picture         =   "frmsettings.frx":0E3E
            Stretch         =   -1  'True
            Top             =   720
            Width           =   495
         End
         Begin VB.Image Image10 
            Height          =   735
            Left            =   5760
            Picture         =   "frmsettings.frx":1B80
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   -72000
         TabIndex        =   129
         Top             =   2160
         Width           =   5295
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -74400
         Top             =   4680
         Visible         =   0   'False
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
         Connect         =   $"frmsettings.frx":32E4
         OLEDBString     =   $"frmsettings.frx":336D
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   -73200
         Top             =   4560
         Visible         =   0   'False
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
         Connect         =   $"frmsettings.frx":33F6
         OLEDBString     =   $"frmsettings.frx":3489
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
      Begin VB.CommandButton Command24 
         Caption         =   "Update"
         Height          =   255
         Left            =   -74640
         TabIndex        =   118
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtotb 
         Height          =   285
         Left            =   -68880
         TabIndex        =   117
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtmcab 
         Height          =   285
         Left            =   -70200
         TabIndex        =   116
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtkab 
         Height          =   285
         Left            =   -71400
         TabIndex        =   115
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtesb 
         Height          =   285
         Left            =   -72240
         TabIndex        =   114
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtavb 
         Height          =   285
         Left            =   -72960
         TabIndex        =   113
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtarb 
         Height          =   285
         Left            =   -73800
         TabIndex        =   112
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtab 
         Height          =   285
         Left            =   -74640
         TabIndex        =   111
         Top             =   1320
         Width           =   615
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   -74400
         Top             =   4560
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   $"frmsettings.frx":351C
         OLEDBString     =   $"frmsettings.frx":35AD
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Slide Show"
         Height          =   255
         Left            =   2040
         TabIndex        =   108
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Office"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Animation"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   100
         Top             =   2520
         Width           =   8175
         Begin VB.CommandButton Command28 
            Caption         =   "Change Pictures"
            Height          =   255
            Left            =   1200
            TabIndex        =   106
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Start"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   101
            Text            =   "1"
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Time:-"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   855
         End
         Begin VB.Image Image7 
            Height          =   1815
            Left            =   4920
            Stretch         =   -1  'True
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pictures"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   96
         Top             =   120
         Width           =   6735
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   120
            TabIndex        =   99
            Top             =   600
            Width           =   2775
         End
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            Height          =   1785
            Left            =   3000
            Pattern         =   "*.jpeg;*.gif;*.emf;*.dib;*.bmp;*.wmf;*.jpg"
            TabIndex        =   98
            Top             =   240
            Width           =   3255
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label35 
            Height          =   255
            Left            =   6840
            TabIndex        =   104
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label34 
            Height          =   255
            Left            =   6840
            TabIndex        =   103
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H8000000E&
         Caption         =   "Path"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   92
         Top             =   2280
         Width           =   8175
         Begin VB.CommandButton Command23 
            Caption         =   "Update"
            Height          =   255
            Left            =   720
            TabIndex        =   95
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtvpath 
            Height          =   285
            Left            =   720
            TabIndex        =   93
            Top             =   840
            Width           =   7215
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ex :- C:\ABCD\abcd\ABC\abcd.exe "
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   480
            Width           =   7935
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enter your new path here as example."
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Path:-"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000E&
         Caption         =   "Office Version"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   87
         Top             =   360
         Width           =   8055
         Begin VB.CommandButton Command25 
            Caption         =   "Update"
            Height          =   255
            Left            =   360
            TabIndex        =   127
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txt16 
            Height          =   285
            Left            =   4560
            TabIndex        =   124
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txt13 
            Height          =   285
            Left            =   3120
            TabIndex        =   123
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txt10 
            Height          =   285
            Left            =   1680
            TabIndex        =   122
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txt7 
            Height          =   285
            Left            =   240
            TabIndex        =   121
            Top             =   1200
            Width           =   615
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H8000000E&
            Caption         =   "2016"
            Height          =   255
            Left            =   4560
            TabIndex        =   91
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H8000000E&
            Caption         =   "2013"
            Height          =   255
            Left            =   3120
            TabIndex        =   90
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H8000000E&
            Caption         =   "2010"
            Height          =   255
            Left            =   1680
            TabIndex        =   89
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H8000000E&
            Caption         =   "2007"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2.Change ""True""/""False"" following text box your office version"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label Label40 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1.Select your version"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Virus Gard"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   56
         Top             =   120
         Width           =   7815
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Other"
            Height          =   255
            Left            =   6000
            TabIndex        =   63
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Avira"
            Height          =   255
            Left            =   1080
            TabIndex        =   62
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "McAfee"
            Height          =   255
            Left            =   4680
            TabIndex        =   61
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kaspaskey"
            Height          =   255
            Left            =   3480
            TabIndex        =   60
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ESET"
            Height          =   255
            Left            =   2640
            TabIndex        =   59
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "AVG"
            Height          =   255
            Left            =   1920
            TabIndex        =   58
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Avast"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label39 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2.Change ""True""/""False"" following text box your virus guard name"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label Label38 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1.Select your virus guard"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   240
            Width           =   6615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skin"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   53
         Top             =   240
         Width           =   7215
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Glass"
            Height          =   375
            Left            =   3600
            TabIndex        =   55
            Top             =   1800
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Classic"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Image Image3 
            Height          =   1335
            Left            =   3600
            Picture         =   "frmsettings.frx":363E
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1935
         End
         Begin VB.Image Image1 
            Height          =   1335
            Left            =   120
            Picture         =   "frmsettings.frx":13480
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Appearance"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   21
         Top             =   2400
         Width           =   7215
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   495
            Left            =   4680
            TabIndex        =   40
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "May 2016"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   615
            Left            =   4680
            TabIndex        =   25
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1215
            Left            =   4680
            TabIndex        =   24
            Top             =   840
            Width           =   2175
         End
         Begin VB.Image Image2 
            Height          =   1815
            Left            =   4560
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fonts"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   16
         Top             =   780
         Width           =   6375
         Begin VB.CommandButton Command22 
            Caption         =   "More Colours >>"
            Height          =   255
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H00008000&
            Height          =   255
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00FFFF00&
            Height          =   255
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   360
            Width           =   255
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "frmsettings.frx":232C2
            Left            =   2520
            List            =   "frmsettings.frx":232CF
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Font Colour:"
            Height          =   255
            Left            =   3840
            TabIndex        =   51
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   255
            Left            =   2520
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Font"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Appearance"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   11
         Top             =   2400
         Width           =   7215
         Begin VB.CheckBox chkseconds 
            BackColor       =   &H80000014&
            Caption         =   "Show Seconds"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtpath 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   3135
         End
         Begin VB.CommandButton cmdbrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "Background Picture:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblampm 
            BackStyle       =   0  'Transparent
            Caption         =   "PM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   6480
            TabIndex        =   38
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lbl1 
            BackStyle       =   0  'Transparent
            Caption         =   "3:00:00"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1095
            Left            =   4560
            TabIndex        =   14
            Top             =   600
            Width           =   2415
         End
         Begin VB.Image imgbackground 
            Height          =   1575
            Left            =   4560
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fonts"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   6375
         Begin VB.CommandButton Command12 
            BackColor       =   &H00008000&
            Height          =   255
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command10 
            Caption         =   "More Colours >>"
            Height          =   255
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFF00&
            Height          =   255
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3840
            MaskColor       =   &H0000FFFF&
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.ComboBox cbosizeclock 
            Height          =   315
            ItemData        =   "frmsettings.frx":232DF
            Left            =   2520
            List            =   "frmsettings.frx":232F2
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cbofontclock 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackColor       =   &H8000000E&
            Caption         =   "Font Colour:"
            Height          =   255
            Left            =   3840
            TabIndex        =   26
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   "-:Size:-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "-:Font:-"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tool Box"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Virus Gard"
         Height          =   195
         Left            =   5880
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player"
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calender"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clock"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concept,Designed and programmed by K.Raveen Malitha for Pisces "
         BeginProperty Font 
            Name            =   "Accord SF"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -72000
         TabIndex        =   134
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Testers :-K.Raveen Malitha,"
         BeginProperty Font 
            Name            =   "Accord SF"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -72000
         TabIndex        =   133
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Image Image9 
         Height          =   855
         Left            =   2040
         Picture         =   "frmsettings.frx":2330A
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Image Image8 
         Height          =   855
         Left            =   120
         Picture         =   "frmsettings.frx":33120
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 2016 Pisces. All rights Reserved "
         BeginProperty Font 
            Name            =   "Accord Light SF"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -75000
         TabIndex        =   86
         Top             =   4800
         Width           =   6495
      End
      Begin VB.Image Image6 
         Height          =   4095
         Left            =   -74880
         Picture         =   "frmsettings.frx":3F4E6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Height          =   5045
         Left            =   -75000
         TabIndex        =   85
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   75
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   74
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   73
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   72
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   71
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Height          =   5050
         Left            =   -75000
         TabIndex        =   70
         Top             =   0
         Width           =   8440
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5050
         Left            =   -75000
         TabIndex        =   69
         Top             =   0
         Width           =   8440
      End
      Begin VB.Image imgtool_box 
         Height          =   735
         Left            =   3960
         Picture         =   "frmsettings.frx":16B528
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgslide_show 
         Height          =   975
         Left            =   5880
         Picture         =   "frmsettings.frx":170D76
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgplayer 
         Height          =   975
         Left            =   3960
         Picture         =   "frmsettings.frx":180B8C
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgcalender 
         Height          =   975
         Left            =   2040
         Picture         =   "frmsettings.frx":18ACBE
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgclock 
         Height          =   975
         Left            =   120
         Picture         =   "frmsettings.frx":1970FC
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Height          =   5050
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   8440
      End
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Beta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   128
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   8640
      X2              =   7560
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   7560
      X2              =   7560
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   7560
      X2              =   6720
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "About"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6730
      TabIndex        =   84
      Top             =   720
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      BorderWidth     =   6
      Height          =   5050
      Left            =   195
      Top             =   1120
      Width           =   8535
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   6
      X1              =   9120
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   6
      X1              =   0
      X2              =   0
      Y1              =   6960
      Y2              =   0
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   4
      X1              =   8880
      X2              =   8880
      Y1              =   6960
      Y2              =   0
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   4
      X1              =   9000
      X2              =   0
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   240
      X2              =   960
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   960
      X2              =   1680
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   1680
      X2              =   2520
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3360
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   3360
      X2              =   4200
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   4200
      X2              =   5160
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5160
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   6720
      X2              =   5880
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   3360
      X2              =   3360
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   1680
      X2              =   1680
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   960
      X2              =   960
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Tool Box"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5880
      TabIndex        =   83
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Office"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5160
      TabIndex        =   82
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Slide Show"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4200
      TabIndex        =   81
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Virus Gard"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      TabIndex        =   80
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Plyer"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   79
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Calender"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1680
      TabIndex        =   78
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Clock"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   77
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Gadgets"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   76
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   240
      X2              =   960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblsettclose 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8400
      TabIndex        =   68
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Share This Software With Your Family And friends"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   67
      Top             =   6240
      Width           =   8895
   End
   Begin VB.Label lblsettitle1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pisces Gadgets "
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   600
      TabIndex        =   66
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label lblsettitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pisces Gadgets "
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   65
      Top             =   80
      Width           =   8895
   End
   Begin VB.Image Image4 
      Height          =   6015
      Left            =   0
      Picture         =   "frmsettings.frx":1A3DC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8895
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   0
      Picture         =   "frmsettings.frx":1BF5A2
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   8895
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String


Private Sub cbofontclock_Click()
f = cbofontclock.Text
ProjectDesktopGagets.frmtime.lbltime.Font = f
ProjectDesktopGagets.frmtime.lblseconds.Font = f
lbl1.Font = f


End Sub

Private Sub cbosizeclock_Click()
X = cbosizeclock.Text
ProjectDesktopGagets.frmtime.lbltime.FontSize = X
ProjectDesktopGagets.frmtime.lblseconds.FontSize = X
lbl1.FontSize = X

End Sub



Private Sub Check1_Click()
If Check1.Value = 1 Then
frmoffice.Show
Else
Unload frmoffice
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
frmslide.Show
Else
Unload frmslide
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
frmtime.Show
Else
Unload frmtime
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
frmdate.Show
Else
Unload frmdate
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
frmplayer.Show
Else
Unload frmplayer
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
frmvirusgards.Show
Else
Unload frmvirusgards
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
frmtools.Show
Else
Unload frmtools
End If
End Sub

Private Sub chkseconds_Click()
If chkseconds.Value = 1 Then
ProjectDesktopGagets.frmtime.lblampm.Left = 2040
ProjectDesktopGagets.frmtime.lblseconds.Visible = True
Else
ProjectDesktopGagets.frmtime.lblampm.Left = 1440
ProjectDesktopGagets.frmtime.lblseconds.Visible = False
End If

End Sub

Private Sub cmd2_Click()
On Error GoTo errhandler:

CommonDialog1.Filter = "Pictuer_Files(*.Jpg/*.Jpg"
CommonDialog1.ShowOpen
Image2.Picture = LoadPicture(CommonDialog1.FileName)
Text2.Text = CommonDialog1.FileName
ProjectDesktopGagets.frmdate.imgback.Picture = Image2.Picture
errhandler:
Exit Sub
End Sub

Private Sub cmdbrowse_Click()
On Error GoTo errhandler:
CommonDialog1.Filter = "Pictuer_Files(*.Jpg/*.Jpg"
CommonDialog1.ShowOpen
imgbackground.Picture = LoadPicture(CommonDialog1.FileName)
txtpath.Text = CommonDialog1.FileName
ProjectDesktopGagets.frmtime.imgback.Picture = imgbackground
errhandler:
Exit Sub
End Sub



Private Sub Combo3_Click()
fd = Combo3.Text
Label6.Font = fd
Label7.Font = fd
Label9.Font = fd
ProjectDesktopGagets.frmdate.lbldate.Font = fd
ProjectDesktopGagets.frmdate.lblmonth.Font = fd
ProjectDesktopGagets.frmdate.lblyear.Font = fd

End Sub



Private Sub Combo4_Click()
fsd = Combo4.Text
Label6.FontSize = fsd
Label7.FontSize = fsd
Label9.FontSize = fsd
ProjectDesktopGagets.frmdate.lblmonth.FontSize = fsd
ProjectDesktopGagets.frmdate.lblyear.FontSize = fsd

End Sub

Private Sub Command1_Click()
lbl1.ForeColor = vbWhite
lblampm.ForeColor = vbWhite
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = vbWhite
ProjectDesktopGagets.frmtime.lbltime.ForeColor = vbWhite
ProjectDesktopGagets.frmtime.lblampm.ForeColor = vbWhite

End Sub

Private Sub Command10_Click()

CommonDialog1.ShowColor
lbl1.ForeColor = CommonDialog1.Color
lblampm.ForeColor = CommonDialog1.Color
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = CommonDialog1.Color
ProjectDesktopGagets.frmtime.lbltime.ForeColor = CommonDialog1.Color
ProjectDesktopGagets.frmtime.lblampm.ForeColor = CommonDialog1.Color


End Sub


Private Sub Command11_Click()
lbl1.ForeColor = vbBlack
lblampm.ForeColor = vbBlack
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = vbBlack
ProjectDesktopGagets.frmtime.lbltime.ForeColor = vbBlack
ProjectDesktopGagets.frmtime.lblampm.ForeColor = vbBlack

End Sub

Private Sub Command12_Click()
lbl1.ForeColor = &H8000& 'dark green
lblampm.ForeColor = &H8000&
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = &H8000&
ProjectDesktopGagets.frmtime.lbltime.ForeColor = &H8000&
ProjectDesktopGagets.frmtime.lblampm.ForeColor = &H8000&
Label10.Visible = True
End Sub

Private Sub Command13_Click()
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
Label9.ForeColor = vbBlack
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command14_Click()
Label6.ForeColor = vbRed
Label7.ForeColor = vbRed
Label9.ForeColor = vbRed
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command15_Click()
Label6.ForeColor = &H80FF& 'orange
Label7.ForeColor = &H80FF& 'orange
Label9.ForeColor = &H80FF& 'orange
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command16_Click()
Label6.ForeColor = vbYellow
Label7.ForeColor = vbYellow
Label9.ForeColor = vbYellow
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command17_Click()
Label6.ForeColor = vbGreen
Label7.ForeColor = vbGreen
Label9.ForeColor = vbGreen
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command18_Click()
Label6.ForeColor = &HFFFF00 'Ceyan
Label7.ForeColor = &HFFFF00  'Ceyan
Label9.ForeColor = &HFFFF00 'Ceyan
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command19_Click()
Label6.ForeColor = vbBlue
Label7.ForeColor = vbBlue
Label9.ForeColor = vbBlue
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command2_Click()
Label6.ForeColor = vbWhite
Label7.ForeColor = vbWhite
Label9.ForeColor = vbWhite
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor

End Sub

Private Sub Command20_Click()
Label6.ForeColor = &HFF00FF    'pink
Label7.ForeColor = &HFF00FF    'pink
Label9.ForeColor = &HFF00FF    'pink
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command21_Click()
Label6.ForeColor = &H8000&  'dark green
Label7.ForeColor = &H8000&  'dark green
Label9.ForeColor = &H8000&  'dark green
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub

Private Sub Command22_Click()
CommonDialog1.ShowColor
Label6.ForeColor = CommonDialog1.Color
Label7.ForeColor = CommonDialog1.Color
Label9.ForeColor = CommonDialog1.Color
ProjectDesktopGagets.frmdate.lbldate.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblmonth.ForeColor = Label6.ForeColor
ProjectDesktopGagets.frmdate.lblyear.ForeColor = Label6.ForeColor
End Sub







Private Sub Command23_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command24_Click()
Adodc3.Recordset.Update
ProjectDesktopGagets.frmGagets.Adodc2.Recordset.Update
End Sub

Private Sub Command25_Click()
Adodc2.Recordset.Update
End Sub

Private Sub Command26_Click()
On Error GoTo errhandler:
CommonDialog2.Filter = "Pictuer_Files(*.Jpg/*.Jpeg/*.bmp/*.dib/*.TIFF/*.JPE/*.JFIF"
CommonDialog2.ShowOpen
Image10.Picture = LoadPicture(CommonDialog2.FileName)
Text6.Text = CommonDialog2.FileName
ProjectDesktopGagets.frmtools.imgback.Picture = Image10.Picture
errhandler:
Exit Sub
End Sub

Private Sub Command27_Click()

Timer1.Enabled = True
Dir1.Enabled = False
Drive1.Enabled = False
File1.Enabled = False
ProjectDesktopGagets.frmslide.Timer1.Enabled = True

End Sub

Private Sub Command28_Click()
Timer1.Enabled = False
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
ProjectDesktopGagets.frmslide.Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
lbl1.ForeColor = vbRed
lblampm.ForeColor = vbRed
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = vbRed
ProjectDesktopGagets.frmtime.lbltime.ForeColor = vbRed
ProjectDesktopGagets.frmtime.lblampm.ForeColor = vbRed

End Sub

Private Sub Command4_Click()
lbl1.ForeColor = &H80FF& 'orange
lblampm.ForeColor = &H80FF&
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = &H80FF&
ProjectDesktopGagets.frmtime.lbltime.ForeColor = &H80FF&
ProjectDesktopGagets.frmtime.lblampm.ForeColor = &H80FF&

End Sub

Private Sub Command5_Click()
lbl1.ForeColor = vbYellow
lblampm.ForeColor = vbYellow
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = vbYellow
ProjectDesktopGagets.frmtime.lbltime.ForeColor = vbYellow
ProjectDesktopGagets.frmtime.lblampm.ForeColor = vbYellow

End Sub

Private Sub Command6_Click()
lbl1.ForeColor = &HFF00& 'light green
lblampm.ForeColor = &HFF00&
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = &HFF00&
ProjectDesktopGagets.frmtime.lbltime.ForeColor = &HFF00&
ProjectDesktopGagets.frmtime.lblampm.ForeColor = &HFF00&

End Sub

Private Sub Command7_Click()
lbl1.ForeColor = &HFFFF00 'Ceyan
lblampm.ForeColor = &HFFFF00
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = &HFFFF00
ProjectDesktopGagets.frmtime.lbltime.ForeColor = &HFFFF00
ProjectDesktopGagets.frmtime.lblampm.ForeColor = &HFFFF00

End Sub

Private Sub Command8_Click()
lbl1.ForeColor = vbBlue
lblampm.ForeColor = vbBlue
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = vbBlue
ProjectDesktopGagets.frmtime.lbltime.ForeColor = vbBlue
ProjectDesktopGagets.frmtime.lblampm.ForeColor = vbBlue

End Sub

Private Sub Command9_Click()
lbl1.ForeColor = &HFF00FF    'pink
lblampm.ForeColor = &HFF00FF
ProjectDesktopGagets.frmtime.lblseconds.ForeColor = &HFF00FF
ProjectDesktopGagets.frmtime.lbltime.ForeColor = &HFF00FF
ProjectDesktopGagets.frmtime.lblampm.ForeColor = &HFF00FF

End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
ProjectDesktopGagets.frmslide.File1.path = Dir1.path
ProjectDesktopGagets.frmslide.Dir1.path = Dir1.path
 
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive

End Sub

Private Sub File1_Click()
a = Dir1.path & "\" & File1.FileName
Image7.Picture = LoadPicture(a)
End Sub

Private Sub Form_Load()
'*********************
'Adodc1.CommandType = adCmdTable
'Adodc1.RecordSource = "Table1"

'Adodc2.CommandType = adCmdTable
'Adodc2.RecordSource = "Table1"

'*****************
'data base
Set txtab.DataSource = Adodc3
txtab.DataField = "Avast"

Set txtarb.DataSource = Adodc3
txtarb.DataField = "Avira"

Set txtavb.DataSource = Adodc3
txtavb.DataField = "AVG"

Set txtesb.DataSource = Adodc3
txtesb.DataField = "ESET"

Set txtkab.DataSource = Adodc3
txtkab.DataField = "Kaspaskey"

Set txtmcab.DataSource = Adodc3
txtmcab.DataField = "McAfee"

Set txtotb.DataSource = Adodc3
txtotb.DataField = "Other"

Set txt7.DataSource = Adodc2
txt7.DataField = "2007"

Set txt10.DataSource = Adodc2
txt10.DataField = "2010"

Set txt13.DataSource = Adodc2
txt13.DataField = "2013"

Set txt16.DataSource = Adodc2
txt16.DataField = "2016"

'end of data base



'virus gard
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavast.Picture Then
 Option3.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavg.Picture Then
 Option4.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgeset.Picture Then
 Option5.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgkaspaskey.Picture Then
 Option6.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgmcafee.Picture Then
 Option7.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavira.Picture Then
 Option8.Value = True
 End If
 If ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgother.Picture Then
 Option9.Value = True
 End If
'end virus gard

'************Slide Show **************
'Dir1.Path = ProjectDesktopGagets.frmslide.Drive1.Drive



'************End of Slide Show **************

'************** Office*************
If ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img7.Picture Then
Option10.Value = True
End If
If ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img10.Picture Then
Option11.Value = True
End If
If ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img13.Picture Then
Option12.Value = True
End If
If ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img16.Picture Then
Option13.Value = True
End If


'*************End of Office*********

' ******** FRMCLOCK ***********
imgbackground.Picture = ProjectDesktopGagets.frmtime.imgback.Picture
cbofontclock.Text = ProjectDesktopGagets.frmtime.lbltime.Font
cbosizeclock.Text = ProjectDesktopGagets.frmtime.lbltime.FontSize
lbl1.ForeColor = ProjectDesktopGagets.frmtime.lbltime.ForeColor
lblampm.ForeColor = lbl1.ForeColor
'******** End of FRMCLOCK **********

'********** Frm Date *********
Combo3.Text = Label6.Font
Combo4.Text = Label6.FontSize
Image2.Picture = ProjectDesktopGagets.frmdate.imgback.Picture
Combo3.Text = ProjectDesktopGagets.frmdate.lblyear.FontName
Combo4.Text = ProjectDesktopGagets.frmdate.lblyear.FontSize
'********* End of Date *********

'***** Font & Font Size *******
Dim X As Integer
For X = 1 To Screen.FontCount
cbofontclock.AddItem Screen.Fonts(X)
Combo3.AddItem Screen.Fonts(X)
Next X
'******** End of Font & Font Size ********

' ******* Gagets Tab *********
If frmtime.Visible = True Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If frmdate.Visible = True Then
Check4.Value = 1
Else
Check4.Value = 0
End If
If frmplayer.Visible = True Then
Check5.Value = 1
Else
Check5.Value = 0
End If
If frmtools.Visible = True Then
Check7.Value = 1
Else
Check7.Value = 0
End If
If frmvirusgards.Visible = True Then
Check6.Value = 1
Else
Check6.Value = 0
End If
If frmoffice.Visible = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If frmslide.Visible = True Then
Check2.Value = 1
Else
Check2.Value = 0
End If

'*********** End of Gagets Tab ************

'************ Seconds label ************
If ProjectDesktopGagets.frmtime.lblseconds.Visible = True Then
chkseconds.Value = 1
Else
chkseconds.Value = 0
End If
'************ End of Seconds label ************

'************  Plyer Tab ************
If ProjectDesktopGagets.frmplayer.imgback.Left = 0 Then
Option1.Value = True
End If
If ProjectDesktopGagets.frmplayer.WindowsMediaPlayer1.Left = 0 Then
Option2.Value = True
End If
'******* End of plyer Tab **********
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsettclose.BackStyle = 0
End Sub

Private Sub Label20_Click()
Line13.X1 = 240
Line13.X2 = 960
Line16.Visible = False
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 0

End Sub

Private Sub Label21_Click()
Line13.X1 = 960
Line13.X2 = 1680
Line16.Visible = True
Line15.Visible = False
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 1
End Sub

Private Sub Label22_Click()
Line13.X1 = 1680
Line13.X2 = 2520
Line16.Visible = True
Line15.Visible = True
Line14.Visible = False
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 2
End Sub

Private Sub Label23_Click()
Line13.X1 = 2520
Line13.X2 = 3360
Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = False
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 3
End Sub

Private Sub Label24_Click()
Line13.X1 = 3360
Line13.X2 = 4200
Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = False
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 7
End Sub

Private Sub Label25_Click()
Line13.X1 = 4200
Line13.X2 = 5160
Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = False
Line9.Visible = True
Line8.Visible = True
Line19.Visible = True
SSTab1.Tab = 5
ProjectDesktopGagets.frmslide.Timer1.Enabled = False
Timer1.Enabled = False



Dir1.path = ProjectDesktopGagets.frmslide.File1.path
Dir1.path = ProjectDesktopGagets.frmslide.Dir1.path
End Sub

Private Sub Label26_Click()
Line13.X1 = 5160
Line13.X2 = 5880
Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = False
Line8.Visible = True
Line19.Visible = True

SSTab1.Tab = 4
End Sub

Private Sub Label27_Click()
Line13.X1 = 5880
Line13.X2 = 6720
Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = False
Line19.Visible = True

SSTab1.Tab = 6

Image10.Picture = ProjectDesktopGagets.frmtools.imgback.Picture
End Sub

Private Sub Label28_Click()
Line13.X1 = 6720
Line13.X2 = 7560

Line16.Visible = True
Line15.Visible = True
Line14.Visible = True
Line12.Visible = True
Line11.Visible = True
Line10.Visible = True
Line9.Visible = True
Line8.Visible = True
Line19.Visible = False

SSTab1.Tab = 8
End Sub

Private Sub lblsettclose_Click()
Unload Me
End Sub

Private Sub lblsettclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsettclose.BackStyle = 1
End Sub



Private Sub lblsettitle1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsettclose.BackStyle = 0
End Sub



Private Sub lblsettitle2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsettclose.BackStyle = 0
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
ProjectDesktopGagets.frmplayer.imgback.Left = 0
ProjectDesktopGagets.frmplayer.imgplay.Left = 0
ProjectDesktopGagets.frmplayer.imgpuase.Left = 600
ProjectDesktopGagets.frmplayer.imgstop.Left = 960
ProjectDesktopGagets.frmplayer.imgopen.Left = 1320

ProjectDesktopGagets.frmplayer.WindowsMediaPlayer1.Left = 2640
ProjectDesktopGagets.frmplayer.Image9.Left = 2640
ProjectDesktopGagets.frmplayer.Image22.Left = 2740
ProjectDesktopGagets.frmplayer.Image23.Left = 2840
ProjectDesktopGagets.frmplayer.Image24.Left = 2940
ProjectDesktopGagets.frmplayer.Image25.Left = 3140
End If
End Sub

Private Sub Option10_Click()
ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img7.Picture

ProjectDesktopGagets.frmGagets.txt2007b.Text = "True"
ProjectDesktopGagets.frmGagets.txt2010b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2013b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2016b.Text = "False"

End Sub

Private Sub Option11_Click()
ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img10.Picture

ProjectDesktopGagets.frmGagets.txt2007b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2010b.Text = "True"
ProjectDesktopGagets.frmGagets.txt2013b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2016b.Text = "False"

End Sub

Private Sub Option12_Click()
ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img13.Picture

ProjectDesktopGagets.frmGagets.txt2007b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2010b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2013b.Text = "True"
ProjectDesktopGagets.frmGagets.txt2016b.Text = "False"

End Sub

Private Sub Option13_Click()
ProjectDesktopGagets.frmoffice.imgback.Picture = ProjectDesktopGagets.frmoffice.img16.Picture

ProjectDesktopGagets.frmGagets.txt2007b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2010b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2013b.Text = "False"
ProjectDesktopGagets.frmGagets.txt2016b.Text = "True"

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
ProjectDesktopGagets.frmplayer.WindowsMediaPlayer1.Left = 0
ProjectDesktopGagets.frmplayer.Image9.Left = 0
ProjectDesktopGagets.frmplayer.Image22.Left = 80
ProjectDesktopGagets.frmplayer.Image23.Left = 630
ProjectDesktopGagets.frmplayer.Image24.Left = 1260
ProjectDesktopGagets.frmplayer.Image25.Left = 2030

ProjectDesktopGagets.frmplayer.imgback.Left = 2640

End If


End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavast.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "Avast"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text




ProjectDesktopGagets.frmGagets.txtavastb.Text = "True"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavg.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "AVG"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text




ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "True"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgeset.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "ESET"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text



ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "True"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgkaspaskey.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "Kaspaskey"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text



ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "True"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If
End Sub

Private Sub Option7_Click()
If Option7.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgmcafee.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "McAfee"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text



ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "True"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgavira.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbBlue

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "Avira"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text



ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "True"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "False"


End If
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
ProjectDesktopGagets.frmvirusgards.Image1.Picture = ProjectDesktopGagets.frmvirusgards.imgother.Picture
ProjectDesktopGagets.frmvirusgards.Label1.ForeColor = vbGreen

Set txtvpath.DataSource = Adodc1
txtvpath.DataField = "Other"
ProjectDesktopGagets.frmvirusgards.txt1.Text = txtvpath.Text


ProjectDesktopGagets.frmGagets.txtavastb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavgb.Text = "False"
ProjectDesktopGagets.frmGagets.txtavirab.Text = "False"
ProjectDesktopGagets.frmGagets.txtesetb.Text = "False"
ProjectDesktopGagets.frmGagets.txtkasb.Text = "False"
ProjectDesktopGagets.frmGagets.txtmcafeeb.Text = "False"
ProjectDesktopGagets.frmGagets.txtotherb.Text = "True"


End If
End Sub

Private Sub Text5_Change()
Timer1.Interval = Text5.Text * 1000
ProjectDesktopGagets.frmslide.Timer1.Interval = Text5.Text * 1000
End Sub

Private Sub Timer1_Timer()
On Error GoTo errhandler:

File1.ListIndex = File1.ListIndex + 1
Image1.Picture = LoadPicture(a)
Label34.Caption = File1.ListCount
Label35.Caption = File1.ListIndex

If Label35.Caption = Label34.Caption - 1 Then
Image1.Picture = LoadPicture(a)
File1.ListIndex = 0
End If
errhandler:
Exit Sub
End Sub

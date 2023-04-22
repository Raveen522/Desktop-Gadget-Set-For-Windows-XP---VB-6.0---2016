VERSION 5.00
Begin VB.Form frmSettingsbut 
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   14745
   ClientTop       =   0
   ClientWidth     =   270
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsettings 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmSettingsbut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdsettings_Click()
frmsettings.Show
End Sub

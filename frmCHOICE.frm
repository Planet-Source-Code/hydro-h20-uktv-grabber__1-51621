VERSION 5.00
Begin VB.Form frmCHOICE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select An Option To Run"
   ClientHeight    =   3255
   ClientLeft      =   2985
   ClientTop       =   4185
   ClientWidth     =   5565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCHOICE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmCHOICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Local Error Resume Next
    
    List1.AddItem "-configure MODEM"
    List1.AddItem "-channels"
    List1.AddItem "-grab MODEM"
    List1.AddItem "-grab days=3 MODEM"
    List1.AddItem "-grab xml=c:\xmltv\uklist.xml MODEM"
    List1.AddItem "-grab days=3 xml=c:\xmltv\uklist.xml MODEM"
    List1.AddItem "-grab days=3 xml=c:\xmltv\uklist.xml MODEM DEBUG"
    List1.AddItem "-grab days=1 xml=c:\xmltv\uklist.xml MODEM"
    List1.AddItem "-grab days=1 xml=c:\xmltv\uklist.xml MODEM DEBUG"
    List1.AddItem "-grab days=3 xml=c:\xmltv\uklist.xml BADMODEM"
    List1.AddItem "-grab days=3 xml=c:\doesnotexist\uklist.xml MODEM"
    List1.AddItem "-rubbish"
    List1.AddItem "-grab xds=c:\xmltv\list.xml"
    List1.AddItem "-grab day=3"
    List1.AddItem "-grab days=0"
    List1.AddItem "-grab days=8"
    List1.AddItem "-grab days=-1"
    List1.AddItem "-grab days=-1.4"
    List1.AddItem "-grab days=1.5"
    List1.AddItem "-grab days=8.5"
    
    
    Me.Move 100, 100
    
End Sub

Private Sub List1_DblClick()
    On Local Error Resume Next
    runOption = List1.Text
    Unload Me
End Sub

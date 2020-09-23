VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmWEB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WEB Page Viewer(Debug Mode)"
   ClientHeight    =   5430
   ClientLeft      =   1320
   ClientTop       =   2430
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHTML 
      BackColor       =   &H80000013&
      Height          =   5175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
   Begin InetCtlsObjects.Inet web 
      Left            =   8760
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Move 100, frmGrabber.Top + frmGrabber.Height + 100
End Sub

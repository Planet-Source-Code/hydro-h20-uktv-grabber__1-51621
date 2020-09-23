VERSION 5.00
Begin VB.Form frmCHANNELS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channel SetUp"
   ClientHeight    =   4560
   ClientLeft      =   2010
   ClientTop       =   2250
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4005
   Begin VB.CommandButton cmdAPPLY 
      Caption         =   "Apply && Exit"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCHANNELS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAPPLY_Click()
    Dim i As Integer
    
    Open App.Path & "\channels.dat" For Output As #1
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = False Then
            Print #1, Chr$(34) & List1.ItemData(i) & Chr$(34) & "," & Chr$(34) & List1.List(i) & Chr$(34) & "," & Chr$(34) & "False" & Chr$(34)
        Else
            Print #1, Chr$(34) & List1.ItemData(i) & Chr$(34) & "," & Chr$(34) & List1.List(i) & Chr$(34) & "," & Chr$(34) & "True" & Chr$(34)
        End If
    Next i
    Close #1
    Unload frmGrabber
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim sCODE As String
    Dim sDSC As String
    Dim sVIEW As String
    
    On Local Error Resume Next
    
    If Len(Dir$(App.Path & "\channels.dat")) = 0 Then
        Unload Me
        End
    End If
    
    Me.Move 100, 100
    
    channelCNT = -1
    Open App.Path & "\channels.dat" For Input As #1
    While EOF(1) = False
        Input #1, sCODE, sDSC, sVIEW
        List1.AddItem sDSC
        List1.ItemData(List1.ListCount - 1) = sCODE
        If sVIEW = "True" Then
            List1.Selected(List1.ListCount - 1) = True
        Else
            List1.Selected(List1.ListCount - 1) = False
        End If
    Wend
        
    Close #1
        
End Sub

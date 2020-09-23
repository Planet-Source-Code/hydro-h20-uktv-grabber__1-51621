VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGrabber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UKTV Grabber By CornStopper"
   ClientHeight    =   1785
   ClientLeft      =   1455
   ClientTop       =   2985
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGrabber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5070
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Sub cmdCLOSE_Click()
    On Local Error Resume Next
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim runSPLIT() As String
    Dim i As Integer
    Dim rCOMMAND As String
    Dim rDAYS As Integer
    Dim rXML As String
    Dim illegalPAR As Boolean
    Dim rSTR As String
    Dim dTEST As String
    Dim modemCONNECT As Boolean
    Dim DisconnectLINE As Boolean
    Dim rRESULT As Boolean
    Dim rSTARTDATE As String
    Dim a As Integer
    
    On Local Error Resume Next
    
    rDAYS = 2
    rXML = App.Path & "\listings.xml"
    rSTARTDATE = Format$(Date, "dd/mmm/yyyy")
    
    If IsDesignTime Then
        frmCHOICE.Show 1
    Else
        runOption = Command()
    End If
    
    If Len(runOption) = 0 Then
        Unload Me
        End
    End If
    
    runOption = Trim$(runOption) & " "
    runSPLIT = Split(runOption, " ")
    
    For i = 0 To UBound(runSPLIT) - 1
        runSPLIT(i) = Trim$(runSPLIT(i))
        If runSPLIT(i) = "-configure" Then
            rCOMMAND = "-configure"
        ElseIf runSPLIT(i) = "-channels" Then
            rCOMMAND = "-channels"
        ElseIf runSPLIT(i) = "-grab" Then
            rCOMMAND = "-grab"
        ElseIf runSPLIT(i) = "debug" Then
            debugON = True
        ElseIf Left$(runSPLIT(i), 4) = "days" Then
            rSTR = Right$(runSPLIT(i), Len(runSPLIT(i)) - 5)
            If IsNumeric(rSTR) = False Then
                illegalPAR = True
                Exit For
            End If
            
            If Val(rSTR) <> CInt(rSTR) Then
                illegalPAR = True
                Exit For
            End If
            
            If Val(rSTR) <= 0 Or Val(rSTR) > 7 Then
                illegalPAR = True
                Exit For
            End If
            
            rDAYS = Val(rSTR)
        ElseIf Left$(runSPLIT(i), 3) = "xml" Then
            rSTR = Right$(runSPLIT(i), Len(runSPLIT(i)) - 4)
            a = InStrRev(rSTR, "\")
            dTEST = rSTR
            If a <> Len(rSTR) Then dTEST = Left$(rSTR, a)
            
            dTEST = Dir$(dTEST, vbDirectory)
            If Len(dTEST) = 0 Then
                illegalPAR = True
                Exit For
            End If
            
            rXML = rSTR
        ElseIf Left$(runSPLIT(i), 9) = "startdate" Then
            rSTR = Right$(runSPLIT(i), Len(runSPLIT(i)) - 10)
        
            If IsDate(rSTR) = False Then
                illegalPAR = True
                Exit For
            End If
            
            rSTARTDATE = Format$(CDate(rSTR), "dd/mmm/yyyy")
        
        ElseIf runSPLIT(i) = "MODEM" Then
            modemCONNECT = True
        Else
            illegalPAR = True
            Exit For
        End If
    Next i
    
    If illegalPAR Then
        Unload Me
        End
    End If
    
    Me.Move 100, 100
    Label1 = ""
    Me.Show
    DoEvents
    
    Select Case rCOMMAND
        Case "-grab", "-configure"
            If modemCONNECT Then
                If Not Online Then
                    For i = 1 To 3
                        Label1 = "Connecting To The Net" & vbCrLf & "Attempt " & i
                        DoEvents
                        ret1 = AutoDial(Me.hwnd, ADF_FORCE_UNATTENDED, True)
                        If ret1 Then Exit For
                    Next i
                    
                    If ret1 = False Then
                        Unload Me
                        End
                    End If
                            
                    DisconnectLINE = True
                End If
            End If
            
            DoEvents
            DoEvents
            Load frmWEB
            DoEvents
            If debugON Then frmWEB.Show
            DoEvents
            DoEvents
            DoEvents
            
            Select Case rCOMMAND
                Case "-grab"
                    rRESULT = Locate_Programmes(rDAYS, rSTARTDATE)
                    If rRESULT = True Then
                        rRESULT = Create_XML(rXML)
                    End If
                Case "-configure"
                    rRESULT = Grab_Channels
            End Select
            
            Unload frmWEB
            
            If modemCONNECT Then
                If DisconnectLINE Then
                    Label1 = "Hanging Up"
                    DoEvents
                    Call Hangup
                End If
            End If
        
            If rRESULT Then
                Select Case rCOMMAND
                    Case "-configure"
                        Load frmCHANNELS
                        Unload Me
                        frmCHANNELS.Show
                        Exit Sub
                    Case "-grab"
                        Unload Me
                        End
                End Select
            End If
        Case "-channels"
            If Len(Dir$(App.Path & "\channels.dat")) = 0 Then
                Unload Me
                End
            End If
            
            Load frmCHANNELS
            Unload Me
            frmCHANNELS.Show
            Exit Sub
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    If debugON Then Exit Sub
    If Len(Dir$(App.Path & "\temp0001.htm")) <> 0 Then Kill App.Path & "\temp0001.htm"
    If Len(Dir$(App.Path & "\temp0002.htm")) <> 0 Then Kill App.Path & "\temp0002.htm"
    If Len(Dir$(App.Path & "\temp0003.htm")) <> 0 Then Kill App.Path & "\temp0003.htm"
End Sub

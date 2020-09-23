Attribute VB_Name = "modCONFIG"
Option Compare Text
Option Explicit

Public Function Grab_Channels() As Boolean
    Dim sLINE As String
    Dim i As Integer
    Dim sCODE As String
    Dim sCHANNEL As String
    Dim sSTART As Date
    Dim urlDATA As String
    
    On Local Error GoTo fcntl
    
    frmGrabber.Label1 = "Configuring Channels"
    DoEvents
    
    sSTART = Now()
    
    If debugON Then frmWEB.txtHTML = "LOADING..." & vbCrLf & "http://www.radiotimes.com/jsp/listings_search.jsp"
    urlDATA = frmWEB.web.OpenURL("http://www.radiotimes.com/jsp/listings_search.jsp")
    DoEvents
    DoEvents
    
sConnect:

    ' have time out here
    While frmWEB.web.StillExecuting = True
        DoEvents
    Wend

    i = DateDiff("s", sSTART, Now())

    If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 And i <= 90 Then GoTo sConnect
    If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 Then GoTo fcntl

    frmGrabber.Label1 = "Site Located..."
    DoEvents

    If debugON Then frmWEB.txtHTML = urlDATA
    sLINE = urlDATA
    
    sLINE = Replace(sLINE, "&amp;", "&", , , vbTextCompare)
    sLINE = Replace(sLINE, "<select", vbCrLf & "<select", , , vbTextCompare)
    sLINE = Replace(sLINE, "<option", vbCrLf & "<option", , , vbTextCompare)
    sLINE = Replace(sLINE, "</select", vbCrLf & "</select", , , vbTextCompare)
    
    Open App.Path & "\temp0001.htm" For Output As #1
    Print #1, sLINE
    Close #1

    ' got channels into temp0001.htm
    ' just readin and parse into channels.dat
'here:
    On Local Error Resume Next
    
    frmGrabber.Label1 = "Parsing Channels..."
    DoEvents
    
    Open App.Path & "\temp0001.htm" For Input As #1
    Open App.Path & "\channels.dat" For Output As #2
    
    Do
        Line Input #1, sLINE
        sLINE = Trim$(sLINE)
        i = 0
        If Left$(sLINE, Len("<select")) = "<select" Then
            i = InStr(1, sLINE, "name=" & Chr$(34) & "channels" & Chr$(34), vbTextCompare)
            If i = 0 Then i = InStr(1, sLINE, "name=channels", vbTextCompare)
        End If
    Loop Until EOF(1) Or i <> 0
    
    If i <> 0 Then
        ' found start of channel list
        Do
            Line Input #1, sLINE
            sLINE = Trim$(sLINE)
'  <option value="92,105,26,132,134,39,40,45,47,48,49,119,122,123,482,483,147,150,156,158,160,1041,1241,921,177,180,182,184,185,197,213,1201,1061,244,1261,249,250,251,252,253,254,255,256,257,258,259,260,248,922,262,264,265,263,300,266,271,274,923,292,423,288,289,801,290,291" >Standard TV Channels
'  <option value=92 >BBC1
'  <option value="95" >BBC3

            If Left$(sLINE, Len("<option value=")) = "<option value=" Then
                sLINE = Right$(sLINE, Len(sLINE) - (Len("<option value=")))
                i = InStr(1, sLINE, ">", vbBinaryCompare)
                If i <> 0 Then
                    sCODE = Left$(sLINE, i - 1)
                    sCODE = Trim$(sCODE)
                    sCODE = Replace(sCODE, Chr$(34), "", , , vbTextCompare)
                    If IsNumeric(sCODE) And InStr(1, sCODE, ",", vbTextCompare) = 0 Then
                        i = InStr(1, sLINE, ">", vbBinaryCompare)
                        sCHANNEL = Right$(sLINE, Len(sLINE) - i)
                        sCHANNEL = Trim$(sCHANNEL)
                        sCHANNEL = Replace(sCHANNEL, "&amp;", "&", , , vbTextCompare)
                        
                        Print #2, Chr$(34) & sCODE & Chr$(34) & "," & Chr$(34) & sCHANNEL & Chr$(34) & "," & Chr$(34) & "False" & Chr$(34)
                    
                    End If
                End If
    
    
            End If
        Loop Until EOF(1) Or Left$(sLINE, Len("</select>")) = "</select>"
    End If
    
    Close #1
    Close #2
    
    Grab_Channels = True
    Exit Function

fcntl:
    frmGrabber.Label1 = "Failed To Locate Channels"
    frmGrabber.cmdCLOSE.Visible = True
    DoEvents
    Grab_Channels = False
End Function

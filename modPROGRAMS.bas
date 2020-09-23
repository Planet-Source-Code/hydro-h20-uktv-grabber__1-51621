Attribute VB_Name = "modPROGRAMS"
Option Compare Text
Option Explicit

Public Function Locate_Programmes(LookUpDays As Integer, stDATE As String) As Boolean
    Dim sCHANNEL As String
    Dim sCODE As String
    Dim sLOAD As String
    Dim sSTART As Date
    Dim i As Integer
    Dim lDAY As Integer
    Dim LookUpDate As String
    Dim lookUpHREF As String
    Dim sLINE As String
    Dim a As Integer
    Dim sProgrammeId As String
    Dim sProgFound As Boolean
    Dim tGrab As Integer
    Dim urlDATA As String
    Dim maxHTM As Integer
    Dim cHTM As Integer
    Dim pVAL As Integer
    
    On Local Error Resume Next
    
    If Len(Dir$(App.Path & "\channels.dat")) = 0 Then
        frmGrabber.Label1 = "Configure Channels"
        frmGrabber.cmdCLOSE.Visible = True
        Locate_Programmes = False
        Exit Function
    End If
    
    If debugON Then
        MkDir App.Path & "\pages"
    End If
    
    Open App.Path & "\channels.dat" For Input As #1
    While EOF(1) = False
        Input #1, sCODE, sCHANNEL, sLOAD
        If CBool(sLOAD) = True Then maxHTM = maxHTM + 1
    Wend
    Close #1
    maxHTM = maxHTM * LookUpDays * 6
    
    frmGrabber.Height = 2670
    DoEvents
    
    Open App.Path & "\channels.dat" For Input As #1
    While EOF(1) = False
        Input #1, sCODE, sCHANNEL, sLOAD
        
        If CBool(sLOAD) = False Then GoTo sNextChannel
        
        For lDAY = 1 To LookUpDays
            LookUpDate = Format$(CDate(stDATE) + (lDAY - 1), "dd/mmm/yyyy")
            
            frmGrabber.Label1 = sCHANNEL & vbCrLf & LookUpDate & vbCrLf & "Locating Site"
            
            For tGrab = 1 To 6
                Select Case tGrab
                    Case 1
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=00%3A00&channels=" & sCODE
                    Case 2
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=04%3A00&channels=" & sCODE
                    Case 3
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=08%3A00&channels=" & sCODE
                    Case 4
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=12%3A00&channels=" & sCODE
                    Case 5
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=16%3A00&channels=" & sCODE
                    Case 6
                        lookUpHREF = "http://www.radiotimes.com/ListingsServlet?event=4&jspGridLocation=%2Fjsp%2Ftv_listings_grid.jsp&jspListLocation=%2Fjsp%2Ftv_listings_single.jsp&jspError=%2Fjsp%2Ferror.jsp&searchDate=" & Format$(CDate(LookUpDate), "dd") & "%2F" & Format$(CDate(LookUpDate), "mm") & "%2F" & Format$(CDate(LookUpDate), "yyyy") & "&searchTime=20%3A00&channels=" & sCODE
                End Select
                        
                
                If debugON Then frmWEB.txtHTML = "LOADING..." & vbCrLf & lookUpHREF
                urlDATA = frmWEB.web.OpenURL(lookUpHREF)
                DoEvents
                DoEvents
                sSTART = Now()
            
sConnect:
                ' have time out here
                While frmWEB.web.StillExecuting = True
                    DoEvents
                Wend
    
                i = DateDiff("s", sSTART, Now())

                If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 And i <= 90 Then GoTo sConnect
                If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 Then GoTo sNextChannel

                ' site found
                frmGrabber.Label1 = sCHANNEL & vbCrLf & LookUpDate & vbCrLf & "Site Located"
                If debugON Then frmWEB.txtHTML = urlDATA
                DoEvents
                sLINE = urlDATA
    
                If debugON Then
                    Open App.Path & "\Pages\" & sCODE & "~" & Replace(LookUpDate, "/", "~") & "~Item" & tGrab & ".htm" For Output As #2
                    Print #2, sLINE
                    Close #2
                End If
                
                sLINE = Replace(sLINE, "<", vbCrLf & "<", , , vbTextCompare)
                sLINE = Replace(sLINE, ">", ">" & vbCrLf, , , vbTextCompare)
                sLINE = Replace(sLINE, "&amp;", "&", , , vbTextCompare)
                
                Open App.Path & "\temp0002.htm" For Output As #2
                Print #2, sLINE
                Close #2
            
            
                frmGrabber.Label1 = sCHANNEL & vbCrLf & LookUpDate & vbCrLf & "Locating Programmes"
                DoEvents
            
                Open App.Path & "\temp0002.htm" For Input As #2
                While EOF(2) = False
                    Line Input #2, sLINE
                    sLINE = Trim$(sLINE)
                    i = InStr(1, sLINE, "programmeid=", vbTextCompare)
                    If i <> 0 Then
'http://www.radiotimes.com/jsp/frameset/index.html?src=http://www.radiotimes.com:80/ListingsServlet?event=10&channelId=92&programmeId=15873649&jspLocation=/jsp/prog_details.jsp&jspError=/jsp/popup_error.jsp
                        sLINE = Right$(sLINE, Len(sLINE) - (i + 11))
                        i = InStr(1, sLINE, "&", vbTextCompare)
                        sProgrammeId = Trim$(Left$(sLINE, i - 1))
                    
'                        lookUpHREF = "http://www.radiotimes.com/jsp/frameset/index.html?src=http://www.radiotimes.com:80/ListingsServlet?event=10&channelId=" & sCODE & "&programmeId=" & sProgrammeId & "&jspLocation=/jsp/prog_details.jsp&jspError=/jsp/popup_error.jsp"
                        lookUpHREF = "http://www.radiotimes.com:80/ListingsServlet?event=10&channelId=" & sCODE & "&programmeId=" & sProgrammeId & "&jspLocation=/jsp/prog_details.jsp&jspError=/jsp/popup_error.jsp"
                        sProgFound = False
                        For a = 0 To LastUKProg - 1
                            If UKProgrammes(a).ChannelCODE = sCODE And UKProgrammes(a).hREF = lookUpHREF Then
                                sProgFound = True
                                Exit For
                            End If
                        Next a
                
                        If sProgFound = False Then
                            ReDim Preserve UKProgrammes(LastUKProg + 1)
                        
                            UKProgrammes(LastUKProg - 1).LookUpDate = LookUpDate
                            UKProgrammes(LastUKProg - 1).Channel = sCHANNEL
                            UKProgrammes(LastUKProg - 1).ChannelCODE = sCODE
                            UKProgrammes(LastUKProg - 1).hREF = lookUpHREF
                            UKProgrammes(LastUKProg - 1).ProgrammeID = sProgrammeId
                        End If
                    End If
                Wend
                Close #2
            
            
                cHTM = cHTM + 1
                pVAL = (cHTM / maxHTM) * 100
                If pVAL > 100 Then pVAL = 100
                frmGrabber.ProgressBar1.Value = pVAL
                DoEvents
            
            Next tGrab
        Next lDAY
    
sNextChannel:
    Wend
    Close #1

    Locate_Programmes = True

End Function

Public Function LastUKProg() As Integer
    On Local Error GoTo fcntl
    
    LastUKProg = UBound(UKProgrammes)
    Exit Function
    
fcntl:
    LastUKProg = 0

End Function

Public Function Create_XML(xmlFILE As String) As Boolean
    Dim progNO As Integer
    Dim sLINE As String
    Dim i As Integer
    Dim sSTART As Date
    Dim a As Integer
    Dim urlDATA As String
    Dim sCODE As String
    Dim sCHANNEL As String
    Dim sLOAD As String
    Dim sXML As String
    Dim pDET As tpPROGDATA
    
    On Local Error GoTo fERROR
                
    frmGrabber.ProgressBar1.Value = 0
    DoEvents
    
    Open xmlFILE For Output As #1
    sXML = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "ISO-8859-1" & Chr$(34) & "?>"
    Print #1, sXML
    sXML = "<!DOCTYPE UKtv SYSTEM " & Chr$(34) & "xmltv.dtd" & Chr$(34) & ">"
    Print #1, sXML
    sXML = "<tv source-info-url=" & Chr$(34) & "http://www.radiotimes.com" & Chr$(34) & " generator-info-name=" & Chr$(34) & "UKTVGrabber By CornStopper" & Chr$(34) & " generator-info-url=" & Chr$(34) & "" & Chr$(34) & ">"
    Print #1, sXML
    
    Open App.Path & "\channels.dat" For Input As #2
    While EOF(2) = False
        Input #2, sCODE, sCHANNEL, sLOAD
    
        If CBool(sLOAD) Then
            sXML = "<channel id=" & Chr$(34) & sCODE & Chr$(34) & ">"
            Print #1, sXML
            sXML = "    <display-name lang=" & Chr$(34) & "en" & Chr$(34) & ">" & sCHANNEL & "</display-name>"
            Print #1, sXML
            sXML = "    <icon src=" & Chr$(34) & Chr$(34) & " />"
            Print #1, sXML
            sXML = "</channel>"
            Print #1, sXML
        End If
    Wend
    Close #2
    
   
    frmGrabber.Height = 2670
    DoEvents
    
    For progNO = 0 To LastUKProg - 1
'    Debug.Print UKProgrammes(i).LookUpDate & " - " & UKProgrammes(i).hREF
        frmGrabber.Label1 = UKProgrammes(progNO).Channel & vbCrLf & UKProgrammes(progNO).LookUpDate & vbCrLf & "Locating Page"
        DoEvents

        
        If debugON Then frmWEB.txtHTML = "LOADING..." & vbCrLf & UKProgrammes(progNO).hREF
        urlDATA = frmWEB.web.OpenURL(UKProgrammes(progNO).hREF)
        DoEvents
        DoEvents
        sSTART = Now()
            
sConnect:
        ' have time out here
        While frmWEB.web.StillExecuting = True
            DoEvents
        Wend
    
        i = DateDiff("s", sSTART, Now())

        If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 And i <= 90 Then GoTo sConnect
        If InStr(1, frmWEB.web.URL, "www.radiotimes.com", vbTextCompare) = 0 Then GoTo sNextProgramme
        

        ' page loaded
        frmGrabber.Label1 = UKProgrammes(progNO).Channel & vbCrLf & UKProgrammes(progNO).LookUpDate & vbCrLf & "Exporting To XML"
        If debugON Then frmWEB.txtHTML = urlDATA
        DoEvents
        sLINE = urlDATA
            
        If debugON Then
            Open App.Path & "\Pages\temp~" & UKProgrammes(progNO).ChannelCODE & "~" & Replace(UKProgrammes(progNO).LookUpDate, "/", "~") & "~" & UKProgrammes(progNO).ProgrammeID & ".htm" For Output As #2
            Print #2, sLINE
            Close #2
        End If
        
        ' parse the xmldata into a type pDET
        pDET = TidyXML(sLINE)
        
        ' write to xml here
        sXML = "<programme start=" & Chr$(34) & Format$(CDate(pDET.progSTART), "yyyymmdd") & Format$(CDate(pDET.progSTART), "HHMMSS") & " +0000" & Chr$(34) & " stop=" & Chr$(34) & Format$(CDate(pDET.progSTOP), "yyyymmdd") & Format$(CDate(pDET.progSTOP), "HHMMSS") & " +0000" & Chr$(34) & " channel=" & Chr$(34) & UKProgrammes(progNO).ChannelCODE & Chr$(34) & ">"
        Print #1, sXML
        sXML = "    <title lang=" & Chr$(34) & "en" & Chr$(34) & ">" & pDET.progTITLE & "</title>"
        Print #1, sXML
        sXML = "    <sub-title lang=" & Chr$(34) & "en" & Chr$(34) & ">" & pDET.progEPISODE & "</sub-title>"
        Print #1, sXML
        sXML = "    <desc lang=" & Chr$(34) & "en" & Chr$(34) & ">" & pDET.progREVIEW & "</desc>"
        Print #1, sXML
        If pDET.progSUBTITLED Then
            sXML = "    <subtitles>" & Chr$(34) & "Yes" & Chr$(34) & "</subtitles>"
            Print #1, sXML
        End If
        If pDET.progWIDESCREEN Then
            sXML = "    <widescreen>" & Chr$(34) & "Yes" & Chr$(34) & "</widescreen>"
            Print #1, sXML
        End If
        If Len(pDET.progCERT) <> 0 Then
            sXML = "    <certificate>" & Chr$(34) & pDET.progCERT & Chr$(34) & "</certificate>"
            Print #1, sXML
        End If
        If Len(pDET.progDIRECTOR) <> 0 Then
            sXML = "    <director>" & Chr$(34) & pDET.progDIRECTOR & Chr$(34) & "</director>"
            Print #1, sXML
        End If
        If Len(pDET.progYEAR) <> 0 Then
            sXML = "    <filmyear>" & Chr$(34) & pDET.progYEAR & Chr$(34) & "</filmyear>"
            Print #1, sXML
        End If
        If UBoundArray(pDET.progCASTLIST) <> 0 Then
            sXML = "    <castlist>"
            Print #1, sXML
            For a = 0 To UBound(pDET.progCASTLIST) - 1
                sXML = "        <castmember>" & Trim$(pDET.progCASTLIST(a)) & "</castmember>"
                Print #1, sXML
            Next a
            sXML = "    </castlist>"
            Print #1, sXML
        End If
        sXML = "</programme>"
        Print #1, sXML
        
sAllLoaded:
        Close #2

sNextProgramme:
        i = ((progNO + 1) / (LastUKProg - 1)) * 100
        If i > 100 Then i = 100
        frmGrabber.ProgressBar1.Value = i
        DoEvents
    Next progNO

    Create_XML = True
    Exit Function
    
fERROR:
    Resume Next
End Function

Public Function UBoundArray(strARRAY() As String) As Long
    On Local Error GoTo fERROR
    
    UBoundArray = UBound(strARRAY)
    Exit Function

fERROR:
    UBoundArray = 0
End Function

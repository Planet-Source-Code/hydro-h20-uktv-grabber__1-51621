Attribute VB_Name = "modGENERAL"
Option Compare Text
Option Explicit

Public Function IsDesignTime() As Boolean
    On Local Error GoTo Is_ERR
    Debug.Print 1 \ 0
    Exit Function
Is_ERR:
    IsDesignTime = True
    Exit Function
End Function

Public Sub test()
    Dim dataIN As String
    Dim pDET As tpPROGDATA
    
    dataIN = String(FileLen(App.Path & "\test.xml"), " ")
    Open App.Path & "\test.xml" For Binary Access Read As #1
    Get #1, , dataIN
    Close #1
    
    pDET = TidyXML(dataIN)
    
End Sub

Public Function TidyXML(xmlDATA As String) As tpPROGDATA
    Dim a As Long
    Dim b As Long
    Dim rDATA As String
    Dim tmpDATA As String
    Dim sDATA() As String
    Dim tmpUKTVProg As tpPROGDATA
    Dim progDATE As Date
    Dim tDATA As String
    Dim sTIME1 As String
    Dim sTIME2 As String
    Dim RunOverYear As Boolean
    Dim i As Integer
    Dim crTXT As String
    Dim allDOTS As Boolean
    Dim pDATA As String
    
    tmpDATA = xmlDATA
    
    For a = 1 To 255
        tmpDATA = Replace(tmpDATA, "&#" & a & ";", Chr$(a), , , vbBinaryCompare)
    Next a
    tmpDATA = Replace(tmpDATA, vbTab, "", , , vbBinaryCompare)
    tmpDATA = Replace(tmpDATA, "&amp;", "&", , , vbTextCompare)
    tmpDATA = Replace(tmpDATA, "&nbsp;", " ", , , vbTextCompare)

    a = InStr(1, tmpDATA, "&progDate=", vbTextCompare)
    If a <> 0 Then progDATE = Mid$(tmpDATA, a + 10, 10)
    a = InStr(1, tmpDATA, "&title=", vbTextCompare)
    If a <> 0 Then
        a = a + 7
        b = InStr(1, tmpDATA, "&channelname=", vbTextCompare)
        tmpUKTVProg.progTITLE = Mid$(tmpDATA, a, b - a)
    End If

'    tmpDATA = Replace(tmpDATA, "<br/>", " ", , , vbTextCompare)
'    tmpDATA = Replace(tmpDATA, "<b>", "", , , vbTextCompare)
'    tmpDATA = Replace(sLINE, "</b>", "", , , vbTextCompare)
'        sLINE = Replace(sLINE, "<br>", vbCrLf, , , vbTextCompare)
'        sLINE = Replace(sLINE, "<", vbCrLf & "<", , , vbTextCompare)
'        sLINE = Replace(sLINE, ">", ">" & vbCrLf, , , vbTextCompare)
    
    ' first search and remove script
    a = InStr(1, tmpDATA, "<script", vbTextCompare)
    While a <> 0
        b = InStr(a, tmpDATA, "</script>", vbTextCompare)
        If b <> 0 Then
            rDATA = Mid$(tmpDATA, a, (b - a) + 9)
            tmpDATA = Replace(tmpDATA, rDATA, "", , , vbBinaryCompare)
        End If
        a = InStr(1, tmpDATA, "<script", vbTextCompare)
    Wend

    a = InStr(1, tmpDATA, "<!-- NORMAL PROGRAMME DETAILS STOP-->", vbTextCompare)
    If a <> 0 Then
        'rDATA = Mid$(tmpDATA, a, (Len(tmpDATA) - a) + 1)
        'tmpDATA = Replace(tmpDATA, rDATA, "", , , vbBinaryCompare)
        tmpDATA = Left$(tmpDATA, a - 1)
    End If

    'remove all tag data
    a = InStr(1, tmpDATA, "<", vbBinaryCompare)
    While a <> 0
        b = InStr(a, tmpDATA, ">", vbTextCompare)
        If b <> 0 Then
            rDATA = Mid$(tmpDATA, a, (b - a) + 1)
            tmpDATA = Replace(tmpDATA, rDATA, "", , , vbBinaryCompare)
        End If
        a = InStr(1, tmpDATA, "<", vbTextCompare)
    Wend
    
    a = InStr(1, tmpDATA, "Related websites", vbTextCompare)
    If a <> 0 Then
        'rDATA = Mid$(tmpDATA, a, (Len(tmpDATA) - a) + 1)
        'tmpDATA = Replace(tmpDATA, rDATA, "", , , vbBinaryCompare)
        tmpDATA = Left$(tmpDATA, a - 1)
    End If

    a = InStr(1, tmpDATA, "Related features", vbTextCompare)
    If a <> 0 Then
        'rDATA = Mid$(tmpDATA, a, (Len(tmpDATA) - a) + 1)
        'tmpDATA = Replace(tmpDATA, rDATA, "", , , vbBinaryCompare)
        tmpDATA = Left$(tmpDATA, a - 1)
    End If
    
    ' extract start / stop times
    a = InStr(1, tmpDATA, "Time:", vbTextCompare)
    a = a + 5
    rDATA = ""
    For b = a To Len(tmpDATA)
        tDATA = Mid$(tmpDATA, b, 1)
        If Asc(tDATA) <> 13 And Asc(tDATA) <> 10 Then
            rDATA = rDATA & tDATA
            If InStr(1, rDATA, " to ", vbTextCompare) <> 0 Then
                If Right$(rDATA, 2) = "am" Or Right$(rDATA, 2) = "pm" Then Exit For
            End If
        End If
    Next b
    rDATA = Trim$(rDATA)
    a = InStr(1, rDATA, "to")
    sTIME1 = Trim$(Left$(rDATA, a - 1))
    sTIME2 = Trim$(Right$(rDATA, Len(rDATA) - (a + 2)))
    If Right$(sTIME1, 2) <> Right$(sTIME2, 2) Then RunOverYear = True
    tmpUKTVProg.progSTART = Format$(progDATE, "dd/mmm/yyyy") & " " & sTIME1
    If RunOverYear And Right$(sTIME1, 2) = "pm" Then
        tmpUKTVProg.progSTOP = Format$(progDATE + 1, "dd/mmm/yyyy") & " " & sTIME2
    Else
        tmpUKTVProg.progSTOP = Format$(progDATE, "dd/mmm/yyyy") & " " & sTIME2
    End If
    
   
    ' parse data into an array
    sDATA = Split(tmpDATA, vbCrLf, , vbBinaryCompare)
    
    ' get episode if exists
    For a = 0 To UBound(sDATA) - 1
        rDATA = Trim$(sDATA(a))
        If rDATA = "episode" Then
            Do
                a = a + 1
                rDATA = Trim$(sDATA(a))
                If Len(rDATA) <> 0 Then tmpUKTVProg.progEPISODE = rDATA
            Loop Until Len(tmpUKTVProg.progEPISODE) <> 0
            Exit For
        End If
    Next a
    
    ' get some other data
    For a = 0 To UBound(sDATA) - 1
        rDATA = Trim$(sDATA(a))
        If Len(rDATA) <> 0 Then
            If InStr(1, rDATA, "Subtitled", vbTextCompare) <> 0 Then tmpUKTVProg.progSUBTITLED = True
            If InStr(1, rDATA, "Widescreen", vbTextCompare) <> 0 Then tmpUKTVProg.progWIDESCREEN = True
            If InStr(1, rDATA, "Certificate:", vbTextCompare) <> 0 Then
                tmpUKTVProg.progCERT = Trim$(Right$(rDATA, Len(rDATA) - 12))
            End If
            If InStr(1, rDATA, "Directed By:", vbTextCompare) <> 0 Then
                tmpUKTVProg.progDIRECTOR = Trim$(Right$(rDATA, Len(rDATA) - Len("Directed By:")))
            End If
            If InStr(1, rDATA, "Filmed in:", vbTextCompare) <> 0 Then
                tmpUKTVProg.progYEAR = Trim$(Right$(rDATA, Len(rDATA) - Len("Filmed in:")))
            End If
        End If
    Next a
        
    ' get review
    For a = 0 To UBound(sDATA) - 1
        rDATA = Trim$(sDATA(a))
        If rDATA = "review" Then
            a = a + 1
            rDATA = Trim$(sDATA(a))
            While a < UBound(sDATA) And rDATA <> "Cast List"
                If Len(rDATA) <> 0 Then
                    If Len(tmpUKTVProg.progREVIEW) <> 0 Then tmpUKTVProg.progREVIEW = tmpUKTVProg.progREVIEW & " "
                    tmpUKTVProg.progREVIEW = tmpUKTVProg.progREVIEW & rDATA
                End If
                a = a + 1
                If a < UBound(sDATA) Then rDATA = Trim$(sDATA(a))
            Wend
            Exit For
        End If
    Next a
    
    ' get cast list
    tDATA = ""
    crTXT = ""
    pDATA = ""
    For a = 0 To UBound(sDATA) - 1
        rDATA = Trim$(sDATA(a))
        If rDATA = "cast list" Then
            For b = a + 1 To UBound(sDATA) - 1
                rDATA = Trim$(sDATA(b))
                If Len(rDATA) <> 0 Then
                    allDOTS = True
                    For i = 1 To Len(rDATA)
                        If Mid$(rDATA, i, 1) <> "." Then
                            allDOTS = False
                            Exit For
                        End If
                    Next i
                    
                    If allDOTS Then
                        tDATA = tDATA & String(30 - Len(pDATA), ".")
                        crTXT = ","
                    Else
                        tDATA = tDATA & rDATA & crTXT
                        crTXT = ""
                    End If
                    pDATA = rDATA
                End If
            Next b
            If Right$(tDATA, 1) <> "," Then tDATA = tDATA & ","
            tmpUKTVProg.progCASTLIST = Split(tDATA, ",")
            Exit For
        End If
    Next a
    
    
    If Len(tmpUKTVProg.progDIRECTOR) <> 0 Then tmpUKTVProg.progREVIEW = Replace(tmpUKTVProg.progREVIEW, "Directed by: " & tmpUKTVProg.progDIRECTOR, "", , , vbTextCompare)
    If Len(tmpUKTVProg.progYEAR) <> 0 Then tmpUKTVProg.progREVIEW = Replace(tmpUKTVProg.progREVIEW, "Filmed in: " & tmpUKTVProg.progYEAR, "", , , vbTextCompare)
    tmpUKTVProg.progREVIEW = Trim$(tmpUKTVProg.progREVIEW)
    
    
    TidyXML = tmpUKTVProg
End Function

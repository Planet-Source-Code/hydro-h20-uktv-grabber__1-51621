Attribute VB_Name = "modTYPES"
Public Type UKTVPrograms
    LookUpDate As String
    ChannelCODE As String
    Channel As String
    hREF As String
    ProgrammeID As String
End Type

Public Type tpPROGDATA
    progTITLE As String
    progEPISODE As String
    progSTART As String
    progSTOP As String
    progREVIEW As String
    progSUBTITLED As Boolean
    progCERT As String
    progYEAR As String
    progWIDESCREEN As Boolean
    progDIRECTOR As String
    progCASTLIST() As String
End Type


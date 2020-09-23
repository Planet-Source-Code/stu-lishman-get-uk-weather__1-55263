Attribute VB_Name = "Module1"
Global SelAcc As String, textline(10000), lastline, Days(5), inf1(5), inf2(5), inf3(5), inf4(5), inf5(5), inf6(5), inf7(5), inf8(5), inf9, inf10, inf11, inf12, inf13, inf14, Area(1000), Code(1000), Location(1000), PartCount(1000), Country(1000), State(1000), LstStationCnt, CurrentArea, AreaSelect
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

' // Returns true if connected to the internet.
Public Function CheckConnection() As Boolean
Dim result As Boolean
    result = InternetGetConnectedState(0&, 0&)  ' Simply test for an internet socket.
    If result = False Then
        CheckConnection = False
    Else
        CheckConnection = True
    End If
End Function

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function
Public Function FileExists(filespec)
Set fso = CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(filespec)) Then FileExists = True Else FileExists = False
End Function


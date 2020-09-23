Attribute VB_Name = "Module1"
Public Type tData
    Key As String
    Value As String
    End Type

Public Function FindArtist(theArtist As String)
    Dim TheData As String, TheSize As Integer, X As Integer, tmpVar As String

    TheSize = Len("a=search&p=1&s=" + theArtist + "&l=artist")
    tmpVar = "POST " & "/cgi-exe/am.cgi" & " HTTP/1.1" & vbCrLf & _
    "Host: www.letssingit.com" & vbCrLf & _
    "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
    "Accept-Encoding: gzip, deflate" & vbCrLf & _
    "Content-Length: " & TheSize & vbCrLf & _
    "Connection: Keep-Alive" & vbCrLf & vbCrLf & _
    "a=search&p=1&s=" + theArtist + "&l=artist" & vbCrLf

    FindArtist = tmpVar
End Function

Public Function BrowserPage(textbox As textbox, Browser As WebBrowser)
Dim doc As Object
Set doc = Browser.Document
With doc

.Open
    .write textbox.Text
    
.Close
End With
End Function

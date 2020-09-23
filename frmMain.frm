VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lyrics Finder v"
   ClientHeight    =   9660
   ClientLeft      =   3270
   ClientTop       =   1905
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   9570
   Begin ComctlLib.StatusBar Status1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   9405
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTest 
      Height          =   330
      Left            =   7170
      TabIndex        =   10
      Text            =   "<PRE style=""font:12px arial"">"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   15
      TabIndex        =   6
      Top             =   -45
      Width           =   5445
      Begin VB.CommandButton Command3 
         Caption         =   "Library"
         Height          =   300
         Left            =   1155
         TabIndex        =   13
         Top             =   810
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "detach \/"
         Height          =   300
         Left            =   60
         TabIndex        =   12
         Top             =   810
         Width           =   1050
      End
      Begin VB.TextBox txtArtist 
         Height          =   330
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   4590
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Default         =   -1  'True
         Height          =   405
         Left            =   3765
         Picture         =   "frmMain.frx":628A
         TabIndex        =   7
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Artist:"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5190
      Left            =   15
      TabIndex        =   4
      Top             =   1110
      Width           =   5445
      Begin MSComDlg.CommonDialog cd 
         Left            =   3840
         Top             =   4110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Save Lyrics to File"
      End
      Begin RichTextLib.RichTextBox txtLyrics 
         Height          =   4905
         Left            =   60
         TabIndex        =   5
         Top             =   195
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8652
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D77C
      End
   End
   Begin VB.TextBox txtTemp2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   9450
      Visible         =   0   'False
      Width           =   2700
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7305
      Top             =   8550
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtNoResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   855
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":D7F3
      Top             =   7500
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7350
      Top             =   7785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.TextBox txtTemp 
      Height          =   2130
      Left            =   3090
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7140
      Visible         =   0   'False
      Width           =   3795
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6240
      Left            =   5505
      TabIndex        =   3
      Top             =   60
      Width           =   4050
      ExtentX         =   7144
      ExtentY         =   11007
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu save 
         Caption         =   "&Save to File"
      End
      Begin VB.Menu print 
         Caption         =   "&Print Lyrics"
      End
      Begin VB.Menu db 
         Caption         =   "&Add to Database"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tUrl As String
Dim SearchState As String
Dim AttachState As Boolean
Option Explicit
Sub Pause(duration)
'This will pause for the duration [duration is in seconds]
Dim Current As Long
Current = Timer
Do Until Timer - Current >= duration
    DoEvents
Loop
End Sub

Private Sub Command1_Click()
Dim WebHost As String

SearchState = ""

On Error Resume Next
WebHost = "www.letssingit.com"

'Clear the source textbox if it isn't
If txtTemp.Text <> "" Then txtTemp.Text = ""

'Close the Winsock incase it's already open with
'another server.
Winsock.Close

'Tell Winsock what server we're connecting to and
'which port we're using
Winsock.RemoteHost = WebHost
Winsock.RemotePort = 80

'Connect to the server
Status1.SimpleText = "Connecting..."
Winsock.Connect
End Sub

Private Sub Command2_Click()
If Command2.Caption = "detach \/" Then
AttachState = True
frmViewer.Show
frmViewer.Height = Frame2.Height
frmViewer.Width = Frame2.Width
Frame2.Visible = False
WebBrowser1.Top = 1185
WebBrowser1.Left = Frame1.Left
WebBrowser1.Width = 4365
frmMain.Height = 8100
txtArtist.Width = 3540
Command1.Left = 2730
Frame1.Width = 4365
frmMain.Width = Frame1.Width + 100
Command2.Caption = "attach"
frmViewer.txtLyrics.TextRTF = txtLyrics.TextRTF
Else
AttachState = False
Frame2.Visible = True
WebBrowser1.Top = 60
WebBrowser1.Left = 5520
WebBrowser1.Width = 4050
txtArtist.Width = 4590
Frame1.Width = 5445
Command1.Left = 3765
frmMain.Width = 9660
frmMain.Height = 6960
Command2.Caption = "detach \/"
frmViewer.Hide
End If
End Sub

Private Sub Command3_Click()
frmtLyrics.Show
End Sub

Private Sub db_Click()
frmtLyrics.Show
frmtLyrics.Width = 7695
frmtLyrics.Command4.Caption = "Edit <-"
frmtLyrics.AddNew
frmtLyrics.txtFields(2).Text = txtLyrics.Text
frmtLyrics.SetFocus
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://www.freewebs.com/jimbohsp/"
mnuMenu.Visible = False
Me.Height = 6960
Me.Caption = Me.Caption & App.Major & "." & App.Minor
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
If State = 3 Then
Status1.SimpleText = "Recieving..."
ElseIf State = 7 Then
Status1.SimpleText = "Done"
End If
End Sub

Private Sub lblLink_Click()

End Sub

Private Sub print_Click()
On Error Resume Next
Printer.Print ""
txtLyrics.SelPrint Printer.hDC
Printer.EndDoc
End Sub

Private Sub save_Click()
cd.Filter = "All Files(*.*)|*.*|Rich Text(*.rtf)|*.rtf"
cd.FilterIndex = 2
cd.ShowSave
If cd.FileName <> "" Then
txtLyrics.SaveFile cd.FileName
End If
End Sub

Private Sub txtLyrics_Change()
If AttachState = True Then
frmViewer.txtLyrics.TextRTF = txtLyrics.TextRTF
End If
End Sub

Private Sub txtLyrics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'if they right click, 1=left, 2=right
    frmMain.PopupMenu mnuMenu 'show popup menu
    Else
    DoEvents
End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim FinalURL
Dim theStart
Dim theStart2
Dim b() As Byte
Dim txt As String
Dim t As Integer


 If SearchState = "ArtistPage" Then
 ' on actual artist's page?
  FinalURL = Right(URL, Len(URL) - 3)
  
    ' cancel navigate?
   If URL <> "" Then
   On Error Resume Next
    If InStr(1, URL, "templyrics", vbTextCompare) = 0 Then
    Cancel = True
    b() = Inet1.OpenURL("http://www.letssingit.com/" & FinalURL, 1)
    txt = ""
    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    txtTemp.Text = txt
    theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD><PRE", vbTextCompare)
    txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
    theStart = InStr(1, txtTemp.Text, "</PRE></TD></TR></TABLE>", vbTextCompare)
    txtTemp.Text = Left(txtTemp.Text, theStart - 1)
    txtTemp.Text = Replace(txtTemp.Text, "<TABLE><TR><TD>", "")
    txtTemp.Text = Replace(txtTemp.Text, txtTest, "")
    txtLyrics.TextRTF = txtTemp.Text
    End If
   End If
    
 ElseIf SearchState = "Results" Then
 ' on search result page?
 FinalURL = Right(URL, Len(URL) - 3)
 'If URL <> "" Then
 Cancel = True
 'End If

     b() = Inet1.OpenURL("http://www.letssingit.com/" & FinalURL, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    txtTemp.Text = txt
    
    'If URL <> "" Then
    On Error Resume Next
     ' a load of parsing bullshit, nothing to be changed
theStart = InStr(1, txtTemp.Text, "</TR>" & vbCrLf & "</TABLE>", vbTextCompare)
txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD>", vbTextCompare)
theStart2 = InStr(1, txtTemp.Text, "</TR></TABLE>", vbTextCompare)
txtTemp.Text = Left(txtTemp.Text, theStart2 - 1)
txtTemp.Text = Replace(txtTemp.Text, vbCrLf, "")

 ' replace all the unnecessary tags
txtTemp.Text = Replace(txtTemp.Text, "</TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "13de", "")
txtTemp.Text = Replace(txtTemp.Text, "<TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "</TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TD>", "")
txtTemp.Text = Replace(txtTemp.Text, "11fc", "")
txtTemp.Text = Replace(txtTemp.Text, "7f5", "")

 ' to look nice
txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"

 ' save temporary html file and load it into webbrowser
BrowserPage txtTemp, WebBrowser1

SearchState = ""

'WebBrowser1.Navigate "C:\templyrics000.html"

SearchState = "ArtistPage"
'End If
 End If
 Exit Sub
End Sub

Private Sub WebBrowser2_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Winsock_Close()
Dim theStart
Dim theStart2

Status1.SimpleText = "Done"

SearchState = ""

'Let all incoming data finish being received
Pause 0.5

'The server closed the connection, now we need
'to close it on this side.
Winsock.Close: Winsock.Tag = "CLOSED"


On Error Resume Next

 ' check to see if Artist is found
 
If InStr(1, txtTemp.Text, "no search results", vbTextCompare) <> 0 Then
 ' nothing found!
 ' show page
BrowserPage txtNoResults, WebBrowser1
 SearchState = "None"
Exit Sub
ElseIf InStr(1, txtTemp.Text, "<TABLE><TR><TD>Showing", vbTextCompare) <> 0 Then
' show search results (parsing)
 theStart = InStr(1, txtTemp.Text, "</TD></TR></TABLE><BR>", vbTextCompare)
 txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
 theStart = InStr(1, txtTemp.Text, "<BR><BR>Select page", vbTextCompare)
 txtTemp.Text = Left(txtTemp.Text, theStart - 1)
 txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"
SearchState = "Results"
 BrowserPage txtTemp, WebBrowser1
Exit Sub
End If

 ' a load of parsing bullshit, nothing to be changed
theStart = InStr(1, txtTemp.Text, "</TR>" & vbCrLf & "</TABLE>", vbTextCompare)
txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD>", vbTextCompare)
theStart2 = InStr(1, txtTemp.Text, "</TR></TABLE>", vbTextCompare)
txtTemp.Text = Left(txtTemp.Text, theStart2 - 1)
txtTemp.Text = Replace(txtTemp.Text, vbCrLf, "")

 ' replace all the unnecessary tags
txtTemp.Text = Replace(txtTemp.Text, "</TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "13de", "")
txtTemp.Text = Replace(txtTemp.Text, "<TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "</TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TD>", "")
txtTemp.Text = Replace(txtTemp.Text, "<SCRIPT>", "")
txtTemp.Text = Replace(txtTemp.Text, "</SCRIPT>", "")
txtTemp.Text = Replace(txtTemp.Text, "11fc", "")
txtTemp.Text = Replace(txtTemp.Text, "7f5", "")
txtTemp.Text = Replace(txtTemp.Text, "e16", "")
txtTemp.Text = Replace(txtTemp.Text, "2c94", "")

 ' to look nice
txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"

SearchState = ""

BrowserPage txtTemp, WebBrowser1

SearchState = "ArtistPage"
End Sub

Private Sub Winsock_Connect()
Dim getString As String, ShortWebSite As String
Winsock.Tag = "OPEN"
On Error Resume Next

'Send the command to the server
Winsock.SendData FindArtist(txtArtist.Text)
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
Status1.SimpleText = "Recieving..."
On Error Resume Next
'Get the incoming data that was sent from the server
If Winsock.Tag = "OPEN" Then Winsock.GetData Buffer

'Add it to the current contents of the HTML Source
'textbox.
txtTemp.Text = txtTemp.Text & Buffer
End Sub


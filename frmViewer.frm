VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmViewer 
   Caption         =   "Lyrics"
   ClientHeight    =   5190
   ClientLeft      =   8970
   ClientTop       =   3210
   ClientWidth     =   4770
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4770
   Begin RichTextLib.RichTextBox txtLyrics 
      Height          =   1140
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   2011
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmViewer.frx":628A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
txtLyrics.Width = Me.Width - 150
txtLyrics.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
End Sub

Private Sub txtLyrics_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then 'if they right click, 1=left, 2=right
    frmMain.PopupMenu frmMain.mnuMenu 'show popup menu
    Else
    DoEvents
End If
End Sub

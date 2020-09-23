VERSION 5.00
Begin VB.Form frmView 
   ClientHeight    =   5220
   ClientLeft      =   7260
   ClientTop       =   3960
   ClientWidth     =   5595
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5595
   Begin VB.TextBox Text1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   5415
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
Text1.Width = Me.Width - 115
Text1.Height = Me.Height - 460
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Cancel = True
End Sub

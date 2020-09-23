VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   4995
   ClientTop       =   5775
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin VB.TextBox t1 
      Height          =   780
      Left            =   1905
      TabIndex        =   0
      Text            =   """"
      Top             =   1155
      Width           =   1635
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu playy 
         Caption         =   "Play Song"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Public Sub RunShellExecute(sTopic As String, _
                           sFile As Variant, _
                           sParams As Variant, _
                           sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub


Private Sub playy_Click()
  Dim sTopic As String
   Dim sFile As String
   Dim sParams As Variant
   Dim sDirectory As Variant
               sTopic = "Open"
               sFile = frmtLyrics.txtFields(1).Text
               sParams = 0&
               sDirectory = 0&

Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)

End Sub

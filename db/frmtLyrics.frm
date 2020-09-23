VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmtLyrics 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lyrics"
   ClientHeight    =   4095
   ClientLeft      =   510
   ClientTop       =   2535
   ClientWidth     =   7605
   Icon            =   "frmtLyrics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmtLyrics.frx":0442
   ScaleHeight     =   4095
   ScaleWidth      =   7605
   Begin MSComDlg.CommonDialog C1 
      Left            =   2775
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Find Song"
      InitDir         =   "C:\Music"
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   270
      Left            =   7080
      TabIndex        =   19
      Top             =   495
      Width           =   435
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit ->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1065
      TabIndex        =   18
      Top             =   3690
      Width           =   1305
   End
   Begin VB.CommandButton Command3 
      Caption         =   "no mask"
      Height          =   300
      Left            =   4575
      TabIndex        =   17
      Top             =   5355
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "show Password"
      Height          =   300
      Left            =   3135
      TabIndex        =   16
      Top             =   5355
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   930
      TabIndex        =   15
      Top             =   5235
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6000
      TabIndex        =   12
      Top             =   3705
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   11
      Top             =   3705
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   10
      Top             =   3705
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   9
      Top             =   3705
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   8
      Top             =   3705
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   3135
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4995
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      DataField       =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Index           =   2
      Left            =   3645
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   810
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      DataField       =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3645
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      DataField       =   "Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3645
      TabIndex        =   1
      Top             =   165
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   3360
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Index           =   3
      Left            =   2175
      TabIndex        =   7
      Top             =   4995
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Lyrics:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2685
      TabIndex        =   6
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2685
      TabIndex        =   5
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Song:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2685
      TabIndex        =   0
      Top             =   165
      Width           =   1005
   End
End
Attribute VB_Name = "frmtLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Keeps the Code clean
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim tButton

Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Command2_Click()
Dim clsDS2 As New clsDS2
MsgBox "Your Password is:   " & vbCrLf & vbCrLf & _
clsDS2.DecryptString(txtFields(3), True), vbInformation, "Password Manager..."
End Sub

Private Sub Command3_Click()
'On Error Resume Next
If Command3.Caption = "no mask" Then
Command3.Caption = "mask"
txtFields(3).PasswordChar = ""
Else
Command3.Caption = "no mask"
txtFields(3).PasswordChar = "*"
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Edit ->" Then
frmtLyrics.Width = 7695
Command4.Caption = "Edit <-"
Else
frmtLyrics.Width = 2550
Command4.Caption = "Edit ->"
End If
End Sub

Private Sub Command5_Click()
If txtFields(1).Locked = False Then

C1.FileName = ""
C1.ShowOpen

 If C1.FileName <> "" Then
 txtFields(1).Text = C1.FileName
  If txtFields(0).Text = "" Then
  txtFields(0).Text = Replace(C1.FileTitle, ".mp3", "", 1, -1)
  End If
 End If

End If
End Sub
Public Sub ClickIt()
cmdAdd_Click
End Sub
Private Sub Form_Initialize()
    Dim comctls As INITCOMMONCONTROLSEX_TYPE  ' identifies the control to register
    Dim retval As Long                        ' generic return value
    With comctls
        .dwSize = Len(comctls)
        .dwICC = ICC_INTERNET_CLASSES
    End With
    retval = InitCommonControlsEx(comctls)
End Sub

Private Sub Form_Load()
frmtLyrics.Width = 2550
frmMenu.Show
frmMenu.Visible = False
  Dim db As Connection
  
    
   On Error GoTo ErrHandler
   
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Lyrics.mdb"
  Set rs = New Recordset
  rs.Open "select Location,Username,Notes,Password from Manager", db, adOpenStatic, adLockOptimistic

  Dim oText As textbox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = rs
  Next


Call Styleme
Call Lockme
Call listme
'works great
App.TaskVisible = False

  mbDataChanged = False

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub

Private Sub Form_Resize()
  On Error Resume Next

End Sub

Private Sub List1_Click()
On Error Resume Next
List1.ToolTipText = List1.Text
Call Search(List1.Text, rs, rs.Fields("Location"))
If txtFields(2) <> "" Then
frmMain.txtLyrics.Text = txtFields(2).Text
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmMenu.edit
End If
End Sub

Private Sub RS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(rs.AbsolutePosition)
End Sub

Private Sub RS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub
Public Function AddNew()
txtFields(3) = "none"
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    Call unLockme
    
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Function
AddErr:
  MsgBox Err.Description
End Function
Private Sub cmdAdd_Click()
txtFields(3) = "none"
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    Call unLockme
    
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

If MsgBox("Delete Entry??", vbCritical + vbYesNo, "Password Manager?") = vbYes Then
  With rs
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
    Call listme
  End With
  Exit Sub
End If

End Sub

Private Sub cmdEdit_Click()
On Error Resume Next

Call unLockme
txtFields(3).Text = ""

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub
  
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs.CancelUpdate
  Call Lockme
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
Dim clsDS2 As New clsDS2
  On Error GoTo UpdateErr
 
'encrypting Password before Safing
txtFields(3).Text = clsDS2.EncryptString(txtFields(3), True)
  
  rs.UpdateBatch adAffectAll
  
 Call listme
 Call Lockme
 
  If mbAddNewFlag Then
    rs.MoveLast              'move to the new record
    
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    rs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
End Sub


Private Sub Styleme()
Dim intfields As Integer

        For intfields = 0 To 3
        MakeFlat txtFields(intfields).hwnd
        Next
        
MakeFlat List1.hwnd
CButton cmdAdd
CButton cmdEdit
CButton cmdUpdate
CButton cmdCancel
CButton cmdDelete
CButton Command1
CButton Command2
CButton Command3
End Sub


Private Sub Lockme()
Dim intfields As Integer

         For intfields = 0 To 3
         frmtLyrics.txtFields(intfields).Locked = True
         Next
End Sub

Private Sub unLockme()
Dim intfields As Integer

         For intfields = 0 To 3
         frmtLyrics.txtFields(intfields).Locked = False
         Next
End Sub

Private Sub listme()
On Error Resume Next
 List1.Clear
 rs.MoveFirst
 While Not rs.EOF
   List1.AddItem rs.Fields("Location")
   rs.MoveNext
 Wend
rs.MoveFirst
End Sub


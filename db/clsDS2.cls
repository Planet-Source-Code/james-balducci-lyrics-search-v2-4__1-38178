VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DS2 (Digitally Secure Encryption 2)
' Copyright © 2001-2002 David Greenwood and David Midkiff.
' Notes:   This Code is ONLY for personal use. To use Digitally Secure
'          Encryption or techiques derived from this code in commercial
'          products contact David Greenwood for authorisation.
'
' Author: David Greenwood <dsguk@lycos.com>
' Contributions/Analysis: David Midkiff <mdj2023@hotmail.com>
'
' ----------
' Distributor Notes:
'
' Notes from David Greenwood:
' <fill in here>
'
' Notes from David Midkiff:
' DS2 is a quite powerful cipher designed with the Visual Basic environment
' in mind. After contributing more efficient string encryption and file
' encryption routines to the original DS1 cipher (created by David Greenwood),
' David Greenwood then released a much more beefed up version of the cipher
' (DS1.5) with my design improvements included. After close analysis I
' discovered several weaknesses which I then fixed resulting in the DS2 cipher.
' You can read about it in the readme file.
'
' Close analysis of this algorithm reveals a very secure and fast design.
' While it appears certain forms of differential attacks may be effective
' on lower-round variants of this algorithm, nothing is certain and I would
' certainly recommend this for use in cryptographic solutions. Key-sizes are
' variable length. Key-size limit is undeterminable (huge) as far as I can tell.

Option Explicit
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Event Progress(Percent As Integer)
Const Rounds = 16

Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Private AddTbl(255, 255) As Byte
Private XTbl(255, 255) As Byte
Private LsTbl(255, 255) As Byte
Private RsTbl(255, 255) As Byte
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    ' Speed string concatenation
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function
Public Function EnHex(Data As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function
Private Function FileExist(FileName As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    FileExist = False
    Exit Function

begin:
    Call FileLen(FileName)
    FileExist = True
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    EncryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO))
        Buffer(LOF(1)) = 32
        Get #FileO, , Buffer()
    Close #FileO
    
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = EncryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    EncryptFile = True
End Function
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputInHex As Boolean) As String
    EncryptString = StrConv(EncryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
    If OutputInHex = True Then EncryptString = EnHex(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextInHex As Boolean) As String
    If IsTextInHex = True Then Text = DeHex(Text)
    DecryptString = StrConv(DecryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    DecryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = DecryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    DecryptFile = True
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Public Function EncryptByte(DS() As Byte, Key() As Byte)
    Dim tmp2() As Byte, p As Integer, i As Long, Bound As Integer, r As Integer
    
    ' Expands keys if too small (weak keys may result)
    If UBound(Key) < 0 Then
        ReDim Key(2): Key(0) = 1: Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0): tmp2(0) = Key(0): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 1 Then
        ReDim tmp2(1): tmp2(0) = Key(0): tmp2(1) = Key(1): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = tmp2(1): Key(2) = 1
    End If

    ' Prepares array for salt
    ReDim tmp2((UBound(DS)) + 5)

    ' Generate salt
    Randomize Timer
    tmp2(0) = Int((Rnd * 254) + 1)
    tmp2(1) = Int((Rnd * 254) + 1)
    tmp2(2) = Int((Rnd * 254) + 1)
    tmp2(UBound(tmp2)) = Int((Rnd * 254) + 1)
    tmp2(UBound(tmp2) - 1) = Int((Rnd * 254) + 1)
    
    ' Fill block with salt
    On Error Resume Next
    Call CopyMem(tmp2(3), DS(0), UBound(DS) + 1)
    ReDim DS(UBound(tmp2)) As Byte
    DS() = tmp2()
    ReDim tmp2(0)
    
    ' Encrypt block
    For r = 1 To Rounds
        Bound = (UBound(Key))
        p = 0
        For i = 0 To UBound(DS) - 1
            If p = Bound Then p = 0
            DS(i) = XTbl(DS(i), AddTbl(DS(i + 1), Key(p)))
            DS(i + 1) = XTbl(DS(i), DS(i + 1))
            DS(i) = XTbl(DS(i), AddTbl(DS(i + 1), Key(p + 1)))
            p = p + 1
        Next
        RaiseEvent Progress((r / Rounds) * 100)
    Next
    
    EncryptByte = DS()
End Function
Public Function DecryptByte(DS() As Byte, Key() As Byte)
    Dim tmp2() As Byte, p As Long, i As Long, Bound As Integer, r As Integer
    
    ' Expands keys if too small
    If UBound(Key) < 0 Then
        ReDim Key(2): Key(0) = 1: Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0): tmp2(0) = Key(0): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 1 Then
        ReDim tmp2(1): tmp2(0) = Key(0): tmp2(1) = Key(1): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = tmp2(1): Key(2) = 1
    End If
    
    ' Decrypt block
    For r = 1 To Rounds
        Bound = (UBound(Key))
        p = (UBound(DS)) Mod (UBound(Key))
        
        For i = (UBound(DS)) To 1 Step -1
            If p = 0 Then p = Bound
            DS(i - 1) = XTbl(DS(i - 1), AddTbl(DS(i), Key(p)))
            DS(i) = XTbl(DS(i - 1), DS(i))
            DS(i - 1) = XTbl(DS(i - 1), AddTbl(DS(i), Key(p - 1)))
            p = p - 1
        Next
        RaiseEvent Progress((r / Rounds) * 100)
    Next
    
    ' Filter out added salt strings
    tmp2() = DS()
    ReDim DS((UBound(tmp2)) - 4) As Byte
    Call CopyMem(DS(0), tmp2(3), UBound(DS))
    ReDim Preserve DS(UBound(DS) - 1) As Byte
    
    DecryptByte = DS()
End Function
Private Sub Class_Initialize()
Dim i As Integer, j As Integer

For i = 0 To 255
    For j = 0 To 255
        XTbl(i, j) = CByte(i Xor j)
        AddTbl(i, j) = CByte((i + j) Mod 255)
    Next
Next
End Sub


'To Decrypt and encrpyt:
'command1_click
'Dim DS2 As New clsDS2
'MsgBox "Your Password is:   " & vbrclf & vbrclf & _
'DS2.DecryptString(txtfields(1), True), vbInformation, "Password Manager..."
'txtfields(1).Text = DS2.EncryptString(txtfields(1), True)




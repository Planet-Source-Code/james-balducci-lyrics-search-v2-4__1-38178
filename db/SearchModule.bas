Attribute VB_Name = "SearchModule"
'********************************************************************
'~Created by Adam Lankford 4/18/2001
'********************************************************************

Public Function Search(parameter As String, rs As ADODB.Recordset, X As Field) As Boolean
    Dim foundFlag As Boolean

    With rs
        If .RecordCount > 0 Then
            .MoveFirst
                For i = 1 To .RecordCount
                    If X = parameter Then
                        foundFlag = True
                        i = .RecordCount
                    End If
                    If foundFlag = False Then
                        .MoveNext
                    End If
                Next i
                If foundFlag = True Then
                   'MsgBox ("Record has been location!")
                Else
                    MsgBox ("No Match in Database!")
                    foundFlag = False
                    .MoveFirst
                End If
        Else
            MsgBox ("There are no records To search!")
            foundFlag = False
        End If
    End With
    Search = foundFlag
'to search
'Call Search(List1.Text, rs, rs.Fields("one"))
End Function





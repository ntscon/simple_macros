Attribute VB_Name = "Module1"
Function leftmost(theseCells As Range) As String
Dim thisLength As Integer
thisLength = theseCells.Count
Dim i As Integer
Dim correct As String
Dim cell As Range

For i = 1 To thisLength
    
    If theseCells.Item(i).Value2 <> "" Then
        correct = theseCells.Item(i).Value2
        Exit For
    End If
Next i

leftmost = correct

End Function
Function rightmost(theseCells As Range) As String
Dim thisLength As Integer
thisLength = theseCells.Count
Dim i As Integer
Dim correct As String
Dim cell As Range

For i = thisLength To 1 Step -1
    
    If theseCells.Item(i).Value2 <> "" Then
        correct = theseCells.Item(i).Value2
        Exit For
    End If
Next i

rightmost = correct

End Function


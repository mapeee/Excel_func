Option Explicit

Function GetNumeric(CellRef As String) As Long
Dim StringLength As Integer
Dim Result As Integer
Dim i As Integer

StringLength = Len(CellRef)
For i = 1 To StringLength
If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
Next i
GetNumeric = Result
End Function

Function GetText(CellRef As String) As String
Dim StringLength As Integer
Dim Result_t As String
Dim i As Integer

StringLength = Len(CellRef)
For i = 1 To StringLength
If Not IsNumeric(Mid(CellRef, i, 1)) Then Result_t = Result_t & Mid(CellRef, i, 1)
Next i
GetText = Result_t
End Function
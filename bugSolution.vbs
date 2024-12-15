Option Explicit

' Function to demonstrate safe comparison
Function SafeCompare(value1, value2)
  If IsNumeric(value1) And IsNumeric(value2) Then
    SafeCompare = CInt(value1) = CInt(value2) 'Explicit conversion and comparison
  ElseIf VarType(value1) = VarType(value2) Then
    SafeCompare = value1 = value2 'Compare only if types are the same
  Else
    SafeCompare = False 'Handle type mismatch appropriately
  End If
End Function

' Example usage
Dim a, b
a = "10"
b = 10

If SafeCompare(a, b) Then
  MsgBox "Values are equal (after safe comparison)"
Else
  MsgBox "Values are not equal"
End If

'Demonstrate early binding for better performance and error detection
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ... use objFSO ... 
Set objFSO = Nothing
Function GetObjectSafe(object_name)
  Dim obj
  On Error Resume Next
  Set obj = GetObject(object_name)
  If Err.Number <> 0 Then
    Err.Clear
    MsgBox "Error: Could not get object '" & object_name & "'. Error code: " & Err.Number, vbExclamation
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

Sub TestGetObjectSafe()
  Dim obj As Object
  Set obj = GetObjectSafe("C:\\Windows\\System32\\notepad.exe")
  If obj Is Nothing Then
    MsgBox "Failed to get object.", vbCritical
  Else
    MsgBox "Object obtained successfully."
    Set obj = Nothing
  End If
End Sub

TestGetObjectSafe()
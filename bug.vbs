Function GetObject(object_name)
  On Error Resume Next
  Set GetObject = GetObject(object_name)
  If Err.Number <> 0 Then
    Err.Clear
    Set GetObject = Nothing
  End If
End Function

Sub TestGetObject()
  Dim obj As Object
  Set obj = GetObject("C:\\Windows\\System32\\notepad.exe")
  If obj Is Nothing Then
    MsgBox "Failed to get object.", vbCritical
  Else
    MsgBox "Object obtained successfully."
    Set obj = Nothing
  End If
End Sub

TestGetObject()
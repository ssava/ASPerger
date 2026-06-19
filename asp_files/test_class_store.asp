<%
Class TestStore
  Public Function GetEntries()
    Dim data
    data = Application.Contents("entries")
    If IsEmpty(data) Then
      Set GetEntries = CreateObject("Scripting.Dictionary")
    Else
      Set GetEntries = data
    End If
  End Function

  Public Function AddEntry(name)
    Dim data, entries, entry, id
    data = Application.Contents("entries")
    If IsEmpty(data) Then
      Set entries = CreateObject("Scripting.Dictionary")
    Else
      Set entries = data
    End If
    id = 1
    Dim keysArr
    keysArr = entries.Keys
    Dim k
    For Each k In keysArr
      If CInt(k) >= id Then id = CInt(k) + 1
    Next
    Set entry = CreateObject("Scripting.Dictionary")
    entry("id") = id
    entry("name") = name
    Call entries.Add(CStr(id), entry)
    Application("entries") = entries
    AddEntry = id
  End Function
End Class

Dim store, eid
Set store = New TestStore
eid = store.AddEntry("hello")
Response.Write "Added: " & eid & "<br>"

Dim entries2, k, html
Set entries2 = store.GetEntries()
html = ""
For Each k In entries2.Keys
  html = html & "key=" & k & " "
Next
Response.Write "Keys: " & html
%>
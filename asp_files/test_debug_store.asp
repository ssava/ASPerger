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

Dim store, i
Set store = New TestStore
i = store.AddEntry("hello")
Response.Write "Add1=" & i & " "

i = store.AddEntry("world")
Response.Write "Add2=" & i & " "

Dim e, k, out
Set e = store.GetEntries()
out = ""
For Each k In e.Keys
  out = out & "[" & k & ":"
  Dim v
  v = e(k)
  out = out & v("name")
  out = out & "]"
Next
Response.Write "Entries=" & out
%>
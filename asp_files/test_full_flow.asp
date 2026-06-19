<%
Dim data, entries
data = Application.Contents("entries")
If IsEmpty(data) Then
  Set entries = CreateObject("Scripting.Dictionary")
Else
  Set entries = data
End If
Dim entry
Set entry = CreateObject("Scripting.Dictionary")
entry("id") = 1
entry("project") = "Test"
Call entries.Add("1", entry)
Application("entries") = entries

Dim entries2
Set entries2 = CreateObject("Scripting.Dictionary")
Dim d
d = Application.Contents("entries")
If Not IsEmpty(d) Then
  Set entries2 = d
End If
Dim k, html
html = ""
For Each k In entries2.Keys
  html = html & "key:" & k & ";val:" & entries2(k)("project") & "|"
Next
Response.Write html
%>
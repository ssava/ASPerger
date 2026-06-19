<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
Dim e
Set e = CreateObject("Scripting.Dictionary")
e("id") = 1
e("name") = "test"
Call d.Add("1", e)
Application("test") = d

Dim rd, k, out
rd = Application.Contents("test")
out = ""
For Each k In rd.Keys
  out = out & "K:" & k & " "
Next
Response.Write out
%>
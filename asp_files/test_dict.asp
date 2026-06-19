<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
d("b") = 2

Dim k
For Each k In d.Keys
  Response.Write "Key: " & k & " = " & d(k) & "<br>"
Next

Response.Write "--- Count: " & d.Count & " ---<br>"

Application("test") = d

Dim d2
d2 = Application.Contents("test")
Response.Write "Keys from stored: "
Dim keys
keys = d2.Keys
For Each k In keys
  Response.Write "[" & k & "=" & d2(k) & "] "
Next
Response.Write "<br>"
%>
<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
Dim k, html
html = ""
For Each k In d.Keys
  html = html & "item:" & d(k) & ";"
Next
Response.Write html
%>
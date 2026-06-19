<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
d("b") = 2
s = "Count: " & d.Count
Response.Write "done"
%>
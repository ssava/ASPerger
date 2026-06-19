<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
d("b") = 2
Dim s
s = "Count: " & d.Count
Response.Write s
%>
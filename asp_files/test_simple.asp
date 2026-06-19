<%
Response.Write "A<br>"
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("x") = 99
Response.Write "B: " & d.Count & "<br>"
Response.Write "C<br>"
%>
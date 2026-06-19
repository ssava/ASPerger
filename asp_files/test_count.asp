<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
d("b") = 2
Response.Write d.Count & "<br>"
%>
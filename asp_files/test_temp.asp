<%
Dim d
Set d = CreateObject("Scripting.Dictionary")
d("a") = 1
d("b") = 2
Dim c
c = d.Count
Response.Write "Count: " & c & "<br>"
%>
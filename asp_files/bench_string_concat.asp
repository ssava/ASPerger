<%
Dim i, s
s = ""
For i = 1 To 5000
    s = s & "x"
Next
%>
Len: <%= Len(s) %>

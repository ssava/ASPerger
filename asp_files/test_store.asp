<%
Dim data
data = Application.Contents("entries")
If IsEmpty(data) Then
  Response.Write "Empty<br>"
Else
  Response.Write "Found: keys=" & data.Keys.Count & "<br>"
End If
%>
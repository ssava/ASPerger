<%
Dim pageTitle
pageTitle = "Arrays Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Array Functions Demo</h2>

<h3>Creating Arrays with Array()</h3>
<%
Dim colors, numbers
colors = Array("Red", "Green", "Blue", "Yellow", "Purple")
numbers = Array(10, 20, 30, 40, 50)
%>
<table>
    <tr><th>Function</th><th>Result</th></tr>
    <tr><td>Array("Red", "Green", ...)</td><td><%= Join(colors, ", ") %></td></tr>
    <tr><td>Array(10, 20, 30, 40, 50)</td><td><%= Join(numbers, ", ") %></td></tr>
</table>

<h3>LBound and UBound</h3>
<table>
    <tr><th>Function</th><th>Value</th></tr>
    <tr><td>LBound(colors)</td><td><%= LBound(colors) %></td></tr>
    <tr><td>UBound(colors)</td><td><%= UBound(colors) %></td></tr>
    <tr><td>LBound(numbers)</td><td><%= LBound(numbers) %></td></tr>
    <tr><td>UBound(numbers)</td><td><%= UBound(numbers) %></td></tr>
</table>

<h3>Iterating an Array with For i</h3>
<ul>
<%
Dim i
For i = 0 To UBound(colors)
    Response.Write("<li>colors(" & i & ") = " & colors(i) & "</li>")
Next
%>
</ul>

<h3>Iterating an Array with For Each</h3>
<ul>
<%
Dim c
For Each c In colors
    Response.Write("<li>" & c & "</li>")
Next
%>
</ul>

<h3>IsArray</h3>
<table>
    <tr><th>Expression</th><th>Result</th></tr>
    <tr><td>IsArray(colors)</td><td><%= IsArray(colors) %></td></tr>
    <tr><td>IsArray("hello")</td><td><%= IsArray("hello") %></td></tr>
    <tr><td>IsArray(42)</td><td><%= IsArray(42) %></td></tr>
    <tr><td>IsArray(Nothing)</td><td><%= IsArray(Nothing) %></td></tr>
</table>

<h3>Multi-Dimensional Array</h3>
<%
Dim matrix(2, 2)
matrix(0, 0) = 1
matrix(0, 1) = 2
matrix(0, 2) = 3
matrix(1, 0) = 4
matrix(1, 1) = 5
matrix(1, 2) = 6
matrix(2, 0) = 7
matrix(2, 1) = 8
matrix(2, 2) = 9
%>
<table>
<%
Dim r, c
For r = 0 To 2
    Response.Write("<tr>")
    For c = 0 To 2
        Response.Write("<td>" & matrix(r, c) & "</td>")
    Next
    Response.Write("</tr>")
Next
%>
</table>

<h3>Dynamic Arrays (ReDim)</h3>
<%
Dim dynArr()
ReDim dynArr(3)
dynArr(0) = "Item A"
dynArr(1) = "Item B"
dynArr(2) = "Item C"
dynArr(3) = "Item D"
Response.Write("<p>ReDim(3): <strong>" & Join(dynArr, ", ") & "</strong></p>")

ReDim Preserve dynArr(5)
dynArr(4) = "Item E"
dynArr(5) = "Item F"
Response.Write("<p>ReDim Preserve(5): <strong>" & Join(dynArr, ", ") & "</strong></p>")

ReDim dynArr(1)
dynArr(0) = "Reset"
dynArr(1) = "Done"
Response.Write("<p>ReDim(1) without Preserve: <strong>" & Join(dynArr, ", ") & "</strong></p>")
%>

<!--#include file="include/footer.asp"-->

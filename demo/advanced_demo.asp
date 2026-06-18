<%
Dim pageTitle
pageTitle = "Advanced Features Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Advanced ASP Classic Features</h2>

<h3>Select Case</h3>
<%
Dim dayNum, dayName
dayNum = Weekday(Now())
Select Case dayNum
    Case 1
        dayName = "Sunday"
    Case 2
        dayName = "Monday"
    Case 3
        dayName = "Tuesday"
    Case 4
        dayName = "Wednesday"
    Case 5
        dayName = "Thursday"
    Case 6
        dayName = "Friday"
    Case 7
        dayName = "Saturday"
    Case Else
        dayName = "Unknown"
End Select
%>
<p>Today is <strong><%= dayName %></strong> (Weekday <%= dayNum %>).</p>

<h4>Select Case with ranges and multiple values</h4>
<%
Dim score, grade
Randomize
score = Int(Rnd * 101)
Select Case score
    Case 90 To 100
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70 To 79
        grade = "C"
    Case 60 To 69
        grade = "D"
    Case Is < 60
        grade = "F"
End Select
%>
<p>Score: <strong><%= score %></strong> &rarr; Grade: <strong><%= grade %></strong></p>

<h3>Constants (Const)</h3>
<%
Const PI = 3.14159
Const APP_NAME = "ASPerger Demo"
Dim radius
radius = 5
%>
<table>
    <tr><td><strong>APP_NAME</strong></td><td><%= APP_NAME %></td></tr>
    <tr><td><strong>PI</strong></td><td><%= PI %></td></tr>
    <tr><td>Circle area (r=<%= radius %>)</td><td><%= PI * radius * radius %></td></tr>
</table>

<h3>Sub Procedures</h3>
<%
Sub DrawLine(count, symbol)
    Dim i
    For i = 1 To count
        Response.Write(symbol)
    Next
    Response.Write("<br>")
End Sub

Sub WriteBox(text)
    Response.Write("<div style=""border:1px solid #999; padding:8px; margin:8px 0; background:#fafafa;"">")
    Response.Write(text)
    Response.Write("</div>")
End Sub
%>
<p>
<%
DrawLine 20, "-"
DrawLine 10, "*="
Call WriteBox("This text is rendered by a Sub procedure.")
%>
</p>

<h3>Type Checking</h3>
<%
Dim vNull, vEmpty, vNum, vStr, vBool
vNum = 42
vStr = "hello"
vBool = True
%>
<table>
    <tr><th>Expression</th><th>IsNull</th><th>IsEmpty</th><th>IsNumeric</th><th>IsDate</th><th>VarType</th><th>TypeName</th></tr>
    <tr>
        <td>vNull</td>
        <td><%= IsNull(vNull) %></td>
        <td><%= IsEmpty(vNull) %></td>
        <td><%= IsNumeric(vNull) %></td>
        <td><%= IsDate(vNull) %></td>
        <td><%= VarType(vNull) %></td>
        <td><%= TypeName(vNull) %></td>
    </tr>
    <tr>
        <td>vEmpty</td>
        <td><%= IsNull(vEmpty) %></td>
        <td><%= IsEmpty(vEmpty) %></td>
        <td><%= IsNumeric(vEmpty) %></td>
        <td><%= IsDate(vEmpty) %></td>
        <td><%= VarType(vEmpty) %></td>
        <td><%= TypeName(vEmpty) %></td>
    </tr>
    <tr>
        <td>42</td>
        <td><%= IsNull(vNum) %></td>
        <td><%= IsEmpty(vNum) %></td>
        <td><%= IsNumeric(vNum) %></td>
        <td><%= IsDate(vNum) %></td>
        <td><%= VarType(vNum) %></td>
        <td><%= TypeName(vNum) %></td>
    </tr>
    <tr>
        <td>"hello"</td>
        <td><%= IsNull(vStr) %></td>
        <td><%= IsEmpty(vStr) %></td>
        <td><%= IsNumeric(vStr) %></td>
        <td><%= IsDate(vStr) %></td>
        <td><%= VarType(vStr) %></td>
        <td><%= TypeName(vStr) %></td>
    </tr>
    <tr>
        <td>True</td>
        <td><%= IsNull(vBool) %></td>
        <td><%= IsEmpty(vBool) %></td>
        <td><%= IsNumeric(vBool) %></td>
        <td><%= IsDate(vBool) %></td>
        <td><%= VarType(vBool) %></td>
        <td><%= TypeName(vBool) %></td>
    </tr>
</table>

<h3>Dynamic Arrays (ReDim, ReDim Preserve)</h3>
<%
Dim dynArr()
ReDim dynArr(2)
dynArr(0) = "First"
dynArr(1) = "Second"
dynArr(2) = "Third"
Response.Write("<p>Initial (3 items): " & Join(dynArr, ", ") & "</p>")

ReDim Preserve dynArr(4)
dynArr(3) = "Fourth"
dynArr(4) = "Fifth"
Response.Write("<p>After Preserve (5 items): " & Join(dynArr, ", ") & "</p>")

ReDim dynArr(1)
dynArr(0) = "Reset"
dynArr(1) = "Lost"
Response.Write("<p>After ReDim without Preserve (2 items, data lost): " & Join(dynArr, ", ") & "</p>")
%>

<h3>Multi-Dimensional Array</h3>
<%
Dim grid(2, 2)
grid(0, 0) = "A1"
grid(0, 1) = "A2"
grid(0, 2) = "A3"
grid(1, 0) = "B1"
grid(1, 1) = "B2"
grid(1, 2) = "B3"
grid(2, 0) = "C1"
grid(2, 1) = "C2"
grid(2, 2) = "C3"
%>
<table>
<%
Dim r, c
For r = 0 To 2
    Response.Write("<tr>")
    For c = 0 To 2
        Response.Write("<td>" & grid(r, c) & "</td>")
    Next
    Response.Write("</tr>")
Next
%>
</table>

<h3>Scripting.Dictionary</h3>
<%
Dim dict, key
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "name", "ASPerger"
dict.Add "version", "1.0"
dict.Add "language", "VBScript"
%>
<table>
    <tr><th>Key</th><th>Value</th></tr>
<%
Dim keys, k
keys = dict.Keys
For Each k In keys
    Response.Write("<tr><td>" & k & "</td><td>" & dict(k) & "</td></tr>")
Next
%>
</table>
<p>Count: <%= dict.Count %> &nbsp;|&nbsp; Exists("version"): <%= dict.Exists("version") %> &nbsp;|&nbsp; Item("name"): <%= dict.Item("name") %></p>
<%
dict.Remove("version")
%>
<p>After Remove("version"), Count: <%= dict.Count %></p>
<%
Set dict = Nothing
%>

<h3>Error Handling</h3>
<%
On Error Resume Next
Dim errResult
errResult = 10 / 0
If Err.Number <> 0 Then
    Response.Write("<p style=""color:#c00;"">Caught error: <strong>#" & Err.Number & "</strong> &mdash; " & Server.HTMLEncode(Err.Description) & "</p>")
    Err.Clear
End If
On Error GoTo 0
Response.Write("<p>Execution continued after the error.</p>")
%>

<h3>Random Numbers</h3>
<%
Dim i, roll
Randomize
%>
<p>Five dice rolls:
<%
For i = 1 To 5
    roll = Int(Rnd * 6) + 1
    Response.Write(roll & " ")
Next
%>
</p>

<!--#include file="include/footer.asp"-->

<%
Dim pageTitle
pageTitle = "About - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>About This Demo</h2>

<p>This application demonstrates various ASP Classic language features supported by <strong>ASPerger</strong>.</p>

<h3>What's Tested Here</h3>
<%
Dim features
features = Array("Variables (Dim, assignment, scoping)", "Conditionals (If/Then/ElseIf/Else, Select Case with ranges)", "Loops (For/Next, For Each, While/Wend, Do/Loop variants, Exit)", "Sub procedures and user-defined Functions", "Constants (Const)", "String functions (Split, Join, Replace, Mid, InStr, UCase, LCase, Trim, Len, Asc, Chr, String, Space, StrComp)", "Date/time functions (Now, DateAdd, DateDiff, FormatDateTime, etc.)", "Array operations (ReDim, ReDim Preserve, multi-dimensional, Filter)", "Scripting.Dictionary", "Error handling (On Error Resume Next, Err object)", "Type checking (IsNull, IsEmpty, VarType, TypeName)", "Random numbers (Randomize, Rnd)", "Include directives", "Request object (Form, QueryString, ServerVariables) + Server.HTMLEncode", "Response.Write, HTML mixed with ASP code blocks")
%>
<ul>
    <%
    For i = 0 To UBound(features)
        Response.Write("<li>" & features(i) & "</li>")
    Next
    %>
</ul>

<h3>Function Demo: Word of the Day</h3>
<%
Function GetGreeting(hour)
    If hour < 12 Then
        GetGreeting = "Good morning"
    ElseIf hour < 18 Then
        GetGreeting = "Good afternoon"
    Else
        GetGreeting = "Good evening"
    End If
End Function

Dim currentHour, greeting
currentHour = Hour(Now())
greeting = GetGreeting(currentHour)
%>
<p><strong><%= greeting %></strong> &mdash; it's <%= currentHour %>:00 here.</p>

<!--#include file="include/footer.asp"-->

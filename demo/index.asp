<%
Dim pageTitle
pageTitle = "Home - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Welcome to the ASPerger Demo App</h2>

<p>This is a simple ASP Classic application to test <strong>ASPerger</strong>, an ASP Classic interpreter and debugger.</p>

<h3>Feature Demo Pages</h3>
<ul>
    <li><a href="date_demo.asp">Date &amp; Time Functions</a> &mdash; Now(), Date(), Time(), DateAdd(), DateDiff()</li>
    <li><a href="form_demo.asp">Form Processing</a> &mdash; Request.Form and query string handling</li>
    <li><a href="loops_demo.asp">Loop Variations</a> &mdash; For Each, While/Wend, Do/Loop, Exit, nesting</li>
    <li><a href="strings_demo.asp">String Functions</a> &mdash; Split, Join, Replace, Mid, InStr, UCase, Trim, etc.</li>
    <li><a href="advanced_demo.asp">Advanced Features</a> &mdash; Select Case, Sub, Const, ReDim, Dictionary, error handling, Rnd</li>
</ul>

<h3>Server Info</h3>
<%
Dim scriptName, serverName, httpMethod, queryString
scriptName = Request.ServerVariables("SCRIPT_NAME")
serverName = Request.ServerVariables("SERVER_NAME")
httpMethod = Request.ServerVariables("REQUEST_METHOD")
queryString = Request.ServerVariables("QUERY_STRING")
%>
<table>
    <tr><td><strong>Script</strong></td><td><%= scriptName %></td></tr>
    <tr><td><strong>Server</strong></td><td><%= serverName %></td></tr>
    <tr><td><strong>Method</strong></td><td><%= httpMethod %></td></tr>
    <tr><td><strong>Query String</strong></td><td><%= queryString %></td></tr>
</table>

<h3>Loop Demo: Multiplication Table</h3>
<table>
    <%
    Response.Write("<tr><th>&times;</th>")
    For col = 1 To 5
        Response.Write("<th>" & col & "</th>")
    Next
    Response.Write("</tr>")
    For row = 1 To 5
        Response.Write("<tr><th>" & row & "</th>")
        For col = 1 To 5
            Response.Write("<td>" & row * col & "</td>")
        Next
        Response.Write("</tr>")
    Next
    %>
</table>

<!--#include file="include/footer.asp"-->

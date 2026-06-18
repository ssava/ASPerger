<%
Dim pageTitle
pageTitle = "String Functions Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>String Function Demos</h2>

<%
Dim sample, sentence, words, result, parts
sample = "  Hello, ASP Classic World!  "
sentence = "The quick brown fox jumps over the lazy dog"
%>

<h3>Basic Inspection</h3>
<table>
    <tr><td><strong>Len("" & sample & "")</strong></td><td><%= Len(sample) %></td></tr>
    <tr><td><strong>InStr(sample, "ASP")</strong></td><td><%= InStr(sample, "ASP") %></td></tr>
    <tr><td><strong>InStr(10, sample, "o")</strong></td><td><%= InStr(10, sample, "o") %></td></tr>
</table>

<h3>Case Conversion</h3>
<table>
    <tr><td><strong>UCase(sample)</strong></td><td><%= UCase(sample) %></td></tr>
    <tr><td><strong>LCase(sample)</strong></td><td><%= LCase(sample) %></td></tr>
</table>

<h3>Trimming</h3>
<table>
    <tr><td><strong>Trim(sample)</strong></td><td>"<%= Trim(sample) %>"</td></tr>
    <tr><td><strong>LTrim(sample)</strong></td><td>"<%= LTrim(sample) %>"</td></tr>
    <tr><td><strong>RTrim(sample)</strong></td><td>"<%= RTrim(sample) %>"</td></tr>
</table>

<h3>Extraction</h3>
<table>
    <tr><td><strong>Left(sample, 5)</strong></td><td>"<%= Left(sample, 5) %>"</td></tr>
    <tr><td><strong>Right(sample, 6)</strong></td><td>"<%= Right(sample, 6) %>"</td></tr>
    <tr><td><strong>Mid(sample, 7, 10)</strong></td><td>"<%= Mid(sample, 7, 10) %>"</td></tr>
</table>

<h3>Replace</h3>
<%
result = Replace(sample, "World", "Universe")
%>
<table>
    <tr><td><strong>Replace(sample, "World", "Universe")</strong></td><td>"<%= result %>"</td></tr>
</table>

<h3>Split and Join</h3>
<%
words = Split(sentence, " ")
parts = Array("one", "two", "three")
%>
<p><strong>Split(sentence, " ")</strong> &rarr; array of <%= UBound(words) + 1 %> words</p>
<ul>
<%
Dim w
For Each w In words
    Response.Write("<li>" & w & "</li>")
Next
%>
</ul>
<p><strong>Join(parts, ", ")</strong> &rarr; "<%= Join(parts, ", ") %>"</p>

<h3>String Comparison</h3>
<%
Dim a, b
a = "apple"
b = "APPLE"
%>
<table>
    <tr><td><strong>StrComp(a, b, vbBinaryCompare)</strong></td><td><%= StrComp(a, b, 0) %> (0=match)</td></tr>
    <tr><td><strong>StrComp(a, b, vbTextCompare)</strong></td><td><%= StrComp(a, b, 1) %> (0=match)</td></tr>
</table>

<h3>Asc / Chr</h3>
<table>
    <tr><td><strong>Asc("A")</strong></td><td><%= Asc("A") %></td></tr>
    <tr><td><strong>Chr(65)</strong></td><td>"<%= Chr(65) %>"</td></tr>
    <tr><td><strong>Chr(38) &amp; Chr(60) &amp; Chr(62)</strong></td><td><%= Chr(38) & Chr(60) & Chr(62) %></td></tr>
</table>

<h3>String(n, char) -- repeat character</h3>
<table>
    <tr><td><strong>String(5, "*")</strong></td><td><%= String(5, "*") %></td></tr>
    <tr><td><strong>String(3, "A")</strong></td><td><%= String(3, "A") %></td></tr>
</table>

<h3>Space</h3>
<table>
    <tr><td><strong>"Hello" &amp; Space(5) &amp; "World"</strong></td><td>"Hello<%= Space(5) %>World"</td></tr>
</table>

<!--#include file="include/footer.asp"-->

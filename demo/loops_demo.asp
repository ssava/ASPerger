<%
Dim pageTitle
pageTitle = "Loops Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Loop Variations</h2>

<h3>For Each ... Next</h3>
<%
Dim fruits, fruit
fruits = Array("Apple", "Banana", "Cherry", "Date", "Elderberry")
%>
<ul>
<%
For Each fruit In fruits
    Response.Write("<li>" & fruit & "</li>")
Next
%>
</ul>

<h3>While ... Wend</h3>
<%
Dim counter
counter = 1
%>
<table>
    <tr><th>Iteration</th><th>Square</th></tr>
<%
While counter <= 5
    Response.Write("<tr><td>" & counter & "</td><td>" & counter * counter & "</td></tr>")
    counter = counter + 1
Wend
%>
</table>

<h3>Do While ... Loop</h3>
<%
Dim dwCounter
dwCounter = 1
%>
<p>
<%
Do While dwCounter <= 5
    Response.Write(dwCounter & " ")
    dwCounter = dwCounter + 1
Loop
%>
</p>

<h3>Do Until ... Loop</h3>
<%
Dim duCounter
duCounter = 1
%>
<p>
<%
Do Until duCounter > 5
    Response.Write(duCounter & " ")
    duCounter = duCounter + 1
Loop
%>
</p>

<h3>Do ... Loop While (runs at least once)</h3>
<%
Dim dlwCounter
dlwCounter = 1
%>
<p>
<%
Do
    Response.Write(dlwCounter & " ")
    dlwCounter = dlwCounter + 1
Loop While dlwCounter <= 5
%>
</p>

<h3>Do ... Loop Until (runs at least once)</h3>
<%
Dim dluCounter
dluCounter = 1
%>
<p>
<%
Do
    Response.Write(dluCounter & " ")
    dluCounter = dluCounter + 1
Loop Until dluCounter > 5
%>
</p>

<h3>Exit For (stop at "Cherry")</h3>
<%
Dim berry
%>
<p>
<%
For Each berry In fruits
    If berry = "Cherry" Then Exit For
    Response.Write(berry & " ")
Next
%>
</p>

<h3>Exit Do (stop when square &gt; 10)</h3>
<%
Dim n
n = 1
%>
<p>
<%
Do While True
    If n * n > 10 Then Exit Do
    Response.Write(n & "&sup2;=" & n * n & " ")
    n = n + 1
Loop
%>
</p>

<h3>Nested Loops: Triangular Pattern</h3>
<%
Dim row, col2
%>
<pre style="font-size:18px; line-height:1.2;">
<%
For row = 1 To 5
    For col2 = 1 To row
        Response.Write("*")
    Next
    Response.Write("<br>")
Next
%>
</pre>

<!--#include file="include/footer.asp"-->

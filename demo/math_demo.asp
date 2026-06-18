<%
Dim pageTitle
pageTitle = "Math Functions Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Math &amp; Numeric Functions Demo</h2>

<h3>Rounding and Integer Functions</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Int</td><td>3.14159</td><td><%= Int(3.14159) %></td></tr>
    <tr><td>Int</td><td>-3.14159</td><td><%= Int(-3.14159) %></td></tr>
    <tr><td>Fix</td><td>3.14159</td><td><%= Fix(3.14159) %></td></tr>
    <tr><td>Fix</td><td>-3.14159</td><td><%= Fix(-3.14159) %></td></tr>
    <tr><td>Round</td><td>3.14159, 2</td><td><%= Round(3.14159, 2) %></td></tr>
    <tr><td>Round</td><td>3.14159, 4</td><td><%= Round(3.14159, 4) %></td></tr>
    <tr><td>Round</td><td>2.5</td><td><%= Round(2.5) %></td></tr>
    <tr><td>Round</td><td>3.5</td><td><%= Round(3.5) %></td></tr>
</table>

<h3>Absolute Value</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Abs</td><td>42</td><td><%= Abs(42) %></td></tr>
    <tr><td>Abs</td><td>-42</td><td><%= Abs(-42) %></td></tr>
    <tr><td>Abs</td><td>0</td><td><%= Abs(0) %></td></tr>
</table>

<h3>Square Root</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Sqr</td><td>9</td><td><%= Sqr(9) %></td></tr>
    <tr><td>Sqr</td><td>2</td><td><%= Sqr(2) %></td></tr>
    <tr><td>Sqr</td><td>100</td><td><%= Sqr(100) %></td></tr>
</table>

<h3>Logarithms and Exponents</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Log</td><td>1</td><td><%= Log(1) %></td></tr>
    <tr><td>Log</td><td>10</td><td><%= Log(10) %></td></tr>
    <tr><td>Log</td><td>100</td><td><%= Log(100) %></td></tr>
    <tr><td>Exp</td><td>1</td><td><%= Exp(1) %></td></tr>
    <tr><td>Exp</td><td>2</td><td><%= Exp(2) %></td></tr>
    <tr><td>Exp</td><td>0</td><td><%= Exp(0) %></td></tr>
</table>

<h3>Trigonometric Functions</h3>
<table>
    <tr><th>Function</th><th>Input (radians)</th><th>Result</th></tr>
    <tr><td>Sin</td><td>0</td><td><%= Sin(0) %></td></tr>
    <tr><td>Sin</td><td>1.5708</td><td><%= Sin(1.5708) %></td></tr>
    <tr><td>Cos</td><td>0</td><td><%= Cos(0) %></td></tr>
    <tr><td>Cos</td><td>3.14159</td><td><%= Cos(3.14159) %></td></tr>
    <tr><td>Tan</td><td>0</td><td><%= Tan(0) %></td></tr>
    <tr><td>Tan</td><td>0.7854</td><td><%= Tan(0.7854) %></td></tr>
    <tr><td>Atn</td><td>0</td><td><%= Atn(0) %></td></tr>
    <tr><td>Atn</td><td>1</td><td><%= Atn(1) %></td></tr>
</table>

<h3>Sign</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Sgn</td><td>42</td><td><%= Sgn(42) %></td></tr>
    <tr><td>Sgn</td><td>0</td><td><%= Sgn(0) %></td></tr>
    <tr><td>Sgn</td><td>-42</td><td><%= Sgn(-42) %></td></tr>
</table>

<h3>Random Numbers</h3>
<%
Randomize
%>
<table>
    <tr><th>Function</th><th>Result</th></tr>
    <tr><td>Rnd</td><td><%= Rnd %></td></tr>
    <tr><td>Rnd</td><td><%= Rnd %></td></tr>
    <tr><td>Int(Rnd * 10) + 1</td><td><%= Int(Rnd * 10) + 1 %></td></tr>
    <tr><td>Int(Rnd * 10) + 1</td><td><%= Int(Rnd * 10) + 1 %></td></tr>
    <tr><td>Int(Rnd * 6) + 1</td><td><%= Int(Rnd * 6) + 1 %></td></tr>
    <tr><td>Int(Rnd * 6) + 1</td><td><%= Int(Rnd * 6) + 1 %></td></tr>
</table>

<h3>Number Formatting</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th></tr>
    <tr><td>Hex</td><td>255</td><td><%= Hex(255) %></td></tr>
    <tr><td>Hex</td><td>16</td><td><%= Hex(16) %></td></tr>
    <tr><td>Hex</td><td>4095</td><td><%= Hex(4095) %></td></tr>
    <tr><td>Oct</td><td>8</td><td><%= Oct(8) %></td></tr>
    <tr><td>Oct</td><td>255</td><td><%= Oct(255) %></td></tr>
    <tr><td>Oct</td><td>64</td><td><%= Oct(64) %></td></tr>
</table>

<h3>Misc Math Operators</h3>
<table>
    <tr><th>Expression</th><th>Result</th></tr>
    <tr><td>10 \ 3 (integer division)</td><td><%= 10 \ 3 %></td></tr>
    <tr><td>10 Mod 3 (modulo)</td><td><%= 10 Mod 3 %></td></tr>
    <tr><td>2 ^ 10 (exponentiation)</td><td><%= 2 ^ 10 %></td></tr>
    <tr><td>5 ^ 3</td><td><%= 5 ^ 3 %></td></tr>
</table>

<!--#include file="include/footer.asp"-->

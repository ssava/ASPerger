<%
Dim pageTitle
pageTitle = "Type Conversion &amp; Checking Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Type Conversion &amp; Type Checking Demo</h2>

<h3>Explicit Type Conversions</h3>
<table>
    <tr><th>Function</th><th>Input</th><th>Result</th><th>TypeName</th></tr>
    <tr>
        <td>CBool</td>
        <td>1</td>
        <td><%= CBool(1) %></td>
        <td><%= TypeName(CBool(1)) %></td>
    </tr>
    <tr>
        <td>CBool</td>
        <td>0</td>
        <td><%= CBool(0) %></td>
        <td><%= TypeName(CBool(0)) %></td>
    </tr>
    <tr>
        <td>CByte</td>
        <td>255</td>
        <td><%= CByte(255) %></td>
        <td><%= TypeName(CByte(255)) %></td>
    </tr>
    <tr>
        <td>CInt</td>
        <td>3.14159</td>
        <td><%= CInt(3.14159) %></td>
        <td><%= TypeName(CInt(3.14159)) %></td>
    </tr>
    <tr>
        <td>CLng</td>
        <td>2147483647</td>
        <td><%= CLng(2147483647) %></td>
        <td><%= TypeName(CLng(2147483647)) %></td>
    </tr>
    <tr>
        <td>CSng</td>
        <td>3.1415926535</td>
        <td><%= CSng(3.1415926535) %></td>
        <td><%= TypeName(CSng(3.1415926535)) %></td>
    </tr>
    <tr>
        <td>CDbl</td>
        <td>3.1415926535</td>
        <td><%= CDbl(3.1415926535) %></td>
        <td><%= TypeName(CDbl(3.1415926535)) %></td>
    </tr>
    <tr>
        <td>CCur</td>
        <td>123.456</td>
        <td><%= CCur(123.456) %></td>
        <td><%= TypeName(CCur(123.456)) %></td>
    </tr>
    <tr>
        <td>CStr</td>
        <td>42</td>
        <td><%= CStr(42) %></td>
        <td><%= TypeName(CStr(42)) %></td>
    </tr>
    <tr>
        <td>CStr</td>
        <td>True</td>
        <td><%= CStr(True) %></td>
        <td><%= TypeName(CStr(True)) %></td>
    </tr>
    <tr>
        <td>CDate</td>
        <td>"2024-07-04"</td>
        <td><%= CDate("2024-07-04") %></td>
        <td><%= TypeName(CDate("2024-07-04")) %></td>
    </tr>
    <tr>
        <td>CDate</td>
        <td>"12:30:00"</td>
        <td><%= CDate("12:30:00") %></td>
        <td><%= TypeName(CDate("12:30:00")) %></td>
    </tr>
</table>

<h3>Type Checking Functions</h3>
<%
Dim vEmpty, vNullVal, vNum, vStr, vDate, vBool, vArr, vObj
vNum = 42
vStr = "Hello"
vDate = Now()
vBool = False
vArr = Array(1, 2, 3)
Set vObj = Server.CreateObject("Scripting.Dictionary")
%>
<table>
    <tr><th>Variable</th><th>IsArray</th><th>IsDate</th><th>IsEmpty</th><th>IsNull</th><th>IsNumeric</th><th>IsObject</th><th>VarType</th><th>TypeName</th></tr>
    <tr>
        <td>vEmpty</td>
        <td><%= IsArray(vEmpty) %></td>
        <td><%= IsDate(vEmpty) %></td>
        <td><%= IsEmpty(vEmpty) %></td>
        <td><%= IsNull(vEmpty) %></td>
        <td><%= IsNumeric(vEmpty) %></td>
        <td><%= IsObject(vEmpty) %></td>
        <td><%= VarType(vEmpty) %></td>
        <td><%= TypeName(vEmpty) %></td>
    </tr>
    <tr>
        <td>vNullVal</td>
        <td><%= IsArray(vNullVal) %></td>
        <td><%= IsDate(vNullVal) %></td>
        <td><%= IsEmpty(vNullVal) %></td>
        <td><%= IsNull(vNullVal) %></td>
        <td><%= IsNumeric(vNullVal) %></td>
        <td><%= IsObject(vNullVal) %></td>
        <td><%= VarType(vNullVal) %></td>
        <td><%= TypeName(vNullVal) %></td>
    </tr>
    <tr>
        <td>vNum (42)</td>
        <td><%= IsArray(vNum) %></td>
        <td><%= IsDate(vNum) %></td>
        <td><%= IsEmpty(vNum) %></td>
        <td><%= IsNull(vNum) %></td>
        <td><%= IsNumeric(vNum) %></td>
        <td><%= IsObject(vNum) %></td>
        <td><%= VarType(vNum) %></td>
        <td><%= TypeName(vNum) %></td>
    </tr>
    <tr>
        <td>vStr ("Hello")</td>
        <td><%= IsArray(vStr) %></td>
        <td><%= IsDate(vStr) %></td>
        <td><%= IsEmpty(vStr) %></td>
        <td><%= IsNull(vStr) %></td>
        <td><%= IsNumeric(vStr) %></td>
        <td><%= IsObject(vStr) %></td>
        <td><%= VarType(vStr) %></td>
        <td><%= TypeName(vStr) %></td>
    </tr>
    <tr>
        <td>vDate (Now)</td>
        <td><%= IsArray(vDate) %></td>
        <td><%= IsDate(vDate) %></td>
        <td><%= IsEmpty(vDate) %></td>
        <td><%= IsNull(vDate) %></td>
        <td><%= IsNumeric(vDate) %></td>
        <td><%= IsObject(vDate) %></td>
        <td><%= VarType(vDate) %></td>
        <td><%= TypeName(vDate) %></td>
    </tr>
    <tr>
        <td>vBool (False)</td>
        <td><%= IsArray(vBool) %></td>
        <td><%= IsDate(vBool) %></td>
        <td><%= IsEmpty(vBool) %></td>
        <td><%= IsNull(vBool) %></td>
        <td><%= IsNumeric(vBool) %></td>
        <td><%= IsObject(vBool) %></td>
        <td><%= VarType(vBool) %></td>
        <td><%= TypeName(vBool) %></td>
    </tr>
    <tr>
        <td>vArr</td>
        <td><%= IsArray(vArr) %></td>
        <td><%= IsDate(vArr) %></td>
        <td><%= IsEmpty(vArr) %></td>
        <td><%= IsNull(vArr) %></td>
        <td><%= IsNumeric(vArr) %></td>
        <td><%= IsObject(vArr) %></td>
        <td><%= VarType(vArr) %></td>
        <td><%= TypeName(vArr) %></td>
    </tr>
    <tr>
        <td>vObj (Dictionary)</td>
        <td><%= IsArray(vObj) %></td>
        <td><%= IsDate(vObj) %></td>
        <td><%= IsEmpty(vObj) %></td>
        <td><%= IsNull(vObj) %></td>
        <td><%= IsNumeric(vObj) %></td>
        <td><%= IsObject(vObj) %></td>
        <td><%= VarType(vObj) %></td>
        <td><%= TypeName(vObj) %></td>
    </tr>
</table>

<h3>CStr with Various Types</h3>
<table>
    <tr><th>Input</th><th>CStr Result</th></tr>
    <tr><td>42 (Number)</td><td>"<%= CStr(42) %>"</td></tr>
    <tr><td>True (Boolean)</td><td>"<%= CStr(True) %>"</td></tr>
    <tr><td>#2024-07-04# (Date)</td><td>"<%= CStr(#2024-07-04#) %>"</td></tr>
    <tr><td>Null</td><td>"<%= CStr(Null) %>"</td></tr>
    <tr><td>Empty</td><td>"<%= CStr(Empty) %>"</td></tr>
</table>

<!--#include file="include/footer.asp"-->

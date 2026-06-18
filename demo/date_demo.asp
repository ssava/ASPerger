<%
Dim pageTitle
pageTitle = "Date &amp; Time Demo - ASPerger Demo"
%>
<!--#include file="include/header.asp"-->

<h2>Date &amp; Time Functions</h2>

<p>Demonstrating VBScript date/time capabilities.</p>

<h3>Current Date/Time</h3>
<table>
    <tr><td><strong>Now()</strong></td><td><%= Now() %></td></tr>
    <tr><td><strong>Date()</strong></td><td><%= Date() %></td></tr>
    <tr><td><strong>Time()</strong></td><td><%= Time() %></td></tr>
</table>

<h3>Date Parts</h3>
<%
Dim dt
dt = Now()
%>
<table>
    <tr><td>Year</td><td><%= Year(dt) %></td></tr>
    <tr><td>Month</td><td><%= Month(dt) %> (<%= MonthName(Month(dt)) %>)</td></tr>
    <tr><td>Day</td><td><%= Day(dt) %></td></tr>
    <tr><td>Weekday</td><td><%= Weekday(dt) %> (<%= WeekdayName(Weekday(dt)) %>)</td></tr>
    <tr><td>Hour</td><td><%= Hour(dt) %></td></tr>
    <tr><td>Minute</td><td><%= Minute(dt) %></td></tr>
    <tr><td>Second</td><td><%= Second(dt) %></td></tr>
</table>

<h3>Date Arithmetic</h3>
<%
Dim future, past, diff
future = DateAdd("d", 7, dt)
past = DateAdd("m", -1, dt)
diff = DateDiff("d", dt, future)
%>
<table>
    <tr><td>+7 days</td><td><%= future %></td></tr>
    <tr><td>-1 month</td><td><%= past %></td></tr>
    <tr><td>Days between now and +7d</td><td><%= diff %></td></tr>
</table>

<h3>Formatting</h3>
<table>
    <tr><td>FormatDateTime (Short)</td><td><%= FormatDateTime(dt, 2) %></td></tr>
    <tr><td>FormatDateTime (Long)</td><td><%= FormatDateTime(dt, 1) %></td></tr>
    <tr><td>FormatDateTime (Time)</td><td><%= FormatDateTime(dt, 4) %></td></tr>
    <tr><td>IsDate check</td><td>"2026-12-25" is date: <%= IsDate("2026-12-25") %></td></tr>
    <tr><td>CDate conversion</td><td><%= CDate("2026-06-17") %></td></tr>
</table>

<!--#include file="include/footer.asp"-->

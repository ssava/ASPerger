<%
Dim i, a, b, c, d, e
a = 3.14159
b = 2.71828
c = 1.41421
d = 0.0
For i = 1 To 10000
    d = d + a * b / c
    e = (a + b) * (c - d) / (a * b + c)
    a = a + 0.001
    b = b - 0.001
Next
%>
Result: <%= d %>

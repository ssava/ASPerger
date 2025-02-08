<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>ASP Classic Test Page</title>
    <style>
        .test-section { margin: 20px; padding: 10px; border: 1px solid #ccc; }
        .test-title { color: #333; }
    </style>
</head>
<body>
    <h1>Test Suite ASP Classic</h1>

    <div class="test-section">
        <h3 class="test-title">1. Response.Write Base (senza parentesi)</h3>
        <%Response.Write "Test di output semplice"%>
    </div>

    <div class="test-section">
        <h3 class="test-title">2. Response.Write con parentesi</h3>
        <%Response.Write("Test con parentesi")%>
    </div>

    <div class="test-section">
        <h3 class="test-title">3. Test Variabili Stringa</h3>
        <%
        Dim strVar
        strVar = "Contenuto della variabile"
        Response.Write(strVar)
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">4. Test Variabili Numeriche</h3>
        <%
        Dim numVar
        numVar = 42
        Response.Write("Il numero è: ")
        Response.Write(numVar)
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">5. Test Condizioni (If-Then)</h3>
        <%
        If numVar > 40 Then
            Response.Write("Il numero è maggiore di 40")
        End If
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">6. Test HTML nei Response.Write</h3>
        <%Response.Write("<strong>Questo dovrebbe essere in grassetto</strong>")%>
    </div>

    <div class="test-section">
        <h3 class="test-title">7. Test Concatenazione Output</h3>
        <%
        Response.Write("Prima parte - ")
        Response.Write("Seconda parte")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">8. Test Commenti</h3>
        <%
        'Questo è un commento
        Response.Write("Testo dopo il commento")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">9. Test Multiple Variables</h3>
        <%
        Dim var1, var2
        var1 = "Prima variabile"
        var2 = "Seconda variabile"
        Response.Write(var1)
        Response.Write(" e ")
        Response.Write(var2)
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">10. Test Caratteri Speciali</h3>
        <%Response.Write("Test con caratteri speciali: è à ò ù")%>
    </div>
</body>
</html>
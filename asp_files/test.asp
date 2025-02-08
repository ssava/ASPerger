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

    <div class="test-section">
        <h3 class="test-title">11. Test Ciclo For</h3>
        <%
        For i = 1 To 5
            Response.Write("Iterazione: " & i & "<br>")
        Next
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">12. Test Ciclo While</h3>
        <%
        Dim counter
        counter = 1
        While counter <= 3
            Response.Write("Conteggio: " & counter & "<br>")
            counter = counter + 1
        Wend
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">13. Test Funzioni</h3>
        <%
        Function AddNumbers(a, b)
            AddNumbers = a + b
        End Function

        Dim result
        result = AddNumbers(5, 3)
        Response.Write("Risultato della funzione: " & result)
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">14. Test Operatori Logici (And, Or)</h3>
        <%
        Dim x, y
        x = True
        y = False

        If x And y Then
            Response.Write("Entrambi sono veri")
        ElseIf x Or y Then
            Response.Write("Almeno uno è vero")
        Else
            Response.Write("Nessuno è vero")
        End If
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">15. Test Chiamata a Funzione</h3>
        <%
        Sub SayHello(name)
            Response.Write("Ciao, " & name & "!")
        End Sub

        Call SayHello("Mondo")
        %>
    </div>
</body>
</html>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>ASP Classic Test Page</title>
    <style>
        .test-section { margin: 20px; padding: 10px; border: 1px solid #ccc; }
        .test-title { color: #333; }
        .success { color: green; }
        .error { color: red; }
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
 <div class="test-section">
        <h3 class="test-title">16. Test Array e ReDim</h3>
        <%
        Dim arr()
        ReDim arr(3)
        arr(0) = "Primo"
        arr(1) = "Secondo"
        arr(2) = "Terzo"
        For i = 0 To 2
            Response.Write(arr(i) & "<br>")
        Next
        
        ' Test ReDim Preserve
        ReDim Preserve arr(4)
        arr(3) = "Quarto"
        Response.Write("Dopo ReDim Preserve: " & arr(3))
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">17. Test Select Case</h3>
        <%
        Dim grade
        grade = "B"
        Select Case grade
            Case "A"
                Response.Write("Eccellente")
            Case "B"
                Response.Write("Buono")
            Case "C"
                Response.Write("Sufficiente")
            Case Else
                Response.Write("Non classificato")
        End Select
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">18. Test With Statement</h3>
        <%
        Class TestClass
            Public value
            Public name
        End Class
        
        Dim obj
        Set obj = New TestClass
        With obj
            .value = 42
            .name = "Test Object"
            Response.Write(.name & ": " & .value)
        End With
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">19. Test Property Get/Let</h3>
        <%
        Class PropertyTest
            Private m_value
            
            Public Property Get Value()
                Value = m_value
            End Property
            
            Public Property Let Value(v)
                m_value = v
            End Property
        End Class
        
        Dim propTest
        Set propTest = New PropertyTest
        propTest.Value = 100
        Response.Write("Property value: " & propTest.Value)
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">20. Test Error Handling</h3>
        <%
        On Error Resume Next
        Dim x
        x = 1 / 0  ' Dovrebbe generare un errore
        If Err.Number <> 0 Then
            Response.Write("Errore catturato: " & Err.Description)
        End If
        On Error Goto 0
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">21. Test Do...Loop Until</h3>
        <%
        Dim counter2
        counter2 = 0
        Do
            Response.Write("Contatore: " & counter2 & "<br>")
            counter2 = counter2 + 1
        Loop Until counter2 > 3
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">22. Test For Each</h3>
        <%
        Dim dict
        Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "a", "Alpha"
        dict.Add "b", "Beta"
        dict.Add "g", "Gamma"
        
        Dim key
        For Each key In dict.Keys
            Response.Write(key & ": " & dict(key) & "<br>")
        Next
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">23. Test Operatori di Confronto</h3>
        <%
        Dim val1, val2
        val1 = 10
        val2 = "10"
        If val1 = val2 Then Response.Write("Uguali (=)<br>")
        If val1 Is val2 Then Response.Write("Stesso oggetto (Is)<br>")
        If val1 >= val2 Then Response.Write("Maggiore o uguale (>=)<br>")
        If val1 <= val2 Then Response.Write("Minore o uguale (<=)<br>")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">24. Test Concatenazione Stringhe</h3>
        <%
        Dim str1, str2, str3
        str1 = "Hello"
        str2 = "World"
        str3 = str1 & " " & str2  ' Usando &
        Response.Write(str3 & "<br>")
        str3 = str1 + " " + str2  ' Usando +
        Response.Write(str3)
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">25. Test Operatori Mod e Integer Division</h3>
        <%
        Dim num1, num2
        num1 = 17
        num2 = 5
        Response.Write("Modulo (17 Mod 5): " & (num1 Mod num2) & "<br>")
        Response.Write("Divisione intera (17 \ 5): " & (num1 \ num2))
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">26. Test Empty, Null, Nothing</h3>
        <%
        Dim testVar, testObj
        Response.Write("Empty: " & IsEmpty(testVar) & "<br>")
        testVar = Null
        Response.Write("Null: " & IsNull(testVar) & "<br>")
        Set testObj = Nothing
        Response.Write("Nothing: " & (testObj Is Nothing))
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">27. Test Eqv e Imp Operators</h3>
        <%
        Dim bool1, bool2
        bool1 = True
        bool2 = False
        Response.Write("True Eqv False: " & (bool1 Eqv bool2) & "<br>")
        Response.Write("True Imp False: " & (bool1 Imp bool2))
        %>
    </div>
</body>
</html>

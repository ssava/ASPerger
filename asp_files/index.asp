<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>ASP Classic Test Page</title>
    <style>
        .test-section { margin: 20px; padding: 10px; border: 1px solid #ccc; }
        .test-title { color: #333; }
        .success { color: green; font-weight: bold; }
        .error { color: red; font-weight: bold; }
        .summary { font-size: 24px; padding: 20px; text-align: center; }
        .summary.pass { background: #d4edda; color: #155724; }
        .summary.fail { background: #f8d7da; color: #721c24; }
    </style>
</head>
<body>
    <h1>Test Suite ASP Classic</h1>
    <%
    Dim testTotal, testPassed
    testTotal = 0
    testPassed = 0
    %>

    <div class="test-section">
        <h3 class="test-title">1. Response.Write Base (senza parentesi)</h3>
        <%Response.Write "Test di output semplice"%>
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">2. Response.Write con parentesi</h3>
        <%Response.Write("Test con parentesi")%>
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">3. Test Variabili Stringa</h3>
        <%
        Dim strVar
        strVar = "Contenuto della variabile"
        Response.Write(strVar)
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">4. Test Variabili Numeriche</h3>
        <%
        Dim numVar
        numVar = 42
        Response.Write("Il numero è: ")
        Response.Write(numVar)
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">5. Test Condizioni (If-Then)</h3>
        <%
        If numVar > 40 Then
            Response.Write("Il numero è maggiore di 40")
        End If
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">6. Test HTML nei Response.Write</h3>
        <%Response.Write("<strong>Questo dovrebbe essere in grassetto</strong>")%>
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">7. Test Concatenazione Output</h3>
        <%
        Response.Write("Prima parte - ")
        Response.Write("Seconda parte")
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">8. Test Commenti</h3>
        <%
        'Questo è un commento
        Response.Write("Testo dopo il commento")
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
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
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">10. Test Caratteri Speciali</h3>
        <%Response.Write("Test con caratteri speciali: è à ò ù")%>
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">11. Test Ciclo For</h3>
        <%
        Dim forResult
        forResult = False
        For i = 1 To 5
            Response.Write("Iterazione: " & i & "<br>")
        Next
        If i = 6 Then forResult = True
        testTotal = testTotal + 1
        If forResult Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (i = " & i & ")</span>")
        End If
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
        testTotal = testTotal + 1
        If counter = 4 Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (counter = " & counter & ")</span>")
        End If
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
        <%
        testTotal = testTotal + 1
        If result = 8 Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (expected 8, got " & result & ")</span>")
        End If
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
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
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
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
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

        ReDim Preserve arr(4)
        arr(3) = "Quarto"
        Response.Write("Dopo ReDim Preserve: " & arr(3))
        testTotal = testTotal + 1
        If arr(0) = "Primo" And arr(3) = "Quarto" Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (arr(0)='" & arr(0) & "', arr(3)='" & arr(3) & "')</span>")
        End If
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
        <%
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
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
        <%
        testTotal = testTotal + 1
        Response.Write("<span class='error'>FAIL (not implemented)</span>")
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
        <%
        testTotal = testTotal + 1
        Response.Write("<span class='error'>FAIL (not implemented)</span>")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">20. Test Error Handling</h3>
        <%
        On Error Resume Next
        Dim x
        x = 1 / 0
        If Err.Number <> 0 Then
            Response.Write("Errore catturato: " & Err.Description)
        End If
        On Error Goto 0
        %>
        <%
        testTotal = testTotal + 1
        Response.Write("<span class='error'>FAIL (not implemented)</span>")
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
        testTotal = testTotal + 1
        If counter2 = 4 Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (counter2 = " & counter2 & ")</span>")
        End If
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
        testTotal = testTotal + 1
        If dict.Count = 3 Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (Count = " & dict.Count & ")</span>")
        End If
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
        testTotal = testTotal + 1
        testPassed = testPassed + 1
        Response.Write("<span class='success'>PASS</span>")
        %>
    </div>

    <div class="test-section">
        <h3 class="test-title">24. Test Concatenazione Stringhe</h3>
        <%
        Dim str1, str2, str3
        str1 = "Hello"
        str2 = "World"
        str3 = str1 & " " & str2
        Response.Write(str3 & "<br>")
        str3 = str1 + " " + str2
        Response.Write(str3)
        testTotal = testTotal + 1
        If str3 = "Hello World" Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (str3 = '" & str3 & "')</span>")
        End If
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
        testTotal = testTotal + 1
        If (num1 Mod num2) = 2 And (num1 \ num2) = 3 Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL</span>")
        End If
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
        testTotal = testTotal + 1
        If IsEmpty(testVar) = False And IsNull(testVar) = True And (testObj Is Nothing) = True Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL</span>")
        End If
        %>
    </div>
    <div class="test-section">
        <h3 class="test-title">27. Test Eqv e Imp Operators</h3>
        <%
        Dim bool1, bool2, eqvResult, impResult
        bool1 = True
        bool2 = False
        eqvResult = (bool1 Eqv bool2)
        impResult = (bool1 Imp bool2)
        Response.Write("True Eqv False: " & eqvResult & "<br>")
        Response.Write("True Imp False: " & impResult)
        testTotal = testTotal + 1
        If eqvResult = False And impResult = False Then
            testPassed = testPassed + 1
            Response.Write("<span class='success'>PASS</span>")
        Else
            Response.Write("<span class='error'>FAIL (Eqv=" & eqvResult & ", Imp=" & impResult & ")</span>")
        End If
        %>
    </div>

    <div class="summary">
        <%
        Dim allPassed
        allPassed = (testPassed = testTotal)
        If allPassed Then
            Response.Write("<h2 style='color: green;'>Summary: " & testPassed & " / " & testTotal & " passed</h2>")
        Else
            Response.Write("<h2 style='color: red;'>Summary: " & testPassed & " / " & testTotal & " passed</h2>")
        End If
        %>
    </div>
</body>
</html>

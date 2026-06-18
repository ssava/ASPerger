<%
Dim pageTitle, submitted, userName, userAge, userColor, encName, encAge, encColor, birthYear, category, msg
pageTitle = "Form Demo - ASPerger Demo"
submitted = (Request.Form("action") = "submit")

If submitted Then
    userName = Request.Form("name")
    userAge = Request.Form("age")
    userColor = Request.Form("color")
End If
%>
<!--#include file="include/header.asp"-->

<h2>Form Processing Demo</h2>

<%
If Not submitted Then
    Response.Write("<p>Fill out the form below to see ASP Classic form handling in action.</p>")
    Response.Write("<form method=""post"" action=""form_demo.asp"">")
    Response.Write("<label>Name: <input type=""text"" name=""name"" id=""name""></label>")
    Response.Write("<label>Age: <input type=""text"" name=""age"" id=""age""></label>")
    Response.Write("<label>Favorite color: <input type=""text"" name=""color"" id=""color""></label>")
    Response.Write("<input type=""hidden"" name=""action"" value=""submit"">")
    Response.Write("<input type=""submit"" class=""button"" value=""Submit"">")
    Response.Write("</form>")
Else
    encName = Server.HTMLEncode(userName)
    encAge = Server.HTMLEncode(userAge)
    encColor = Server.HTMLEncode(userColor)
    Response.Write("<h3>Thank You, " & encName & "!</h3>")
    Response.Write("<table class=""hover"">")
    Response.Write("<thead><tr><th>Field</th><th>Value</th></tr></thead>")
    Response.Write("<tbody>")
    Response.Write("<tr><td>Name</td><td>" & encName & "</td></tr>")
    Response.Write("<tr><td>Age</td><td>" & encAge & "</td></tr>")
    Response.Write("<tr><td>Color</td><td>" & encColor & "</td></tr>")
    Response.Write("</tbody></table>")

    If IsNumeric(userAge) Then
        birthYear = Year(Now()) - CInt(userAge)
        If CInt(userAge) < 18 Then
            category = "minor"
        ElseIf CInt(userAge) < 65 Then
            category = "adult"
        Else
            category = "senior"
        End If
        msg = "You were likely born around " & birthYear & "."
        msg = msg & " You are categorized as " & category & "."
        Response.Write("<p>" & msg & "</p>")
    Else
        Response.Write("<p>Age value doesn't appear to be numeric.</p>")
    End If

    Response.Write("<p><a href=""form_demo.asp"">&laquo; Back to form</a></p>")
End If
%>

<!--#include file="include/footer.asp"-->

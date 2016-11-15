<head><title>Private Chat - Login Required</title></head><body>
<center><h1>Private Chat - Login Required</h1></center>
<%
Dim strTempName
Response.Expires = -1000
Response.Buffer = True
Session("UserLoggedIn") = ""
If Request.Form("login") = "true" Then
    CheckLogin
Else
    ShowLogin
End If

Sub ShowLogin
%>
<!-- --------------------\/ Login Form \/---------------------- -->
<form name=form1 action=logchat.asp method=post>
<center><table><tr><td align="right">
<b>User Name : </b><input type=text name=username>
</center></td></tr><tr><td align="right">
<b>Password : </b><input type=password name=userpwd>
</center></td></tr><tr><td><center>
<input type=hidden name=login value=true>
<input type=submit value="Let me in!">
</center></td></tr></table></center></form>
<%
End Sub

Sub CheckLogin
Dim adoConn, strConn, RS, blnNoUserMatch
strConn = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
    "DBQ=" & Server.MapPath("\prashanthi1\db\user.mdb")
blnNoUserMatch = True
Set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.Open(strConn)
Set RS = adoConn.Execute("SELECT * FROM userinfo")
Do While NOT RS.EOF
    If LCase(Request.Form("username")) = LCase(RS("name")) Then
        blnNoUserMatch = False
        If Request.Form("userpwd") = RS("pwd") Then
            strTempName = RS("username")
            LoginOK
            Exit Do
        Else
            Response.Write("<center><h2>Invalid Password for " & _
            RS("User") & ". Login Failed.</H2><br />Password is " & _
            "case-sensitive. Make sure your Caps-Lock is not on." & _
            "</center><br /><br />")
            ShowLogin
            Exit Do
        End If
    Else
        RS.MoveNext
    End If
Loop
If blnNoUserMatch Then 
    Response.Write("<center><h2>User Not Found. Login Failed." & _
    "</h2></center><br /><br />")
    ShowLogin
End If
Set RS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Sub LoginOK
Session("UserLoggedIn") = "true"
Session("User") = strTempName
Session("ToName") = ""
Response.Redirect "chat.asp"
End Sub
%>
________________________________________________________________
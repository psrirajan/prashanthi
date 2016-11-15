<%
Dim conn,rs,SQL,RecsAffected
Dim objconn,strQ

'Set conn=Server.CreateObject("ADODB.Connection")
'conn.Mode=adModeRead
'conn.ConnectionString = aConnectionString
'conn.Open
'Set rs =Server.CreateObject("ADODB.Recordset")



If Request.Form("username") <> "" Then
Set objconn = Server.CreateObject("ADODB.Connection")
Set objrs = Server.CreateObject("ADODB.Recordset")
connect="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"

Set objco = Server.CreateObject("ADODB.Connection")
Set objsel = Server.CreateObject("ADODB.Recordset")
con="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"


'response.write conn
'response.write strQ


'   objconn.Open(connect)

'response.write "Err code" & Err.Number 
'response.write "Err.Description " & Err.Description

'set objrs=objconn.execute(strQ)

Dim password,password2,error_flag,error_msg
error_flag="False" '' Setting error flag
error_msg="" '' Error message 
uname=Request("username")
oldpassword=Request("oldpassword")
password=Request("password")
password2=Request("password2")

dim RExp : set RExp = new RegExp
with RExp
.Pattern = "^[a-zA-Z0-9]{3,8}$"
.IgnoreCase = True
.Global = True
end with


If (not (RExp.test(password) and RExp.test(password2))) then
error_flag="True" 
error_msg = "<br><font color='red'>Please enter valid user name</font> "
End if 
'Response.write "err flag"  & error_flag
If (password<>password2 ) then
error_flag="True"
error_msg = error_msg + "<br><font color='red'>Passwords are not matching</font>"
End if 


if error_flag="False" then
str = "select * from userinfo " 
'response.write "str select" & str
objco.open(con)

If Err.Number= 0 Then
objsel.open str, objco 


Do while not objsel.EOF 
'Response.write Request.Form("username")
'Response.write objsel("name")

If Request.Form("username") = objsel("name") then 
  sql="update userinfo set pwd = '"& Request.Form("password") & " ' " & " where name = '" & Request.Form("username") &"' "
  'response.write "sql : " & sql

  objconn.open(connect) 
  objrs.open sql ,objconn
  Response.Write "<br><br><font color='red'>Successfully changed Password </font>"
Else

 if objsel.EOF then
  'objsel.close
  set objsel=nothing
  'objco.close
  'set objco=nothing
  error_msg = "The name entered doesnot exist"
  response.write "<font color='red'>" & error_msg & "</font>"
  response.write "<a href='chgpwd.asp'>Go Back</a><br><br>"
  Response.end
 End if
  End If
objsel.MoveNext
loop

End If 

'objrs.close
Set objrs = Nothing
objconn.Close
Set conn = Nothing
End If
Else
 'response.write "Err msg : " & Err.Description

End If
%>



<html>
<head>
<link rel=StyleSheet href="web.css" type="text/css">
</head>
<body>

<form method=post action=chgpwd.asp>
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" align=center width=400>
<tr><td><%=error_msg%></td></tr>
<tr><td>User Name : </td>
<td bgcolor="green"><input type=text name="username"></td>

<tr><td colspan2 align=center><b><font color="blue">Change Password</font></b></td></tr>

<tr><td bgcolor="green"><font color="white"> Password</font></td><td><input type=password name=password></tr>

<tr><td bgcolor="green"><font color="white">Re-enter Password</font></td><td><input type=password name=password2></tr>

<tr><td colspan=2 align=center><input type=submit value="Update Password"></td></tr>
</table>
</div>
</form>
</body>
</html>
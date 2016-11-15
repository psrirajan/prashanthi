<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Code to prevent the page from being cached
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

'Check if the user is logged in. If not redirect them to the login page with an error message.
'If Session("VarName") = "" Then 
'	Response.Redirect("login.asp?errmsg=Please login")
'End If

Set objconn=Server.CreateObject("ADODB.Connection")

conn="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"

'conn.Provider="Microsoft.Jet.OLEDB.4.0"
'conn.Open Server.MapPath("user.mdb")

objconn.open(conn)
'Check if the form has been submitted
'If Request.Form("Submit") = "Update" Then
	
	'Simple string validation function. This can be expanded on.
	Function validateStr(str)
		temp = str
		temp = Trim(temp)
		temp = Replace(temp,"'","''")
		
		validateStr = temp
	End Function
	
	'Simple email validation function. e.g. abc123@domain.ext
	Function validateEmail(email)
		isValidE = True
		set regEx = New RegExp
		
		regEx.IgnoreCase = False
		
		regEx.Pattern = "^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
		isValidE = regEx.Test(email)
		
		validateEmail = isValidE
	End Function
	
	'Grab the form data
	user_name = validateStr(Request.Form("user_name"))
	user_pass = validateStr(Request.Form("user_pass"))
	confirm_pass = validateStr(Request.Form("confirm_pass")) 
	user_email = validateStr(Request.Form("user_email"))
   identity = Request.Form("identity")
'response.write "name :" & user_name & "<br />"
'response.write "pass :" & user_pass & "<br />"
'response.write "conf pass:"  & confirm_pass & "<br />"
'response.write "email :" & user_email & "<br />"
	
'response.write "identity : " & identity & "<br />"


	'Data validation
If user_name  = "" Then
 errMsg = errMsg & " -Please enter a name <br />"
End If

	If user_pass = "" Then
  errMsg = errMsg & "-Please enter a password<br />"
	ElseIf user_pass <> confirm_pass Then
		errMsg = errMsg & "-Passwords do not match<br />"
	End If
'response.write "err : " & errMsg  & "<br />"	
	If user_email = "" Then
		errMsg = errMsg & "-Please enter an email address<br />"
	ElseIf validateEmail(user_email) = False Then
		errMsg = errMsg & "-Email addrress is not valid<br />"
	End If
 
	'response.write "err : " & errMsg  & "<br />"

	'If there are no errors then continue with the update
	If Len(errMsg) = 0 Then	
		sql = "INSERT INTO userinfo(name,pwd,email,ident) VALUES ( '" & user_name  & "' ," &_  
           " '" & user_pass & "' ," &_
         "  '" & user_email & "' , " &  identity & "  )"
        
         
			   
'response.write sql
'response.end
		objconn.Execute sql

'response.end
		'errMsg = "Details updated"
	End If

'End If

'Set rsUserDetails = Server.CreateObject("ADODB.Recordset")
'sql = "SELECT * FROM userinfo WHERE name = '" & Session("VarName") & "'"
'rsUserDetails.Open sql, conn
'If Not rsUserDetails.EOF Then
'	user_pass = rsUserDetails("pwd")
'	user_email = rsUserDetails("user_email")
'End If
'rsUserDetails.Close
%>
<html>
<head>
<style>
  .errorMessage {
    color : #F00;
  }
  .errorItem {
    background : #F99;
  }
</style>

<link rel=StyleSheet href="web.css"
type="text/css">
</head>
<body>
 <div align="center"><center>
<%
IF errMsg <> "" THEN
  %>
  <p class="errorMessage"> <b><%=errMsg%></b></p>
  

  <%
ELSE
  %>
  <h3>Thank you!</h3>
  </body>
  </html>
  <%
 Session("ident")=Request.Form("identity")
  Response.End
END IF
%>
</center>
<form ACTION="reg.asp" name="saveform" METHOD="POST" align="center">

<div align="center"><center>
<table align="left" border="0" width="704" cellspacing="0" cellpadding="0">
<tr>
<td ><div align="right"><table border="0" height="59" width="226"
bgcolor="#FFFFFF" cellspacing="1" cellpadding="0">
<tr>
<td width="154" height="19" bgcolor="green">
<div align="right"><p><font color="#FFFFFF">Username:</font></td>
<td width="133" height="19" bgcolor="green">
<div align="left"><p><input NAME="user_name" id="user_name"
VALUE SIZE="8" MAXLENGTH="16" tabindex="1"></td>
<td width="64" height="19" bgcolor="#C0C0C0" align="center">
<div align="center"><center><p><a
href="javascript:alert('The username must be between 4 and 16 characters long.')"><small><small>Help</small></small></a></td>
</tr>
<tr align="center">
<td width="154" height="17" bgcolor="green"><div align="right"><p>
<font color="#FFFFFF">Password:</font></td>
<td height="17" width="133" bgcolor="green"><div align="left"><p>
<input type="password" name="user_pass" id="user_pass" size="8" tabindex="2" maxlength="8"></td>

<td height="17" bgcolor="#C0C0C0" align="center"><a
href="javascript:alert('The password must be between 4 and 8 characters long.')"><small><small>Help</small></small></a></td>
</tr>
<tr align="center">
<td width="154" height="17" bgcolor="green"><div align="right"><p>
<font color="#FFFFFF">Confirm Password:</font></td>
<td height="17" width="133" bgcolor="green"><div align="left"><p>
<input type="password" name="confirm_pass" id="confirm_pass" size="8" tabindex="2" maxlength="8"></td>
</tr>
<tr align="center">
<td width="154" height="17" bgcolor="green"><div align="right"><p>
<font color="#FFFFFF">Email:</font></td>

<td height="17" width="133" bgcolor="green"><div align="left"><p>
<input name="user_email" type="text" id="user_email"></td>
</tr>
<tr align="center">
<td width="154" height="17" bgcolor="green"><div align="right"><p>
<font color="#FFFFFF">Identification:</font></td>

<td style="align:left" height="17" width="133" bgcolor="green"><p>
<div style="text-align:left;">
<select name="identity"  style="left">
<option value="1">Doctor</option>
<option value="2">Patient</option>
</select>
</div>


</td>
</tr>
<tr align="center">
<td width="154" height="1" bgcolor="green"></td>
<td width="133" height="1" bgcolor="green"><div align="left"><p>
<input TYPE="submit" NAME="FormsButton2" VALUE="Register"  bgcolor="#C0C0C0" tabindex="3"
style="font-family: Verdana; font-size: 8pt"></td>
<td width="64" height="1" align="center"><a
href="javascript:alert('Click to save the details')"><small><small>Help</small></small></a></td>
</tr>
</table>
</div></td>

</tr>

</table>
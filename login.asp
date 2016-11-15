<%
Session("VarName")=""
Session("Identity")=""
error_message = ""
found_error = false

sub traperror(error)
	found_error = true
	error_message = error_message & error & ", "
		response.write(vbCrLf & "<table width=""75%"" border=""1"" bordercolor=""red"">")
		response.write(vbCrLf & "  <tr>")
		response.write(vbCrLf & "    <td>")
		response.write(vbCrLf & "      <font class=""main"">")
		response.write(vbCrLf & "      <h4>Error:</h4>")
		response.write(vbCrLf & "      Error #" & err.number & "<br />")
		response.write(vbCrLf & "      Error Source: " & err.source & "<br />")
		response.write(vbCrLf & "      Error Description: " & err.description & "<br />")
		response.write(vbCrLf & "      <br />")
		response.write(vbCrLf & "      <h5>There has been an error. Please try again later.</h5>")
		response.write(vbCrLf & "      <h5><i>If the problem persists, please </i><a href=""/info/contact.asp?error_num=" & err.number & "&error_src=" & err.source & "&err_description=" & err.description & "&sub=err"" class=""main"">contact</a><i> the webmaster</i></h5>")
		response.write(vbCrLf & "      </font>")
		response.write(vbCrLf & "    </td>")
		response.write(vbCrLf & "  </tr>")
		response.write(vbCrLf & "</table>")
end sub

on error resume next
%>



<html>
<head>
<link rel=StyleSheet href="web.css" type="text/css">
<link rel="stylesheet" type="text/css" href="tswtabs.css" />
</head>
<body>
<form ACTION="loginresult.asp" name="saveform" METHOD="POST" align="center">
<div align="center"><center><table border="0" width="704" cellspacing="0" cellpadding="0">
<table align="Left">
<tr>
<td width="385"><div align="right"><table border="0" height="59" width="226"
bgcolor="#FFFFFF" cellspacing="1" cellpadding="0">
<tr>
<td width="154" height="19" bgcolor="green">
<div align="right"><p><font color="#FFFFFF">Username:</font></td>
<td width="133" height="19" bgcolor="green">
<div align="left"><p><input NAME="username"
VALUE SIZE="8" MAXLENGTH="16" tabindex="1"></td>
<td width="64" height="19" bgcolor="#C0C0C0" align="center">
<div align="center"><center><p><a
href="javascript:alert('The username must be between 4 and 16 characters long.')"><small><small>Help</small></small></a></td>
</tr>
<tr align="center">
<td width="154" height="17" bgcolor="green"><div align="right"><p>
<font color="#FFFFFF">Password:</font></td>
<td height="17" width="133" bgcolor="green"><div align="left"><p><input type="password"
name="password" size="8" tabindex="2" maxlength="8"></td>
<td height="17" bgcolor="#C0C0C0" align="center"><a
href="javascript:alert('The password must be between 4 and 8 characters long.')"><small><small>Help</small></small></a></td>
</tr>
<tr align="center">
<td width="154" height="1" bgcolor="green"></td>
<td width="133" height="1" bgcolor="green">
<div id="tswcsstabs" align="left">
<ul>
<li>
<input TYPE="submit"
NAME="FormsButton2" VALUE="Login"  tabindex="3"
style="font-family: Verdana; font-size: 8pt"></td>
</li>
</ul>
<td width="64" height="1" bgcolor="#C0C0C0" align="center"><a
href="javascript:alert('Click to save the details')"><small><small>Help</small></small></a></td>



</tr>
<tr>
<div id="tswcsstabs">
<ul>
<li><a href="chgpwd.asp" target="content">Change Password </a></li>
</ul>
</div>
</tr>
</table>


</center></div>

</form>



</body>
</html>

<%
  If err.number <> 0  Then
    traperror err.description
  End If
%>
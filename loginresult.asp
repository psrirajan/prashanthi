<% 
Response.Buffer = True
'response.end

Err.Clear
On Error Resume Next

Dim objconn,conn,strQ


Set objconn = Server.CreateObject("ADODB.Connection")
Set objrs = Server.CreateObject("ADODB.Recordset")
conn="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"
'response.write conn
strQ = "SELECT * from userinfo where name = '" & Request.Form("username")
strQ = strQ & "' and pwd = '" & Request.Form("password") & "'"

'response.write conn
'response.write strQ
'response.End

'If not objconn.open then
   objconn.Open(conn)
'Else
'response.write "Err code" & Err.Number 
'response.write "Err.Description " & Err.Description
'End If
'On Error GoTo 0

'objrs.CursorType=adOpenStatic

set objrs=objconn.execute(strQ)
 
'response.write "cnt = " & objrs(0).Value
If Err.Number <> 0 Then
    Response.Write (Err.Description& "<br><br>")
  Response.Write("This means there is most likely a problem " & vbCrLf)
  'Response.Write("""ConnectionString"" info that you specified.<br>" & vbCrLf)
  Response.End
  
Else

'response.end
%>
<html>
<head>

<link rel=StyleSheet href="web.css" type="text/css">
<link rel="stylesheet" type="text/css" href="tswtabs.css" />
</head>
<body>
<%
 'Response.Write "objname = " & objrs("name")
 'Response.Write "name = " & Request.Form("username")
'Response.Write "objpwd = " & objrs("pwd")
 'Response.Write "password = " & Request.Form("password")
%>

<% If not objrs.EOF then 
 if (objrs("name") = Request.Form("username")  and objrs("pwd") = Request.Form("password") ) then %>
 
  <% Session("VarName") = Request.Form("username")
     Session("Identity") = objrs("ident") 
     Response.write "Ident : " & Session("Identity")

 
Else
  'response.write "user : " & Request.Form("username") & "pwd : " & Request.Form("password") 
  Session("VarName") = Request.Form("username")
  Session("Identity") = objrs("ident")
  'Response.write "Ident : " & Session("Identity")
End if %>
<TABLE BORDER=0>
<TR>    

</td>
</tr>
 <% 

  

Select Case Session("Identity")
 
    case 0 :

  %>
<table>
<Tr>
<td>
  Welcome Administrator 
</td>
</tr> 
<tr>   <TD><img src="green_flag_16.png" > - Appointment made </TD>
  <TD><img src="red_flag_16.png" > Appointment not made </TD> </tr>
</table>


<div id="tswcsstabs">
<ul>
<li><a href="AdminDoctor.asp?flg=d" target="content">Doctors</a></li>
</ul>
<ul>
<li><a href="AdminDoctor.asp?flg=p" target="content">Patients</a></li>
</ul>

</div>
<% 
'response.End %>

<table>

</tr>
</table>


 <a href="logout.asp" target="content"><img src="logout.jpg"></a>
<% 
'response.write "drname : " & objAdminrs("drname")
 'response.End %>
   <% Case   1  %>

<TABLE>
<TR>
    <TD> <FONT COLOR=BLUE> Welcome  Dr.<%=Session("VarName")%>  </FONT>
</td></tr>
</TABLE>

<div id="tswcsstabs">
<ul>
<li><a href="links.asp" target="content">Links</a></td></li>
<li><a href="http://www.google.com/calendar/hosted/brinkster.net" target="content">Calendar</a></td></li>
<li><a href="appoint.asp" target="content">Appointment</a></td></li>
<li><a href="logout.asp" target="content">Log Out</a></td></li>
</ul>
</div>
     

<% Case 2 %>
<TABLE>
<TR>
    <TD> <FONT COLOR=BLUE> Welcome  <%=Session("VarName")%>  </FONT>
</td></tr>
</TABLE>
<div id="tswcsstabs">
<ul>
<li><a href="evtcal.asp" target="content">Calendar</a></td></li>
<li><a href="appoint.asp" target="content">Appointment</a></td></li>
<li><a href="logout.asp" target="content">Log Out</a></td></li>
</ul>
</div>

<% End Select %>
   

</p></p></p>

</body>
<%
objconn.close

Set objconn = nothing
%>
</html>


<FONT COLOR=RED> 
<% Else
   Response.write "Your username or password does not match"  %>
</FONT>  
</p>
<a href="login.asp"> Login  </a>
<%  Response.End %>

<% 
End If
End If  

%>
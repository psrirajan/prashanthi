<%
Dim objAdmconn,Adminconn,strQ
Set objAdminconn = Server.CreateObject("ADODB.Connection")
Set objAdminrs = Server.CreateObject("ADODB.Recordset")
Adminconn="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"
'response.write Adminconn
strQ = "SELECT distinct name from userinfo where ident = 1"
strQ1 = "SELECT distinct name from userinfo where ident = 2"
'response.write strQ
objAdminconn.Open(Adminconn)
'response.write "objAdminconn Err code " & Err.number
'response.End
If Err.number = 0 then
  set objAdminrs = objAdminconn.Execute(strQ)
 ' response.write "EOF T/F" & objAdminrs.EOF 
End If 


%>


 <%
Select Case Request.Querystring("flg")  
Case "d" : 
Response.Write "Doctors List"
 If Not objAdminrs.EOF then
 do until objAdminrs.EOF 
%>

<table>
<tr>
<td>
</td>
</tr>
<tr>
  <td width="20%"> <font color="blue"><%=objAdminrs.Fields.Item("name")%></font>
</td>
</tr>  
<%  objAdminrs.MoveNext 
    loop
    End if
%>
</table>
<%
Case "p" :

  Response.Write "Patients list"
 set objAdminrs = objAdminconn.Execute(strQ1)

 If Not objAdminrs.EOF then
 do until objAdminrs.EOF 
%>

<table>
<tr>
  <td width="20%"> <font color="blue"><%=objAdminrs.Fields.Item("name")%></font>
</td>
</tr>  
<%  objAdminrs.MoveNext 
    loop
    End if
%>
</table> 
<%
End Select
%>
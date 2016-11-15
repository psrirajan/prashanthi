<%@ Language="VBScript" %>


<%
pnam = Request.QueryString("pname")
apptdate =Request.QueryString("apptdate")
appttime =Request.QueryString("appttime")
'Response.write "pname : " & 
if Request.QueryString("pname")= ""  then
   ' Response.write "No patient selected for deletion"
else
   DeleteApptrec pnam,apptdate,appttime
end if 

   Function DeleteApptrec(name,apdate,aptime)
      'Response.write "Patient name = " & name
Dim delsql
delsql = "Delete from appoints where ptname =  '" &  name & "'  and appointdate = '" & apdate & "' and appointtime =  '" &   aptime  & "' "
 Response.write delsql
 'sqldelete = "Delete from appoints where ptname =  '" &  name & "'  and appointdate = '" & apdate & "' and appointtime =  '" &   aptime  & "' "


     Response.End 
   End Function
%>

<% 

'if Session("VarName") <> NULL then
If Session("Identity") =  2 then 
If Request.Form("Submit") = "Make Appointment" then 
	  Response.Write " Submit : " & Request.Form("Submit") & "<br>"
   Response.Write "Dr.Name : " & Request.Form("drselect")
  Dim appdate
Dim makeappconn,objmakeappconn,objmakersappoint,connmakeapppoint
Set objmakeappconn = Server.CreateObject("ADODB.Connection")
Set objmakeapprs = Server.CreateObject("ADODB.Recordset")

makeappconn="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"
objmakeappconn.open(makeappconn)

  appdate = Request.Form("demo1")
  ' Response.Write "Appointment Date : " & appdate
   'Response.Write "Appointment Time : " & Request.Form("group1")
   Dim StrSelect
   StrSelect = "Insert INTO appoints (drname,ptname,appointdate,appointtime) "
   StrSelect = StrSelect & " values ( '" & Request.Form("drselect") & "',"
   StrSelect = StrSelect & "  '" & Session("VarName") & "', " 
   StrSelect = StrSelect & "  '" & Request.Form("demo1") &"'  ,"
   
    StrSelect = StrSelect & "  '" & Request.Form("group1") &"') " 
 'Response.Write StrSelect   
set objmakeapprs=objmakeappconn.execute(StrSelect)
' Response.End
End If 
End If

'  Response.Write "Dr.Name : "  & Session("VarName")
 ' Response.Write "Appoint Date : " & Request.Form("demo1")
 Dim apptime 
 apptime = Request.Form("group1")
  'Response.Write "Appoint Time : " & apptime
 ' Response.End
'else
'  Response.Write "Empty"
'  Response.End 
'end if
Dim apdate
if Request.Form("demo1") = "" then
 apdate= "" 
Else
 apdate = Request.Form("demo1")
'response.write "Date : " &   apdate
'response.
 'response.write "Date : " &  format(apdate,'mm/dd/yyyy')
End if

Err.Clear
On Error Resume Next
Dim Errmsg
Errmsg=""
Dim objconn,conn,strQ

Dim objappoint,objrsappoint,connapppoint,strQappoint

Set objconn = Server.CreateObject("ADODB.Connection")
Set objrs = Server.CreateObject("ADODB.Recordset")
conn="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"

strQ = "SELECT * from userinfo where ident = 1 " 
'Response.write "Identity : " & Session("Identity")
 'response.write "conn :  " & conn
'Response.write " strQ : " & strQ

'Response.End


'response.write strQ

   objconn.Open(conn)

'response.write "Err code" & Err.Number 
'response.write "Err.Description " & Err.Description
'On Error GoTo 0



set objrs=objconn.execute(strQ)

'response.end
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="tswtabs.css" />
<link rel="stylesheet" type="text/css" href="path/to/calendarview.css" />
<script language="javascript" type="text/javascript" src="datetimepicker.js">
</script>


</head>
<body>
<form name="appointment" method="post" action="appoint.asp">
<% if Session("VarName") = "" then %>
<div id="tswcsstabs">
<ul>
<li><a href="login.asp" target="content">Login</a></li>
</ul>
</div>
<% else %>
<div id="tswcsstabs">
<ul>
<li><a href="links.asp" target="content">Links</a></li>
<li><a href="evtcal.asp" target="content">Calendar</a></li>
<li><a href="appoint.asp" target="content">Appointment</a></li>
<li><a href="logout.asp" target="content">Log Out</a></li>
</ul>
</div>

<% end if %>

<%Response.Write "ID : " & Session("Identity")
'  Response.End  %>

<% Select Case Session("Identity")
  Case  0
    Set objappoint = Server.CreateObject("ADODB.Connection")
    strAdmin = "select * from userinfo "
    set rsAdmin = Server.CreateObject("ADODB.Recordset")
    connappoint="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"
    response.write strAdmin 
     response.End  
    objappoint.open(connappoint)
    set rsAdmin =objappoint.execute(strAdmin)

Case    1 

Set objappoint = Server.CreateObject("ADODB.Connection")
Set objrsappoint = Server.CreateObject("ADODB.Recordset")
connappoint="DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & server.mappath("prashanthi1\..\db\user.mdb") & ";"  & "Persist Security Info=False"
    
   'If apdate = "" then

'strAppoint = "select b.drname ,b.ptname , a.email ,b.appointdate,b.appointtime from userinfo a, appoints b where a.name=b.drname  "
strAppoint ="select b.drname ,b.ptname , a.email ,b.appointdate,b.appointtime from userinfo a, appoints b where a.name=b.ptname and b.drname='" & Session("VarName") & "'"

strAppoint = strAppoint & " and a.email in (select email from userinfo where email=a.email and a.ident = 2) order by drname "
'strAppoint= strAppoint & " and b.drname='" & Session("VarName") & "'"
'    strAppoint = strAppoint & ptname = Session 
	'Else 
  '	 strAppoint=strAppoint & " and " & " appointdate ='" &  Request.Form("demo1") & "'"
  'End if
'response.write "str appoint : " & strAppoint 
'response.End

objappoint.open(connappoint)
set objrsappoint=objappoint.execute(strAppoint)


%>

</p>
<img src="green_flag_16.png" > - Appointment made 
  <img src="red_flag_16.png" > Appointment not made
</p></p>
<table width="60%" border="2" border-spacing=2 border-color="black">

<% 'response.write objrsappoint.EOF
 

  

If objrsappoint.EOF = True then %>
     <img src="red_flag_16.png" >    
  <% Errmsg =  "No appointments booked "
  Response.Write Errmsg
  Response.End
Else %>
<tr bgcolor="lightgrey">
<td align="center">Doctor Name  </td>
<td align="center">Patient Name  </td>
<td align="center">Patient Email  </td>
<td align="center">Appointment Date </td>
<td align="center">Appointment Time </td>
<td width="20%"> <font color="green">Cancel</font></td>
</tr>
<%
cnt = 0
do until objrsappoint.EOF
%>
</tr>
<tr>
  <td width="20%"> <font color="blue"><%=objrsappoint.Fields.Item("drname")%></font></td>
  <td width="20%"><a style="color: green; text-decoration: none; background: yellow" href="patdetail.asp"><font color="blue"><%=objrsappoint.Fields.Item("ptname")%></font></a></td>
  <td width="20%"> <font color="blue"><%=objrsappoint.Fields.Item("email")%></font></td>


<td width="20%"><font color="blue"><%=objrsappoint.Fields.Item("appointdate")%></font></td>
<td width="20%"><font color="blue"><%=objrsappoint.Fields.Item("appointtime")%></font><img src="green_flag_16.png" ></td>
<td><a href="appoint.asp?pname=<%=objrsappoint.Fields.Item("ptname")%>&apptdate=<%=objrsappoint.Fields.Item("appointdate")%>&appttime=<%=objrsappoint.Fields.Item("appointtime")%>"><IMG SRC="cancel.jpg" ALT="some text" WIDTH=32 HEIGHT=32> </a></td>

<%
cnt=cnt+1  


drname(count)=objrsappoint.Fields.Item("drname")
patname(count)=objrsappoint.Fields.Item("ptname")
email=objrsappoint.Fields.Item("email")
appointdate=objrsappoint.Fields.Item("appointdate")
appointtime=objrsappoint.Fields.Item("appointtime")

objrsappoint.MoveNext %>

<%
  loop
 

%>

</td>

</tr>
<tr>
<td width="100">

</td>
</tr>
</table>
<% End If  %>
<% 

response.write "tot rec: " & cnt
'Response.End %>




<% Case  2  %>
     <table>
  <tr>
 
      <td>Patient's Name : </td>
      <td>

      <%=Session("VarName")%></td>
     </tr>

     <tr>
     <td>Doctor's Name : </td>
     <td>

<% 'response.End %>
    <select name="drselect" size="1">
      <option value="Choose a Doctor" selected="selected">Choose a  Doctor</option>
  <%
     do until  objrs.EOF 
  %>
  <option value="<%=objrs.Fields.Item("name")%>"><%=objrs.Fields.Item("name")%>
  </option>

  
<% 
  objrs.MoveNext
  loop
%>
</td>
</tr>
<tr>
<td>
Appointment Date : </td>
<td>
<input id="demo1" name="demo1" type="text" value="" size="25">
<a href="javascript:NewCal('demo1','ddmmyyyy')"><img src="cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>

<tr>
<td>Appointment Time : </td>
<td>
<div align="center"><br>
<input type="radio" name="group1" value="7 pm" checked> 7 pm
<input type="radio" name="group1" value="8 pm"> 8 pm 
<input type="radio" name="group1" value="9 pm"> 9 pm 
</div> 
</td>
</tr>
<tr>
<td>
  <input type="submit" name="Submit" onclick="get_radio_value()" value="Make Appointment">
</td>
</tr>
<% 'End If %>
<% End Select %>


</br></br>

 
</form>
</body>
 <%
    
objconn.close
set objconn=nothing
objappoint.close
set objappoint=nothing
%>

</html>
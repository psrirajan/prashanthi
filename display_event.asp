<%@ LANGUAGE=VBSCRIPT %>
<%Option Explicit%>

<%
Dim DB_CONNECTIONSTRING
Dim objRecordset
Dim Added
%>

	<!--#include file="adovbs.inc"-->

<%
'DB_CONNECTIONSTRING = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.Mappath("users.mdb") & ";"

'Set objRecordset = Server.CreateObject("ADODB.Recordset")
'objRecordset.Open "calendar", DB_CONNECTIONSTRING, adOpenStatic, adLockPessimistic, adCmdTable

'If Request.Form("btnAdd") = "Add Event" Then

'        objRecordset.AddNew

	'-- Add records to database from form --

'			objRecordset.Fields("Subject") = Request.Form("txtSubject")
	'		objRecordset.Fields("Message") = Request.Form("Message")
		'	objRecordset.Fields("Day") = Request.Form("selDay")
			'objRecordset.Fields("Month") = Request.Form("selMonth")
		'	objRecordset.Fields("Year") = Request.Form("selYear")
		'	objRecordset.Fields("AddedBy") = Request.Form("txtAddedBy")
		'	objRecordset.Fields("DateAdded") = Now()

'		objRecordset.Update

'		Added = "True"

'End If

'objRecordset.Close
'Set objRecordset = Nothing

If Added = "True" Then
	Response.Redirect("calendar.asp")
End If
%>

<html>

<center>

<form method="post" action="add_event.asp">

<table border="0" cellpadding="0" cellspacing="0">
 <tr>
  <td>Day: 

	<select name="selDay">

	<option value="1">1</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5</option>
	<option value="6">6</option>
	<option value="7">7</option>
	<option value="8">8</option>
	<option value="9">9</option>
	<option value="10">10</option>
	<option value="11">11</option>
	<option value="12">12</option>
	<option value="13">13</option>
	<option value="14">14</option>
	<option value="15">15</option>
	<option value="16">16</option>
	<option value="17">17</option>
	<option value="18">18</option>
	<option value="19">19</option>
	<option value="20">20</option>
	<option value="21">21</option>
	<option value="22">22</option>
	<option value="23">23</option>
	<option value="24">24</option>
	<option value="25">25</option>
	<option value="26">26</option>
	<option value="27">27</option>
	<option value="28">28</option>
	<option value="29">29</option>
	<option value="30">30</option>
	<option value="31">31</option>

	</select></td>

  <td>Month: 

	<select name="selMonth">

	<OPTION VALUE="1">January</option>
	<OPTION VALUE="2">February</option>
	<OPTION VALUE="3">March</option>
	<OPTION VALUE="4">April</option>
	<OPTION VALUE="5">May</option>
	<OPTION VALUE="6">June</option>
	<OPTION VALUE="7">July</option>
	<OPTION VALUE="8">August</option>
	<OPTION VALUE="9">September</option>
	<OPTION VALUE="10">October</option>
	<OPTION VALUE="11">November</option>
	<OPTION VALUE="12">December</option>

	</select></td>

  <td>Year: 

	<select name="selYear">

	<option value="1999">1999</option>
	<option value="2000">2000</option>
	<option value="2001">2001</option>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>

	</select></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center">Subject: <input type="text" name="txtSubject" size="35"></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center"><textarea name="Message" cols="40" rows="10"></textarea></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center">Added By: <input type="text" name="txtAddedBy" size="35" value="<%= Session("FirstName") & " " & Session("LastName")%>"></td>
 </tr>
</table>

<p>

<input type="submit" name="btnAdd" value="Add Event">&nbsp;&nbsp;<input type="Reset" name="btnReset" value="Clear">

</form>

</center>

</body>
</html>
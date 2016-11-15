<%@ LANGUAGE=VBSCRIPT %>
<%Option Explicit%>

<%
Dim DB_CONNECTIONSTRING
Dim objRecordset
Dim Updated
%>

	<!--#include file="adovbs.inc"-->

<%
If Request.Form("btnCalendar") = "Back To Calendar" Then
	Response.Redirect("calendar.asp")
End If

DB_CONNECTIONSTRING = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.Mappath("users.mdb") & ";"

Set objRecordset = Server.CreateObject("ADODB.Recordset")
objRecordset.Open "calendar", DB_CONNECTIONSTRING, adOpenStatic, adLockPessimistic, adCmdTable

	If Session("EventToEdit") <> 0 Then
		If Not objRecordset.EOF Then

			objRecordset.MoveFirst
			Do Until objRecordset.Fields("ID") = Session("EventToEdit")
				objRecordset.MoveNext
			Loop
	End If
		End If

If Request.Form("btnEdit") = "Update" Then

	'-- Add records to database from form --

			objRecordset.Fields("Subject") = Request.Form("txtSubject")
			objRecordset.Fields("Message") = Request.Form("Message")
			objRecordset.Fields("Day") = Request.Form("txtDay")
			objRecordset.Fields("Month") = Request.Form("txtMonth")
			objRecordset.Fields("Year") = Request.Form("txtYear")
			objRecordset.Fields("AddedBy") = Request.Form("txtAddedBy")
			objRecordset.Fields("DateAdded") = Now()

		objRecordset.Update

		Updated = "True"

End If

If Request.Form("btnDelete") = "Delete" Then

	'--Delete event from database --

		objRecordset.Delete adAffectCurrent

	Response.Redirect("calendar.asp")


End If
%>

<html>

<center>

<%
If Updated = "True" Then
	Response.Write("<font size=4><b><font color=" & Chr(34) & "red" & Chr(34) & ">" & objRecordset.Fields("Subject") & "</font> has been updated</b></font><p>")
End If
%>

<form method="post" action="edit_event.asp">

<table border="0" cellpadding="0" cellspacing="0">
 <tr>
  <td>Day: <input type="text" name="txtDay" size="4" value="<%= objRecordset.Fields("Day")%>"></td>
  <td>Month: <input type="text" name="txtMonth" size="10" value="<%= objRecordset.Fields("Month")%>"></td>
  <td>Year: <input type="text" name="txtYear" size="5" value="<%= objRecordset.Fields("Year")%>"></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center">Subject: <input type="text" name="txtSubject" size="35" value="<%= objRecordset.Fields("Subject")%>"></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center"><textarea name="Message" cols="40" rows="10"><%= objRecordset.Fields("Message")%></textarea></td>
 </tr>
 <tr>
  <td colspan="3">&nbsp;</td>
 </tr>
 <tr>
  <td colspan="3" align="center">Added By: <input type="text" name="txtAddedBy" size="35" value="<%= objRecordset.Fields("AddedBy")%>"></td>
 </tr>
</table>

<p>

<input type="submit" name="btnEdit" value="Update">&nbsp;&nbsp;
<input type="submit" name="btnDelete" value="Delete">&nbsp;&nbsp;
<input type="Reset" name="btnReset" value="Reset">&nbsp;&nbsp;
<input type="submit" name="btnCalendar" value="Back To Calendar">

</form>

</center>

</body>
</html>

<%
objRecordset.Close
Set objRecordset = Nothing
%>
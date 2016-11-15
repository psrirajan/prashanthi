<%@ LANGUAGE=VBSCRIPT %>
<%Option Explicit%>


	<!--#include file="adovbs.inc"-->

<%
If Request.Form("AddEvent") = "Add Event" Then
	Response.Redirect("add_event.asp")
End If

If Request.Form("EditEvent") = "Edit Event" Then
	Response.Redirect("edit_event.asp")
End If

Dim DB_CONNECTIONSTRING
Dim objRecordset

DB_CONNECTIONSTRING = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.Mappath("users.mdb") & ";"

Set objRecordset = Server.CreateObject("ADODB.Recordset")
objRecordset.Open "calendar", DB_CONNECTIONSTRING, adOpenStatic, adLockPessimistic, adCmdTable
%>

<%
Function GetDaysInMonth(iMonth, iYear)
	Select Case iMonth
		Case 1, 3, 5, 7, 8, 10, 12
			GetDaysInMonth = 31
		Case 4, 6, 9, 11
			GetDaysInMonth = 30
		Case 2
			If IsDate("February 29, " & iYear) Then
				GetDaysInMonth = 29
			Else
				GetDaysInMonth = 28
			End If
	End Select
End Function

Function GetWeekdayMonthStartsOn(iMonth, iYear)
	GetWeekdayMonthStartsOn = WeekDay(CDate(iMonth & "/1/" & iYear))
End Function

Function SubtractOneMonth(dDate)
Dim iDay, iMonth, iYear	
	iDay = Day(dDate)
	iMonth = Month(dDate)
	iYear = Year(dDate)

	If iMonth = 1 Then
		iMonth = 12
		iYear = iYear - 1
	Else
		iMonth = iMonth - 1
	End If
	
	If iDay > GetDaysInMonth(iMonth, iYear) Then iDay = GetDaysInMonth(iMonth, iYear)

	SubtractOneMonth = CDate(iMonth & "-" & iDay & "-" & iYear)
End Function

Function AddOneMonth(dDate)
Dim iDay, iMonth, iYear	
	iDay = Day(dDate)
	iMonth = Month(dDate)
	iYear = Year(dDate)

	If iMonth = 12 Then
		iMonth = 1
		iYear = iYear + 1
	Else
		iMonth = iMonth + 1
	End If
	
	If iDay > GetDaysInMonth(iMonth, iYear) Then iDay = GetDaysInMonth(iMonth, iYear)

	AddOneMonth = CDate(iMonth & "-" & iDay & "-" & iYear)
End Function


Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table

If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
	Else
		dDate = Date()

		If Request.QueryString.Count <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
	End If
End If

iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(Month(dDate), Year(dDate))
%>

<html>

<center>

<table border="1" cellspacing="0" cellpadding="1">
	<tr>
		<td bgcolor="blue" align="center" colspan="7">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td align="right"><b><A HREF="calendar.asp?date=<%= SubtractOneMonth(dDate) %>" style="color: #FFFFFF">&lt;--</A></b></TD>
					<td align="center"><font color="#FFFFFF"><B><%= MonthName(Month(dDate)) & "  " & Year(dDate) %></B></font></TD>
					<td align="left"><b><A HREF="calendar.asp?date=<%= AddOneMonth(dDate) %>" style="color: #FFFFFF">--&gt;</A></b></TD>
					<td align="right" valign="top"><a href="http://www.4guysfromrolla.com"><img src="close.gif" border="0"></a></td>
				</tr>
			</TABLE>
		</td>
	</tr>
	<tr bgcolor="blue">
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Sun</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Mon</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Tue</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Wed</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Thu</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Fri</font></B></TD>
		<td ALIGN="center" width="80"><B><font color="#FFFFFF">Sat</font></B></TD>
	</tr>
<%
If iDOW <> 1 Then
	Response.Write(vbTab & "<tr>" & vbCrLf)
	iPosition = 1
	Do While iPosition < iDOW
		Response.Write(vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf)
		iPosition = iPosition + 1
	Loop
End If

	'-- Write days of month in proper day slots --

iCurrent = 1
iPosition = iDOW

Do While iCurrent <= iDIM


	'-- open the table row --

	If iPosition = 1 Then
		Response.Write(vbTab & "<tr>" & vbCrLf)
	End If


	'-- Write the date and subject --

		Response.Write(vbTab & vbTab & "<td align=left valign=top height=60><b>" & iCurrent & "</b>")



	If Not objRecordset.BOF Then
		objRecordset.MoveFirst
		Do Until objRecordset.EOF

	If objRecordset.Fields("Year") = Year(dDate) Then
	    If objRecordset.Fields("Month") = Month(dDate) Then

		If objRecordset.Fields("Day") = iCurrent Then
		 Response.Write("<br><font size=2><a href=" & Chr(34) & "display_event.asp?ID=" & objRecordset.Fields("ID") & Chr(34) & ">" & objRecordset.Fields("Subject") & "</a></font><br>")
		End If

	    End If
	End If

			objRecordset.MoveNext

		Loop

	End If



		Response.Write("</td>" & vbCrLf)


	'-- Close the table row --

	If iPosition = 7 Then
		Response.Write vbTab & "</tr>" & vbCrLf
		iPosition = 0
	End If

	
	'-- Increment variables --

	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
Loop

If iPosition <> 1 Then
	Do While iPosition <= 7
		Response.Write(vbTab & vbTab & "<td>&nbsp;</td>" & vbCrLf)
		iPosition = iPosition + 1
	Loop
	Response.Write vbTab & "</TR>" & vbCrLf
End If
%>
</table>


<%
objRecordset.Close
Set objRecordset = Nothing
%>

<%
	Response.Write("<form action=" & Chr(34) & "calendar.asp" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & ">" & Chr(10))
	Response.Write("<input type=" & Chr(34) & "submit" & Chr(34) & " name=" & Chr(34) & "AddEvent" & Chr(34) & " value=" & Chr(34) & "Add Event" & Chr(34) & ">" & "&nbsp;&nbsp;")
	Response.Write("</form>")
%>

</center>

</body>
</html>
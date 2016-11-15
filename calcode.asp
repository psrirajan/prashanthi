<%
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Created 2001 by Jason Benassi
'email jbenassi@futurecasting.com
'
'To change the month, Pass in a date with POST or GET of a valid date
'to a variable named incDate
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Function GetDateName (monthNumber)

	Select Case monthNumber
	
		Case 1
			GetDateName = "January"	
		Case 2
			GetDateName = "February"
		Case 3
			GetDateName = "March"
		Case 4
			GetDateName = "April"
		Case 5
			GetDateName = "May"
		Case 6
			GetDateName = "June"
		Case 7
			GetDateName = "July"
		Case 8
			GetDateName = "August"
		Case 9
			GetDateName = "September"
		Case 10
			GetDateName = "October"
		Case 11
			GetDateName = "November"
		Case 12
			GetDateName = "December"
	End Select

End Function

	Dim incDate	, boxSize
	
	'Pixel size for the cell boxes
	boxSize = "70"
	incDate = Request("incDate")
		If Not incDate <> "" Then
		
			'Get the current months first day
			incDate = DatePart ("m" , Date) & "/01/" & DatePart("yyyy" , Date)
		End If	
%>


<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Calendar</title>
</head>

<body style="font-family: Arial">
<table border="1" width="751" bgcolor="#DEDFC6">
    
    <tr><td>
        <p align="center"><b><%=GetDateName (DatePart ("m" , incDate)) & " " &  DatePart("yyyy" , incDate)%></b></p>
      </td></tr>
    <td width="100%">
       <table border="1" width="750" bordercolorlight="#319A63" bordercolordark="#FFFFFF">
  <tr>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Sunday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Monday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Tuesday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Wednesday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Thursday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Friday</b></font></td>
    <td width="<%=boxSize%>" align="left" nowrap><font face="Arial Narrow" size="2"><b>Saturday</b></font></td>
  </tr>
  
<%
'Calendar loop
		
		Dim w , d , count, dispDay, currDay
		dateCount = incDate
		count = 1
		dispDay = 1
		For w = 1 to 6
		%>
  <tr>
<%		
  				For d = 1 to 7
					
						If Not (count => DatePart("w" , incDate)) And (count < 8) Then

%>
	<td width="<%=boxSize%>" height="<%=boxSize%>" valign="top" align="left">&nbsp</td>
<%
						count = count + 1
					
						Else

							
							If IsDate( DatePart("m" , incDate) & "/" & dispDay & "/" & DatePart("yyyy" , incDate)) Then	
								
								currDay = DatePart("m" , incDate) & "/" & dispDay & "/" & DatePart("yyyy" , incDate)
										
							If currDay = CStr(Date) Then
								
%>			
	<td bgcolor ="green" width="<%=boxSize%>" align="left" height="<%=boxSize%>" valign="top"><%=dispDay%></td>
<%
							Else
%>			
	<td width="<%=boxSize%>" align="left" height="<%=boxSize%>" valign="top"><%=dispDay%></td>
<%
							End If
							count = count + 1
							dispDay = dispDay + 1				
						
							Else
%>
	<td width="<%=boxSize%>" align="left" height="<%=boxSize%>" valign="top">&nbsp</td>
<%
							End If
						
						End If
						
					
					
					Next	
%>
  </tr> 
<%		
		Next
%>

</table>
</td>
</table>

</body>

</html>
<form method="get" action="AddEvent.asp">

   <br>Currently Scheduled Events for <b><%
    Response.Write MonthName(nMonth) & "&nbsp;&nbsp;" & nYear
    %></b>:<p>
   <table bgcolor="gray" border=1 cellpadding=3>
   <tr bgcolor="Blue"><td><font color="white"><b>Day</b></font></td>
      <td colspan=2><font color="white"><b>Event</b></font></td></tr>

   <% if rs.EOF then
      Response.Write "<tr><td colspan=3 bgcolor=""#ffffc0"">No events listed</td></tr>"
   end if

   while not rs.EOF
      Response.Write "<tr bgcolor=""#ffffc0""><td>" & rs("nDay") & "</td><td>"
      Response.Write rs("vcEvent") & "</td><td><input type=""button"" value=""Remove"""
      Response.Write " onClick=""window.location.href='RemoveEvent.asp?nMonth=" & nMonth
      Response.Write "&nYear=" & nYear & "&idSchedule=" & rs("idSchedule") & "'""></td></tr>"
      rs.MoveNext
   wend
   %></table>

   <p><br>
   <table bgcolor="gray" border=1 cellpadding=3>
      <tr bgcolor="Blue"><td><font color="white"><b>Add New Event:</b></font></td></tr>
      <tr bgcolor="#ffffc0"><td>Event:
         <input type="text" size=30 maxlength=100 name="Event"> Day: <select name="nDay"><%
         ' Set the date to the first of the current month
         dtDate = DateSerial(nYear, nMonth, 1)

         dtTemp = dtDate
         do
            Response.Write "<option value=""" & Day(dtTemp) & """>" & Day(dtTemp)
            dtTemp = DateAdd("d", 1, dtTemp)
         loop until (Month(dtTemp) <> CInt(nMonth)) %></select>
         <input type="hidden" name="nMonth" value="<%= nMonth %>">
         <input type="hidden" name="nYear" value="<%= nYear %>">
         <input type="Submit" value="Add Event"></td></tr>
   </table>

https://www.google.com/calendar/embed?src=psrirajan%40gmail.com&ctz=Asia/Calcutta
   </form>
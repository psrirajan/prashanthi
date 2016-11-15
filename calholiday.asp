<HTML>
<HEAD>
<TITLE>Calendar</TITLE>
</HEAD>
<STYLE>
<!--
TD.DAYS {
	font-family : sans-serif;font-size : 8pt;font-style : normal;letter-spacing : 0.1mm;border : thin ridge Aqua ;background-color : #E0FFFF;
}
TD.HEAD {
	font-family : sans-serif;font-size : 8pt;font-style : normal;font-weight : bold;letter-spacing : 0.1mm;border : thin ridge ;background-color : #A0A0A0;
}
TD.MAIN {
	border : medium double Double;
}
TD.TODAY {
	font-family : sans-serif;font-size : 12pt;font-weight : bold;font-color : blue;letter-spacing : 0.1mm;border : thin groove #FFFACD;background-color : #FFFFCC;
}
TD.PUBLIC {
	font-family : sans-serif;font-size : 8pt;font-weight : bold;font-color : blue;letter-spacing : 0.1mm;border : thin ridge Aqua ;background-color : #AAFFCC;
}
-->
</STYLE>
<BODY>

<FORM ACTION="CALENDAR.ASP" METHOD="POST">
<SELECT NAME="MONTH">
<%
	'****************************************************************
	' This is a Calendar that can be used to enter a single memo on a date
	' cell.I have used it on a staff intranet page and shows important events
	' That are coming up. Also can be used as a leave planner etc.
	' The size can be varied and a list of public holidays highlighted.
	' Today's date is also highlighted.
	' memo.txt is used to load the info in the form of 
	' yyyy/mm/dd;Memo Text
	'
	' 2001/03/29;Supper at Ted's
	' 2001/04/12;Movies at 21:00
	' 2001/04/15;Return Videos
	'
	' memo.gif is stored in /images folder of the reference page
	' The memo text is added as alt text of the gif and will be seen when 
	' the mouse is held over the gif
	' any suggestions/questions info@cool.co.za
	'****************************************************************
	
	'****************************************************************
	'Build Select boxes and fill  them with the months and years 
	'that the user can select
	'****************************************************************

For I = 1 to 12                                         
   If I <> MONTH(NOW) then
     Response.Write "<OPTION VALUE=" &I &">"  &MonthName(I)
   else
     Response.Write "<OPTION VALUE=" &I &" SELECTED>"  &MonthName(I)  
   end if  
Next
%>
</SELECT>
<SELECT NAME="YEAR">
<%
  Startdate = DateAdd("yyyy",-5,NOW)
   For I = 0 to 10
   dtCalc = YEAR(DateAdd("yyyy",I,StartDate))
    If dtCalc <> YEAR(NOW) then
     Response.Write "<OPTION VALUE=" &dtCalc &">"  &dtCalc
     else
     Response.Write "<OPTION VALUE=" &dtCalc &" SELECTED>"  &dtCalc
    end if 
   Next  
 
%>
</SELECT>
<BR>
<!-- The cell dimensions & abbrviation checkbox
 can be hardcoded and were used only during development-->
WIDTH:<INPUT TYPE="TEXT" VALUE="75" NAME="CELLW" STYLE="WIDTH=25px">
HEIGHT:<INPUT TYPE="TEXT" VALUE="50" NAME="CELLH" STYLE="WIDTH=25px">
<BR>
Abbreviate Day Names <INPUT TYPE=CHECKBOX NAME="ABB" VALUE="1"><BR>
<INPUT TYPE="SUBMIT" VALUE="BUILD">
</FORM>
<%
 
	'****************************************************************
	' If the form is being visited for the first time this code will 
	'                           not execute
	'****************************************************************


 If Len(Request.Form("MONTH")) then 

	   '*******************************************************************
	   '  Open text file and extract memo information and store in arrays
	   '*******************************************************************
	  
	   Redim aMemoInfo(0)
	   Redim aMemoDate(0)
	   arrIndex = 0
	   
	   Set FSO = CreateObject("Scripting.FileSystemObject")
	   Set File = FSO.GetFile(Server.MapPath("\ASP_Source\Calendar") & "/memo.txt")
	   Set TextStream = File.OpenAsTextStream(1)
	     Do While Not TextStream.AtEndOfStream
	       S = TextStream.ReadLine 
	       aMemoDate(arrIndex) = Left(S,InStr(1,S,";")-1)
	       aMemoInfo(arrIndex) = Right(S,Len(S)-InStr(1,S,";"))
	       arrIndex = arrIndex + 1
	       Redim Preserve aMemoInfo(arrIndex)
	       Redim Preserve aMemoDate(arrIndex)
	     Loop
		   
	   TextStream.Close
	   Redim Preserve aMemoInfo(arrIndex-1)
	   Redim Preserve aMemoDate(arrIndex-1)
	
        '*************************************
        '   Editable list of Public holidays
        '*************************************

          arrPublic = array("2001/02/27","2001/03/12","2001/04/15","2001/04/30")
    
        '*************************************
        '       Get the cell dimensions
        '*************************************
          'Const CELL_WIDTH  = 100
          'Const CELL_HEIGHT = 100
          CELL_WIDTH  = Request.Form("CELLW")
          CELL_HEIGHT = Request.Form("CELLH")

        '*************************************
        '  Must the Day Names be abbreviated
        '*************************************

	  'blnAbbr     = TRUE
          blnAbbr      = Request.Form("ABB")      
          
          

	dtSelected = Request.Form("YEAR") &"/" &Request.Form("MONTH") &"/1"
	datetime   = dtSelected
	
	'******************************************************************
	'    Work out the number of days in the month that was submitted
	'******************************************************************
	datetime   = dateadd("d", -datepart("d",datetime)+1,datetime)
	datetime   = dateadd("m", 1, datetime)
	datetime   = dateadd("d", -1, datetime)
	intDays    = datepart("d",datetime)



	'****************************************************************
	'                      CALENDAR HEADER
	'****************************************************************

	%>

<CENTER>	
<TABLE>
<TR><TD CLASS=MAIN>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 >
<TR><TD COLSPAN=7 CLASS=HEAD ALIGN=CENTER><FONT COLOR=WHITE><%=MONTHNAME(Request.Form("MONTH")) &" - " &Request.Form("YEAR")%></TD></TR><TR>
  <%
  For I = 1 to 7 
  Response.Write "<TD ALIGN=CENTER CLASS=HEAD><FONT COLOR=WHITE>" &WeekdayName(I,blnAbbr) &"</TD>"
  Next
  %>
</TR><TR>

<%     

'****************************************************************
'Use intCounter to see how many cells have been written out
'****************************************************************
 
    For  I = 1 to Weekday(dtSelected)-1
          Response.Write "<TD VALIGN=TOP ALIGN=RIGHT WIDTH="&CELL_WIDTH &" HEIGHT="&CELL_HEIGHT&">&nbsp;</TD>"
          intCounter = intCounter + 1
       Next   

        If intCounter-1  >= 7 then 
           intCounter = 0 
           Response.Write "</TR><TR>"
        End if

        For I = 1 to intDays
         
         blnISPublic = FALSE 
         intCounter = intCounter + 1
         newdate = Request.Form("YEAR") &"/" &Request.Form("MONTH") &"/" &I
         
          For J = 0 to UBOUND(arrPublic)
            If CDate(NewDate) = CDate(arrPublic(J)) then 
              blnIsPublic = TRUE
            End if
          Next
          
          If CDate(NewDate) = DATE then 
            strCLASS = "TODAY"
            elseif blnIsPublic then 
              strCLASS = "PUBLIC"
            else  
             strCLASS = "DAYS"
          end if  
          
	'****************************************************************
	'                    Write out data
	'****************************************************************
          
          For J = 0 to UBOUND(aMemoDate)
            If CDate(aMemoDate(J)) = CDate(NEWDATE) then 
              strIMG = "<CENTER><IMG SRC='./images/memo.gif' ALT='"&aMemoInfo(J) &"'>"                
               exit For
             else
              strIMG = "&nbsp;"
            End if   
         Next
            
	  if intCounter mod 7 <> 0 then             
	   Response.Write "<TD CLASS=" &strCLASS &" VALIGN=TOP ALIGN=RIGHT WIDTH="&CELL_WIDTH &" HEIGHT="&CELL_HEIGHT &" BGCOLOR=" &strBGCOLOR &" OnClick='ALERT'>" &I &strIMG &"</TD>"
	  else
	   Response.Write "<TD CLASS=" &strCLASS &" VALIGN=TOP ALIGN=RIGHT WIDTH="&CELL_WIDTH &" HEIGHT="&CELL_HEIGHT &" BGCOLOR=" &strBGCOLOR &">" &I &strIMG &"</TD></TR><TR>"
	  end if   
	 	  
	Next  
%>
</TR>
</TABLE></TD></TR></TABLE>
<CENTER>
<% end if%> 
</BODY>
</HTML>
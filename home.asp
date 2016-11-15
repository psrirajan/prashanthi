<html>

<head>
<meta name="google-site-verification" content="2ED4shRRteQ4yB6I54ZL3wR5SLLFuDklNohnx5NbaI8" />
<title>Home</title>
<link rel="stylesheet" type="text/css" href="tswtabs.css" />

	<title>Whats New To YOU</title>
</head>
<body>


<table align="center" >
<tr>
<td align="center" bgcolor="lightyellow">
<DIV class=clock-pos>
<SPAN class=clock-opm id=clock><FONT style="bold">00:00:00</FONT></SPAN>
        </DIV>
</td>

</tr>
</table>	
<table>

</table>	
<SCRIPT language=JavaScript1.2>
<!--

	function getTimeString() {
	  time = getFormat("H") + ":" + getFormat("m") + ":" + getFormat("s");
	  return time;
	}

	function setDigits(num, howmany) {
	  str = num.toString();
	  if (str.length >= howmany) return str;
	  for (i = str.length; i < howmany; i++) str = "0" + str;
	  return str;
	}

	function getFormat(ch) {
	  now = new Date();
	  if(ch=='H') { 
	  if(now.getHours()>=13) {return now.getHours()-12; }
	  else if(now.getHours()==0) {return now.getHours()+12; }
	  else return now.getHours(); }
	  else if(ch=='m') { return setDigits(now.getMinutes(),2); }
	  else if(ch=='s') { return setDigits(now.getSeconds(),2); }
	  else return "";
	}
function changeText(id, str) {
  if(document.layers) { // if Navigator 4.0+ 
    with(document[id].document) {
      open();
      write(str);
      close();
    }
  } else { // Internet Explorer 4.0+ 
    document.all[id].innerHTML = str;
  }
}

function startClock() {
  ClockID = setInterval("changeText('clock',getTimeString())",500);
}

//if (document.all||document.layers) startClock();
if (document.all) startClock();

//-->
</SCRIPT>
<% if Session("VarName") = "" then %>
<div id="tswcsstabs">
<ul>
<li><a href="login.asp" target="content">Login</a></li>
</ul>
<ul>
<li><a href="reg.asp" target="content">Register</a></li>
</ul>

</div>
<% else 
If Session("Identity") =  2 then %>
<div id="tswcsstabs">
<ul>
<li><a href="links.asp" target="content">Links</a></li>
<li><a href="evtcal.asp" target="content">Calendar</a></li>
<li><a href="appoint.asp" target="content">Appointment</a></li>
<li><a href="logout.asp" target="content">Log Out</a></li>
</ul>
</div>
<% End If %>
<% end if %>

</body>


</html>
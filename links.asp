<html>
<head>
<link rel="stylesheet" type="text/css" href="tswtabs.css" />
<style>
a:link{ text-decoration: none;
	color: green;
}
	
a:visited{ text-decoration: none;
	color: orange;
}

a:hover{ text-decoration: none;
	color: blue;
	font-weight: bolder;
	letter-spacing: 2px;
}


</style>
<link rel="stylesheet" type="text/css" href="web.css" />

  <meta   http-equir="content-type"     
    
  content="text/html";charset="gb2312">   
  <script   language="Javascript">   
 
function openDoc(filename,target){

//to ensure no global variable conflict with other script
var strWinHandle = target + "objDocWin"; 

//just focus to the corresponding window if it is already open
if (window[strWinHandle] && !window[strWinHandle].closed){
window[strWinHandle].focus();
return false;
}

 function   doword(str)   
  {   
  var   WordApp=new   ActiveXObject("Word.Application");   
  WordApp.Application.Visible=true;   
  var   Doc=WordApp.Documents.Add("docs/prashanthi.docx",true);   
  }   
 function open()
{
 window.open("http://prashanthi1.brinkster.net/prashanthi1/docs/prashanthi.docx")
}
  </script>  
</head>
<body>
<div id="tswcsstabs">
<ul>
<li><a href="links.asp" target="content">Links</a></li>
<li><a href="evtcal.asp" target="content">Calendar</a></li>
<li><a href="appoint.asp" target="content">Appointment</a></li>
<li><a href="logout.asp" target="content">Log Out</a></li>
</ul>
</div>
</p></p>

<table border=2 bglor=green>
<tr>
<td>
 Documents
</td>
</tr>
</table>

<div>
<ul>
<li>
<a href="http://prashanthi1.brinkster.net/docs/prashanthi.docx" type="application/ms-word" title="Word Document XXX
kb">Vijaya health center</a>
  
</li>
<li>
<iframe src="http://docs.google.com/viewer?url=http%3A%2F%2Fprashanthi1.brinkster.net%2Fprashanthi.docx&embedded=true" width="600" height="780" style="border: none;"></iframe>
  


</li>

</div>



--------------------------------------



</body>
</html>
<% Option Explicit %>
<html>
<!-- ******************************************************* -->
<!-- * Web-based private chat room project
<!-- * By Jeff Brazill. If you have questions, PLEASE CHECK
<!-- * MY FAQ PAGE FIRST! It is located at:
<!-- * http://www15.brinkster.com/vbwizard/chatfaq.htm
<!-- * IF!... you don't find your answer there, you may contact me on my
<!-- * contact page at: http://www15.brinkster.com/vbwizard/contact.asp
<!-- * First published 06/19/2002 @ www.brinkster.com
<!-- * Updated 09/25/2002: strSQLDel fixed to handle no messages
<!-- * Updated 12/07/2002: Fixed improper message display - RS("msg")
<!-- * THIS FILE: default.asp
<!-- ******************************************************* -->
<head><title>Private Chat</title></head>
<body><center><h1>Private Chat</h1><br />

<!-- --------------------\/ Message Form \/---------------------- -->
<form method="POST" action="chat.asp"><table><tr><td>
<%
If Request.Form("ToName")<>"" Then
    Session("ToName") = Request.Form("ToName")
End If
Response.Write("<b>To: </b><input type='text' name='ToName' " & _
    "value='" & Session("ToName") & "' size=13 maxlength=20>")
%>
      
<a href='houseclean.asp'>Erase All Messages</a></td></tr><tr><td>
<textarea name="comments" wrap="physical" rows=6 cols=60 value="">
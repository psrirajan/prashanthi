<%

posted = request.form ("submit")

if posted = "Submit" then

 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' Customize the following 5 lines with your own information. ''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

vtoaddress = "Email@Domain.com" ' Change this to the email address you will be receiving your notices.

vmailhost = "mymail.brinkster.com"  ' Change this to mail.yourDomain or leave as is.

vfromaddress = "You@Yourdomain.com" ' Change this to the email address you will use to send and authenticate with.

vfrompwd = "password" ' Change this to the above email addresses password.

vsubject = "ASP Contact Form" 'Change this to your own email message subject.

 

'''''''''''''''''''''''''''''''''''''''''''

'' DO NOT CHANGE ANYTHING PAST THIS LINE ''

'''''''''''''''''''''''''''''''''''''''''''

vfromname = request.form ("TName")

vbody = request.form ("TBody")

vrplyto = request.form ("TEmail")

vmsgbody = vfromname &"<br>"& vrplyto &"<br>"& vbody

 

Set objEmail = Server.CreateObject("Persits.MailSender")

 

objEmail.Username = vfromaddress

objEmail.Password = vfrompwd

objEmail.Host = vmailhost

objEmail.From = vfromaddress

objEmail.AddAddress vtoaddress

objEmail.Subject = vsubject

objEmail.Body = vmsgbody

objEmail.IsHTML = True

objEmail.Send



vErr = Err.Description
if vErr <> "" then
    response.write vErr & "<br><br>There was an error on this page."
else
    response.write "Thank you, your message has been sent."
End If
 

Set objEmail = Nothing

 

response.write "Thank you, your message has been sent."

end if

%>

 

<html>
<head>
<link rel=StyleSheet href="web.css" type="text/css">
</head>
<body>

<form name="SendEmail01" method="post">

<table border=0>

<tr>

            <td>Name:</td>

            <td><input type="text" name="TName" size="30"></td>

</tr>

<tr>

            <td>Email:</td>

            <td><input type="text" name="TEmail" size="30"></td>

</tr>

<tr>

            <td>Body:</td>

            <td><textarea rows="4" name="TBody" cols="30"></textarea></td>

</tr>

<tr>

            <td><input type="submit" name="Submit" value="Submit"></td>

</tr>
</table>
</form>

</body></html>
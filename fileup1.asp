<%
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''' Replace the word "upload" in the line below '''
''' with your own special password				'''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
specialpassword = "rambha1"










if request.form("pwdlogin") = "login" then
		pwd = request.form("pass")
	else
		pwd = ""
end if

if pwd = specialpassword then
	session("username")= "fileuploader"
	Session.Timeout=3
	response.redirect "fileup.asp"
Else
	session("username") = "empty"
	Session.Timeout=1
	response.redirect "fileup.asp"
end if
%>
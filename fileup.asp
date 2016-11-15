<html>
<%
'Setting the script timeout to 30 seconds.
Server.ScriptTimeout = 30

''''''''''''''''''''''''''''''''''''
''''' Please read this section '''''
''''''''''''''''''''''''''''''''''''
'1) Use this application for file uploading with the ASPUpload component.
'2) You will need to edit the Username, password, path, and maxfilesize settings just below these instructions.
'3) You do NOT need to request write permissions to be set.
	'You are already authenticating with a user that has permissions.
'4) The path folder must exist at the same level this page is running in.
'5) Don't forget to open the fileup1.asp page and set the secret password for the login.
'6) Be sure to edit the list of approved file types for upload below. 
	'This is below the section that says "Edit this next line."
'7) This application WILL overwrite other files with the same name.

brinkusername = "prashanthi1" 	'Change this to your Brinkster User Name
brinkpwd= "rambha1" 		'Change this to your Brinkster Password
path = "upload" 			'Change this to the folder name inside your root directory you want to upload to.
maxfilesize = "8192" 		'Change this to the maximum file size you want to be uploaded. 1024 = 1mb 8192 = 8MB

t = lcase(left(Request.ServerVariables("CONTENT_TYPE"), 19))
if t = "multipart/form-data" then

			Set Upload = Server.CreateObject("Persits.Upload.1")
			Upload.LogonUser "brinkster", ""& brinkusername &"", ""&brinkpwd&""
			Upload.Save


	For Each File in Upload.Files
			fileok = file.filename
			fileok = lcase(right(fileok,4))
			
			select case fileok
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Edit this next 1 line. Add / remove any file types you want allowed '''
''' Add the last 4 characters of the file name. If the file extesion is '''
''' 4 characters already, do not add a "." example :					'''
''' "html" and not ".html"												'''
''' Don't forget to separate each file type with a ","				    '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						case ".gif",".jpg",".png","jpeg", ".php", ".asp", "html", ".htm", ".pdf", ".txt"
								
								if File.Size > maxfilesize then
									response.write file.size &"<br>"
									response.write file.filename &" is larger than " & maxfilesize & " We cannot accept the file upload.<br>"
								else
									savepath = Server.MapPath("\"& path) &"\"& File.FileName
									File.SaveAs savepath
									Response.Write file.filename &" uploaded.<br>"
								end if
						case else
								response.write file.filename &" - File Extension not allowed.<br>"
			end select 
	Next
end if

sub uploadform()
		response.write"" &_
		"<Form Method='post' enctype='multipart/form-data'>" &_
		"<input type='file' size='40' name='file1'><br>" &_
		"<input type='file' size='40' name='file2'><br>" &_
		"<input type='file' size='40' name='file3'><br>" &_
		"<input type='file' size='40' name='file4'><br>" &_
		"<input type=submit value='Upload!'>" &_
		"</form>"

set upload = nothing


''This section will list all of the files in your folder.

response.write "List of files in the "& path & " folder. <br>"
Dim objFSO, objFile, objFolder
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder (Server.MapPath("\"& path) &"\")

response.write "<table border=0>"
For Each objFile in objFolder.Files
Response.Write "<tr><td>"& objFile.Name & "</td><td><a style=text-decoration:none href='"& path &"\"& objfile.name&"' target=_blank>Open File</a><br></td></tr>"
Next
response.write "</table>"
Set objFolder = Nothing
Set objFSO = Nothing		

end sub

sub loginform()
		response.write"" &_
		"<Form method='post' action='fileup1.asp'>" &_
		"<input type=password size=20 name='pass'><BR>" &_
		"<input type=submit value='login' name='pwdlogin'>" &_
		"</form>"
end sub

'if session("username") = "prashanthi1" then
	call uploadform
'else
'	response.write "Please login."
'	call loginform
'end if

%>
</html>
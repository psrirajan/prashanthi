<%
''Current Directory must have writeable permissions for this sample to work.
''It is suggested to NOT run this in the root directory 
'' and NOT set writeable permissions to the root directory of your website.


vqs = request.querystring("Action")
response.write "" &_
"<a href='?action=cf'>Create File</a>&nbsp;&nbsp;&nbsp;" &_
"<a href='?action=cd'>Create Directory</a>&nbsp;&nbsp;&nbsp;" &_
"<a href='?action=df'>Delete File</a>&nbsp;&nbsp;&nbsp;" &_
"<a href='?action=dd'>Delete Directory</a>&nbsp;&nbsp;&nbsp;" &_
"<br><br>"


select case vqs
  Case "cf"
    call createfile
  Case "cd"
    call createdir
  Case "df"
    call delfile
  Case "dd"
    call deldir
End Select


sub createfile()
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  if not fs.FileExists(server.mappath("test.txt")) then
    set fname=fs.CreateTextFile(server.mappath("test.txt"),true)
    fname.WriteLine("Hello World!")
    fname.Close
    response.write "File Created. " & server.mappath("test.txt")&"<br>"
  else
    response.write "File Exists Already. " & server.mappath("test.txt")
  end if
  set fs=nothing
  set fname=nothing
end sub




sub createdir()
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  if fs.FolderExists(server.mappath("test")) then
    response.write "Directory Already Exists - must remove before creation. " & server.mappath("test")&"<br>"
  else
    set f=fs.CreateFolder (server.mappath("test")) 
    response.write "Directory Created. " & server.mappath("test")&"<br>"
  end if  




  set f=nothing
  set fs=nothing
end sub


sub delfile()
  Set fs=Server.CreateObject("Scripting.FileSystemObject")
  if fs.FileExists(server.mappath("test.txt")) then
    fs.DeleteFile(server.mappath("test.txt"))
    response.write "File Delete. " & server.mappath("test.txt")
  else
    response.write "File Does Not Exist. " & server.mappath("test.txt")
  end if
  set fs=nothing
end sub


sub deldir()
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  if fs.FolderExists(server.mappath("test")) then
    fs.DeleteFolder(server.mappath("test"))
    response.write "Directory Delete. " & server.mappath("test")
  end if
  set fs=nothing
end sub
%>
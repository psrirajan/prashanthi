<html>
<head>
<link rel=StyleSheet href="web.css"
type="text/css">
</head>
<body>

<%
Dim objConn, objRS, str
'On Error GoTo DisplayErrorInfo


'Dim oConn, sConnString
'Set oConn = Server.CreateObject("ADODB.Connection")
'sConnString = "DRIVER={Microsoft Access Driver (*.mdb)};" & _ 
'    "DBQ=" & Server.MapPath("\MemberName\db\dbname.mdb") & ";"
'oConn.Open(sConnString)


response.write "Welcome"
Set objConn = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.RecordSet")

str= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" &  Request.ServerVariables("APPL_PHYSICAL_PATH") & "\p\r\a\db\user.mdb;"
response.write str

objConn.open(str)


strQ = "SELECT * from loginfo where username = '" & Request.Form("username")
strQ = strQ & "' and pwd = '" & Request.Form("password") & "'"

response.write strQ

'objRS.Open strQ, objConn

'Sub DisplayErrorInfo()
		 'For Each errorObject In objRS.objConn.Errors
  '  Response.Write "Description: " & errorObject.Description & Chr(10) & Chr(13) & _
 '          "Number:      " & Hex(errorObject.Number)
'  Next

'End Sub
'If (objRS.EOF) then = "" then
'  response.write("Not found")
'else 
' response.write("found")
'end if
'If Request.Form("password") > "" Then
'If objRS.EOF Then
'Response.Write "<font color=#ff0000><Center><b>**Failed Login**</b></Center></font>"
'Else
'Response.Write "<font color=#008000><Center><b>**Successful Login**</b></Center></font>"
'End If
'Else
'Response.Write "<font color=#ff8000><Center><b>**No login attempted**</b></Center></font>"
'End If

'objRS.close
objConn.close
'Set objRS = Nothing
'Set objConn = Nothing
%>

</body>
</html>
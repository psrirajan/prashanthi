<%
Dim objConn, objRS

Set objConn = Server.CreateObject("ADODB.Connection")
conString = "DBQ=\prahanthi1\db\loginfo.mdb"
response.write conString
'objConn.Open "DRIVER={Microsoft Access Driver (*.mdb)};" & conString

'Set objRS = Server.CreateObject("ADODB.RecordSet")
'strQ = "SELECT * from loginfo where username = '" & Request.Form("username")
'strQ = strQ & "' and password = '" & Request.Form("pwd") & "'"

%>
<html>
<head>
<script langauge="vb" runat="server">

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim sConnString As String
        Dim sSQL As String
        Dim oConn As System.Data.OleDb.OleDbConnection
        Dim oComm As System.Data.OleDb.OleDbCommand
        Dim oDR As System.Data.OleDb.OleDbDataReader
        Dim i As Integer
        Dim sColor As String
        Dim sTemp As String
        Dim sMapPath As String

        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="
        '
        ' Jared's sMapPath for PREMIUM Users
        '
        'sMapPath = Left(Server.MapPath("\"), Len(Server.MapPath("\")) - 1)
        'sMapPath = Mid(sMapPath, 1, InStrRev(sMapPath, "\") - 1) & "\prashanthi1\db\Example1.mdb;"
        '
        ' sMapPath for General users.  Replace "markjerde" with your directory name
        '
        sMapPath = Server.MapPath("\prashanthi1\db\Example1.mdb")
        '
        ' ... back to Jared's code
        '
        sConnString = sConnString & sMapPath
        oConn = New System.Data.OleDb.OleDbConnection(sConnString)

        oConn.Open()

        Response.Write("<font face=Verdana size=2>Connection Opened.<br><br>")

        sSQL = "SELECT ProductName, QuantityPerUnit, UnitPrice, UnitsInStock FROM Products"
        oComm = New System.Data.OleDb.OleDbCommand(sSQL, oConn)
        oDR = oComm.ExecuteReader()

        Response.Write("<table border=1 cellpadding=1 cellspacing=1 style='font-family:arial; font-size:10pt;'>")
        Response.Write("<tr bgcolor=black style='color:white;'><td>Product Name</td>")
        Response.Write("<td>Quantity Per Unit</td>")
        Response.Write("<td align=right>Price</td>")
        Response.Write("<td>In Stock</td></tr>")
        
        sColor = "white"
        Do While oDR.Read

            If sColor = "silver" Then
                sColor = "white"
            Else
                sColor = "silver"
            End If

            Response.Write("<tr bgcolor='" & sColor & "'>")
            Response.Write("<td>" & oDR.Item("ProductName") & "</td>")
            Response.Write("<td>" & oDR.Item("QuantityPerUnit") & "</td>")
            Response.Write("<td align=right>$" & oDR.Item("UnitPrice") & "</td>")
            Response.Write("<td align=right>" & oDR.Item("UnitsInStock") & "</td></tr>")

        Loop

        Response.Write("</table><br><br>")

        oDR.Close()
        oConn.Close()

        Response.Write("Connection Closed.<br><br>")
        Response.Write(Now() & "</font><br><br>")


    End Sub

</script>
</head>

<body>
</body>
</html>
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Globalization
Module Module2

    Sub Main()
        Dim oTest As Object
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim server, conString, exePath, passWord As String
        Dim fld As String = ""
        Dim valu As String = ""
        While xmlReader.Read
            Select Case xmlReader.NodeType
                Case XmlNodeType.Element
                    fld = xmlReader.Name
                Case XmlNodeType.Text
                    valu = xmlReader.Value
            End Select
            If fld = "SERVER" Then server = valu
            If fld = "EXEPATH" Then exePath = valu
            If fld = "PD" Then passWord = valu
        End While
        ''conString = "Server=" & server & ";Initial Catalog=RCClient;User Id=sa;Password=" & passWord & ""
        conString = "Server=" & server & ";Initial Catalog=RCClient;Integrated Security=True" & ""
        Dim dbase As String = ""
        Dim client As String = ""
        Dim sqlUserId As String = ""
        Dim sqlPassword As String = ""
        Dim xmlPath, ClientConString, sql, marketing, errorLog As String
        Dim con, con2 As SqlConnection
        Dim cmd As SqlCommand
        Dim rdr, rdr2 As SqlDataReader
        Dim cnt As Int32 = 0
        Dim BeginDate, EndDate, todaysDate As Date
        Dim tbl As New DataTable
        Dim row As DataRow
        tbl.Columns.Add("Client")
        tbl.Columns.Add("Server")
        tbl.Columns.Add("dBase")
        tbl.Columns.Add("xmlPath")
        tbl.Columns.Add("sqlUserId")
        tbl.Columns.Add("sqlPassword")
        tbl.Columns.Add("BeginDate")
        tbl.Columns.Add("EndDate")
        tbl.Columns.Add("errorLog")
        tbl.Columns.Add("Marketing")
        con = New SqlConnection(conString)
        con.Open()
        Console.WriteLine("RCClient is now open")
        sql = "SELECT Client_Id, [Database] AS dbase, Server, SQLUserId, SQLPassword, XMLs AS Path, errorLog, Marketing FROM CLIENT_MASTER " & _
            "WHERE Status = 'Active' ORDER BY Client_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            client = rdr("Client_Id")
            dbase = rdr("dbase")
            server = rdr("Server")
            sqlUserId = rdr("SQLUserId")
            sqlPassword = rdr("SQLPassword")
            xmlPath = rdr("Path")
            oTest = rdr("errorLog")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then errorLog = oTest Else errorLog = ""
            oTest = rdr("Marketing")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then marketing = CStr(oTest) Else marketing = ""
            ''ClientConString = "Server=" & server & ";Initial Catalog=" & dbase & ";User Id=" & sqlUserId & ";Password=" & sqlPassword & ""
            ClientConString = "Server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True" & ""
            con2 = New SqlConnection(ClientConString)
            con2.Open()
            todaysDate = CDate(Date.Today)
            sql = "SELECT sDate, eDate FROM Calendar WHERE '" & todaysDate & "' BETWEEN sDate AND eDate AND PrdWk > 0"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                BeginDate = rdr2("sDate")
                EndDate = rdr2("eDate")
            End While
            con2.Close()
            con2.Dispose()
            row = tbl.NewRow
            row("Client") = client
            row("Server") = server
            row("dBase") = dbase
            row("xmlPath") = xmlPath
            row("sqlUserId") = sqlUserId
            row("sqlPassword") = sqlPassword
            row("BeginDate") = CDate(BeginDate)
            row("EndDate") = CDate(EndDate)
            row("errorLog") = errorLog
            row("Marketing") = marketing
            tbl.Rows.Add(row)
        End While
        con.Close()
        con.Dispose()
        For Each row In tbl.Rows
            client = row("Client")
            server = row("Server")
            dbase = row("dBase")
            xmlPath = row("xmlPath")
            sqlUserId = row("sqlUserId")
            sqlPassword = row("sqlPassword")
            BeginDate = row("BeginDate")
            EndDate = row("EndDate")
            errorLog = row("errorLog")
            marketing = row("Marketing")
            ''If client = "PARGIF" Then
            Console.WriteLine("Calling XML_Update for " & client & " exepath=" & exePath)
            Dim p As New ProcessStartInfo
            p.FileName = exePath & "\XMLUpdate.exe"
            p.Arguments = client & ";" & server & ";" & dbase & ";" & xmlPath & ";" & sqlUserId & ";" &
                sqlPassword & ";" & BeginDate & ";" & EndDate & ";" & errorLog & ";" & marketing
            Console.WriteLine("Calling Client_XML_Update " & p.Arguments)
            p.UseShellExecute = True
            p.WindowStyle = ProcessWindowStyle.Normal
            Dim proc As Process = Process.Start(p)
            Threading.Thread.Sleep(6000)
            ''End If

        Next
    End Sub
End Module

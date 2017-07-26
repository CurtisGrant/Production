Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Module Module1
    Private con, rccon As SqlConnection
    Private cmd As SqlCommand
    Private rdr As SqlDataReader
    Private client, conString, sql, err, errorPath, xmlPath, thePath, fileName As String
    Private oTest, oTest2 As Object
    Private thisEdate As Date
    Private infoReader As System.IO.FileInfo

    Sub Main()
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim rcServer, rcConString, rcExePath, rcPassword As String
        Dim server, dbase, userid, password As String
        Dim program, process As String
        Dim fld As String = ""
        Dim valu As String = ""
        Dim today As Date = Date.now
        Dim createDate As DateTime
        Dim fileSize As Integer
        Dim yesterday As Date = DateAdd(DateInterval.Day, -1, Date.Today)
        Dim DOW As Integer = today.DayOfWeek + 1           ' bump it up by 1 so we can check the previous day's processes
        Dim thisDay As Integer
        Dim tbl As New DataTable
        Dim row, foundRow As DataRow
        While xmlReader.Read
            Select Case xmlReader.NodeType
                Case XmlNodeType.Element
                    fld = xmlReader.Name
                Case XmlNodeType.Text
                    valu = xmlReader.Value
                Case XmlNodeType.EndElement
                    'Console.WriteLine("</" & xmlReader.Name)
            End Select
            If fld = "SERVER" Then rcServer = valu
            If fld = "EXEPATH" Then rcExePath = valu
            If fld = "PD" Then rcPassword = valu
        End While
        Dim clientTbl As New DataTable
        clientTbl.Columns.Add("Client_Id", GetType(System.String))
        clientTbl.Columns.Add("Server", GetType(System.String))
        clientTbl.Columns.Add("Database", GetType(System.String))
        clientTbl.Columns.Add("UserId", GetType(System.String))
        clientTbl.Columns.Add("Password", GetType(System.String))
        clientTbl.Columns.Add("errorLog", GetType(System.String))
        clientTbl.Columns.Add("XMLPath", GetType(System.String))
        clientTbl.Columns.Add("Item4Cast", GetType(System.String))
        clientTbl.Columns.Add("Marketing", GetType(System.String))
        clientTbl.Columns.Add("SalesPlan", GetType(System.String))
        rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
        rccon = New SqlConnection(rcConString)
        rccon.Open()
        sql = "SELECT Client_Id, [Database], Server, SQLUserId, SQLPassword, XMLs, errorLog, Item4Cast, Marketing, SalesPlan " & _
            "FROM CLIENT_MASTER WHERE Status = 'Active' ORDER BY Client_Id"
        cmd = New SqlCommand(sql, rccon)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = clientTbl.NewRow
            row("Client_Id") = rdr("Client_Id")
            row("Server") = rdr("Server")
            row("Database") = rdr("Database")
            row("UserId") = rdr("SQLUserId")
            row("Password") = rdr("SQLPassword")
            row("XMLPath") = rdr("XMLs")
            row("errorLog") = rdr("errorLog")
            row("Item4Cast") = rdr("Item4Cast")
            row("Marketing") = rdr("Marketing")
            row("SalesPlan") = rdr("SalesPlan")
            clientTbl.Rows.Add(row)
        End While
        rccon.Close()

        For Each clientRow As DataRow In clientTbl.Rows
            client = clientRow("Client_Id")
            server = clientRow("Server")
            dbase = clientRow("Database")
            userid = clientRow("UserId")
            password = clientRow("Password")
            xmlPath = clientRow("XMLPath")
            errorPath = clientRow("errorLog")
            conString = "Server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
            tbl = New DataTable
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Process"
            tbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = tbl.Columns("Process")
            tbl.PrimaryKey = PrimaryKey
            tbl.Columns.Add("Program")
            tbl.Columns.Add("DOW")
            tbl.Columns.Add("XMLFile")
            tbl.Columns.Add("OK")
            Dim xmlTbl As DataTable = New DataTable
            Dim colmn As New DataColumn
            colmn.DataType = System.Type.GetType("System.String")
            colmn.ColumnName = "xmlFile"
            xmlTbl.Columns.Add(colmn)
            Dim primKey(1) As DataColumn
            primKey(0) = xmlTbl.Columns("xmlFile")
            xmlTbl.PrimaryKey = primKey
            '
            '                Check for an error log file in the xml folder
            '
            Dim errLog As String = xmlPath & "\errlog.txt"
            If System.IO.File.Exists(errLog) Then
                err = "Found an error log file for " & client & " in " & xmlPath
                Call Write_Error(err)
            End If
            '
            '                 Check for an error log in the  ERRORS folder
            '
            errLog = errorPath & "\errLog.txt"
            If System.IO.File.Exists(errLog) Then
                err = "Found an error log file for " & client & " in " & errorPath
                Call Write_Error(err)
            End If

            rccon.Open()
            sql = "SELECT Program, Module, Process, DOW, XMLFile FROM Client_Processes WHERE Client_Id = '" & client & "' "

            cmd = New SqlCommand(sql, rccon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = tbl.NewRow
                row("Process") = rdr("Process")
                row("Program") = rdr("Program")
                row("DOW") = rdr("DOW")
                oTest = rdr("XMLFile")

                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    thePath = xmlPath & "\" & rdr("XMLFile")
                    If System.IO.File.Exists(thePath) Then
                        createDate = File.GetLastWriteTime(thePath)
                        infoReader = My.Computer.FileSystem.GetFileInfo(thePath)
                        fileSize = infoReader.Length
                        If createDate < yesterday Then
                            err = rdr("XMLFile") & " Not as expected - Created on " & createDate
                            Call Write_Error(err)
                        End If
                        If fileSize < 300 Then
                            err = rdr("XMLFile") & " May not contain enough records - file size is only " & fileSize & " bytes!"
                            Call Write_Error(err)
                        End If
                    Else
                        err = "Did not find " & rdr("XMLFile") & " in the XML folder"
                        Call Write_Error(err)
                    End If
                End If
                tbl.Rows.Add(row)
            End While
            rccon.Close()

            con = New SqlConnection(conString)
            con.Open()
            sql = "SELECT eDate FROM Calendar WHERE Convert(date,GetDate()) BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisEdate = rdr("eDate")
            End While
            con.Close()

            con.Open()
            sql = "SELECT Program, Module, Process FROM Process_Log WHERE Date BETWEEN '" & yesterday & "' AND '" & today & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                process = rdr("process")
                program = rdr("Program")
                foundRow = tbl.Rows.Find(process)
                If Not IsNothing(foundRow) Then
                    foundRow("OK") = "OK"
                End If
            End While
            con.Close()

            Dim allIsGood As Boolean = True
            For Each row In tbl.Rows
                err = row("Program") & " " & row("Process") & " " & row("DOW") & " " & row("OK")
                oTest = row("DOW")
                If IsDBNull(oTest) Or IsNothing(oTest) Then
                    oTest2 = row("OK")
                    If IsDBNull(oTest2) Then
                        allIsGood = False
                        Call Write_Error(err)
                    End If
                Else
                    thisDay = CInt(oTest) + 1
                    If thisDay = DOW Then                    ' a process is supposed to run this day of the week
                        oTest2 = row("OK")
                        If IsDBNull(oTest2) Then
                            Call Write_Error(err)
                            allIsGood = False
                        End If
                    End If
                End If
            Next

            con.Open()
            sql = "SELECT ISNULL(SUM(End_OH * Retail),0) AS OnHand FROM Item_Inv WHERE eDate = '" & thisEdate & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("OnHand")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If rdr("OnHand") < 2000000 Then
                        err = "Inventory may be off. Current On Hand at retail is only " & Format(CDec(rdr("OnHand")), "$###,###,###.00")
                        Call Write_Error(err)
                    End If
                End If
            End While
            con.Close()
        Next
        
    End Sub

    Private Sub Write_Error(ByVal err As String)
        If My.Computer.FileSystem.FileExists(errorPath & "\failure.txt") Then
            Dim fs1 As FileStream = New FileStream(errorPath & "\failure.txt", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.WriteLine(client & " " & err)
            s1.Close()
            fs1.Close()
        Else
            Using sw As StreamWriter = File.CreateText(errorPath & "\failure.txt")
                sw.WriteLine(err)
            End Using
        End If
    End Sub

End Module

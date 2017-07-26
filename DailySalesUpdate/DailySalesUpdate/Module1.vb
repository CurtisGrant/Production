Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Module Module1
    Private con, con2, rcCon As SqlConnection
    Private cmd As SqlCommand
    Private rdr, rcrdr As SqlDataReader
    Public rcErrorPath, xmlFileToProcess As String
    Private client, sql As String
    Private oTest As Object
    Private tbl As DataTable
    Sub Main()
        Try
            Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
            Dim rcServer, rcConString, rcExePath, rcPassword As String
            Dim server, dbase, user, pw As String
            Dim fld As String = ""
            Dim valu As String = ""
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
                If fld = "ERRORPATH" Then rcErrorPath = valu
                If fld = "EXEPATH" Then rcExePath = valu
                If fld = "PD" Then rcPassword = valu
            End While
            rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
            rcCon = New SqlConnection(rcConString)
            rcCon.Open()
            sql = "SELECT Client_Id, Server, [Database], SQLUserID, SQLPassword, XMLs FROM Client_Master WHERE Status = 'Active' "
            cmd = New SqlCommand(sql, rcCon)
            rcrdr = cmd.ExecuteReader
            While rcrdr.Read
                server = rcrdr("Server")
                dbase = rcrdr("Database")
                user = rcrdr("SQLUserId")
                pw = rcrdr("SQLPassword")
                Dim conString As String = "Server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
                con = New SqlConnection(conString)
                con2 = New SqlConnection(conString)
                Dim xmls As String = rcrdr("XMLs")
                ''  Console.WriteLine("Processing " & rdr("Client_Id")) 
                Dim okToProcess As Boolean = False
                xmlFileToProcess = xmls & "\DailySales.xml"

                con.Open()
                sql = "SELECT Value FROM Controls WHERE ID = 'OptionalSoftware' AND Parameter = 'SalesChart'"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    If rdr("Value") = "YES" Then okToProcess = True
                End While
                con.Close()

                If okToProcess AndAlso System.IO.File.Exists(xmlFileToProcess) Then
                    Call Process_Sales(xmls)
                End If
            End While
            rcCon.Close()
        Catch ex As Exception
            If rcCon.State = ConnectionState.Open Then rcCon.Close()
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim theMessage As String = ex.Message
            Dim el As New DailySalesUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "DailySalesUpdate, Main")
        End Try
    End Sub
    Private Sub Process_Sales(ByVal XMLPath As String)
        Try
            Dim stopWatch As New Stopwatch
            stopWatch.Start()
            tbl = New DataTable
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Dept"
            tbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = tbl.Columns("Dept")
            tbl.PrimaryKey = PrimaryKey
            tbl.Columns.Add("Cost", GetType(System.Decimal))
            tbl.Columns.Add("Sales", GetType(System.Decimal))
            tbl.Columns.Add("Tickets", GetType(System.Decimal))
            tbl.Columns.Add("Qty", GetType(System.Decimal))
            tbl.Columns.Add("LastUpdate", GetType(System.DateTime))
            Dim tktTbl As New DataTable
            column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Tkt"
            tktTbl.Columns.Add(column)
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = tktTbl.Columns("Tkt")
            tktTbl.PrimaryKey = PrimaryKey2
            Dim dept As String
            Dim cost, retail, qty As Decimal
            Dim ticket As String
            Dim prevTicket As String = ""
            Dim tickets As Integer
            Dim ttickets As Integer = 0
            Dim dte, lastUpdate As DateTime
            Dim tsales As Decimal = 0
            Dim tcost As Decimal = 0
            Dim tqty As Integer = 0
            Dim tplan As Decimal = 0
            Dim yr, prd, wk As Integer
            Dim today As Date = Date.Today
            Dim thisDOW As Integer = today.DayOfWeek + 1
            If thisDOW = 0 Then Exit Sub

            con.Open()
            sql = "SELECT ID AS Dept FROM Departments ORDER BY Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            Dim row As DataRow
            While rdr.Read
                row = tbl.NewRow
                row("Dept") = rdr("Dept")
                row("Cost") = 0
                row("Sales") = 0
                row("Tickets") = 0
                row("Qty") = 0
                tbl.Rows.Add(row)
            End While
            con.Close()
            con.Open()
            sql = "SELECT Year_Id, Prd_Id, PrdWk FROM Calendar WHERE '" & today & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                yr = rdr("Year_Id")
                prd = rdr("Prd_Id")
                wk = rdr("PrdWk")
            End While
            con.Close()

            con.Open()
            sql = "DELETE FROM DaySales"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim xmlFile As XmlReader
            ''Dim Path As String = XMLPath & "\DailySales.xml"
            Dim testFile As New IO.FileInfo(xmlFileToProcess)
            Dim fileCreated As DateTime = File.GetCreationTime(xmlFileToProcess)
            Dim siz As Long = testFile.Length
            If siz < 600 Then
                Dim theMessage As String = "DailySales.xml is empty"
                Dim el As New DailySalesUpdate.ErrorLogger
                el.WriteToErrorLog(theMessage, "", "UpdateSalesTime,Process_UnpostedSales")
                Exit Sub
            End If
            Dim foundRow, trow, tkrow As DataRow
            xmlFile = Xml.XmlReader.Create(xmlFileToProcess, New XmlReaderSettings())
            ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                dt.DefaultView.Sort = "TKT_NO"
                For Each row In dt.Rows
                    oTest = row("SKU")
                    If Not IsDBNull(oTest) AndAlso oTest <> "TOTALS" Then
                        oTest = row("TRANS_DATE")
                        If IsDBNull(oTest) Then dte = "1900-01-01" Else dte = oTest
                        dte = dte.ToShortDateString()
                        If dte = today Then
                            oTest = row("QTY")
                            If IsDBNull(oTest) Then qty = 0 Else qty = CDec(row("QTY"))
                            tqty += qty
                            oTest = row("RETAIL")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then retail = CDec(oTest) Else retail = 0
                            tsales += qty * retail
                            oTest = row("COST")
                            If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                            tcost += qty * cost
                            oTest = row("DEPT")
                            If IsDBNull(oTest) Then dept = "UNKNOWN" Else dept = CStr(oTest)
                            oTest = row("TKT_NO")
                            If IsDBNull(oTest) Then ticket = "UNKNOWN" Else ticket = CStr(oTest)
                            foundRow = tbl.Rows.Find(dept)
                            If Not IsNothing(foundRow) Then
                                foundRow("Sales") += qty * retail
                                foundRow("Cost") += qty * cost
                                foundRow("Qty") += qty
                                foundRow("LastUpdate") = row("EXTRACT_DATE")
                            Else
                                trow = tbl.NewRow
                                trow("Dept") = dept
                                trow("Sales") = qty * retail
                                trow("Cost") = qty * cost
                                trow("Qty") = qty
                                trow("Tickets") = 1
                                ''oTest = row("EXTRACT_DATE")
                                ''If Not IsDBNull(oTest) Then trow("LastUpdate") = CDate(oTest)
                                trow("LastUpdate") = Date.Now
                                tickets = 1
                                tbl.Rows.Add(trow)
                            End If
                            foundRow = tktTbl.Rows.Find(ticket)
                            If qty > 0 Then
                                If IsNothing(foundRow) Then
                                    tkrow = tktTbl.NewRow
                                    tkrow(0) = ticket
                                    tktTbl.Rows.Add(tkrow)
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            tickets = tktTbl.Rows.Count

            Dim dataView As New DataView(tbl)
            dataView.Sort = "Dept ASC"
            Dim sortedTBL As DataTable = dataView.ToTable()

            con.Open()
            sql = "DELETE FROM DaySales"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            Dim plan, sales As Integer
            For Each row In sortedTBL.Rows
                dept = row("Dept")
                plan = 0
                tplan += plan
                sales = row("Sales")
                cost = row("Cost")
                oTest = row("LastUpdate")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lastUpdate = oTest
                tickets = tickets
                ttickets = tickets
                qty = row("Qty")
                sql = "INSERT INTO DaySales (Dept, [Plan], Sales, Cost, Tickets, Qty, LastUpdate) " & _
                    "SELECT '" & dept & "'," & plan & "," & sales & "," & cost & "," & tickets & "," & qty & ",'" & lastUpdate & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                '' MsgBox(row("Dept") & " Sales " & row("Sales") & " Plan " & row("Plan"))
            Next
            sql = "INSERT INTO DaySales (Dept, [Plan], Sales, Cost, Tickets, Qty, LastUpdate) " & _
                "SELECT ' ALL'," & tplan & "," & tsales & "," & tcost & "," & ttickets & "," & tqty & ",'" & lastUpdate & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            Dim rightnow As Date = Date.Today
            ''cmd = New SqlCommand("dbo.sp_Sales", con)
            ''cmd.CommandType = CommandType.StoredProcedure
            ''cmd.Parameters.Add("@thisDate", SqlDbType.DateTime).Value = rightnow
            ''cmd.ExecuteNonQuery()
            con.Close()

            stopWatch.Stop()
            Dim ts As TimeSpan = stopWatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim ms As String = "Last File Date: " & fileCreated
            con.Open()
            sql = "INSERT INTO Process_Log(Date, Program, Module, Process, Message, Duration) " & _
                "SELECT '" & DateAndTime.Now & "', 'DailySalesUpdate',1,'Process_Sales','" & ms & "','" & et & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            '
            '
            ''  System.IO.File.Delete(xmlFileToProcess)
            '
            '
        Catch ex As Exception

            If con.State = ConnectionState.Open Then con.Close()
            Dim theMessage As String = ex.Message
            Dim el As New DailySalesUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "UpdateSalesTime,Process_UnpostedSales")
        End Try
    End Sub
End Module

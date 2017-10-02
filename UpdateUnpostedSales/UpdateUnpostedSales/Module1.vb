Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Public errorPath As String
    Private thePath, xmlPath, conString, rcConString, sql As String
    Private rcCon, con As SqlConnection
    Private cmd As SqlCommand
    Private xmlReader As XmlTextReader
    Private xmlWriter As XmlTextWriter
    Private oTest As Object
    Sub Main()
        Try
            Dim fileSZ As Integer = 0
            Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
            Dim rcServer, rcExePath, RCUserId, rcPassword, client, server, sqluserid, sqlpassword, dbase As String
            Dim fld As String = ""
            Dim valu As String = ""
            Dim rdr As SqlDataReader
            Dim cnt As Integer = 0
            Dim clientTbl As New DataTable
            clientTbl.Columns.Add("Client", GetType(System.String))
            clientTbl.Columns.Add("Server", GetType(System.String))
            clientTbl.Columns.Add("Database", GetType(System.String))
            clientTbl.Columns.Add("XMLs", GetType(System.String))
            clientTbl.Columns.Add("UserId", GetType(System.String))
            clientTbl.Columns.Add("Password", GetType(System.String))
            clientTbl.Columns.Add("ErrorPath", GetType(System.String))
            Dim row As DataRow

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
                If fld = "USERID" Then RCUserId = valu
                If fld = "PD" Then rcPassword = valu
            End While
            rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
            rcCon = New SqlConnection(rcConString)
            rcCon.Open()
            sql = "SELECT Client_Id, Server, [Database], SQLUserId, SQLPassword, XMLs, ErrorLog FROM Client_Master " &
                "WHERE Status = 'Active'"
            cmd = New SqlCommand(sql, rcCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = clientTbl.NewRow
                row("Client") = rdr("Client_Id")
                row("Server") = rdr("Server")
                row("Database") = rdr("Database")
                row("XMLs") = rdr("XMLs")
                row("UserId") = rdr("SQLUserId")
                row("Password") = rdr("SQLPassword")
                row("ErrorPath") = rdr("ErrorLog")
                clientTbl.Rows.Add(row)
            End While
            rcCon.Close()
            rcCon.Dispose()

            For Each clientRow As DataRow In clientTbl.Rows
                client = clientRow("Client")
                server = clientRow("Server")
                dbase = clientRow("Database")
                sqluserId = clientRow("UserId")
                sqlPassword = clientRow("Password")
                oTest = clientRow("ErrorPath")
                If IsDBNull(oTest) Or IsNothing(oTest) Then
                    Console.WriteLine("No path found for the error log. Add ErrorLog path to Client_Master for " & client & " and try again")
                    Console.ReadLine()
                    Exit Sub
                End If
                errorPath = CStr(oTest)
                xmlPath = clientRow("XMLs")
                conString = "server=" & server & ";Initial Catalog=" & dbase & ";User Id=" & sqluserid & ";Password=" & sqlpassword & ""
                Console.WriteLine("Processing Header")
                thePath = xmlPath & "\UnpostedSalesHDR.xml"
                If FileSize(thePath) > 0 Then Call Process_Header(thePath, conString)
                Console.WriteLine("Processing Detail")
                thePath = xmlPath & "\UnpostedSalesDTL.xml"
                If FileSize(thePath) > 0 Then Call Process_Detail(thePath, conString)
            Next
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If rcCon.State = ConnectionState.Open Then rcCon.Close()
            Dim el As New UpdateUnpostedSales.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Inital set up")
        End Try

    End Sub

    Private Sub Process_Header(ByVal thePath, ByVal constring)
        Try
            con = New SqlConnection(constring)
            con.Open()
            ''sql = "IF OBJECT_ID('tempdb.#t1','U') IS NOT NULL DROP TABLE #t1; " &
            sql = "CREATE TABLE _t1(transId varchar(30), strId varchar(20), locId varchar(20), station varchar(20), drawer varchar(20), " &
                "tDate datetime, cust varchar(20), cust_typ varchar(10), ord_typ varchar(10), tkt varchar(20), rep varchar(20), " &
                "cntry varchar(20), due decimal(18,4), terms varchar(20), paid decimal(20), stat varchar(10), totl decimal(18,4), edte datetime)"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()

            Dim transId, strId, locId, station, drawer, cust, cust_typ, ord_typ, tkt, rep, cntry, terms, stat As String
            Dim totl, due, paid As Decimal
            Dim tdate, edte As DateTime
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim cnt As Integer = 0
            Dim xmlFile As XmlReader
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("TRANS_ID")
                    If IsDBNull(oTest) Then transId = "UNKNOWN" Else transId = CStr(oTest)
                    oTest = row("STR_ID")
                    If IsDBNull(oTest) Then strId = "UNKNOWN" Else strId = CStr(oTest)
                    oTest = row("LOCATION")
                    If IsDBNull(oTest) Then locId = "UNKNOWN" Else locId = CStr(oTest)
                    oTest = row("STATION")
                    If IsDBNull(oTest) Then station = "UNKNOWN" Else station = CStr(oTest)
                    oTest = row("DRAWER")
                    If IsDBNull(oTest) Then drawer = "UNKNOWN" Else drawer = CStr(oTest)
                    oTest = row("TRANS_DATE")
                    If IsDBNull(oTest) Then tdate = "1/1/1900" Else tdate = CDate(oTest)
                    oTest = row("CUST_NO")
                    If IsDBNull(oTest) Then cust = "UNKNOWN" Else cust = CStr(oTest)
                    oTest = row("CUST_TYPE")
                    If IsDBNull(oTest) Then cust_typ = "UNKNOWN" Else cust_typ = CStr(oTest)
                    oTest = row("TRANS_TYPE")
                    If IsDBNull(oTest) Then ord_typ = "U" Else ord_typ = CStr(oTest)
                    oTest = row("TKT_NO")
                    If IsDBNull(oTest) Then tkt = "UNKNOWN" Else tkt = CStr(oTest)
                    oTest = row("SALES_REP")
                    If IsDBNull(oTest) Then rep = "UNKNOWN" Else rep = CStr(oTest)
                    oTest = row("SHIP_COUNTRY")
                    If IsDBNull(oTest) Then cntry = "" Else cntry = CStr(oTest)
                    oTest = row("AMOUNT_DUE")
                    If IsDBNull(oTest) Then due = 0 Else due = CDec(oTest)
                    oTest = row("TERMS")
                    If IsDBNull(oTest) Then terms = "UNKNOWN" Else terms = CStr(oTest)
                    oTest = row("AMT_APPLIED")
                    If IsDBNull(oTest) Then paid = 0 Else paid = CDec(oTest)
                    oTest = row("STATUS")
                    If IsDBNull(oTest) Then stat = "UNKNOWN" Else stat = CStr(oTest)
                    oTest = row("ORDER_TOTAL")
                    If IsDBNull(oTest) Then totl = 0 Else totl = CDec(oTest)
                    oTest = row("EXTRACT_DATE")
                    If IsDBNull(oTest) Then edte = "1/1/1900" Else edte = CDate(oTest)
                    sql = "INSERT INTO _t1(transId, strId, locId, station, drawer, tDate, cust, cust_typ, ord_typ, tkt, rep, " &
                        "cntry, due, terms, paid, stat, totl, edte) " &
                        "SELECT '" & transId & "','" & strId & "','" & locId & "','" & station & "','" & drawer & "','" & tdate & "','" &
                        cust & "','" & cust_typ & "','" & ord_typ & "','" & tkt & "','" & rep & "','" & cntry & "'," & due & ",'" &
                        terms & "'," & paid & ",'" & stat & "'," & totl & ",'" & edte & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                Next
                sql = "MERGE UnpostedSalesHDR AS t USING #t1 AS s ON s.transId = t.TRANS_ID AND s.tDate = t.TRANS_DATE AND s.strId = t.STR_ID " &
                        "WHEN NOT MATCHED BY TARGET " &
                            "THEN INSERT(TRANS_ID, STR_ID, LOC_ID, STATION, DRAWER, TRANS_DATE, CUST_NO, CUST_TYPE, SALES_TYPE, TKT_NO, " &
                                "SALES_REP, SHIP_COUNTRY, TERMS, ORDER_TOTAL, AMOUNT_DUE, AMOUNT_PAID, STATUS, EXTRACT_DATE) " &
                            "VALUES(s.transId, s.StrId, s.LocId, s.station, s.drawer, s.tdate, s.cust, s.cust_typ, s.ord_typ, s.tkt, " &
                            "s.rep, s.cntry, s.terms, s.totl, s.due, s.paid, s.stat, s.edte) " &
                        "WHEN MATCHED " &
                            "THEN UPDATE SET t.AMOUNT_DUE = s.due, t.AMOUNT_PAID = s.paid, t.STATUS = s.stat " &
                        "WHEN NOT MATCHED BY SOURCE THEN DELETE;"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If rcCon.State = ConnectionState.Open Then rcCon.Close()
            Dim el As New UpdateUnpostedSales.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Unposted Sales Header")
        End Try
    End Sub

    Private Sub Process_Detail(ByVal thePath, ByVal constring)
        Try
            con = New SqlConnection(constring)
            con.Open()
            sql = "IF OBJECT_ID('tempdb.#t1','U') IS NOT NULL DROP TABLE #t1; " &
                "CREATE TABLE #t1(transId varchar(30), seqNo integer, strId varchar(20), locId varchar(20), sku varchar(90), " &
                "tdate datetime, qty decimal(18,4), cost decimal(18,4), retail decimal(18,4), mkdn decimal(18,4), reason varchar(30), " &
                "coupon varchar(30), edte datetime, transtype varchar(10), linetype varchar(10))"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()

            Dim transId, strId, locId, sku, coupon, reason, transtype, linetype As String
            Dim qty, cost, retail, mkdn As Decimal
            Dim seqNo As Integer
            Dim tdate, edte As DateTime
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim cnt As Integer = 0
            Dim xmlFile As XmlReader
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("SKU")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        oTest = row("TRANS_ID")
                        If IsDBNull(oTest) Then transId = "UNKNOWN" Else transId = CStr(oTest)
                        oTest = row("SEQ_NO")
                        If IsDBNull(oTest) Then seqNo = 0 Else seqNo = CInt(oTest)
                        oTest = row("STR_ID")
                        If IsDBNull(oTest) Then strId = "UNKNOWN" Else strId = CStr(oTest)
                        oTest = row("LOCATION")
                        If IsDBNull(oTest) Then locId = "UNKNOWN" Else locId = CStr(oTest)
                        oTest = row("SKU")
                        If IsDBNull(oTest) Then sku = "UNKNOWN" Else sku = CStr(oTest)
                        oTest = row("TRANS_DATE")
                        If IsDBNull(oTest) Then tdate = "1/1/1900" Else tdate = CDate(oTest)
                        oTest = row("QTY")
                        If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                        oTest = row("COST")
                        If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                        oTest = row("RETAIL")
                        If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                        oTest = row("MARKDOWN")
                        If IsDBNull(oTest) Then mkdn = 0 Else mkdn = CDec(oTest)
                        oTest = row("MKDN_REASON")
                        If IsDBNull(oTest) Then reason = "UNKNOWN" Else reason = CStr(oTest)
                        oTest = row("COUPON_CODE")
                        If IsDBNull(oTest) Then coupon = "UNKNOWN" Else coupon = CStr(oTest)
                        oTest = row("EXTRACT_DATE")
                        If IsDBNull(oTest) Then edte = "1/1/1900" Else edte = CDate(oTest)
                        oTest = row("TRANS_TYPE")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then transtype = "" Else transtype = CStr(oTest)
                        oTest = row("LINE_TYPE")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then linetype = "" Else linetype = CStr(oTest)
                        sql = "INSERT INTO #t1(transId, seqNo, strId, locId, sku, tdate, qty, cost, retail, mkdn, reason, coupon, edte," & _
                            "transtype, linetype) " &
                            "SELECT '" & transId & "'," & seqNo & ",'" & strId & "','" & locId & "','" & sku & "','" & tdate & "'," &
                            qty & "," & cost & "," & retail & "," & mkdn & ",'" & reason & "','" & coupon & "','" & edte & "','" &
                            transtype & "','" & linetype & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                    End If
                Next
                sql = "MERGE UnpostedSalesDTL AS t USING #t1 AS s ON s.transId = t.TRANS_ID AND s.seqNo = t.SEQ_NO " &
                        "AND s.tDate = t.TRANS_DATE AND s.strId = t.STR_ID AND s.locId = t.LOC_ID AND s.sku = t.SKU " &
                    "WHEN NOT MATCHED BY TARGET THEN " &
                        "INSERT(TRANS_ID, TRANS_TYPE, SEQ_NO, STR_ID, LOC_ID, SKU, TRANS_DATE, QTY, COST, RETAIL, MARKDOWN, MKDN_REASON, " &
                            "COUPON_CODE, EXTRACT_DATE, LINE_TYPE) " &
                        "VALUES (s.transId, s.transtype, s.seqNo, s.strId, s.locId, s.sku, s.tdate, s.qty, s.cost, s.retail, s.mkdn, s.reason, " &
                            "s.coupon, s.edte, s.linetype) " &
                    "WHEN MATCHED THEN " &
                        "UPDATE SET t.QTY = s.qty, t.COST = s.cost, t.RETAIL = s.retail, t.MARKDOWN = s.mkdn " &
                    "WHEN NOT MATCHED BY SOURCE THEN DELETE;"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
                con.Dispose()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If rcCon.State = ConnectionState.Open Then rcCon.Close()
            Dim el As New UpdateUnpostedSales.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Unposted Sales Detail")
        End Try
    End Sub
    Private Function FileSize(ByVal fileName As String)
        Dim length As Integer = 0
        If System.IO.File.Exists(fileName) Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(fileName)
            length = infoReader.Length
            If length = 0 Then
                Dim el As New UpdateUnpostedSales.ErrorLogger
                el.WriteToErrorLog(fileName, "HAS NO RECORDS", "Check File Size")
            End If
            Dim fileDate As DateTime = File.GetLastWriteTime(fileName)
            Dim hoursDiff = DateDiff(DateInterval.Hour, fileDate, Date.Now)
            If hoursDiff > 26 Then
                Dim el As New UpdateUnpostedSales.ErrorLogger
                el.WriteToErrorLog(fileName & " IS OLDER THAN 24 HOURS", "", "Check File Datetime")
            End If
        Else
            If fileName <> xmlPath & "\ErrLog.txt" Then
                Dim el As New UpdateUnpostedSales.ErrorLogger
                el.WriteToErrorLog(fileName & " NOT FOUND", "", "Check File Exists")
            End If
        End If
        Return length
    End Function
End Module

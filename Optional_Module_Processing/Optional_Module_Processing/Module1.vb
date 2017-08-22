Imports System.Data.SqlClient
Imports System.Xml
'
'
'
'
'           Summary files need to by changed to Inv_Summary an Sales Summary
'
'
'
'
'
Module Module1
    Public errorLog, client, errorPath As String
    Public mcon As SqlConnection
    Private con, con2, con3, con4, con5 As SqlConnection
    Private cmd As SqlCommand
    Private rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Private sql, store, dept, buyer, clss As String
    Private stopWatch As Stopwatch
    Private aYearAgo, Begin_Date, End_Date, thisWeek, thisdate, thisEndDate, finalWeek As Date
    Private thisDateTime As DateTime
    Private cnt As Integer
    Private oTest As Object
    Private conString, rcConString, rcExePath, xmlPath, server, dbase, sqluserId, sqlPassword As String
    Private thisYear, lastYear As Integer
    Private thisStore, thisBuyer, thisPlan As String
    Sub Main()
        '---------------------------------------------------------- Sales Plan ------------------------------------------------------------
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim rcServer, rcPassword As String
        stopWatch = New Stopwatch

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
            If fld = "EXEPATH" Then rcExePath = valu
            If fld = "PD" Then rcPassword = valu
        End While
        rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
        ''rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;User Id=sa;Password=" & rcPassword & ""

        Dim clientTbl As New DataTable
        clientTbl.Columns.Add("Client", GetType(System.String))
        clientTbl.Columns.Add("Server", GetType(System.String))
        clientTbl.Columns.Add("Database", GetType(System.String))
        clientTbl.Columns.Add("XMLs", GetType(System.String))
        clientTbl.Columns.Add("UserId", GetType(System.String))
        clientTbl.Columns.Add("Password", GetType(System.String))
        clientTbl.Columns.Add("ErrorPath", GetType(System.String))
        clientTbl.Columns.Add("Item4Cast", GetType(System.String))
        clientTbl.Columns.Add("Marketing", GetType(System.String))
        clientTbl.Columns.Add("SalesPlan", GetType(System.String))
        Dim row As DataRow

        con = New SqlConnection(rcConString)
        mcon = New SqlConnection(rcConString)
        con.Open()
        sql = "SELECT Client_Id, Server, [Database], XMLs, SQLUserId, SQLPassword, ErrorLog, Item4Cast, Marketing, SalesPlan " & _
            "FROM Client_Master WHERE Status = 'Active'"
        cmd = New SqlCommand(sql, con)
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
            errorLog = rdr("ErrorLog")
            row("Item4Cast") = rdr("Item4Cast")
            row("Marketing") = rdr("Marketing")
            row("SalesPlan") = rdr("SalesPlan")
            clientTbl.Rows.Add(row)
        End While
        con.Close()
        con.Dispose()

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
            xmlPath = clientRow("XMLs")
            ''conString = "server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
            conString = "server=" & server & ";Initial Catalog=" & dbase & ";User Id=" & sqluserId & ";Password=" & sqlPassword & ""
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con4 = New SqlConnection(conString)
            con5 = New SqlConnection(conString)
            oTest = clientRow("SalesPlan")
            If Not IsDBNull(oTest) Then
                If clientRow("SalesPlan") = "Y" Then
                    Dim finalWeek As Date
                    con.Open()
                    sql = "SELECT MAX(eDate) FROM Weekly_Summary"
                    Dim cmd0 As New SqlCommand(sql, con)
                    cmd0.CommandTimeout = 120
                    Dim rdr0 As SqlDataReader = cmd0.ExecuteReader
                    While rdr0.Read
                        finalWeek = rdr0(0)
                    End While
                    con.Close()

                    con.Open()
                    sql = "SELECT eDate FROM Calendar WHERE CONVERT(date,GETDATE()) BETWEEN sDate AND eDate AND Week_Id > 0"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        End_Date = rdr("eDate")
                        thisEndDate = rdr("eDate")
                    End While
                    aYearAgo = DateAdd(DateInterval.Month, -12, End_Date)
                    con.Close()







                    ''End_Date = "8/20/2016"
                    ''thisEndDate = "8/20/2016"







                    Try
                        stopWatch.Start()
                        Console.WriteLine("Updating Sales Plan data for " & client)
                        Dim year, period, week, yrprd, prdwk As Integer
                        Dim planamt, planmkdn As Decimal
                        Dim sdate, edate As Date
                        con.Open()
                        sql = "UPDATE Weekly_Summary SET Plan_Sales = 0, Plan_Mkdn = 0 WHERE eDate >= '" & End_Date & "' "
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()

                        con.Open()
                        sql = "SELECT c.Plan_Year AS Year_Id, c.Str_Id, c.Prd_Id, s.Week_Id, cl.sDate, cl.eDate, cl.YrPrd, PrdWk, c.Dept, c.Buyer, c.Class, " & _
                            "ISNULL(s.AMT * c.Pct * b.Pct,0) AS PlanAmt, ISNULL(s.AMT * c.Pct * b.Pct,0) * ISNULL(s.Mkdn_Pct,0) AS PlanMkdn  FROM Class_Pct c " & _
                            "LEFT JOIN Buyer_Pct b ON b.Plan_Year = c.Plan_Year AND b.Str_Id = c.Str_Id AND b.Prd_Id = c.Prd_Id " & _
                                "AND b.Dept = c.Dept AND b.Buyer = c.Buyer " & _
                            "LEFT JOIN Sales_Plan s ON s.Year_Id = c.Plan_Year AND s.Str_Id = c.Str_id AND s.Dept = c.Dept AND s.Prd_Id = c.Prd_Id " & _
                            "LEFT JOIN Calendar cl ON cl.Year_Id = s.Year_Id AND cl.Prd_Id = s.Prd_Id AND cl.PrdWk = s.Week_Id AND cl.PrdWk > 0 " & _
                            "WHERE cl.eDate >= '" & End_Date & "' AND s.Status = 'Active'"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        cnt = 0
                        While rdr.Read
                            cnt += 1
                            If cnt Mod 1000 = 0 Then Console.WriteLine("Updated Plan_Sales for " & cnt & " records")
                            year = rdr("Year_Id")
                            store = rdr("Str_Id")
                            period = rdr("Prd_Id")
                            week = rdr("Week_Id")
                            sdate = rdr("sDate")
                            edate = rdr("eDate")
                            prdwk = rdr("PrdWk")
                            yrprd = rdr("YrPrd")
                            dept = rdr("Dept")
                            buyer = rdr("Buyer")
                            clss = rdr("Class")
                            planamt = rdr("PlanAmt")
                            planmkdn = rdr("PlanMkdn")
                            sql = "IF NOT EXISTS (SELECT Plan_Sales FROM Weekly_Summary WHERE Year_Id = " & year & " " & _
                                "AND Str_Id = '" & store & "' AND Prd_Id = " & period & " AND Week_Id = " & week & " " & _
                                "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "') " & _
                                "INSERT INTO Weekly_Summary (Str_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, eDate, " & _
                                    "YrPrd, Week_Num, Plan_Sales, Plan_Mkdn) " & _
                                "SELECT '" & store & "'," & year & "," & period & "," & prdwk & ",'" & dept & "','" & buyer & "','" &
                                clss & "','" & sdate & "','" & edate & "'," & yrprd & "," & week & "," & planamt & "," & planmkdn & " " & _
                                "FROM Calendar WHERE eDate = '" & edate & "' AND Week_Id > 0 " & _
                                "ELSE " & _
                                "UPDATE Weekly_Summary SET Plan_Sales = " & planamt & ", Plan_Mkdn = " & planmkdn & " " & _
                                "WHERE Str_Id = '" & store & "' AND Year_Id = " & year & " AND Prd_Id = " & period & " " & _
                                "AND Week_Id = " & week & " AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' "
                            con2.Open()
                            cmd = New SqlCommand(sql, con2)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                            con2.Close()
                        End While
                        con.Close()
                        Call Update_Process_Log("2", "Update Plan Sales", "", "")
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Plan__Sales")
                    End Try

                    Try
                        stopWatch.Start()
                        Console.WriteLine("Updating Plan_WksOH for " & client)
                        con.Open()
                        sql = "UPDATE w SET w.Plan_WksOH = b.Plan_WksOH FROM Weekly_Summary w " & _
                                "JOIN Buyer_PCT b ON b.Str_Id = w.Str_Id AND w.Year_Id = Plan_Year " & _
                                "AND b.Prd_Id = w.Prd_Id AND b.Dept = w.Dept AND b.Buyer = w.Buyer"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()

                        thisWeek = End_Date
                        Dim nextWeek As Date = DateAdd(DateInterval.Day, 7, thisWeek)          ' thisweek is the End Date passed to this program
                        Do Until thisWeek >= finalWeek                                         ' finalWeek is the last eDate in OTB

                            Console.WriteLine("Updating Plan_Inv_Retail for " & client & " week " & thisWeek)

                            sql = "UPDATE Weekly_Summary SET Weekly_Summary.Plan_Inv_Retail = ( " & _
                                "SELECT SUM(ISNULL(o2.Plan_Sales,0)) + SUM(ISNULL(o2.Plan_Mkdn,0)) FROM Weekly_Summary AS o2 " & _
                                "WHERE o2.Str_Id = Weekly_Summary.Str_Id AND o2.Dept = Weekly_Summary.Dept " & _
                                "AND o2.Buyer = Weekly_Summary.Buyer AND o2.Class = Weekly_Summary.Class " & _
                                "AND o2.sDate BETWEEN Weekly_Summary.sDate AND DATEADD(week,Weekly_Summary.Plan_WksOH-1,Weekly_Summary.sDate)) " & _
                                "WHERE Weekly_Summary.Buyer <> '*' AND Weekly_Summary.Class <> '*' " & _
                                "AND Weekly_Summary.eDate = '" & thisWeek & "'"
                            con.Open()
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                            con.Close()
                            thisWeek = DateAdd(DateInterval.Day, 7, thisWeek)
                            nextWeek = DateAdd(DateInterval.Day, 7, nextWeek)
                        Loop
                        Call Update_Process_Log("2", "Update Plan_WksOH", "", "")
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Plan_WksOH")
                    End Try

                    Try
                        stopWatch.Start()
                        thisWeek = End_Date
                        Dim prevWeek As Date = DateAdd(DateInterval.Day, -7, thisWeek)
                        Do Until thisWeek >= finalWeek                                         ' finalWeek is the last eDate in Weekly_Summary
                            '                                                                   thisEndDate is this Weeks End Date
                            Console.WriteLine("Updating Projected_Inv and Weekly_Summary for " & client & " week " & thisWeek)

                            If thisWeek = thisEndDate Then
                                sql = "UPDATE Weekly_Summary SET Projected_Inv = ISNULL(Act_Inv_Retail,0) " & _
                                    "+ ISNULL(OnOrder,0) " & _
                                    "- ISNULL(Plan_Sales,0) " & _
                                    "+ ISNULL(Act_Sales,0) " & _
                                    "- ISNULL(Plan_Mkdn,0) " & _
                                    "+ ISNULL(Act_Mkdn,0) WHERE eDate = '" & thisEndDate & "'"
                            Else                                                                                            ' create a new record if one does not exists
                                sql = "UPDATE f SET f.Projected_Inv = ISNULL(p.Projected_Inv,0) " & _
                                    "+ ISNULL(f.OnOrder,0) " & _
                                    "- ISNULL(f.Plan_Sales,0) " & _
                                    "- ISNULL(f.Plan_Mkdn,0) " & _
                                    "FROM Weekly_Summary AS f INNER JOIN Weekly_Summary AS p " & _
                                    "ON f.Str_Id = p.Str_Id AND p.Dept = f.Dept " & _
                                    "AND p.Buyer = f.Buyer AND p.Class = f.Class " & _
                                    "WHERE f.eDate = '" & thisWeek & "' " & _
                                    "AND p.eDate = '" & prevWeek & "' "
                            End If
                            con.Open()
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()

                            sql = "DECLARE @store varchar(10), @dept varchar(10), @buyer varchar(10), @class varchar(10), " & _
                                     "@sDate date, @eDate date, @projected decimal(18,4) " & _
                                 "DECLARE @tbl table(ID int NOT NULL identity(1,1), store varchar(10), dept varchar(10), " & _
                                     "buyer varchar(10), class varchar(10), sdate date, edate date, projected decimal(18,4)) " & _
                                 "INSERT @tbl(store, dept, buyer, class, sdate, edate, projected) " & _
                                 "SELECT Str_Id, Dept, Buyer, Class, sDate, eDate, Projected_Inv FROM Weekly_Summary w " & _
                                     "WHERE eDate = '" & thisWeek & "' " & _
                                 "DECLARE @numrows integer = (SELECT COUNT(*) FROM @tbl) " & _
                                 "DECLARE @cnt integer = @numrows " & _
                                 "WHILE @cnt > 0 " & _
                                 "BEGIN " & _
                                     "SELECT @store = store FROM @tbl WHERE ID = @cnt " & _
                                     "SELECT @dept = dept FROM @tbl WHERE ID = @cnt " & _
                                     "SELECT @class = class FROM @tbl WHERE ID = @cnt " & _
                                     "SELECT @Buyer = buyer FROM @tbl WHERE ID = @cnt " & _
                                     "SELECT @eDate = edate FROM @tbl WHERE ID = @cnt " & _
                                     "SELECT @projected = projected FROM @tbl WHERE ID = @cnt " & _
                                     "IF @projected > 0 " & _
                                     "BEGIN " & _
                                     "IF NOT EXISTS (SELECT eDate FROM Weekly_Summary WHERE Str_Id = @store AND Dept = @dept " & _
                                         "AND Class = @class	AND Buyer = @buyer AND Class = @class AND eDate = DateAdd(day,7,@eDate)) " & _
                                     "INSERT INTO Weekly_Summary (Str_Id, Dept, Buyer, Class, sDate, eDate, Year_Id, Prd_Id, YrPrd, Week_Id, Week_Num) " & _
                                     "SELECT @store, @dept, @buyer, @class, sDate, DateAdd(day,7,@edate), Year_Id, " & _
                                            "Prd_Id, yrprd, PrdWk, Week_Num FROM Calendar " & _
                                         "WHERE eDate = DateAdd(day,7,@edate) AND Week_Id > 0 " & _
                                     "END " & _
                                     "SET @cnt = @cnt -1 " & _
                                     "END"
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                            con.Close()
                            prevWeek = thisWeek
                            thisWeek = DateAdd(DateInterval.Day, 7, thisWeek)
                        Loop
                        Call Update_Process_Log("2", "Update Projected_Inv", "", "")
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Projected_Inv")
                    End Try

                    Try
                        stopWatch.Start()
                        Console.WriteLine("Updating Projected_WksOH for " & client)
                        Dim actInv As Int32 = 0
                        Dim planSales As Int32 = 0
                        Dim numWeeks As Int16
                        Dim eDatex As Date
                        Dim oTest As Object
                        con.Open()
                        sql = "Select Str_Id, Dept, Buyer, Class, eDate, Act_Inv_Retail FROM Weekly_Summary " & _
                            "WHERE Act_Inv_Retail > 0 AND eDate = '" & End_Date & "' " & _
                            "ORDER BY eDate, Str_Id, Dept, Buyer, Class"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        cnt = 0
                        While rdr.Read
                            cnt += 1
                            If cnt Mod 100 = 0 Then
                                Console.WriteLine(eDatex & " " & store & " " & dept & " " & buyer & " " & clss)
                            End If
                            store = rdr("Str_Id")
                            dept = rdr("Dept")
                            buyer = rdr("Buyer")
                            clss = rdr("Class")
                            eDatex = rdr("eDate")
                            oTest = rdr("Act_Inv_Retail")
                            If IsNumeric(oTest) Then actInv = oTest Else actInv = 0
                            planSales = 0
                            numWeeks = 0
                            con2.Open()
                            sql = "SELECT ISNULL(Plan_Sales,0) + ISNULL(Plan_Mkdn,0) As Sales FROM Weekly_Summary " & _
                                "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND " & _
                                "Class = '" & clss & "' AND eDate > '" & eDatex & "' ORDER by eDate"
                            Dim cmd26 As New SqlCommand(sql, con2)
                            cmd26.CommandTimeout = 120
                            Dim rdr26 As SqlDataReader = cmd26.ExecuteReader
                            While rdr26.Read
                                oTest = rdr26(0)
                                If IsNumeric(oTest) Then
                                    planSales += oTest
                                    numWeeks += 1
                                End If
                                If planSales >= actInv Then
                                    con3.Open()
                                    sql = "UPDATE Weekly_Summary SET Projected_WksOH = " & numWeeks & " WHERE Str_Id = '" & store & "' " & _
                                    "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & eDatex & "'"
                                    cmd = New SqlCommand(sql, con3)
                                    cmd.ExecuteNonQuery()
                                    con3.Close()
                                    Exit While
                                End If
                            End While
                            con2.Close()
                        End While
                        con.Close()
                        Call Update_Process_Log("2", "Update Projected_WksOH", "", "")
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Projected_WksOH")
                    End Try

                    Try
                        stopWatch.Start()
                        Console.WriteLine("Updating Act_WksOH for: " & client)
                        Dim theDate As Date
                        Dim sRow As DataRow
                        Dim store As String
                        con.Open()
                        sql = "SELECT DISTINCT Str_Id FROM Weekly_Summary ORDER BY Str_Id"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            store = rdr("Str_Id")
                            con2.Open()
                            sql = "SELECT DISTINCT Dept FROM Weekly_Summary WHERE Str_Id = '" & store & "' ORDER By Dept"
                            cmd = New SqlCommand(sql, con2)
                            rdr2 = cmd.ExecuteReader
                            While rdr2.Read
                                Dim dept = rdr2("Dept")
                                Console.WriteLine("Updating Act_WksOH for: " & client & " " & dept)
                                con3.Open()
                                sql = "SELECT DISTINCT Buyer FROM Weekly_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' Order BY Buyer"
                                cmd = New SqlCommand(sql, con3)
                                rdr3 = cmd.ExecuteReader
                                While rdr3.Read
                                    Dim buyer As String = rdr3("Buyer")
                                    sql = "SELECT DISTINCT Class FROM Weekly_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' ORDER BY Class"
                                    con4.Open()
                                    cmd = New SqlCommand(sql, con4)
                                    rdr4 = cmd.ExecuteReader
                                    While rdr4.Read
                                        clss = rdr4("Class")
                                        Dim invTable As New DataTable
                                        Dim inv, totl, val As Decimal
                                        Dim indx2 As Integer
                                        invTable.Columns.Add("eDate", Type.GetType("System.DateTime"))
                                        invTable.Columns.Add("Inv", Type.GetType("System.Decimal"))

                                        con5.Open()
                                        sql = "SELECT eDate, ISNULL(Act_Inv_Retail,0) AS Inv FROM Weekly_Summary " & _
                                            "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
                                            "AND eDate BETWEEN '" & aYearAgo & "' AND '" & End_Date & "' " & _
                                            "AND Class = '" & clss & "' AND ISNULL(Act_Inv_Retail,0) > 0 AND Act_WksOH IS NULL ORDER BY eDate"
                                        cmd = New SqlCommand(sql, con5)
                                        rdr5 = cmd.ExecuteReader
                                        While rdr5.Read
                                            ''Console.WriteLine("Updating Weekly_Summary for " & store & " " & dept & " " & buyer & " " &
                                            ''                  clss & "  " & rdr5("eDate"))
                                            row = invTable.NewRow
                                            row("eDate") = rdr5("eDate")
                                            inv = rdr5("Inv")
                                            row("Inv") = rdr5("Inv")
                                            invTable.Rows.Add(row)
                                        End While
                                        con5.Close()

                                        Dim salesTable As New DataTable
                                        Dim Column As DataColumn = New DataColumn
                                        Column.DataType = System.Type.GetType("System.DateTime")
                                        Column.ColumnName = "eDate"
                                        salesTable.Columns.Add(Column)
                                        Dim PrimaryKey(1) As DataColumn
                                        PrimaryKey(0) = salesTable.Columns("eDate")
                                        salesTable.PrimaryKey = PrimaryKey
                                        salesTable.Columns.Add("Sales", Type.GetType("System.Decimal"))

                                        con5.Open()
                                        sql = "SELECT eDate, ISNULL(Act_Sales,0) Sales FROM Weekly_Summary " & _
                                             "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
                                             "AND Class = '" & clss & "' ORDER BY eDate"
                                        cmd = New SqlCommand(sql, con5)
                                        rdr5 = cmd.ExecuteReader
                                        While rdr5.Read
                                            sRow = salesTable.NewRow
                                            sRow("eDate") = rdr5("eDate")
                                            sRow("Sales") = rdr5("Sales")
                                            oTest = rdr5("eDate")
                                            val = rdr5("Sales")
                                            salesTable.Rows.Add(sRow)
                                        End While
                                        con5.Close()
                                        For Each iRow As DataRow In invTable.Rows
                                            theDate = iRow("eDate")
                                            inv = iRow("Inv")
                                            totl = 0
                                            cnt = 0
                                            indx2 = FindRow(salesTable, theDate)
                                            If indx2 > -1 Then
                                                Do While totl < inv And indx2 < salesTable.Rows.Count
                                                    sRow = salesTable(indx2)
                                                    oTest = sRow("eDate")
                                                    val = sRow("Sales")
                                                    totl += val
                                                    indx2 += 1
                                                    cnt += 1
                                                Loop
                                                If totl >= inv Then
                                                    con5.Open()
                                                    sql = "UPDATE Weekly_Summary SET Act_WksOH = " & cnt & " WHERE Str_Id = '" & store & "' " & _
                                                        "AND Dept = '" & dept & "' AND Class = '" & clss & "' AND Buyer = '" & buyer & "' " & _
                                                        "AND eDate = '" & theDate & "'"
                                                    cmd = New SqlCommand(sql, con5)
                                                    cmd.ExecuteNonQuery()
                                                    con5.Close()
                                                End If
                                            End If
                                        Next
                                    End While
                                    con4.Close()
                                End While
                                con3.Close()
                            End While
                            con2.Close()
                        End While
                        con.Close()
                        Call Update_Process_Log("2", "Update Act_WksOH", "", "")
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Act_WksOH")
                    End Try

                    Try
                        thisdate = Date.Today
                        Dim dayofweek As Integer = thisdate.DayOfWeek
                        If dayofweek = 6 Then                                                    ' Saturday

                            con.Open()
                            sql = "SELECT eDate FROM Calendar WHERE eDate = (SELECT MAX(eDate) FROM Calendar WHERE eDate < '" & aYearAgo & "' " & _
                                "AND Week_Id = 0) AND Week_Id = 0"
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                Begin_Date = rdr("eDate")
                            End While
                            con.Close()

                            Dim dateTbl As New DataTable
                            Dim column As New DataColumn()
                            column.DataType = System.Type.GetType("System.DateTime")
                            column.ColumnName = "eDate"
                            dateTbl.Columns.Add(column)
                            Dim primaryKey(1) As DataColumn
                            primaryKey(0) = dateTbl.Columns("eDate")
                            dateTbl.PrimaryKey = primaryKey
                            dateTbl.Columns.Add("Year", GetType(System.Int16))
                            dateTbl.Columns.Add("Period", GetType(System.Int16))
                            Dim buyerTbl As New DataTable
                            buyerTbl.Columns.Add("Buyer")
                            Dim selectDate As Date
                            Dim wks As Integer

                            con.Open()
                            sql = "SELECT eDate, Year_Id, Prd_Id  FROM Calendar WHERE Week_Id = 0 AND Prd_Id > 0 " & _
                                "AND eDate >= '" & Begin_Date & "' "
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                row = dateTbl.NewRow
                                row("eDate") = rdr("eDate")
                                row("Year") = rdr("Year_Id")
                                row("Period") = rdr("Prd_Id")
                                dateTbl.Rows.Add(row)
                            End While
                            con.Close()

                            con.Open()
                            Dim deptTbl As New DataTable
                            deptTbl.Columns.Add("Dept")
                            sql = "SELECT Dept FROM Departments WHERE Status = 'Active'"
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                row = deptTbl.NewRow
                                row("Dept") = rdr("Dept")
                                deptTbl.Rows.Add(row)
                            End While
                            con.Close()

                            con.Open()
                            sql = "SELECT Buyer FROM Buyers WHERE Status = 'Active'"
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                row = buyerTbl.NewRow
                                row("Buyer") = rdr("Buyer")
                                buyerTbl.Rows.Add(row)
                            End While
                            con.Close()
                            Dim thisDept As String
                            Dim thisYear As Integer
                            Dim rr As Integer = 0
                            Dim thisPeriod As Integer
                            For Each deptRow In deptTbl.Rows
                                thisDept = deptRow("Dept")
                                Dim thisBuyer As String
                                Dim onHand As Decimal
                                Dim totl As Decimal
                                Dim cnt As Integer
                                For Each dateRow In dateTbl.Rows
                                    selectDate = dateRow("eDate")
                                    thisPeriod = dateRow("Period")
                                    thisYear = dateRow("Year")
                                    For Each buyerRow In buyerTbl.Rows
                                        thisBuyer = buyerRow("Buyer")
                                        Console.WriteLine("Updating Buyer_PCT " & thisDept & " " & thisBuyer)
                                        onHand = 0
                                        totl = 0
                                        cnt = 0
                                        con.Open()
                                        sql = "SELECT ISNULL(SUM(End_OH * Retail),0) As Inv FROM Item_Inv i " & _
                                            "JOIN Item_Master m ON m.Item_No = i.Item_No " & _
                                            "WHERE eDate = '" & selectDate & "' " & _
                                            "AND Str_Id = '1' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "'"
                                        cmd = New SqlCommand(sql, con)
                                        rdr = cmd.ExecuteReader
                                        While rdr.Read
                                            onHand = rdr("Inv")
                                        End While
                                        con.Close()

                                        con.Open()
                                        sql = "SELECT eDate, ISNULL(SUM(Sales_Retail),0) As Sales FROM Item_Detail d " & _
                                            "JOIN Item_Master m ON m.Item_No = d.Item_No " & _
                                            "WHERE Str_Id = '1' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                                            "AND eDate >= '" & selectDate & "' " & _
                                            "GROUP BY eDate ORDER By eDate"
                                        cmd = New SqlCommand(sql, con)
                                        rdr = cmd.ExecuteReader
                                        While rdr.Read
                                            rr += 1
                                            cnt += 1
                                            totl += rdr("Sales")
                                            If totl >= onHand Then
                                                wks = cnt
                                                Exit While
                                            End If
                                        End While
                                        con.Close()

                                        If totl > onHand Then
                                            con.Open()
                                            sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Str_Id = '1' AND Dept = '" & thisDept & "' " & _
                                                "AND Buyer = '" & thisBuyer & "' AND Plan_Year = " & thisYear & " AND Prd_Id = " & thisPeriod & ") " & _
                                                "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Act_WksOH) " & _
                                                "SELECT " & thisYear & ",'1'," & thisPeriod & ",'" & thisDept & "','" & thisBuyer & "'," & wks & " " & _
                                                "ELSE " & _
                                                "UPDATE Buyer_Pct SET Act_WksOH = " & wks & " WHERE Plan_Year= " & thisYear & " " & _
                                                "AND Str_Id = '1' AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' " & _
                                                "AND Buyer = '" & thisBuyer & "'"
                                            cmd = New SqlCommand(sql, con)
                                            cmd.ExecuteNonQuery()
                                            con.Close()
                                        End If
                                    Next
                                Next
                            Next
                            Call Update_Process_Log("2", "Update Buyer_PCT.Act_WksOH", "", "")

                            Process.Start(rcExePath & "\ScoreItems.exe")

                        End If
                    Catch ex As Exception
                        Dim theMessage As String = ex.Message
                        Console.WriteLine(ex.Message)
                        Dim el As New Optional_Module_Processing.ErrorLogger
                        el.WriteToErrorLog(ex.Message, ex.StackTrace, "Optional_Module_Processing, Update Buyer_PCT.Act_WksOH")
                    End Try
                    '          Update PCT tables for closed periods
                    '
                    Call Check_For_Completed_Period()
                    '
                    '
                End If
            End If
            Console.WriteLine("Done with Sales Plan")
            oTest = clientRow("Item4Cast")

            If Not IsDBNull(oTest) Then
                If clientRow("Item4Cast") = "Y" Then
                    thisdate = Date.Today
                    Dim dayofweek As Integer = thisdate.DayOfWeek
                    If dayofweek = 6 Then
                        Process.Start(rcExePath & "\ForecastItems.exe")
                    End If
                End If
            End If
            oTest = clientRow("Marketing")
            If Not IsDBNull(oTest) Then
                If clientRow("Marketing") = "Y" Then
                    Call Process_Customers(xmlPath, con)
                End If
            End If

        Next                          ' Get the next Client
        '
       
    End Sub

    Private Sub Process_Customers(ByVal thePath, ByVal con)
        Try
            stopWatch.Start()
            Console.WriteLine("Processing " & client & " Customer records")
            thePath = xmlPath & "\Customers.xml"
            Dim tbl As DataTable = New DataTable
            Dim dset As DataSet = New DataSet
            Dim xmlFile As XmlReader
            Dim row As DataRow
            Dim id As String
            Dim dtTbl = New DataTable
            dset = New DataSet
            Dim column As New DataColumn
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Cust_No"
            dtTbl.Columns.Add(column)
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = dtTbl.Columns("Cust_No")
            dtTbl.PrimaryKey = PrimaryKey2
            dtTbl.Columns.Add("Name", GetType(System.String))
            dtTbl.Columns.Add("fName", GetType(System.String))
            dtTbl.Columns.Add("lName", GetType(System.String))
            dtTbl.Columns.Add("Addr1", GetType(System.String))
            dtTbl.Columns.Add("Addr2", GetType(System.String))
            dtTbl.Columns.Add("Addr3", GetType(System.String))
            dtTbl.Columns.Add("City", GetType(System.String))
            dtTbl.Columns.Add("State", GetType(System.String))
            dtTbl.Columns.Add("Zip", GetType(System.String))
            dtTbl.Columns.Add("Phone1", GetType(System.String))
            dtTbl.Columns.Add("Phone2", GetType(System.String))
            dtTbl.Columns.Add("Cell1", GetType(System.String))
            dtTbl.Columns.Add("Cell2", GetType(System.String))
            dtTbl.Columns.Add("eMail", GetType(System.String))
            dtTbl.Columns.Add("eMail2", GetType(System.String))
            dtTbl.Columns.Add("Type", GetType(System.String))
            dtTbl.Columns.Add("Balance", GetType(System.Decimal))
            dtTbl.Columns.Add("Loyalty", GetType(System.Int32))
            dtTbl.Columns.Add("okToMail", GetType(System.String))
            dtTbl.Columns.Add("okToEmail", GetType(System.String))
            dtTbl.Columns.Add("fDate", GetType(System.DateTime))
            dtTbl.Columns.Add("lDate", GetType(System.DateTime))
            dtTbl.Columns.Add("lUpdate", GetType(System.DateTime))
            dtTbl.Columns.Add("Spouse", GetType(System.String))
            dtTbl.Columns.Add("Birthday", GetType(System.DateTime))
            Dim custNo, name, fName, lName, Addr1, Addr2, Addr3, city, state, zip, spouse, phone1, phone2,
                email, type, oktomail, oktoemail, cell1, cell2 As String
            Dim cnt As Integer = 0
            Dim birthday, fDate, lDate, lUpdate As String
            Dim bal, points As String
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            dset.ReadXml(xmlFile)
            If ((dset.Tables.Count > 0) AndAlso dset.Tables(0).Rows.Count > 0) Then
                tbl = dset.Tables(0)
                For Each row In tbl.Rows
                    cnt += 1
                    If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                    custNo = Trim(Left(Replace(row("CUST_NO"), "'", "''"), 30))
                    name = Trim(Left(Replace(row("NAME"), "'", "''"), 50))
                    fName = Trim(Left(Replace(row("FIRST_NAME"), "'", "''"), 30))
                    lName = Trim(Left(Replace(row("LAST_NAME"), "'", "''"), 30))
                    Addr1 = Trim(Left(Replace(row("ADDRESS_1"), "'", "''"), 40))
                    Addr2 = Trim(Left(Replace(row("ADDRESS_2"), "'", "''"), 40))
                    Addr3 = Trim(Left(Replace(row("ADDRESS_3"), "'", "''"), 40))
                    city = Trim(Left(Replace(row("CITY"), "'", "''"), 30))
                    state = Trim(Left(Replace(row("STATE"), "'", "''"), 10))
                    zip = Trim(Left(Replace(row("ZIP"), "'", "''"), 15))
                    spouse = Trim(Left(Replace(row("SPOUSE"), "'", "''"), 30))
                    birthday = Trim(Left(Replace(row("BIRTHDAY"), "'", "''"), 20))
                    phone1 = Trim(Left(Replace(row("PHONE_1"), "'", "''"), 15))
                    phone2 = Trim(Left(Replace(row("PHONE_2"), "'", "''"), 15))
                    email = Trim(Left(Replace(row("EMAIL"), "'", "''"), 50))
                    type = Trim(Left(Replace(row("TYPE"), "'", "''"), 10))
                    bal = Trim(Left(Replace(row("BALANCE"), "'", "''"), 20))
                    fDate = Trim(Left(Replace(row("FIRST_SALE_DATE"), "'", "''"), 20))
                    lDate = Trim(Left(Replace(row("LAST_SALE_DATE"), "'", "''"), 20))
                    points = Trim(Left(Replace(row("LOYALTY_POINTS"), "'", "''"), 10))
                    oktoemail = Trim(Left(Replace(row("OK_TO_EMAIL"), "'", "''"), 1))
                    oktomail = Trim(Left(Replace(row("OK_TO_MAIL"), "'", "''"), 1))
                    cell1 = Trim(Left(Replace(row("CELL_1"), "'", "''"), 15))
                    cell2 = Trim(Left(Replace(row("CELL_2"), "'", "''"), 15))
                    lUpdate = Trim(Left(Replace(row("LAST_UPDATE"), "'", "''"), 50))
                    Dim row2 As DataRow = dtTbl.NewRow
                    row2("Cust_No") = custNo
                    row2("Name") = name
                    row2("fName") = fName
                    row2("lName") = lName
                    row2("Addr1") = Addr1
                    row2("Addr2") = Addr2
                    row2("Addr3") = Addr3
                    row2("City") = city
                    row2("State") = state
                    row2("Zip") = zip
                    row2("Spouse") = spouse
                    If Not IsDBNull(birthday) And IsNothing(birthday) Then
                        If birthday <> "" And birthday Like "##/##/##" Then
                            row2("Birthday") = CDate(birthday)
                        End If
                    End If
                    row2("Phone1") = phone1
                    row2("Phone2") = phone2
                    row2("eMail") = email
                    row2("Type") = type
                    If IsNumeric(bal) Then row2("Balance") = CDec(bal)
                    If Not IsNothing(fDate) And fDate <> "" Then row2("fDate") = CDate(fDate)
                    If Not IsNothing(lDate) And lDate <> "" Then row2("lDate") = CDate(lDate)
                    If Not IsNothing(lUpdate) And lUpdate <> "" Then row2("lUpdate") = CDate(lUpdate)
                    If IsNumeric(points) Then row2("Loyalty") = CInt(points)
                    row2("OkToMail") = oktomail
                    row2("OkToEmail") = oktoemail
                    row2("Cell1") = cell1
                    row2("Cell2") = cell2
                    dtTbl.Rows.Add(row2)
                Next
                con.Open()
                sql = "DELETE FROM CUSTOMERS"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()

                Dim connection As SqlConnection = New SqlConnection(conString)
                Console.WriteLine("Bulk copy records to Customers")
                Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
                connection.Open()
                bulkCopy.DestinationTableName = "dbo.Customers"
                bulkCopy.BulkCopyTimeout = 120
                bulkCopy.WriteToServer(dtTbl)
                connection.Close()

                Console.WriteLine("Ranking Customers")
                con.Open()
                sql = "declare @onefifth integer, @twofifths integer, @threefifths integer, @fourfifths integer " & _
               "update customers set recency = NULL, frequency = NULL, monetary = NULL " & _
               "select row_number() over (order by count(*) asc) as row, cust_no, count(*) as tickets into #t1 from tickets group by cust_no " & _
               "select row_number() over (order by sum(amt) asc) as row, cust_no, sum(amt) as amt into #t2 from tickets group by cust_no " & _
               "select row_number() over (order by max(date) asc) as row, cust_no, max(date) as date into #t3 from tickets group by cust_no " & _
                "set @onefifth=(select count(*)/5 from #t1) " & _
               "set @twofifths=@onefifth * 2 " & _
               "set @threefifths=@onefifth * 3 " & _
               "set @fourfifths=@onefifth * 4 " & _
               "update c set Recency=(case when row < @onefifth then 1 when row >= @onefifth and row < @twofifths then 2 " & _
                   "when row  >= @twofifths and row < @threefifths then 3 " & _
                   "when row >= @threefifths and row < @fourfifths then 4 else 5 end)  from customers c " & _
                   "join #t3 t on t.cust_no = c.cust_no " & _
               "update c set Frequency=(case when row < @onefifth then 1 when row >= @onefifth and row < @twofifths then 2 " & _
                   "when row  >= @twofifths and row < @threefifths then 3 " & _
                   "when row >= @threefifths and row < @fourfifths then 4 else 5 end)  from customers c " & _
                   "join #t1 t on t.cust_no=c.cust_no " & _
               "update c set Monetary=(case when row < @onefifth then 1 when row >= @onefifth and row < @twofifths then 2 " & _
                   "when row  >= @twofifths and row < @threefifths then 3 " & _
                   "when row >= @threefifths and row < @fourfifths then 4 else 5 end)  from customers c " & _
                   "join #t2 t on t.cust_no=c.cust_no "
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()
                con.Close()
            End If

            Dim m As String = "Created " & cnt & " records"
            Call Update_Process_Log("3", "Create Customers", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New Optional_Module_Processing.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Customers")
        End Try
    End Sub

    Function FindRow(ByVal dt As DataTable, ByVal eDate As String) As Integer
        For i As Integer = 0 To dt.Rows.Count - 1
            If dt.Rows(i)("eDate") = eDate Then Return i
        Next
        Return -1
    End Function

    Private Sub Check_For_Completed_Period()
        Dim thisPeriod, thisYrPrd, lastYrPrd As Integer
        Dim lastPeriodUpdated As Integer = 0
        Dim firstPeriodToUpdate As Integer = 0
        Dim lastPeriodToUpdate As Integer = 0
        Dim storeTbl As New DataTable
        storeTbl.Columns.Add("Store")
        Dim srow As DataRow
        con.Open()
        sql = "SELECT Year_ID, Prd_Id, YrPrd FROM Calendar WHERE CONVERT(Date,GETDATE()) BETWEEN sDate AND eDate AND Prd_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            thisYear = rdr("Year_Id")
            thisPeriod = rdr("Prd_Id")
            thisYrPrd = rdr("YrPrd")
        End While
        con.Close()

        con.Open()
        sql = "SELECT MAX(YrPrd) AS YrPrd FROM Calendar WHERE YrPrd < " & thisYrPrd & " "
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("YrPrd")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then lastYrPrd = CInt(oTest)
            End If
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT Str_Id FROM Buyer_Pct WHERE Year_Id = " & lastYear & " "
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            thisStore = rdr("Str_Id")
            srow = storeTbl.NewRow
            srow("Store") = rdr("Str_Id")
            storeTbl.Rows.Add(srow)
        End While

        For Each srow In storeTbl.Rows
            thisStore = srow("Store")
            lastPeriodUpdated = 0
            con2.Open()
            sql = "SELECT MAX(YrPrd) AS Prd_Id FROM Buyer_PCT b " & _
                "JOIN Calendar c ON c.Year_Id = b.Year_Id AND c.Prd_Id = b.Prd_Id AND c.Week_Id = 0 " & _
                "WHERE Str_Id = '" & thisStore & "' AND ISNULL(Act_Pct,0) > 0"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                oTest = rdr2("Prd_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then lastPeriodUpdated = CInt(oTest)
                End If
            End While
            con2.Close()

            If lastPeriodUpdated = 0 Then
                con2.Open()
                sql = "SELECT MAX(YrPrd) AS YrPrd FROM Calendar WHERE YrPrd < " & lastYrPrd & " AND Prd_Id > 0"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    oTest = rdr("YrPrd")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then lastPeriodUpdated = CInt(oTest)
                    End If
                End While
                con2.Close()
            End If

            con2.Open()
            sql = "SELECT MIN(YrPrd) AS Prd_Id FROM Calendar WHERE YrPrd > " & lastPeriodUpdated & " AND Prd_Id > 0 AND Week_Id = 0"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                oTest = rdr2("Prd_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then firstPeriodToUpdate = CInt(oTest)
            End While
            con2.Close()

            con2.Open()
            sql = "SELECT MAX(YrPrd) AS Prd_Id FROM Calendar WHERE YrPrd < " & thisYrPrd & " AND Week_Id = 0"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                oTest = rdr2("Prd_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lastPeriodToUpdate = CInt(oTest)
            End While
            con2.Close()


            If firstPeriodToUpdate > 0 And thisYrPrd > 0 And firstPeriodToUpdate < thisYrPrd Then
                For period As Integer = firstPeriodToUpdate To lastPeriodToUpdate
                    con2.Open()
                    cmd = New SqlCommand("sp_OptionalModule_PeriodClose", con2)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
                    cmd.Parameters.Add("@yrprd", SqlDbType.Int).Value = period
                    cmd.ExecuteNonQuery()
                    con2.Close()
                Next
            End If
        Next
        con.Close()
    End Sub


    ''    code moved to sp_OptionalModule_PeriodClose

    ''Private Sub Update_Closed_Period(ByVal period As Integer)
    ''    Dim thisUser As String = Environment.UserName
    ''    Dim tblDept As DataTable = New DataTable
    ''    tblDept.Columns.Add("Dept")
    ''    Dim row As DataRow
    ''    con3.Open()
    ''    sql = "SELECT ID FROM Departments WHERE Status = 'Active'"
    ''    cmd = New SqlCommand(sql, con3)
    ''    rdr = cmd.ExecuteReader
    ''    While rdr.Read
    ''        row = tblDept.NewRow
    ''        row(0) = rdr(0)
    ''        tblDept.Rows.Add(row)
    ''    End While
    ''    con3.Close()

    ''    Dim thisDept As String
    ''    Dim thisEdate As Date
    ''    Dim onHand As Decimal = 0
    ''    Dim planSales As Int32 = 0
    ''    Dim wks As Integer = 0
    ''    Dim dteTbl As New DataTable
    ''    dteTbl.Columns.Add("eDate")
    ''    dteTbl.Columns.Add("PrdWk")
    ''    Dim dRow As DataRow
    ''    con3.Open()
    ''    sql = "SELECT eDate, PrdWK FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " "
    ''    cmd = New SqlCommand(sql, con3)
    ''    rdr = cmd.ExecuteReader
    ''    While rdr.Read
    ''        dRow = dteTbl.NewRow
    ''        dRow("eDate") = rdr("eDate")
    ''        dRow("prdWk") = rdr("PrdWk")
    ''        dteTbl.Rows.Add(dRow)
    ''        If rdr("PrdWK") = 0 Then thisEdate = rdr("eDate")
    ''    End While
    ''    con3.Close()
    ''    For Each dRow In tblDept.Rows
    ''        thisDept = dRow(0)
    ''        con3.Open()                                     ' Buyer_Pct records
    ''        sql = "SELECT Buyer, ISNULL(SUM(Act_Sales),0) AS Sales, ISNULL(SUM(Act_Inv_Retail),0) AS inv FROM Weekly_Summary " & _
    ''            "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND eDate = '" & thisEdate & "' " & _
    ''            "GROUP BY Buyer"
    ''        cmd = New SqlCommand(sql, con3)
    ''        rdr = cmd.ExecuteReader
    ''        While rdr.Read
    ''            thisBuyer = rdr("Buyer")
    ''            onHand = rdr("inv")
    ''            con4.Open()
    ''            sql = "SELECT eDate, ISNULL(SUM(Plan_Sales),0) AS sales FROM Weekly_Summary WHERE Str_Id = '" & thisStore & "' " & _
    ''                "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' AND eDate > '" & thisEdate & "' " & _
    ''                "GROUP BY eDate"
    ''            cmd = New SqlCommand(sql, con4)
    ''            rdr2 = cmd.ExecuteReader
    ''            While rdr2.Read
    ''                planSales += rdr2("sales")
    ''                wks += 1
    ''            End While
    ''            '
    ''            '
    ''            '   Gary believes Act_Pct should be updated instead of "Pct" here, DayOfWeekPct and Class_PCT
    ''            '
    ''            '
    ''            '
    ''            If planSales >= onHand Then
    ''                con5.Open()
    ''                sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Year_Id = " & thisYear & " AND Str_Id = '" & thisStore & "' " & _
    ''                    "AND Prd_Id = " & period & " AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "') " & _
    ''                    "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Projected_WksOH) " & _
    ''                    "SELECT " & thisYear & ", '" & thisStore & "', " & period & ", '" & thisDept & "', '" & thisBuyer & "', " & wks & " " & _
    ''                    "ELSE " & _
    ''                    "UPDATE Buyer_PCT SET Projected_WksOH = " & wks & " WHERE Plan_Year = " & thisYear & " " & _
    ''                    "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
    ''                    "AND Prd_Id = " & period & " "
    ''                cmd = New SqlCommand(sql, con5)
    ''                cmd.ExecuteNonQuery()
    ''                con5.Close()
    ''                Exit While
    ''            End If
    ''        End While
    ''        con4.Close()
    ''        con3.Close()

    ''        con3.Open()
    ''        sql = "CREATE TABLE #t2 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(10), " & _
    ''                "sales decimal(20), pct decimal(18,4)) " & _
    ''            "INSERT INTO #t2 (Str_Id, Prd_Id, Dept, Buyer, Sales) " & _
    ''            "SELECT Str_Id, Prd_Id, Dept, Buyer, ISNULL(SUM(Act_Sales),0) AS Sales FROM Weekly_Summary " & _
    ''            "GROUP BY Str_Id, Prd_Id, Dept, Buyer " & _
    ''            "CREATE TABLE #t2a (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), sales decimal(20)) " & _
    ''            "INSERT INTO #t2a (Str_Id, Prd_Id, Dept, sales) " & _
    ''            "SELECT Str_Id, Prd_Id, Dept, SUM(Sales) FROM #t2 " & _
    ''            "GROUP BY Str_Id, Prd_Id, Dept " & _
    ''            "UPDATE t2 SET t2.pct = CASE WHEN t2a.Sales <> 0 THEN t2.sales / t2a.sales ELSE 0 END FROM #t2 t2 " & _
    ''            "JOIN #t2a t2a ON t2a.Str_Id = t2.Str_Id AND t2a.Prd_Id = t2.Prd_Id AND t2a.Dept = t2.Dept " & _
    ''            "UPDATE b SET b.Act_PCT = t.pct FROM Buyer_PCT b " & _
    ''            "JOIN #t2 t ON t.Str_Id = b.Str_Id AND t.Prd_Id = b.Prd_Id AND t.Dept = b.Dept " & _
    ''            "AND t.Buyer = b.Buyer WHERE b.Plan_Year = " & thisYear & " AND b.Prd_Id = " & period & " " 
    ''        cmd = New SqlCommand(sql, con3)
    ''        cmd.ExecuteNonQuery()

    ''        sql = "CREATE TABLE #t1a (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
    ''            "Class varchar(20), Sales decimal(18,4), PCT Decimal(18,4)) " & _
    ''            "INSERT INTO #t1a (Str_Id, Prd_Id, Dept, Buyer, Class, Sales, PCT) " & _
    ''            "SELECT Str_Id, Prd_Id, Dept, CASE WHEN Buyer IS NULL OR Buyer = '' THEN 'OTHER' ELSE Buyer END AS Buyer, " & _
    ''                "Class, ISNULL(SUM(Act_Sales),0) AS Sales, 0 FROM Weekly_Summary " & _
    ''            "WHERE Str_Id = '" & thisStore & "' AND Prd_Id = " & period & " AND Act_Sales <> 0 " & _
    ''            "GROUP BY Str_Id, Prd_Id, Dept, Buyer, Class " & _
    ''            "CREATE TABLE #t6 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), Sales decimal(18,4)) " & _
    ''            "INSERT INTO #t6 (Str_Id, Prd_Id, Dept, Buyer, Sales) " & _
    ''            "SELECT Str_Id, Prd_Id, Dept, Buyer, SUM(Sales) As Sales " & _
    ''                "FROM #t1a WHERE Sales <> 0 GROUP BY Str_Id, Prd_Id, Dept, Buyer " & _
    ''            "UPDATE t1 SET t1.PCT = t1.Sales / t6.Sales FROM #t1a t1 " & _
    ''            "JOIN #t6 t6 ON t6.Dept = t1.Dept AND t6.Buyer = t1.Buyer " & _
    ''            "UPDATE c SET c.Act_PCT = t.PCT FROM Class_PCT c " & _
    ''            "JOIN #t1a t ON t.Str_Id = c.Str_Id AND t.Prd_Id = c.Prd_Id AND t.Dept = c.Dept " & _
    ''            "AND t.Buyer = c.Buyer AND t.Class = c.Class"
    ''        cmd = New SqlCommand(sql, con3)
    ''        cmd.ExecuteNonQuery()

    ''        sql = "CREATE TABLE #t1 (Str_Id varchar(10), Year_Id integer, Prd_Id integer NOT NULL, Day integer, " & _
    ''                "Sales decimal(18,4), PCT Decimal(18,4)) " & _
    ''            "INSERT INTO #t1 (Str_Id, Year_Id, Prd_Id, Day, Sales, PCT) " & _
    ''            "SELECT LOCATION, Year_Id, Prd_Id, DATEPART(dw,TRANS_DATE) AS Day, ISNULL(SUM(QTY * RETAIL),0) AS Sales, 0 " & _
    ''            "FROM Daily_Transaction_Log l " & _
    ''            "JOIN Calendar c ON CONVERT(Date,TRANS_DATE) BETWEEN sDate AND eDate AND Week_Id = 0 " & _
    ''            "WHERE QTY <> 0 AND RETAIL > 0 AND LOCATION = '" & thisStore & "' " & _
    ''            "AND TYPE = 'Sold' AND c.Year_id = " & thisYear & " AND c.Prd_Id = " & period & " " & _
    ''            "GROUP BY LOCATION, Year_Id, Prd_Id, DATEPART(dw,TRANS_DATE) " & _
    ''            "DECLARE @total DECIMAL(18,4) " & _
    ''            "SET @total = (SELECT SUM(Sales) FROM #t1) " & _
    ''            "UPDATE t SET t.PCT = t.Sales / @total FROM #t1 t " & _
    ''            "UPDATE d SET d.PCT = t.PCT FROM #t1 t " & _
    ''            "JOIN DayOfWeekPct d ON d.Year_Id = t.Year_Id AND d.Str_Id = t.Str_Id " & _
    ''            "AND d.Prd_Id = t.Prd_Id AND d.Day = t.Day "
    ''        cmd = New SqlCommand(sql, con3)
    ''        cmd.ExecuteNonQuery()
    ''        con.Close()
    ''        con3.Close()

    ''        For Each dteRow In dteTbl.Rows
    ''            thisEdate = dteRow("eDate")                          ' Week records
    ''            thisWeek = dteRow("PrdWk")
    ''            lastYear = thisYear - 1
    ''            con4.Open()
    ''            sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint,@amt decimal " & _
    ''                "DECLARE @rnd AS integer, @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4), @mamt decimal " & _
    ''                "SELECT @rnd = " & rnd & " " & _
    ''            "SET @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
    ''                "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
    ''            "SET @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
    ''                "AND Prd_Id = " & period & ") " & _
    ''            "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
    ''            "SET @amt = CONVERT(int,(@pct * @allsales)) " & _
    ''            "SET @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
    ''                "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
    ''            "SET @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
    ''                "AND Prd_Id = " & period & ") " & _
    ''            "SELECT @mpct = COALESCE(@msales / (NULLIF(@sales,0) + NULLIF(@msales,0)),0) WHERE @msales <> 0 " & _
    ''            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
    ''            "AND Dept = '" & thisDept & "' AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
    ''            "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User, Status, " & _
    ''            "Mkdn_Pct) " & _
    ''            "SELECT '" & thisPlan & "','" & thisStore & "','" & thisDept & "'," & thisYear & "," & period & "," & _
    ''                thisWeek & ",@amt,@pct,'" & today & "','" & thisUser & "','Active',@mpct " & _
    ''            "ELSE " & _
    ''            "UPDATE Sales_Plan SET Amt = ROUND(@amt / @rnd,0) * @rnd, Pct = @pct, Last_Update = '" & today & "', " & _
    ''                "Mkdn_Pct = @mpct, Last_User = '" & thisUser & "' " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
    ''                "AND Year_Id = " & thisYear & " " & _
    ''                "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & " AND Plan_id = '" & thisPlan & "'"
    ''            cmd = New SqlCommand(sql, con4)
    ''            cmd.ExecuteNonQuery()
    ''            con4.Close()
    ''        Next

    ''        con3.Open()                                            ' Dept record
    ''        sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint, @amt bigint " & _
    ''           "SELECT @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Weekly_Summary WHERE Str_Id = '1' AND Dept = 'AC' " & _
    ''               "AND Year_Id = " & thisYear & ") " & _
    ''           "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Weekly_Summary WHERE Str_Id = '1' AND Year_Id = " & thisYear & ") " & _
    ''           "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
    ''           "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
    ''                   "AND Year_Id = " & thisYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
    ''                   "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "') " & _
    ''                "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User, Status) " & _
    ''                   "SELECT '" & thisPlan & "'," & thisYear & ",'" & thisStore & "','" & thisDept & "',0,0,@sales,@pct,'" &
    ''                       today & "', '" & thisUser & "','Active' " & _
    ''               "ELSE " & _
    ''                "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
    ''                   "WHERE Year_Id = " & thisYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
    ''                   "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "'"
    ''        cmd = New SqlCommand(sql, con3)
    ''        cmd.ExecuteNonQuery()
    ''        con3.Close()

    ''        con3.Open()                                         ' Period record
    ''        sql = "DECLARE @pct decimal(18,4), @sales decimal(18,4), @allsales decimal(18,4) " & _
    ''            "DECLARE @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4) " & _
    ''            "SELECT @sales = (SELECT ISNULL(SUM(Act_sales),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " &
    ''                    lastYear & " AND Prd_Id = " & period & ") " & _
    ''            "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Prd_Id = " & period & ") " & _
    ''            "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
    ''            "SELECT @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " &
    ''                    lastYear & " AND Prd_Id = " & period & ") " & _
    ''            "SELECT @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Weekly_Summary " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Dept = '" & thisDept & "') " & _
    ''            "SELECT @mpct = COALESCE(@msales / (NULLIF(@sales,0) + NULLIF(@msales,0)),0) WHERE @msales <> 0 " & _
    ''            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
    ''                "AND Year_Id = " & thisYear & " AND Prd_Id = " & period & " AND Week_Id = 0 " & _
    ''                "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "') " & _
    ''            "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Mkdn_Pct, Last_Update, Last_User, Status) " & _
    ''            "SELECT '" & thisPlan & "'," & thisYear & ",'" & thisStore & "','" & thisDept & "'," & period & ",0,@sales," & _
    ''                "@pct,@mpct,'" & today & "', '" & thisUser & "','Active' " & _
    ''            "ELSE " & _
    ''            "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Mkdn_Pct = @mpct, Last_Update = '" & today & "', " & _
    ''                "Last_User = '" & thisUser & "' " & _
    ''                "WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " AND Week_Id = 0 " & _
    ''            "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "' "
    ''        cmd = New SqlCommand(sql, con3)
    ''        cmd.ExecuteNonQuery()
    ''        con3.Close()
    ''    Next
    ''End Sub

    Private Sub Update_Buyer_Pct(ByVal period As Integer)
        '
        '                   this code updates Act_WksOH in Buyer_PCT
        '

        Dim thisDept As String = dept
        Dim thisEdate As Date
        Dim onHand As Decimal = 0
        Dim actSales As Int32 = 0
        Dim wks As Integer = 0
        Dim deptTbl As New DataTable
        deptTbl.Columns.Add("Dept")
        Dim row As DataRow
        con3.Open()
        sql = "SELECT DISTINCT Dept FROM Buyer_Pct WHERE Plan_Year = " & thisYear & " AND Str_Id = '" & thisStore & "' " & _
            "AND Prd_Id = " & period & " "
        cmd = New SqlCommand(sql, con3)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = deptTbl.NewRow
            row("Dept") = rdr("Dept")
            deptTbl.Rows.Add(row)
        End While
        con3.Close()

        For Each row In deptTbl.Rows
            thisDept = row("Dept")
            actSales = 0
            wks = 0
            con3.Open()
            sql = "SELECT MAX(eDate) As eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " AND Week_Id > 0"
            cmd = New SqlCommand(sql, con3)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisEdate = rdr("eDate")
            End While
            con3.Close()

            con3.Open()
            sql = "SELECT w.Buyer, ISNULL(SUM(Act_Inv_Retail),0) AS Inv FROM Weekly_Summary w " & _
                    "JOIN Buyers b ON b.ID = w.Buyer " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND b.Status = 'Active' " & _
                    "AND eDate = '" & thisEdate & "' " & _
                    "GROUP BY w.Buyer"
            cmd = New SqlCommand(sql, con3)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisBuyer = rdr("Buyer")
                onHand = rdr("inv")
                If onHand > 0 Then
                    con4.Open()
                    sql = "SELECT eDate, ISNULL(SUM(Act_Sales),0) AS sales FROM Weekly_Summary w " & _
                        "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND eDate > '" & thisEdate & "' " & _
                        "GROUP BY eDate"
                    cmd = New SqlCommand(sql, con4)
                    rdr2 = cmd.ExecuteReader
                    While rdr2.Read
                        actSales += rdr2("sales")
                        wks += 1
                        If actSales >= onHand Then
                            con5.Open()
                            sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Plan_Year = " & thisYear & " AND Str_Id = '" & thisStore & "' " & _
                                "AND Prd_Id = " & period & " AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "') " & _
                                "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Act_WksOH) " & _
                                "SELECT " & thisYear & ", '" & thisStore & "', " & period & ", '" & thisDept & "', '" & thisBuyer & "', " & wks & " " & _
                                "ELSE " & _
                                "UPDATE Buyer_PCT SET Act_WksOH = " & wks & " WHERE Plan_Year = " & thisYear & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                                "AND Prd_Id = " & period & " "
                            cmd = New SqlCommand(sql, con5)
                            cmd.ExecuteNonQuery()
                            con5.Close()
                            actSales = 0
                            wks = 0
                            con4.Close()
                            GoTo 50
                        End If
                    End While
                    con4.Close()
                End If
50:         End While
            con3.Close()
        Next
        
    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        stopWatch.Stop()
        thisDateTime = CDate(Now)
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "Optional_Module_Processing"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & thisDateTime & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
End Module

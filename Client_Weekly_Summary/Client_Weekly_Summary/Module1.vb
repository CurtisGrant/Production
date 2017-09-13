Imports System.Data.SqlClient
Imports System.Xml
Imports System.Globalization
Module Module1
    Public client As String
    Public errorLog As String
    Public rcCON As SqlConnection
    Private server, dbase, thePath, xmlPath, sqlUserId, sqlPassword, conString, rcConString, sql, msg, rcExePath As String
    Private Begin_Date, End_Date, todaysdate, thisdate, thisenddate, thisweek, finalweek, dte, xmlBeginDate, xmlEndDate As Date
    Private aYearAgo As Date
    Private thisDateTime As DateTime
    Private con, con2, con3, con4, con5, mcon As SqlConnection
    Private cmd As SqlCommand
    Private location, store, dept, buyer, clss, dateStr As String
    Private rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Private cnt, thisYear As Integer
    Private oTest As Object
    Private stopWatch As Stopwatch
    Private Has_Sales_Plan As Boolean

    Sub Main()
        Console.WriteLine("Client_Weekly_Summary")
        Has_Sales_Plan = False
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


        Try
            Dim args As String() = Environment.GetCommandLineArgs()
            Dim txtArray() As String = args(1).Split(";")
            client = txtArray(0)
            server = txtArray(1)               ' Get from Client_Master
            dbase = txtArray(2)                ' Get from Client_Master
            xmlPath = txtArray(3)              ' Get from Client_Master
            sqlUserId = txtArray(4)            ' Get from Client_Master
            sqlPassword = txtArray(5)          ' Get from Client_Master
            dateStr = txtArray(6)
            errorLog = txtArray(7)             ' Get from Client_Master
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New Client_Weekly_Summary.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Separate Arguments")
        End Try


        ''client = "PARGIF"
        ''server = "LP-CURTIS"
        ''dbase = "PARGIF"
        ''xmlPath = "c:\retailclarity\xmls\PARGIF"
        ''sqlUserId = "sa"
        ''sqlPassword = "PGadm01!!"
        ''dateStr = "8/19/2017"
        ''errorLog = "c:\retailclarity\errors"


        Try
            rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
            rcCON = New SqlConnection(rcConString)
            rcCON.Open()
            sql = "SELECT Server, [Database], SQLUserId, SQLPassword, XMLs, ErrorLog, SalesPlan FROM Client_Master " &
                "WHERE Client_ID = '" & client & "'"
            cmd = New SqlCommand(sql, rcCON)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Server")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then server = CStr(oTest)

                oTest = rdr("Database")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dbase = CStr(oTest)

                oTest = rdr("SQLUserId")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then sqlUserId = CStr(oTest)

                oTest = rdr("SQLPassword")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then sqlPassword = CStr(oTest)

                oTest = rdr("XMLs")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then xmlPath = CStr(oTest)

                oTest = rdr("ErrorLog")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then errorLog = CStr(oTest)

                oTest = rdr("SalesPlan")
                If Not IsDBNull(oTest) Then
                    If oTest = "Y" Then Has_Sales_Plan = True
                End If
            End While
            rcCON.Close()
            ''con.Dispose()

        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Console.ReadLine()
        End Try

        ''Dim dateStr As String
        ''dateStr = "11/5/2016"



        ''conString = "server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
        conString = "server=" & server & ";Initial Catalog=" & dbase & ";User Id=" & sqlUserId & ";Password=" & sqlPassword & ""
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)

        If IsNothing(dateStr) Or dateStr = "" Then
            Call Update_Message("No Dates to Process")
            End
        End If
        Dim finalWeek As Date
        con.Open()
        sql = "SELECT MAX(eDate) FROM Sales_Summary"
        Dim cmd0 As New SqlCommand(sql, con)
        cmd0.CommandTimeout = 120
        Dim rdr0 As SqlDataReader = cmd0.ExecuteReader
        While rdr0.Read
            If Not IsDBNull(rdr0(0)) Then finalWeek = rdr0(0)
        End While
        con.Close()
        Dim eDateArray As String() = dateStr.Split(",")
        Dim value As String
        For Each value In eDateArray
            Try
                xmlEndDate = CDate(value)
                ' put the inventory code in a loop by loc_id instead of str_id
                xmlBeginDate = DateAdd(DateInterval.Day, -6, xmlEndDate)
                thisweek = CDate(xmlEndDate)
                Begin_Date = CDate(xmlBeginDate)
                End_Date = CDate(xmlEndDate)
                thisYear = DatePart(DateInterval.Year, End_Date)
                con = New SqlConnection(conString)
                con2 = New SqlConnection(conString)
                con3 = New SqlConnection(conString)
                con4 = New SqlConnection(conString)
                con5 = New SqlConnection(conString)
                mcon = New SqlConnection(conString)

                con.Open()
                Dim todaysdate As Date = CDate(Date.Today)
                sql = "SELECT eDate FROM Calendar WHERE '" & todaysdate & "' BETWEEN sDate AND eDate AND PrdWk > 0"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    thisenddate = CDate(rdr("eDate"))
                End While
                aYearAgo = DateAdd(DateInterval.Day, -365, thisenddate)
                con.Close()

            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Initialization")
                If con.State = ConnectionState.Open Then con.Close()
            End Try




            ''  GoTo 100




            Try
                stopWatch.Start()
                Console.WriteLine("Dates from " & Begin_Date & " thru " & End_Date)
                Console.WriteLine("Updating On Order for " & client)
                '
                '
                '                                                                  ' create Inv_Summary records where needed
                Dim dte As Date
                con.Open()
                sql = "SELECT DISTINCT c.eDate, h.Loc_ID, m.Dept, m.Buyer, m.Class " & _
                    "FROM PO_Detail AS d JOIN PO_Header h ON h.PO_NO = d.PO_NO " & _
                    "INNER JOIN Item_Master AS m ON m.Sku = d.Sku " & _
                    "JOIN Calendar c ON Due_Date BETWEEN c.sDate AND c.eDate AND c.Week_Id > 0 " & _
                    "WHERE Qty_Due > 0 AND Due_Date >= '" & End_Date & "'"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                con2.Open()
                While rdr.Read
                    dte = rdr("eDate")
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    sql = "IF NOT EXISTS (SELECT eDate FROM Inv_Summary " & _
                        "WHERE Loc_Id = '" & location & "' AND Dept = '" & dept & "' " & _
                        "AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & dte & "') " & _
                        "INSERT INTO Inv_Summary (Loc_Id, Dept, Buyer, Class, Year_ID, Prd_Id, Week_Id, sDate, eDate, YrPrd, Week_Num) " & _
                        "SELECT '" & location & "','" & dept & "','" & buyer & "','" & clss & "',Year_Id, Prd_Id, PrdWk, sDate, " &
                       "eDate, YrPrd, Week_Id FROM Calendar WHERE PrdWk > 0 AND eDate = '" & dte & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()
                Dim p As String = "Create New Sales Summary Records"
                Dim m As String = ""
                Call Update_Process_Log("1", p, m, "")
                '
                '
                '
                '------------------------------ CLEAR OUT ON ORDER FOR PAST 6 MONTHS, THEN UPDATE ON ORDER FROM PO_Detail -------------------
                con.Open()
                con2.Open()
                Dim clearDate As Date = DateAdd(DateInterval.Month, -6, thisenddate)
                Dim amt As Decimal
                cnt = 0
                sql = "UPDATE Inv_Summary SET OnOrder = 0 WHERE eDate >= '" & clearDate & "' "
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                sql = "SELECT c.eDate, h.Loc_ID, m.Dept, m.Buyer, m.Class, SUM(ISNULL(Qty_Due,0) * ISNULL(Curr_Retail,0)) AS Amt " & _
                    "FROM PO_Detail AS d JOIN PO_Header h ON h.PO_NO = d.PO_NO " & _
                    "INNER JOIN Item_Master AS m ON m.Sku = d.Sku " & _
                    "JOIN Calendar c ON Due_Date BETWEEN c.sDate AND c.eDate AND c.Week_Id > 0 " & _
                    "WHERE Qty_Due > 0 " & _
                    "GROUP BY c.eDate, h.Loc_Id, m.Dept, m.Buyer, m.Class " & _
                    "order by h.Loc_id DESC, m.class DESC, edate"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                While rdr.Read
                    cnt += 1
                    If cnt Mod 100 = 0 Then Console.WriteLine("Updating OnOrder " & cnt)
                    dte = rdr("eDate")
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    amt = rdr("Amt")
                    sql = "IF NOT EXISTS (SELECT OnOrder FROM Inv_Summary " & _
                        "WHERE Loc_Id = '" & location & "' " & _
                        "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' " & _
                        "AND eDate = '" & dte & "') " & _
                        "INSERT INTO Inv_Summary (Loc_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, " & _
                        "eDate, YrPrd, Week_Num, OnOrder) " & _
                        "SELECT '" & location & "',Year_Id, Prd_Id, PrdWk, '" & dept & "','" & buyer & "','" & clss & "',sDate,'" & _
                        dte & "',YrPrd,Week_Num," & amt & " FROM Calendar WHERE '" & dte & "' BETWEEN sDate AND eDate AND Week_Id > 0 " & _
                        "ELSE " & _
                        "UPDATE Inv_Summary SET OnOrder = " & amt & " FROM Inv_Summary " & _
                   "WHERE Loc_Id = '" & location & "' AND eDate = '" & dte & "' AND Dept = '" & dept & "' " & _
                   "AND Buyer = '" & buyer & "' AND Class = '" & clss & "' "
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()
                p = "Update OnOrder"
                Call Update_Process_Log("1", p, "", "")
                '-----------------------------------  CHANGE THIS WEEKS ON ORDER TO EVERYTHING STILL DUE --------------------------------------
                con.Open()
                con2.Open()
                sql = "SELECT Loc_ID, Dept, Buyer, Class, SUM(ISNULL(OnOrder,0)) AS Amt " & _
                     "FROM Inv_Summary " & _
                     "WHERE eDate <= '" & thisenddate & "' " & _
                     "GROUP BY Loc_Id, Dept, Buyer, Class"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                While rdr.Read
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    amt = rdr("Amt")
                    sql = "UPDATE w SET OnOrder = " & amt & " FROM Inv_Summary w " & _
                        "WHERE w.eDate = '" & thisenddate & "' AND Loc_Id = '" & location & "' " & _
                        "AND Dept = '" & dept & "' AND Class = '" & clss & "' AND Buyer = '" & buyer & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()

                Call Update_Process_Log("1", "Update OnOrder Past Due", "", "")
                '------------------------------------------------------------------------------------------------------------------------------
            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, On Order")
                If con.State = ConnectionState.Open Then con.Close()
                If con2.State = ConnectionState.Open Then con2.Close()
            End Try

            '------------------------------------------------------- Inventory @ cost -----------------------------------------------------------------
            Try
                stopWatch.Start()
                Console.WriteLine("Updating inventory at cost for " & client)
                Dim location As String
                con.Open()
                sql = "SELECT DISTINCT eDate, ISNULL(Loc_ID,'NA') Loc_Id, ISNULL(Dept,'NA') Dept, ISNULL(Buyer,'NA') Buyer, " & _
                    "ISNULL(Class,'NA') Class " & _
                    "FROM Item_Inv AS d " & _
                    "INNER JOIN Item_Master AS m ON m.Sku = d.Sku " & _
                    "WHERE End_OH > 0 AND eDate >= '" & End_Date & "'"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                con2.Open()
                While rdr.Read
                    dte = rdr("eDate")
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    sql = "IF NOT EXISTS (SELECT eDate FROM Inv_Summary WHERE Loc_Id = '" & location & "' AND Dept = '" & dept & "' " & _
                        "AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & dte & "')" & _
                        "INSERT INTO Inv_Summary (Loc_Id, Dept, Buyer, Class, Year_ID, Prd_Id, Week_Id, sDate, eDate, Week_Num) " & _
                        "SELECT '" & location & "','" & dept & "','" & buyer & "','" & clss & "',Year_Id, Prd_Id, PrdWk, sDate, " &
                       "eDate, Week_Id FROM Calendar WHERE PrdWk > 0 AND eDate = '" & dte & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()

                con.Open()
                sql = "SELECT eDate, Loc_Id, Dept, Buyer, Class, SUM(ISNULL(End_OH,0) * ISNULL(Cost,0)) AS OnHand INTO #t1 FROM Item_Inv i " & _
                      "JOIN Item_Master m ON m.Sku = i.Sku WHERE i.eDate >= '" & Begin_Date & "' AND i.eDate <= '" & End_Date & "' " & _
                      "GROUP BY eDate, Loc_Id, Dept, Buyer, Class " & _
                      "UPDATE w SET w.Act_Inv_Cost = t.OnHand FROM Inv_Summary w " & _
                      "JOIN #t1 t ON t.eDate = w.edate AND t.Loc_Id = w.Loc_Id AND t.Dept = w.Dept " & _
                      "AND t.Buyer = w.Buyer AND t.Class = w.Class"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()

                sql = "SELECT eDate, Loc_Id, Dept, Buyer, Class, SUM(ISNULL(Max_OH,0) * ISNULL(Retail,0)) AS OnHand INTO #t2 FROM Item_Inv i " & _
               "JOIN Item_Master m ON m.Sku = i.Sku WHERE i.eDate >= '" & Begin_Date & "' AND i.eDate <= '" & End_Date & "' " & _
               "GROUP BY eDate, Loc_Id, Dept, Buyer, Class " & _
               "UPDATE w SET w.Max_OH_Cost = t.OnHand FROM Inv_Summary w " & _
               "JOIN #t2 t ON t.eDate = w.edate AND t.Loc_Id = w.Loc_Id AND t.Dept = w.Dept " & _
               "AND t.Buyer = w.Buyer AND t.Class = w.Class"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()
                con.Close()

                Call Update_Process_Log("1", "Update Act_Inv_Cost", "", "")

            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Inventory At Cost")
            End Try
            ''-------------------------------------------- Inventory @ Retail --------------------------------------------------------------
            Try
                stopWatch.Start()
                Console.WriteLine("Updating inventory at retail for " & client)
                con.Open()
                sql = "SELECT eDate, Loc_Id, Dept, Buyer, Class, SUM(ISNULL(End_OH,0) * ISNULL(Retail,0)) AS OnHand INTO #t1 FROM Item_Inv i " & _
                "JOIN Item_Master m ON m.Sku = i.Sku WHERE i.eDate >= '" & Begin_Date & "' AND i.eDate <= '" & End_Date & "' " & _
                "GROUP BY eDate, Loc_Id, Dept, Buyer, Class " & _
                "UPDATE w SET w.Act_Inv_Retail = t.OnHand FROM Inv_Summary w " & _
                "JOIN #t1 t ON t.eDate = w.edate AND t.Loc_Id = w.Loc_Id AND t.Dept = w.Dept " & _
                "AND t.Buyer = w.Buyer AND t.Class = w.Class"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()

                con.Close()
                Call Update_Process_Log("1", "Update Act_Inv_Retail", "", "")
            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Inventory At Retail")
            End Try
            ''------------------------------------------------ Sales --------------------------------------------------------------------
100:        Try
                Dim location As String
                stopWatch.Start()
                Console.WriteLine("Updating sales for " & client)

                '   Update Act_Sales
                con.Open()
                con2.Open()
                sql = "SELECT DISTINCT eDate, Str_ID, Loc_Id, Dept, Buyer, Class " & _
                    "FROM Item_Sales AS d " & _
                    "INNER JOIN Item_Master AS m ON m.Sku = d.Sku " & _
                    "WHERE Sold <> 0 AND eDate >= '" & End_Date & "'"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                While rdr.Read
                    dte = rdr("eDate")
                    store = rdr("Str_Id")
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    sql = "IF NOT EXISTS (SELECT eDate FROM Sales_Summary WHERE Str_Id = '" & store & "' AND Loc_Id = '" & location & "' " & _
                        "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & dte & "')" & _
                        "INSERT INTO Sales_Summary (Str_Id, Loc_Id, Dept, Buyer, Class, Year_ID, Prd_Id, Week_Id, sDate, eDate, Week_Num) " & _
                        "SELECT '" & store & "','" & location & "','" & dept & "','" & buyer & "','" & clss & "'," & _
                        "Year_Id, Prd_Id, PrdWk, sDate, eDate, Week_Id FROM Calendar WHERE PrdWk > 0 AND eDate = '" & dte & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()

                con.Open()
                con2.Open()
                sql = "SELECT Str_Id, Loc_Id, Dept, Buyer, Class, eDate, SUM(ISNULL(Sales_Retail,0)) As sales, " & _
                    "SUM(ISNULL(Sales_Cost,0)) AS cost, SUM(ISNULL(Markdown,0)) as mkdn " & _
                    "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                    "WHERE sDate >= '" & Begin_Date & "' AND eDate <= '" & End_Date & "' AND Sales_Retail <> 0 " & _
                    "GROUP BY Str_Id, Loc_Id, Dept, Buyer, Class, eDate"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                Dim amt, cost, mkdn As Decimal
                While rdr.Read
                    dte = rdr("eDate")
                    store = rdr("Str_Id")
                    location = rdr("Loc_Id")
                    dept = rdr("Dept")
                    buyer = rdr("Buyer")
                    clss = rdr("Class")
                    amt = rdr("sales")
                    cost = rdr("cost")
                    mkdn = rdr("mkdn")
                    sql = "UPDATE Sales_Summary SET Act_Sales = " & amt & ", Act_Sales_Cost = " & cost & ", Act_Mkdn = " & mkdn & " " & _
                   "WHERE Str_Id='" & store & "' AND Loc_Id = '" & location & "' AND eDate = '" & dte & "' AND Dept = '" & dept & "' " & _
                   "AND Buyer = '" & buyer & "' AND Class = '" & clss & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                End While
                con.Close()
                con2.Close()

                Call Update_Process_Log("1", "Update Sales", "", "")


            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Sales")
            End Try
            '----------------------------------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------- Sales Plan ------------------------------------------------------------
            '----------------------------------------------------------------------------------------------------------------------------------
            Try
                If Has_Sales_Plan Then
                    stopWatch.Start()
                    Console.WriteLine("Updating Sales Plan data for " & client)
                    Dim year, period, week, yrprd, prdwk As Integer
                    Dim planamt, planmkdn As Decimal
                    Dim sdate, edate As Date
                    Dim loc As String
                    Dim oTest As Object
                    con.Open()
                    sql = "UPDATE Sales_Summary SET Plan_Sales = 0, Plan_Mkdn = 0 WHERE eDate >= '" & End_Date & "' "
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                    con.Close()

                    con.Open()
                    sql = "SELECT c.Plan_Year AS Year_Id, c.Str_Id, s.Inv_Loc, c.Prd_Id, s.Week_Id, cl.sDate, cl.eDate, cl.YrPrd, PrdWk, " & _
                        "c.Dept, c.Buyer, c.Class, ISNULL(s.AMT * c.Pct * b.Pct,0) AS PlanAmt, " & _
                        "ISNULL(s.AMT * c.Pct * b.Pct,0) * ISNULL(s.Mkdn_Pct,0) AS PlanMkdn FROM Class_Pct c " & _
                        "LEFT JOIN Buyer_Pct b ON b.Plan_Year = c.Plan_Year AND b.Str_Id = c.Str_Id AND b.Prd_Id = c.Prd_Id " & _
                            "AND b.Dept = c.Dept AND b.Buyer = c.Buyer " & _
                        "LEFT JOIN Sales_Plan s ON s.Year_Id = c.Plan_Year AND s.Str_Id = c.Str_id AND s.Dept = c.Dept " & _
                        "AND s.Prd_Id = c.Prd_Id " & _
                        "LEFT JOIN Calendar cl ON cl.Year_Id = s.Year_Id AND cl.Prd_Id = s.Prd_Id AND cl.PrdWk = s.Week_Id " & _
                        "AND cl.PrdWk > 0 " & _
                        "LEFT JOIN Stores s ON s.ID = c.Str_Id " & _
                        "WHERE cl.eDate >= '" & End_Date & "' AND s.Status = 'Active'"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    cnt = 0
                    con2.Open()
                    While rdr.Read
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine("Updated Plan_Sales for " & cnt & " records")
                        year = rdr("Year_Id")
                        store = rdr("Str_Id")
                        oTest = rdr("Inv_loc")
                        If IsDBNull(oTest) Then oTest = store
                        loc = CStr(oTest)
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
                        sql = "IF NOT EXISTS (SELECT Plan_Sales FROM Sales_Summary WHERE Year_Id = " & year & " " & _
                            "AND Str_Id = '" & store & "' AND Loc_Id = '" & loc & "' AND AND Prd_Id = " & period & " " & _
                            "AND Week_Id = " & week & " AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "') " & _
                            "INSERT INTO Sales_Summary (Str_Id, Loc_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, eDate, " & _
                                "YrPrd, Week_Num, Plan_Sales, Plan_Mkdn) " & _
                            "SELECT '" & store & "','" & loc & "','" & year & "," & period & "," & prdwk & ",'" & dept & "','" & buyer & "','" &
                            clss & "','" & sdate & "','" & edate & "'," & yrprd & "," & week & "," & planamt & "," & planmkdn & " " & _
                            "ELSE " & _
                            "UPDATE Sales_Summary SET Plan_Sales = " & planamt & ", Plan_Mkdn = " & planmkdn & " " & _
                            "WHERE Str_Id = '" & store & "' AND Loc_Id = '" & loc & "' AND Year_Id = " & year & " AND Prd_Id = " & period & " " & _
                            "AND Week_Id = " & week & " AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' "
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                    End While
                    con.Close()
                    con2.Close()

                    Call Update_Process_Log("2", "Plan Sales", "", "")

                    Console.WriteLine("Updating Plan_WksOH for " & client)
                    con.Open()
                    sql = "UPDATE w SET w.Plan_WksOH = b.Plan_WksOH FROM Sales_Summary w " & _
                            "JOIN Buyer_PCT b ON b.Str_Id = w.Str_Id AND w.Year_Id = Plan_Year " & _
                            "AND b.Prd_Id = w.Prd_Id AND b.Dept = w.Dept AND b.Buyer = w.Buyer"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                    con.Close()

                    thisweek = xmlEndDate
                    Dim nextWeek As Date = DateAdd(DateInterval.Day, 7, thisweek)          ' thisweek is the End Date passed to this program
                    Do Until thisweek >= finalWeek                                         ' finalWeek is the last eDate in OTB

                        Console.WriteLine("Updating Plan_Inv_Retail for " & client & " week " & nextWeek)

                        sql = "UPDATE Sales_summary SET Sales_Summary.Plan_Inv_Retail = ( " & _
                            "SELECT SUM(ISNULL(o2.Plan_Sales,0)) + SUM(ISNULL(o2.Plan_Mkdn,0)) FROM Sales_summary AS o2 " & _
                            "WHERE o2.Str_Id = Sales_summary.Str_Id AND o2.Dept = Sales_summary.Dept " & _
                            "AND o2.Buyer = Sales_summary.Buyer AND o2.Class = Sales_summary.Class " & _
                            "AND o2.sDate BETWEEN Sales_summary.sDate AND DATEADD(week,Sales_summary.Plan_WksOH-1,Sales_summary.sDate)) " & _
                            "WHERE Sales_summary.Buyer <> '*' AND Sales_summary.Class <> '*' " & _
                            "AND Sales_summary.eDate = '" & thisweek & "'"
                        con.Open()
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                        thisweek = DateAdd(DateInterval.Day, 7, thisweek)
                        nextWeek = DateAdd(DateInterval.Day, 7, nextWeek)
                    Loop

                    Call Update_Process_Log("2", "Update Plan_Inv_Retail", "", "")

                    thisweek = xmlEndDate                                      ' End_Date is the XML End Date passed to this program
                    Dim prevWeek As Date = DateAdd(DateInterval.Day, -7, thisweek)
                    Do Until thisweek >= finalWeek                                         ' finalWeek is the last eDate in Sales_summary
                        '                                                                   thisEndDate is this Weeks End Date
                        Console.WriteLine("Updating Projected_Inv and Inv_Summary for " & client & " week " & thisweek)

                        If thisweek = thisenddate Then
                            ''sql = "UPDATE Inv_Summary SET Projected_Inv = ISNULL(Act_Inv_Retail,0) " & _
                            ''    "+ ISNULL(OnOrder,0) " & _
                            ''    "- ISNULL(Plan_Sales,0) " & _
                            ''    "+ ISNULL(Act_Sales,0) " & _
                            ''    "- ISNULL(Plan_Mkdn,0) " & _
                            ''    "+ ISNULL(Act_Mkdn,0) WHERE eDate = '" & thisenddate & "'"
                            sql = "UPDATE Inv_Summary SET Projected_Inv = ISNULL(Act_Inv_Retail,0) " & _
                                "+ ISNULL(OnOrder,0) WHERE eDate = '" & thisenddate & "'"
                        Else
                            ''sql = "UPDATE f SET f.Projected_Inv = ISNULL(p.Projected_Inv,0) " & _
                            ''    "+ ISNULL(f.OnOrder,0) " & _
                            ''    "- ISNULL(f.Plan_Sales,0) " & _
                            ''    "+ ISNULL(f.Act_Sales,0) " & _
                            ''    "- ISNULL(f.Plan_Mkdn,0) " & _
                            ''    "+ ISNULL(f.Act_Mkdn,0) " & _
                            ''    "FROM Weekly_Summary AS f INNER JOIN Weekly_Summary AS p " & _
                            ''    "ON f.Str_Id = p.Str_Id AND p.Dept = f.Dept " & _
                            ''    "AND p.Buyer = f.Buyer AND p.Class = f.Class " & _
                            ''    "WHERE f.eDate = '" & thisweek & "' " & _
                            ''    "AND p.eDate = '" & prevWeek & "'"
                            sql = "UPDATE f SET f.Projected_Inv = ISNULL(p.Projected_Inv,0) " & _
                            "+ ISNULL(f.OnOrder,0) FROM Inv_Summary f " & _
                            "INNER JOIN Inv_Summary p ON p.Str_Id = f.Str_Id AND p.Dept = f.Dept " & _
                            "AND p.BUYER = f.Buyer AND f.Class = p.Class " & _
                            "WHERE f.eDate = '" & thisweek & "' AND p.eDate = '" & prevWeek & "'"
                        End If
                        con.Open()
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                        prevWeek = thisweek
                        thisweek = DateAdd(DateInterval.Day, 7, thisweek)
                    Loop

                    Call Update_Process_Log("2", "Update Projected_Inv", "", "")

                    Console.WriteLine("Updating Projected_WksOH for " & client)
                    Dim actInv As Int32 = 0
                    Dim planSales As Int32 = 0
                    Dim numWeeks As Int16
                    Dim eDatex As Date
                    con.Open()
                    sql = "Select Loc_Id, Dept, Buyer, Class, eDate, Act_Inv_Retail FROM Inv_Summary " & _
                        "WHERE Act_Inv_Retail > 0 AND eDate >= '" & Begin_Date & "' AND eDate <= '" & End_Date & "' " & _
                        "ORDER BY eDate, Str_Id, Dept, Buyer, Class"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        location = rdr("Loc_Id")
                        dept = rdr("Dept")
                        buyer = rdr("Buyer")
                        clss = rdr("Class")
                        eDatex = rdr("eDate")
                        oTest = rdr("Act_Inv_Retail")
                        If IsNumeric(oTest) Then actInv = oTest Else actInv = 0
                        Console.WriteLine(eDatex & " " & location & " " & dept & " " & buyer & " " & clss & " " & oTest)
                        planSales = 0
                        numWeeks = 0
                        con2.Open()
                        sql = "SELECT SUM(ISNULL(Plan_Sales,0) + ISNULL(Plan_Mkdn,0)) As Sales FROM Sales_Summary " & _
                            "WHERE Loc_Id = '" & location & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND " & _
                            "Class = '" & clss & "' AND eDate > '" & eDatex & "' "
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
                                sql = "UPDATE Inv_Summary SET Projected_WksOH = " & numWeeks & " WHERE Loc_Id = '" & location & "' " & _
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

                    Call Update_Process_Log("2", "Update Projeted_WksOH", "", "")

                    Call Update_Act_WksOH()
                End If
            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Update Plan Week On Hand")
            End Try
            '-------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------- End Sales Plan Code --------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------
            Try

                con = New SqlConnection(rcConString)
                con.Open()
                sql = "UPDATE Client_Master SET Last_Summary_Update = '" & Date.Now & "' WHERE Client_Id = '" & client & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
                con.Dispose()
            Catch ex As Exception
                Dim theMessage As String = ex.Message
                Console.WriteLine(ex.Message)
                Dim el As New Client_Weekly_Summary.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Client_Weekly_Summary, Update Client_Master")
            End Try
            Console.WriteLine("Client_Weekly_Summary complete!")
        Next
        '
        '                                                           Process Optional Modules
        '
        Process.Start(rcExePath & "\Optional_Module_Processing.exe")

    End Sub

    Private Sub Update_Message(m)
        stopWatch.Stop()
        thisDateTime = CDate(Now)
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim title As String = "Client_Weekly_Summary"
        Dim message As String = m & " Completed in " & ts.Minutes & " minutes and " & ts.Seconds & " seconds"
        mcon.Open()
        sql = "INSERT INTO Message (mDate, Title, Message) " & _
            "SELECT '" & thisDateTime & "','" & title & "','" & message & "'"
        cmd = New SqlCommand(sql, mcon)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        mcon.Close()
    End Sub

    Private Sub Update_Act_WksOH()
        stopWatch.Start()
        Dim theDate As Date
        Dim row, sRow As DataRow
        con.Open()
        sql = "SELECT DISTINCT Str_Id FROM Inv_Summary ORDER BY Loc_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            store = rdr("Str_Id")
            con2.Open()
            sql = "SELECT DISTINCT Dept FROM Inv_Summary WHERE Loc_Id = '" & store & "' ORDER By Dept"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                Dim dept = rdr2("Dept")
                con3.Open()
                sql = "SELECT DISTINCT Buyer FROM Inv_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' Order BY Buyer"
                cmd = New SqlCommand(sql, con3)
                rdr3 = cmd.ExecuteReader
                While rdr3.Read
                    Dim buyer As String = rdr3("Buyer")
                    sql = "SELECT DISTINCT Class FROM Inv_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' ORDER BY Class"
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
                        sql = "SELECT eDate, ISNULL(Act_Inv_Retail,0) AS Inv FROM Inv_Summary i " & _
                            "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                            "WHERE s.Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
                            "AND eDate BETWEEN '" & aYearAgo & "' AND '" & thisenddate & "' " & _
                            "AND Class = '" & clss & "' AND ISNULL(Act_Inv_Retail,0) > 0 AND Act_WksOH IS NULL ORDER BY eDate"
                        cmd = New SqlCommand(sql, con5)
                        rdr5 = cmd.ExecuteReader
                        While rdr5.Read
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
                        sql = "SELECT eDate, ISNULL(Act_Sales,0) Sales FROM Sales_Summary " & _
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
                                    sql = "UPDATE i SET Act_WksOH = " & cnt & " FROM Inv_Summary s " & _
                                        "JOIN Stores s ON s.Inv_Loc = i.Loc_Id WHERE s.Str_Id = '" & store & "' " & _
                                        "AND Dept = '" & dept & "' AND Class = '" & clss & "' AND Buyer = '" & buyer & "' " & _
                                        "AND eDate = '" & theDate & "'"
                                    cmd = New SqlCommand(sql, con5)
                                    cmd.ExecuteNonQuery()
                                    con5.Close()
                                End If
                            End If
                        Next
                        ' MsgBox("Ended idx2=" & indx2 & " rows=" & salesTable.Rows.Count & " " & store & " " & dept & " " & buyer & " " & clss)
                    End While
                    con4.Close()
                    ' MsgBox("Endedxx " & store & " " & dept & " " & buyer & " " & clss)
                End While
                con3.Close()
            End While
            con2.Close()
        End While
        con.Close()

        Call Update_Process_Log("1", "Update Act_WksOH", "", "")
    End Sub

    Function FindRow(ByVal dt As DataTable, ByVal eDate As String) As Integer
        For i As Integer = 0 To dt.Rows.Count - 1
            If dt.Rows(i)("eDate") = eDate Then Return i
        Next
        Return -1
    End Function

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        stopWatch.Stop()
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "Client_Weekly_Summary"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & Date.Now & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Process_Optional_Modules()
        Dim rcCon As New SqlConnection(rcConString)
        rcCon.Open()
        sql = "SELECT SalesPlan FROM Client_Master WHERE Client_Id = '" & client & "'"
        cmd = New SqlCommand(sql, rcCon)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("SalesPlan") = "Y" Then
                Call Check_For_Completed_Period()
            End If
        End While
    End Sub

    Private Sub Check_For_Completed_Period()
        Dim prd As Integer = 0
        Dim lastPlanPrd, maxPrd As Integer
        con.Open()
        sql = "SELECT MAX(Prd_Id) as Prd_Id FROM Sales_Plan WHERE Status = 'Active' AND Year_Id = " & thisYear & " " & _
            "AND ISNULL(Amt,0) > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            lastPlanPrd = rdr("Prd_Id")
        End While
        con.Close()

        con.Open()

    End Sub

End Module

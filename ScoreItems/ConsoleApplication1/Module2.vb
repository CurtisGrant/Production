Imports System
Imports System.Data.SqlClient
Imports System.Xml
Imports System.IO
Module Module2
    Private con, con2, con3 As SqlConnection
    Private cmd As SqlCommand
    Private mrdr, rdr, rdr2 As SqlDataReader
    Public client As String
    Public errorLog As String
    Private sql As String
    Private oTest As Object
    Private todaysEDate, lySDate, thisSDate, thisEDate, twoWeeksAgo, lastUpdate, startDate As Date
    Private weeksBack As Integer
    Private thisStore As String
    Private scoreTbl, scoreTbl2, scoreTbl3, storeTbl, dateTbl, atbl, btbl, ctbl, tbl, tbl2, universeTbl As DataTable
    Private row, row2, row3, foundRow As DataRow
    Private conString As String
    Private stopwatch, stopWatch2 As Stopwatch

    Sub Main()
        stopwatch = New Stopwatch
        stopWatch2 = New Stopwatch
        stopwatch.Start()
        stopWatch2.Start()
        dateTbl = New DataTable
        dateTbl.Columns.Add("sDate", GetType(System.DateTime))
        dateTbl.Columns.Add("eDate", GetType(System.DateTime))
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
                Case XmlNodeType.EndElement
                    'Console.WriteLine("</" & xmlReader.Name)
            End Select
            If fld = "SERVER" Then server = valu
            If fld = "EXEPATH" Then exePath = valu
            If fld = "PD" Then passWord = valu
            If fld = "CLIENTID" Then client = valu
        End While
        conString = "Server=" & server & ";Initial Catalog=RCClient;User Id=sa;Password=" & passWord & ""
        '' con = New SqlConnection(conString)
        Dim dbase, sqlUserId, sqlPassword As String
        Dim masterCon As New SqlConnection(conString)
        Dim mcon2 As New SqlConnection(conString)
        masterCon.Open()
        sql = "SELECT Client_Id, Server, [Database], SQLUserId, SQLPassword, ErrorLog, Last_Scored FROM Client_Master " & _
            "WHERE Status = 'Active' AND ItemScoring = 'Y'"
        ''sql = "select client_id, server, [database], sqluserid, sqlpassword, errorlog, last_scored from client_master where client_id='demo3'"
        cmd = New SqlCommand(sql, masterCon)
        mrdr = cmd.ExecuteReader
        While mrdr.Read
            Client = mrdr("Client_Id")
            server = mrdr("Server")
            dbase = mrdr("Database")
            sqlUserId = mrdr("sqlUserId")
            sqlPassword = mrdr("SQLPassword")
            errorLog = mrdr("ErrorLog")
            conString = "Server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)

            Call Setup()




            ''Dim r As DataRow = dateTbl.NewRow
            ''r("eDate") = "11/5/2016"
            ''dateTbl.Rows.Add(r)





            stopwatch.Start()
            Dim typ As String
            typ = "Sku"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "Dept"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "Class"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Get_GrossMargin_Params(typ)
            stopwatch.Start()
            typ = "Sku"
            Call Get_Turns_Params(typ)
            stopwatch.Start()
            typ = "Dept"
            Call Get_Turns_Params(typ)
            stopwatch.Start()
            typ = "Class"
            Call Get_Turns_Params(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Get_Turns_Params(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Get_Turns_Params(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Get_Turns_Params(typ)

            stopwatch.Start()
            typ = "Sku"
            Call Get_UnitCost_Params(typ)
            stopwatch.Start()
            typ = "Dept"
            Call Get_UnitCost_Params(typ)
            stopwatch.Start()
            typ = "Class"
            Call Get_UnitCost_Params(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Get_UnitCost_Params(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Get_UnitCost_Params(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Get_UnitCost_Params(typ)

            stopwatch.Start()
            typ = "Sku"
            Call Score_GrossMargin(typ)
            stopwatch.Start()
            typ = "Dept"
            Call Score_GrossMargin(typ)
            stopwatch.Start()
            typ = "Class"
            Call Score_GrossMargin(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Score_GrossMargin(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Score_GrossMargin(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Score_GrossMargin(typ)

            stopwatch.Start()
            typ = "Sku"
            Call Score_Turns(typ)
            stopwatch.Start()
            typ = "Dept"
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Call Score_Turns(typ)
            stopwatch.Start()
            typ = "Class"
            Call Score_Turns(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Score_Turns(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Score_Turns(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Score_Turns(typ)

            stopwatch.Start()
            typ = "Sku"
            Call Score_Cost(typ)
            stopwatch.Start()
            typ = "Dept"
            Call Score_Cost(typ)
            stopwatch.Start()
            typ = "Class"
            Call Score_Cost(typ)
            stopwatch.Start()
            typ = "Buyer"
            Call Score_Cost(typ)
            stopwatch.Start()
            typ = "PLine"
            Call Score_Cost(typ)
            stopwatch.Start()
            typ = "Vendor_Id"
            Call Score_Cost(typ)
50:
            mCon2.Open()
            sql = "UPDATE Client_Master SET Last_Scored='" & Date.Now & "' WHERE Client_Id = '" & Client & "'"
            cmd = New SqlCommand(sql, mCon2)
            cmd.ExecuteNonQuery()
            mCon2.Close()

            If con.State = ConnectionState.Open Then con.Close()
            con.Open()
            stopwatch2.Stop()
            Dim thisDateTime As DateTime = CDate(Now)
            Dim ts As TimeSpan = stopwatch2.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim title As String = "Score Items"
            Dim message As String = "Scored Items in " & ts.Hours & " hours " & ts.Minutes & " minutes and " & ts.Seconds & " seconds"
            sql = "INSERT INTO Message (mDate, Title, Message) " & _
                "SELECT '" & thisDateTime & "','" & title & "','" & message & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

        End While
        masterCon.Close()


    End Sub

    Private Sub Setup()
        Try
            Dim pct As Decimal
            scoreTbl = New DataTable
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Code"
            scoreTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = scoreTbl.Columns("Code")
            scoreTbl.PrimaryKey = PrimaryKey
            scoreTbl.Columns.Add("Break$", GetType(System.Decimal))
            scoreTbl.Columns.Add("%GM", GetType(System.Decimal))
            scoreTbl.Columns.Add("%SKU", GetType(System.Decimal))
            scoreTbl.Columns.Add("CumGM%", GetType(System.Decimal))
            scoreTbl.Columns.Add("Skus", GetType(System.Int32))
            storeTbl = New DataTable
            storeTbl.Columns.Add("Store", GetType(System.String))
            Dim eDate, nextEdate As Date
            con.Open()
            '
            '                   Get max edate to determine how many weeks need to be scored and store the edates in dateTbl
            '
            sql = "SELECT MAX(eDate) AS eDate FROM Item_Sales WHERE Score IS NOT NULL"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("eDate")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then eDate = CDate(oTest)
            End While
            con.Close()

            con.Open()
            sql = "SELECT eDate FROM Calendar WHERE CONVERT(Date,GETDATE()) BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysEDate = rdr("eDate")
            End While
            con.Close()

            If eDate = "12:00:00 AM" Then eDate = DateAdd(DateInterval.Day, -14, todaysEDate)


            nextEdate = DateAdd(DateInterval.Day, 7, eDate)
            If nextEdate <= todaysEDate Then
                Do Until nextEdate = todaysEDate
                    row = dateTbl.NewRow
                    row("sDate") = DateAdd(DateInterval.Day, -6, nextEdate)
                    row("eDate") = nextEdate
                    dateTbl.Rows.Add(row)
                    nextEdate = DateAdd(DateInterval.Day, 7, nextEdate)
                Loop
            End If


            For Each row In dateTbl.Rows
                oTest = row("sDate") & " " & row("eDate")
            Next
            '
            '                               Get the percentages that correspond to scores A,B,C,D & E
            '
            Dim arr() As String = {"A", "B", "C", "D", "E"}
            '' Dim arr2() As Decimal = {0.5, 0.3, 0.15, 0.04, 0.01}
            For Each element As String In arr
                row = scoreTbl.NewRow
                row(0) = element
                row(1) = 0
                scoreTbl.Rows.Add(row)
            Next
            Dim tpct As Decimal = 0
            For i As Integer = 0 To 4
                con.Open()
                sql = "SELECT Weeks, " & arr(i) & "Pct AS Pct FROM Score_Setup_History WHERE ID = " & _
                    "(SELECT MAX(ID) FROM Score_Setup_History)"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    oTest = rdr("Pct")
                    If Not IsDBNull(oTest) Then pct = CDec(oTest)
                    scoreTbl.Rows(i).Item("%GM") = pct
                    tpct += pct
                End While
                scoreTbl.Rows(i).Item("CUMGM%") = tpct
                con.Close()
            Next
            '
            '                                     Get the number of weeks we want in our calculations
            '

            con.Open()
            sql = "SELECT Weeks FROM Score_Setup_History WHERE ID = " & _
                "(SELECT MAX(ID) FROM Score_Setup_History)"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                startDate = DateAdd(DateInterval.WeekOfYear, rdr("Weeks") * -1, todaysEDate)
                weeksBack = rdr("Weeks")
            End While
            con.Close()
            '
            '                         Get all the stores
            '
            con.Open()
            sql = "SELECT Str_Id FROM Stores WHERE Status = 'Active' ORDER BY Str_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = storeTbl.NewRow
                row("Store") = rdr("Str_Id")
                storeTbl.Rows.Add(row)
            End While
            con.Close()

            Call Update_Process_Log(1, "Setup", "", "")

        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Setup")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Get_GrossMargin_Params(typ)                                     ' Compute paramaters for Gross Margin
        Try
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")
                For Each drow As DataRow In dateTbl.Rows
                    thisEDate = drow("eDate")
                    Console.WriteLine("Computing  Gross Margin period " & thisEDate)
                    Dim tgm As Integer = 0
                    Dim summgn As Integer = 0
                    Dim tcnt As Integer = 0
                    Dim sumcnt As Integer = 0
                    Dim records As Integer = 0
                    Dim val As Decimal
                    '
                    '                                       Get the date range
                    '
                    con.Open()
                    sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE Week_Id > 0 AND eDate = '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSDate = rdr("sDate")
                        thisEDate = rdr("eDate")
                    End While
                    con.Close()

                    tbl = New DataTable
                    tbl.Columns.Add(typ, GetType(System.String))
                    tbl.Columns.Add("gm", GetType(System.Int16))
                    tbl.Columns.Add("pctmgn", GetType(System.Decimal))
                    tbl.Columns.Add("pctsku", GetType(System.Decimal))


                    atbl = New DataTable
                    atbl.Columns.Add(typ, GetType(System.String))
                    atbl.Columns.Add("GrossMargin", GetType(System.Decimal))
                    atbl.Columns.Add("Pct", GetType(System.Decimal))
                    atbl.Columns.Add("iPct", GetType(System.Decimal))
                    '
                    '                        Zero Score Table Gross Margin
                    '
                    For i As Integer = 0 To scoreTbl.Rows.Count - 1
                        scoreTbl.Rows(i).Item(1) = 0
                        scoreTbl.Rows(i).Item(5) = 0
                    Next
                    '
                    '                        Get the gross margin for every sku within the date range and stick it in a table in descending GM sequence
                    ' 
                    con.Open()
                    sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Retail - Sales_Cost),0) AS grossmargin " & _
                            "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                            "WHERE Str_id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' " & _
                            "AND eDate <= '" & thisEDate & "' AND Sales_Retail > 0 " & _
                            "GROUP BY m." & typ & " ORDER BY grossmargin DESC"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        row = atbl.NewRow
                        oTest = rdr("typ")
                        row(typ) = rdr("typ")
                        val = rdr("grossmargin")
                        If val > 0 Then
                            tcnt += 1
                            summgn += val
                        End If
                        row("GrossMargin") = val
                        atbl.Rows.Add(row)
                    End While
                    con.Close()

                    records = atbl.Rows.Count                ' Get out if there are no records to process
                    If records = 0 Then Exit Sub

                    Dim cnt As Integer = 0
                    Dim pct As Decimal
                    Dim view As New DataView(atbl)
                    view.Sort = "GrossMargin DESC"                     ' sort the table in descending gross margin sequence
                    For Each rw As DataRowView In view
                        cnt += 1
                        oTest = rw(0)
                        val = rw("grossmargin")
                        tgm += val                                     ' add gross margin to total gross margin
                        pct = tgm / summgn                             ' compute this items percentage of the total gross margin
                        rw("Pct") = pct                                ' save the result in a table
                        rw("iPct") = cnt / tcnt                        ' save this items percentage of the total record count
                        For j As Integer = 0 To 4                      ' compare this items GM percentage to the percentage that correspond with A-E
                            'oTest = scoreTbl(j).Item(1)                ' and save its GM in the score table when its >=
                            'oTest = scoreTbl(j).Item(4)
                            Dim it As Decimal = scoreTbl(j).Item(4)
                            oTest = scoreTbl.Rows(j).Item("Break$")
                            If pct >= scoreTbl.Rows(j).Item("CumGM%") And scoreTbl.Rows(j).Item("Break$") = 0 Then
                                scoreTbl.Rows(j).Item("Break$") = val
                            End If
                        Next
                    Next
                    Dim Acnt, bCnt, cCnt, Dcnt, Ecnt, gm As Decimal
                    Acnt = 0
                    bCnt = 0
                    cCnt = 0
                    Dcnt = 0
                    Ecnt = 0
                    '
                    '                            Loop thru aTbl and determine how many A's, B's, C's, D's and E's we have based on each rows GM
                    '

                    For Each row As DataRow In atbl.Rows
                        'If typ = "Dept" Then
                        '    oTest = scoreTbl.Rows(0).Item(1)
                        'End If
                        gm = row("grossmargin")
                        If gm >= scoreTbl.Rows(0).Item("Break$") Then Acnt += 1
                        If gm >= scoreTbl.Rows(1).Item("Break$") And gm < scoreTbl.Rows(0).Item("Break$") Then bCnt += 1
                        If gm >= scoreTbl.Rows(2).Item("Break$") And gm < scoreTbl.Rows(1).Item("Break$") Then cCnt += 1
                        If gm >= scoreTbl.Rows(3).Item("Break$") And gm < scoreTbl.Rows(2).Item("Break$") Then Dcnt += 1
                        If gm >= scoreTbl.Rows(4).Item("Break$") And gm < scoreTbl.Rows(3).Item("Break$") Then Ecnt += 1
                    Next
                    '
                    '                           Save the record count and percentage of total record count in the scoreTbl by Score
                    '
                    If tcnt > 0 Then
                        scoreTbl.Rows(0).Item("%Sku") = CDec(Acnt / tcnt)
                        scoreTbl.Rows(1).Item("%Sku") = CDec(bCnt / tcnt)
                        scoreTbl.Rows(2).Item("%Sku") = CDec(cCnt / tcnt)
                        scoreTbl.Rows(3).Item("%Sku") = CDec(Dcnt / tcnt)
                        scoreTbl.Rows(4).Item("%Sku") = CDec(Ecnt / tcnt)
                    End If
                    scoreTbl.Rows(0).Item("Skus") = Acnt
                    scoreTbl.Rows(1).Item("Skus") = bCnt
                    scoreTbl.Rows(2).Item("Skus") = cCnt
                    scoreTbl.Rows(3).Item("Skus") = Dcnt
                    scoreTbl.Rows(4).Item("Skus") = Ecnt
                    Dim id, type, code As String
                    Dim amt, ipct As Decimal
                    Dim now As DateTime = Date.Now
                    For Each row As DataRow In scoreTbl.Rows
                        amt += row(2)
                    Next
                    con.Open()
                    For Each row As DataRow In scoreTbl.Rows
                        id = "GrossMargin"
                        ''type = typ
                        code = row(0)
                        oTest = row(1)
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = row(1) Else amt = 0
                        oTest = row(2)
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(2) Else pct = 0
                        oTest = row(3)
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ipct = row(3) Else ipct = 0
                        Dim dbTyp As String = ""
                        Select Case typ
                            Case "Sku"
                                dbTyp = "Item"
                            Case "Dept"
                                dbTyp = "Dept"
                            Case "Class"
                                dbTyp = "Class"
                            Case "Buyer"
                                dbTyp = "Buyer"
                            Case "Vendor_Id"
                                dbTyp = "Vendor"
                            Case "PLine"
                                dbTyp = "PLine"
                        End Select
                        sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                            "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "') " & _
                            "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                            "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','" & dbTyp & "','" & _
                            code & "'," & amt & "," & pct & "," & ipct & ",'" & now & "' " & _
                            "ELSE " & _
                            "UPDATE Score_Params SET Value1 = " & amt & ", Value2 = " & pct & ", Value3 = " & ipct & ", Last_Update = '" & now & "' " & _
                            "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "' " & _
                            "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                    Next
                    con.Close()
                Next
            Next
            Dim m As String = "Get GrossMargin " & typ
            Call Update_Process_Log(1, m, "", "")

        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Get GrossMargin Parameters " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub
    Private Sub Get_Turns_Params(typ)
        '                                                    Compute paramaters from Turns
        Try
            Dim dbTyp As String
            Select Case typ
                Case "Sku"
                    dbTyp = "Item"
                Case "Dept"
                    dbTyp = "Dept"
                Case "Class"
                    dbTyp = "Class"
                Case "Buyer"
                    dbTyp = "Buyer"
                Case "Vendor_Id"
                    dbTyp = "Vendor_Id"
                Case "PLine"
                    dbTyp = "PLine"
            End Select

            Console.WriteLine("Computing Turns Parameters " & typ)
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")

                For Each drow As DataRow In dateTbl.Rows
                    thisEDate = drow("eDate")
                    Console.WriteLine("Computing Turns period " & thisEDate)
                    con.Open()
                    sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE eDate = '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSDate = rdr("sDate")
                        thisEDate = rdr("eDate")
                    End While
                    con.Close()

                    scoreTbl2 = New DataTable
                    scoreTbl2.Columns.Add("Dept", GetType(System.String))
                    Dim PrimaryKey2(1) As DataColumn
                    PrimaryKey2(0) = scoreTbl2.Columns("Dept")
                    scoreTbl2.PrimaryKey = PrimaryKey2
                    scoreTbl2.Columns.Add("F", GetType(System.Decimal))
                    scoreTbl2.Columns.Add("M", GetType(System.Decimal))
                    scoreTbl2.Columns.Add("S", GetType(System.Decimal))
                    scoreTbl2.Columns.Add("records", GetType(System.Int32))

                    Dim now As DateTime = Date.Now
                    If typ = "Sku" Then
                        con.Open()
                        sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            row2 = scoreTbl2.NewRow
                            oTest = rdr("Dept")
                            row2(0) = rdr("Dept")
                            scoreTbl2.Rows.Add(row2)
                        End While
                        con.Close()

                        btbl = New DataTable
                        btbl.Columns.Add("cnt", GetType(System.Int32))
                        btbl.Columns.Add("dept", GetType(System.String))
                        btbl.Columns.Add("Sku", GetType(System.String))
                        btbl.Columns.Add("invcost", GetType(System.Decimal))
                        btbl.Columns.Add("wksoh", GetType(System.Int32))
                        btbl.Columns.Add("COGS", GetType(System.Decimal))
                        btbl.Columns.Add("turns", GetType(System.Decimal))
                        Dim cnt, wks, fCnt, mCnt As Integer
                        Dim val, cogs, invcost As Decimal
                        Dim dept, item As String
                        Dim prevDept As String = ""
                        Dim foundRow As DataRow
                        con.Open()
                        sql = "SELECT m.Dept, d.Sku, CONVERT(Decimal,0) AS invcost, ISNULL(SUM(Sales_Cost),0) As COGS, " & _
                          "CONVERT(Integer,0) AS wksOH INTO #t1 FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                          "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                          "GROUP BY Str_Id, m.Dept, d.Sku " & _
                          "SELECT m.Dept, i.Sku, ISNULL(SUM(Max_OH * Curr_Retail),0) AS invcost, " & _
                          "COUNT(*) As wksOH INTO #t2 FROM Item_Inv i " & _
                          "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                          "JOIN Item_Master m ON m.Sku = i.Sku " & _
                          "WHERE s.Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                          "GROUP BY Str_Id, m.Dept, i.Sku " & _
                          "UPDATE #t1 SET #t1.invcost = #t2.invcost, #t1.wksOH = #t2.wksOH FROM #t1 " & _
                          "JOIN #t2 ON #t2.Dept = #t1.Dept AND #t2.Sku = #t1.Sku " & _
                          "SELECT Dept, Sku, invcost, wksoh, COGS FROM #t1 ORDER BY Dept"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            dept = rdr("Dept")
                            If dept <> prevDept Then
                                If prevDept = "" Then
                                    prevDept = dept
                                Else
                                    foundRow = scoreTbl2.Rows.Find(prevDept)
                                    foundRow(4) = cnt
                                    cnt = 0
                                    prevDept = dept
                                End If
                            End If
                            ''oTest = rdr("COGS")
                            ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            ''    If IsNumeric(oTest) Then
                            ''        If CDec(oTest) > 0 Then cnt += 1
                            ''    End If
                            ''End If
                            If rdr("COGS") > 0 Then
                                cnt += 1
                            End If
                            row = btbl.NewRow
                            row("cnt") = cnt
                            row("dept") = rdr("Dept")
                            row("Sku") = rdr("Sku")
                            row("invcost") = rdr("invcost")
                            row("wksoh") = rdr("wksOH")
                            row("COGS") = rdr("COGS")
                            row("turns") = 0
                            btbl.Rows.Add(row)
                            item = rdr("Sku")

                        End While
                        con.Close()

                        foundRow = scoreTbl2.Rows.Find(prevDept)              ' Get out if there is no data to process
                        If IsNothing(foundRow) Then Exit Sub

                        foundRow(4) = cnt
                        prevDept = ""
                        For Each row As DataRow In btbl.Rows
                            wks = row("wksoh")
                            cogs = row("COGS")
                            invcost = row("invcost")
                            If wks > 0 And invcost > 0 Then
                                val = (row("cogs") / (invcost / row("wksoh"))) * (52 / row("wksoh"))
                                row("turns") = val
                            Else : row("turns") = 0
                            End If
                        Next
                        cnt = 0
                        fCnt = 0
                        mCnt = 0
                        Dim view As New DataView(btbl)
                        view.Sort = "Dept ASC, turns DESC, Sku ASC"
                        For Each rw As DataRowView In view
                            cnt += 1
                            ''If cnt Mod 1000 = 0 Then
                            ''    Console.WriteLine(cnt & "  " & val)
                            ''End If
                            dept = rw("dept")
                            val = rw("turns")
                            item = rw("Sku")
                            If cnt = fCnt Then
                                foundRow = scoreTbl2.Rows.Find(prevDept)
                                foundRow(1) = val
                            End If
                            If cnt = mCnt Then
                                foundRow = scoreTbl2.Rows.Find(prevDept)
                                foundRow(2) = val
                                foundRow(3) = 0.01
                            End If
                            If dept <> prevDept Then
                                foundRow = scoreTbl2.Rows.Find(dept)
                                fCnt = foundRow(4) / 3
                                mCnt = fCnt * 2
                                cnt = 1
                                prevDept = dept
                            End If
                        Next
                        '                                   Save to Sql table
                        Dim id, type, code As String
                        Dim pct, pct2, pct3 As Decimal


                        con.Open()
                        For Each row As DataRow In scoreTbl2.Rows
                            id = "Turns"
                            type = "Item"
                            code = row(0)
                            oTest = row(1)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(1) Else pct = 0
                            oTest = row(2)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct2 = row(2) Else pct2 = 0
                            oTest = row(3)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct3 = row(3) Else pct3 = 0
                            sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                                "AND eDate = '" & thisEDate & "' AND Type = 'Item' AND Code = '" & code & "') " & _
                                "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                                "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','Item','" & _
                                code & "'," & pct & "," & pct2 & "," & pct3 & ",'" & now & "' " & _
                                "ELSE " & _
                                "UPDATE Score_Params SET Value1 = " & pct & ", Value2 = " & pct2 & ", Value3 = " & pct3 & ", Last_Update = '" & now & "' " & _
                                "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = '" & code & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                        Next
                        con.Close()
                    Else
                        Dim cnt As Integer = 0
                        btbl = New DataTable
                        btbl.Columns.Add("dept", GetType(System.String))
                        btbl.Columns.Add("turns", GetType(System.String))
                        con.Open()
                        If typ = "Buyer" Or typ = "Class" Or typ = "Dept" Then
                            sql = "SELECT w." & typ & " AS Dept, ISNULL(SUM(Act_Sales - Act_Sales_Cost),0) / (ISNULL(SUM(Act_Inv_Cost),0) / 52) AS Turns " & _
                            "FROM Sales_Summary w " & _
                            "JOIN Stores s ON s.Str_Id = w.Str_Id " & _
                            "JOIN Inv_Summary i ON i.Loc_Id = s.Inv_Loc AND i.Dept = w.Dept AND i.Buyer = w.Buyer " & _
                            "AND i.Class = w.Class AND i.eDate = w.eDate " & _
                            "WHERE w.Str_Id = '" & thisStore & "' AND w.sDate >= '" & thisSDate & "' AND w.eDate <= '" & thisEDate & "' " & _
                            "AND Act_Inv_Cost > 0 GROUP BY w." & typ & " ORDER BY Turns DESC"
                        Else
                            sql = "IF OBJECT_ID('tempdb.dbo.#t1', 'U') IS NOT NULL DROP TABLE #t1 " & _
                            "IF OBJECT_ID('tempdb.dbo.#t2', 'U') IS NOT NULL DROP TABLE #t2 " & _
                            "SELECT m." & typ & " AS Dept,i.Sku, ISNULL(SUM(Max_OH * Curr_Cost),0) AS invcost, CONVERT(Decimal,0) AS COGS, " & _
                            "COUNT(*) AS wksoh, CONVERT(Decimal,0) AS turns INTO #t1 FROM Item_Inv i " & _
                            "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                            "JOIN Item_Master m ON m.Sku = i.Sku " & _
                            "WHERE s.Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "GROUP BY Str_Id, i.Sku, m." & typ & " " & _
                            "DELETE FROM #t1 WHERE invcost = 0 " & _
                            "SELECT m." & typ & " AS Dept, d.Sku, CONVERT(Decimal,0) AS invcost, ISNULL(SUM(Sales_Cost),0) AS COGS, " & _
                            "CONVERT(Decimal,0) AS wksoh INTO #t2 FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                            "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "GROUP BY Str_Id, d.Sku, m." & typ & " " & _
                            "UPDATE #t1 SET #t1.COGS = #t2.COGS FROM #t1 " & _
                            "JOIN #t2 ON #t2.Dept = #t1.Dept AND #t2.Sku = #t1.Sku " & _
                            "UPDATE #t1 SET turns = CASE WHEN invcost > 0 AND wksoh > 0 THEN (COGS/(invcost/wksoh))*52/wksoh ELSE 0 END " & _
                            "SELECT Dept, AVG(turns) AS turns FROM #t1 GROUP BY Dept ORDER BY turns DESC"
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                        End If
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            row = btbl.NewRow
                            row("dept") = rdr("Dept")
                            row("turns") = rdr("Turns")
                            btbl.Rows.Add(row)
                            cnt += 1
                        End While
                        con.Close()

                        If cnt = 0 Then Exit Sub ' Get out if there are no records to process

                        Dim fcnt As Integer = cnt / 3
                        Dim mcnt As Integer = fcnt * 2
                        Dim fast, medium As Decimal
                        cnt = 0
                        For Each row As DataRow In btbl.Rows
                            cnt += 1
                            If cnt = fcnt Then fast = row("turns")
                            If cnt = mcnt Then medium = row("turns")
                        Next
                        con.Open()
                        sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' " & _
                            "AND Type = '" & dbTyp & "' AND Code = '" & dbTyp & "' AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "') " & _
                            "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
                            "SELECT 'Turns','" & thisStore & "','" & dbTyp & "','" & dbTyp & "','" & thisSDate & "','" & thisEDate & "'," & _
                            fast & "," & medium & "," & 0.01 & ",'" & now & "' " & _
                            "ELSE " & _
                            "UPDATE Score_Params SET Value1 = " & fast & ", Value2 = " & medium & ", Last_Update = '" & now & "' " & _
                            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = '" & dbTyp & "' AND Code = '" & dbTyp & "' " & _
                            "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Next
            Next
            Dim m As String = "Get Turns " & typ
            Call Update_Process_Log(1, m, "", "")
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Get Turns Parameters " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Get_UnitCost_Params(typ)
        Try
            Dim records As Integer = 0
            Dim dbTyp As String
            Select Case typ
                Case "Sku"
                    dbTyp = "Item"
                Case "Dept"
                    dbTyp = "Dept"
                Case "Class"
                    dbTyp = "Class"
                Case "Buyer"
                    dbTyp = "Buyer"
                Case "Vendor_Id"
                    dbTyp = "Vendor_Id"
                Case "PLine"
                    dbTyp = "Pline"
            End Select
            Console.WriteLine("Computing Unit Cost parameters")
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")
                For Each drow As DataRow In dateTbl.Rows
                    thisEDate = drow("eDate")
                    Console.WriteLine("Computing period " & thisEDate)
                    scoreTbl3 = New DataTable
                    scoreTbl3.Columns.Add("Dept", GetType(System.String))
                    Dim PrimaryKey3(1) As DataColumn
                    PrimaryKey3(0) = scoreTbl3.Columns("Dept")
                    scoreTbl3.PrimaryKey = PrimaryKey3
                    scoreTbl3.Columns.Add("H", GetType(System.Decimal))
                    scoreTbl3.Columns.Add("M", GetType(System.Decimal))
                    scoreTbl3.Columns.Add("L", GetType(System.Decimal))
                    scoreTbl3.Columns.Add("records", GetType(System.Int32))

                    con.Open()
                    sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE eDate = '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSDate = rdr("sDate")
                        thisEDate = rdr("eDate")
                    End While
                    con.Close()

                    If typ = "Sku" Then
                        con.Open()
                        sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            row3 = scoreTbl3.NewRow
                            row3(0) = rdr("Dept")
                            scoreTbl3.Rows.Add(row3)
                        End While
                        con.Close()

                        ctbl = New DataTable
                        ctbl.Columns.Add("Dept", GetType(System.String))
                        ctbl.Columns.Add("Sku", GetType(System.String))
                        ctbl.Columns.Add("avgCost", GetType(System.Decimal))
                        ctbl.Columns.Add("score", GetType(System.String))
                        Dim dept, prevDept As String
                        Dim cnt As Integer
                        Dim val, cost, div, sold As Decimal
                        Dim foundRow As DataRow
                        prevDept = ""
                        con.Open()
                        sql = "SELECT Dept, d.Sku, ISNULL(SUM(Sold),0) AS sold, ISNULL(SUM(Sales_Cost),0) AS cost " & _
                            "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                            "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "GROUP BY Dept, d.Sku ORDER BY Dept"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            dept = rdr("Dept")
                            If dept <> prevDept Then
                                If prevDept = "" Then
                                    prevDept = dept
                                Else
                                    foundRow = scoreTbl3.Rows.Find(prevDept)
                                    foundRow(4) = cnt
                                    cnt = 0
                                    prevDept = dept
                                End If
                            End If
                            sold = rdr("sold")
                            cost = rdr("cost")
                            val = 0
                            If Not IsDBNull(sold) And Not IsNothing(sold) And Not IsDBNull(cost) And Not IsNothing(cost) Then
                                If IsNumeric(sold) And IsNumeric(cost) Then
                                    If sold > 0 Then val = cost / sold
                                End If
                            End If
                            If val > 0 Then cnt += 1
                            row = ctbl.NewRow
                            row("Dept") = dept
                            row("avgCost") = val
                            ctbl.Rows.Add(row)
                        End While
                        con.Close()

                        records = ctbl.Rows.Count                          ' Get out if threr are no records to process
                        If records = 0 Then Exit Sub

                        Dim hCnt, mCnt As Integer
                        Dim view As New DataView(ctbl)
                        view.Sort = "Dept ASC, avgCost DESC"                       ' was Dept ASC, avgCosr DESC

                        cnt = 0
                        prevDept = ""
                        For Each rw As DataRowView In view
                            cnt += 1
                            dept = rw("Dept")
                            val = rw("avgCost")
                            If cnt = hCnt Then
                                foundRow = scoreTbl3.Rows.Find(dept)
                                foundRow(1) = val
                            End If
                            If cnt = mCnt Then
                                foundRow = scoreTbl3.Rows.Find(dept)
                                foundRow(2) = val
                                foundRow(3) = 0.01
                            End If
                            If dept <> prevDept Then
                                foundRow = scoreTbl3.Rows.Find(dept)
                                oTest = foundRow(4)
                                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                    If IsNumeric(oTest) Then
                                        hCnt = foundRow(4) / 3
                                        mCnt = hCnt * 2
                                        cnt = 1
                                    End If
                                End If
                                prevDept = dept
                            End If
                        Next

                        Dim id, type, code, score As String
                        Dim amt, pct, pct2, pct3 As Decimal
                        Dim now As DateTime = Date.Now

                        con.Open()
                        For Each row As DataRow In scoreTbl3.Rows
                            id = "UnitCost"
                            code = row(0)
                            oTest = row(1)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(1) Else pct = 0
                            oTest = row(2)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct2 = row(2) Else pct2 = 0
                            oTest = row(3)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct3 = row(3) Else pct3 = 0

                            sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                                "AND eDate = '" & thisEDate & "' AND Type = 'Item' AND Code = '" & code & "') " & _
                                "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                                "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','Item','" & _
                                code & "'," & pct & "," & pct2 & "," & pct3 & ",'" & now & "' " & _
                                "ELSE " & _
                                "UPDATE Score_Params SET Value1 = " & pct & ", Value2 = " & pct2 & ", Value3 = " & pct3 & ", Last_Update = '" & now & "' " & _
                                "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = '" & code & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                        Next
                        con.Close()
                    Else
                        ctbl = New DataTable
                        ctbl.Columns.Add("cost", GetType(System.String))
                        Dim cnt As Integer = 0
                        con.Open()
                        sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Cost),0) / ISNULL(SUM(Sold),0) AS cost " & _
                            "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                            "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "AND Sold > 0 GROUP BY m." & typ & " ORDER BY cost DESC"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            row = ctbl.NewRow
                            oTest = rdr("cost")
                            row("cost") = rdr("cost")
                            ctbl.Rows.Add(row)
                            cnt += 1
                        End While
                        con.Close()

                        records = ctbl.Rows.Count                  ' Get out if there are no records to process
                        If records = 0 Then Exit Sub

                        Dim hcnt As Integer = cnt / 3
                        Dim mcnt As Integer = hcnt * 2
                        Dim high, medium As Decimal
                        cnt = 0
                        For Each row As DataRow In ctbl.Rows
                            cnt += 1
                            If cnt = hcnt Then high = row("cost")
                            If cnt = mcnt Then medium = row("cost")
                        Next
                        con.Open()
                        sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' " & _
                            "AND Type = '" & dbTyp & "' AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "') " & _
                            "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
                            "SELECT 'UnitCost','" & thisStore & "','" & dbTyp & "','" & dbTyp & "','" & thisSDate & "','" & thisEDate & "'," & _
                            high & "," & medium & "," & 0.01 & ",'" & Now & "' " & _
                            "ELSE " & _
                            "UPDATE Score_Params SET Value1 = " & high & ", Value2 = " & medium & ", Last_Update = '" & Now & "' " & _
                            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = '" & dbTyp & "' " & _
                            "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Next
            Next
            Dim m As String = "Get UnitCost " & typ
            Call Update_Process_Log(1, m, "", "")
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Get UnitCost Parameters " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Score_GrossMargin(typ)
        Try

            Dim type As String = ""
            Dim tblName As String
            Select Case typ
                Case "Sku"
                    type = "Item"
                Case "Dept"
                    type = "Dept"
                    tblName = "Department_Scores"
                Case "Class"
                    type = "Class"
                    tblName = "Class_Scores"
                Case "Buyer"
                    type = "Buyer"
                    tblName = "Buyer_Scores"
                Case "PLine"
                    type = "PLine"
                    tblName = "PLine_Scores"
                Case "Vendor_Id"
                    tblName = "Vendor_Scores"
                    type = "Vendor_Id"
            End Select
            Dim thisSdate, thisEdate, bDate As Date
            Dim code, item, score As String
            Dim val1, gm As Decimal
            Dim cnt, rw As Integer
            Dim foundRow As DataRow
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")
                For Each row As DataRow In dateTbl.Rows
                    thisEdate = row("eDate")
                    Console.WriteLine("Gross Margin Scoring " & typ & " " & thisEdate)
                    con.Open()
                    sql = "SELECT DISTINCT eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE eDate = '" & thisEdate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSdate = rdr("sDate")
                        thisEdate = rdr("eDate")
                    End While
                    con.Close()

                    bDate = DateAdd(DateInterval.Day, -21, thisEdate)
                    scoreTbl = New DataTable
                    scoreTbl.Columns.Add("Code", GetType(System.String))
                    scoreTbl.Columns.Add("Amount", GetType(System.Decimal))

                    con.Open()
                    sql = "SELECT Code, Value1 FROM Score_Params WHERE Str_ID = '" & thisStore & "' " & _
                        "AND ID = 'GrossMargin' AND Type = '" & type & "' AND sDate = '" & thisSdate & "' AND eDate = '" & thisEdate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        row = scoreTbl.NewRow
                        code = rdr("Code")
                        val1 = rdr("Value1")
                        row("Code") = code
                        row("Amount") = val1
                        scoreTbl.Rows.Add(row)
                    End While
                    con.Close()

                    universeTbl = New DataTable
                    universeTbl.Columns.Add(typ)
                    Dim priKey(1) As DataColumn
                    priKey(0) = universeTbl.Columns(typ)
                    universeTbl.PrimaryKey = priKey
                    con.Open()
                    sql = "SELECT DISTINCT m." & typ & " FROM Item_Inv d JOIN Item_Master m ON m.Sku = d.Sku " & _
                        "WHERE Max_OH > 0 AND eDate BETWEEN '" & bDate & "' AND '" & thisEdate & "' "
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        item = rdr(typ)
                        row = universeTbl.NewRow
                        row(0) = item
                        universeTbl.Rows.Add(row)
                    End While
                    con.Close()

                    If typ = "Sku" Then
                        con.Open()
                        sql = "Update Item_Sales SET Score = NULL WHERE Str_Id = '" & thisStore & "' AND eDate = '" & thisEdate & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If

                    con.Open()
                    con2.Open()
                    sql = "SELECT m." & typ & ", ISNULL(SUM(Sales_Retail - Sales_Cost),0) AS grossmargin " & _
                    "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku WHERE Str_Id = '" & thisStore & "' " & _
                    "AND sDate >= '" & thisSdate & "' AND eDate <= '" & thisEdate & "' " & _
                    "GROUP BY m." & typ & " "
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        score = "-"
                        item = rdr(typ)
                        foundRow = universeTbl.Rows.Find(item)
                        If Not IsNothing(foundRow) Then
                            gm = rdr("grossmargin")
                            If Not IsNothing(scoreTbl) Then
                                If gm >= scoreTbl.Rows(0).Item("Amount") Then score = "A"
                                If gm >= scoreTbl.Rows(1).Item("Amount") And gm < scoreTbl.Rows(0).Item(1) Then score = "B"
                                If gm >= scoreTbl.Rows(2).Item("Amount") And gm < scoreTbl.Rows(1).Item(1) Then score = "C"
                                If gm >= scoreTbl.Rows(3).Item("Amount") And gm < scoreTbl.Rows(2).Item(1) Then score = "D"
                                If gm >= scoreTbl.Rows(4).Item("Amount") And gm < scoreTbl.Rows(3).Item(1) Then score = "E"
                            End If
                            If typ = "Sku" Then
                                sql = "UPDATE Item_Sales SET Score = '" & score & "' WHERE Str_Id ='" & thisStore & "' " & _
                                    "AND Sku = '" & item & "' AND eDate = '" & thisEdate & "'"
                            Else
                                sql = "IF NOT EXISTS (SELECT * FROM " & tblName & " WHERE Str_Id = '" & thisStore & "' " & _
                                    "AND ID = '" & item & "' AND eDate = '" & thisEdate & "') " & _
                                    "INSERT INTO " & tblName & "(ID, Str_Id, eDate, Score) " & _
                                    "SELECT '" & item & "','" & thisStore & "','" & thisEdate & "','" & score & "' " & _
                                    "ELSE " & _
                                    "UPDATE " & tblName & " SET Score = '" & score & "' " & _
                                    "WHERE ID = '" & item & "' AND Str_Id = '" & thisStore & "' AND eDate = '" & thisEdate & "'"
                            End If
                            cmd = New SqlCommand(sql, con2)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                        End If
                    End While
                    con.Close()
                    con2.Close()
                Next
            Next
            Dim m As String = "Score GrossMargin " & typ
            Call Update_Process_Log(1, m, "", "")
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Score GrossMargin " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Score_Turns(typ)
        Try
            Dim type, tblName As String
            Select Case typ
                Case "Sku"
                    type = "Item"
                Case "Dept"
                    type = "Dept"
                    tblName = "Department_Scores"
                Case "Class"
                    type = "Class"
                    tblName = "Class_Scores"
                Case "Buyer"
                    type = "Buyer"
                    tblName = "Buyer_Scores"
                Case "PLine"
                    type = "PLine"
                    tblName = "PLine_Scores"
                Case "Vendor_Id"
                    tblName = "Vendor_Scores"
                    type = "Vendor_Id"
            End Select
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")
                For Each row As DataRow In dateTbl.Rows
                    If con.State = ConnectionState.Open Then con.Close()
                    Dim x As Integer = dateTbl.Rows.Count
                    thisEDate = row("eDate")
                    Console.WriteLine("Scoring Turns " & " " & typ & " " & thisEDate & " " & thisStore)

                    con.Open()
                    sql = "SELECT eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE eDate = '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSDate = rdr("sDate")
                        thisEDate = rdr("eDate")
                    End While
                    con.Close()

                    Dim bDate As Date = DateAdd(DateInterval.Day, -21, thisEDate)
                    Dim scoreTbl2 = New DataTable
                    scoreTbl2.Columns.Add("Code", GetType(System.String))
                    Dim primaryKey(1) As DataColumn
                    primaryKey(0) = scoreTbl2.Columns("Code")
                    scoreTbl2.PrimaryKey = primaryKey
                    scoreTbl2.Columns.Add("Value1", GetType(System.Decimal))
                    scoreTbl2.Columns.Add("Value2", GetType(System.Decimal))
                    scoreTbl2.Columns.Add("Value3", GetType(System.Decimal))

                    con.Open()
                    sql = "SELECT DISTINCT Code, Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' " & _
                        "AND Type = '" & type & "' AND Str_Id = '" & thisStore & "' AND eDate = '" & thisEDate & "' " & _
                        "ORDER BY Code"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    Dim cnt As Integer = scoreTbl2.Rows.Count
                    While rdr.Read
                        oTest = rdr("Code")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            Dim rw As DataRow = scoreTbl2.NewRow
                            rw("Code") = rdr("Code")
                            rw("Value1") = rdr("Value1")
                            rw("Value2") = rdr("Value2")
                            rw("Value3") = rdr("Value3")
                            scoreTbl2.Rows.Add(rw)
                        End If
                    End While
                    con.Close()

                    Dim foundRow As DataRow
                    tbl2 = New DataTable
                    tbl2.Columns.Add("edate", GetType(System.DateTime))
                    tbl2.Columns.Add("dept", GetType(System.String))
                    tbl2.Columns.Add("Sku", GetType(System.String))
                    tbl2.Columns.Add("invcost", GetType(System.Decimal))
                    tbl2.Columns.Add("wksoh", GetType(System.Decimal))
                    tbl2.Columns.Add("COGS", GetType(System.Decimal))
                    Dim dept, item, prevDept, score As String
                    Dim wks, cogs, invcost, val As Decimal
                    ''Dim cnt As Integer
                    cnt = 0

                    universeTbl = New DataTable
                    universeTbl.Columns.Add(type)
                    Dim priKey(1) As DataColumn
                    priKey(0) = universeTbl.Columns(type)
                    universeTbl.PrimaryKey = priKey

                    con.Open()
                    sql = "SELECT DISTINCT m." & typ & " AS type FROM Item_Inv d JOIN Item_Master m ON m.Sku = d.Sku " & _
                        "WHERE Max_OH > 0 AND eDate BETWEEN '" & bDate & "' AND '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        item = rdr("type")
                        row = universeTbl.NewRow
                        row(0) = item
                        universeTbl.Rows.Add(row)
                    End While
                    con.Close()

                    If typ = "Sku" Then
                        con.Open()
                        sql = "SELECT d.eDate, m.Dept, d.Sku, CONVERT(Decimal,0) AS invcost, ISNULL(SUM(Sales_Cost),0) As COGS, " & _
                            "CONVERT(Integer,0) AS wksOH, ISNULL(SUM(Sales_Retail),0) - ISNULL(SUM(Sales_Cost),0) AS GrossProfit " & _
                            "INTO #t1 FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                            "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "GROUP BY Str_Id, d.eDate, m.Dept, d.Sku " & _
                            "SELECT i.eDate, m.Dept, i.Sku, ISNULL(SUM(Max_OH * Curr_Retail),0) AS invcost, " & _
                            "CONVERT(decimal(10,2),0) AS COGS, COUNT(*) As wksOH, CONVERT(decimal(10,2),0) AS GrossProfit " & _
                            "INTO #t2 FROM Item_Inv i JOIN Item_Master m ON m.Sku = i.Sku " & _
                            "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                            "WHERE s.Str_Id = '" & thisStore & "' AND Max_OH > 0 " & _
                            "AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "GROUP BY Str_Id, i.eDate, m.Dept, i.Sku " & _
                            "MERGE #t1 AS t USING #t2 s ON s.eDate = t.eDate AND s.Dept = t.Dept AND s.Sku = t.Sku " & _
                            "WHEN NOT MATCHED BY TARGET THEN " & _
                            "INSERT(eDate, Dept, Sku, InvCost, COGS, WksOH, GrossProfit) " & _
                            "VALUES(s.eDate, s.Dept, s.Sku, s.InvCost, s.COGS, s. WksOH, s.GrossProfit) " & _
                            "WHEN MATCHED THEN UPDATE SET t.InvCost = s.InvCost, t.wksOH = s.wksOH; " & _
                            "SELECT eDate, Dept, Sku, invcost, wksoh, COGS FROM #t1 ORDER BY Dept, Sku"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            cnt += 1
                            Dim rw As DataRow = tbl2.NewRow
                            rw("edate") = rdr("eDate")
                            rw("dept") = rdr("Dept")
                            rw("Sku") = rdr("Sku")
                            rw("invcost") = rdr("invcost")
                            rw("wksoh") = rdr("wksoh")
                            rw("COGS") = rdr("COGS")
                            tbl2.Rows.Add(rw)
                        End While
                        con.Close()




                        'For Each r As DataRow In scoreTbl2.Rows
                        '    oTest = r("Code")
                        '    oTest = r(1)
                        'Next




                        con.Open()
                        prevDept = ""
                        cnt = 0
                        Console.WriteLine("Computing Turns")
                        For Each row2 As DataRow In tbl2.Rows
                            cnt += 1
                            If cnt Mod 1000 = 0 Then Console.WriteLine(cnt & " " & thisEDate & " Records updated with Turns Score")
                            thisEDate = row2("edate")
                            dept = row2("dept")
                            item = row2("Sku")
                            foundRow = universeTbl.Rows.Find(item)
                            If Not IsNothing(foundRow) Then
                                wks = row2("wksoh")
                                cogs = row2("COGS")
                                invcost = row2("invcost")
                                If wks > 0 And invcost > 0 Then
                                    foundRow = scoreTbl2.Rows.Find(dept)
                                    val = (row2("cogs") / (row2("invcost") / row2("wksoh"))) * (52 / row2("wksoh"))
                                    If val >= foundRow("Value1") Then score = "F"
                                    If val >= foundRow("Value2") And val < foundRow("Value1") Then score = "M"
                                    If val > 0 And val < foundRow("Value2") Then score = "S"
                                    If val <= 0 Then score = "-"
                                    sql = "UPDATE Item_Sales SET Score = Score + '" & score & "', Turns = " & val & " WHERE Str_Id = '" & thisStore & "' " & _
                                        "AND Sku = '" & item & "' AND eDate = '" & thisEDate & "'"
                                    cmd = New SqlCommand(sql, con)
                                    cmd.CommandTimeout = 120
                                    cmd.ExecuteNonQuery()
                                End If
                            End If
                        Next
                        sql = "UPDATE Item_Sales SET Score = Score + '-' WHERE eDate = '" & thisEDate & "' AND LEN(Score) = 1"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                    Else
                        Dim id As String
                        con.Open()

                        If typ = "PLine" Or typ = "Vendor_Id" Then
                            sql = "SELECT DISTINCT m." & typ & " FROM Item_Inv d JOIN Item_Master m ON m.Sku = d.Sku " & _
                                "JOIN Stores s ON s.Inv_Loc = d.Loc_Id " & _
                                "WHERE s.Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                                "AND Max_OH * Cost > 0 ORDER BY m." & typ & " "
                        Else
                            sql = "SELECT w." & typ & ", ISNULL(SUM(Act_Sales - Act_Sales_Cost),0) / (ISNULL(SUM(Act_Inv_Cost),0) / 52) AS turns " & _
                                "FROM Sales_Summary w " & _
                                "JOIN Stores s ON s.Str_Id = w.Str_Id " & _
                                "JOIN Inv_Summary i ON i.Loc_Id = s.Inv_Loc AND i.Dept = w.Dept AND i.Buyer = w.Buyer " & _
                                "AND i.Class = w.Class AND i.eDate = w.eDate " & _
                                "WHERE Act_Inv_Cost > 0 AND w.Str_Id = '" & thisStore & "' " & _
                                "AND i.sDate >= '" & thisSDate & "' AND i.eDate <= '" & thisEDate & "' GROUP BY w." & typ & " "
                        End If

                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            id = rdr(typ)

                            Console.WriteLine(id)

                            foundRow = scoreTbl2.Rows.Find(typ)        ' Get out if there are no records to process
                            If IsNothing(foundRow) Then Exit Sub

                            oTest = foundRow(0)
                            If typ = "PLine" Or typ = "Vendor_Id" Then
                                con2.Open()
                                If typ = "Vendor_Id" Then
                                    sql = "SELECT dbo.fnTurns('" & thisStore & "','" & id & "','" &
                                        DateAdd(DateInterval.Day, 6, thisSDate) & "','" & thisEDate & "') AS Turns"
                                Else
                                    sql = "SELECT dbo.fnPLineTurns('" & thisStore & "','" & id & "','" &
                                        DateAdd(DateInterval.Day, 6, thisSDate) & "','" & thisEDate & "') AS Turns"
                                End If
                                cmd = New SqlCommand(sql, con2)
                                cmd.CommandTimeout = 120
                                rdr2 = cmd.ExecuteReader
                                While rdr2.Read
                                    oTest = rdr2("Turns")
                                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                        If IsNumeric(oTest) Then val = oTest Else val = 0
                                    End If

                                End While
                                con2.Close()
                            Else
                                val = rdr("turns")
                            End If

                            If val >= foundRow("Value1") Then score = "F"
                            If val >= foundRow("Value2") And val < foundRow("Value1") Then score = "M"
                            If val > 0 And val < foundRow("Value2") Then score = "S"
                            If val <= 0 Then score = "-"
                            con2.Open()
                            sql = "IF NOT EXISTS (SELECT ID FROM " & tblName & " WHERE ID = '" & id & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND eDate = '" & thisEDate & "') " & _
                                "INSERT INTO " & tblName & " (ID, Str_Id, eDate, Score) " & _
                                "SELECT '" & id & "','" & thisStore & "','" & thisEDate & "','-" & score & "' " & _
                                "ELSE " & _
                                "UPDATE " & tblName & " SET Score = LEFT(Score,1) + '" & score & "' WHERE ID = '" & id & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND eDate = '" & thisEDate & "'"
                            cmd = New SqlCommand(sql, con2)
                            cmd.ExecuteNonQuery()
                            con2.Close()
                        End While
                        con.Close()
                    End If
                Next
            Next
            If con.State = ConnectionState.Open Then con.Close()
            Dim m As String = "Score Turns " & typ
            Call Update_Process_Log(1, m, "", "")
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Score Turns " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Score_Cost(typ)
        Try
            Dim type, tblName As String
            Select Case typ
                Case "Sku"
                    type = "Item"
                Case "Dept"
                    type = "Dept"
                    tblName = "Department_Scores"
                Case "Class"
                    type = "Class"
                    tblName = "Class_Scores"
                Case "Buyer"
                    type = "Buyer"
                    tblName = "Buyer_Scores"
                Case "PLine"
                    type = "PLine"
                    tblName = "PLine_Scores"
                Case "Vendor_Id"
                    type = "Vendor_Id"
                    tblName = "Vendor_Scores"
            End Select
            For Each srow In storeTbl.Rows
                thisStore = srow("Store")
                For Each row As DataRow In dateTbl.Rows
                    Console.WriteLine("Scoring Cost " & " " & typ & " " & thisEDate)
                    thisEDate = row("eDate")
                    con.Open()
                    sql = "SELECT DISTINCT eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                        "WHERE eDate = '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisSDate = rdr("sDate")
                        thisEDate = rdr("eDate")
                    End While
                    con.Close()

                    Dim bDate As Date = DateAdd(DateInterval.Day, -21, thisEDate)
                    Dim scoreTbl3 = New DataTable
                    scoreTbl3.Columns.Add("Code", GetType(System.String))
                    Dim primaryKey3(1) As DataColumn
                    primaryKey3(0) = scoreTbl3.Columns("Code")
                    scoreTbl3.PrimaryKey = primaryKey3
                    scoreTbl3.Columns.Add("sDate")
                    scoreTbl3.Columns.Add("eDate")
                    scoreTbl3.Columns.Add("Value1")
                    scoreTbl3.Columns.Add("Value2")
                    scoreTbl3.Columns.Add("Value3")

                    con.Open()
                    sql = "SELECT Str_Id, Code, sDate, eDate, Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' " & _
                        "AND Type = '" & type & "' AND Str_Id = '" & thisStore & "' AND eDate = '" & thisEDate & "' ORDER BY Code"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        Dim rw As DataRow = scoreTbl3.NewRow
                        rw("Code") = rdr("Code")
                        rw("sDate") = rdr("sDate")
                        rw("eDate") = rdr("eDate")
                        rw("Value1") = rdr("Value1")
                        rw("Value2") = rdr("Value2")
                        rw("Value3") = rdr("Value3")
                        scoreTbl3.Rows.Add(rw)
                    End While
                    con.Close()

                    Dim cnt As Integer = 0
                    Dim foundRow As DataRow
                    Dim item, dept, score As String
                    Dim cost, sold, val As Decimal
                    universeTbl = New DataTable
                    universeTbl.Columns.Add(typ)
                    Dim priKey(1) As DataColumn
                    priKey(0) = universeTbl.Columns(typ)
                    universeTbl.PrimaryKey = priKey
                    con.Open()
                    sql = "SELECT DISTINCT m." & typ & " FROM Item_Inv d JOIN Item_Master m ON m.Sku = d.Sku " & _
                        "WHERE Max_OH > 0 AND eDate BETWEEN '" & bDate & "' AND '" & thisEDate & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        item = rdr(typ)
                        row = universeTbl.NewRow
                        row(0) = item
                        universeTbl.Rows.Add(row)
                    End While
                    con.Close()

                    If typ = "Sku" Then
                        con.Open()
                        con2.Open()
                        sql = "SELECT DISTINCT Sku FROM Item_Sales WHERE sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                            "AND Sales_Retail > 0 ORDER BY Sku"
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        sql = "SELECT Dept, m.Sku, ISNULL(Curr_Cost,0) AS cost " & _
                            "FROM Item_Master m " & _
                            "ORDER BY Dept, m.Sku"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            dept = rdr("Dept")
                            item = rdr("Sku")
                            cost = rdr("cost")
                            val = 0
                            foundRow = universeTbl.Rows.Find(item)
                            If Not IsNothing(foundRow) Then
                                If Not IsDBNull(sold) And Not IsNothing(sold) And Not IsDBNull(cost) And Not IsNothing(cost) Then
                                    If IsNumeric(cost) And cost > 0 Then
                                        val = cost
                                        foundRow = scoreTbl3.Rows.Find(dept)
                                        If Not IsNothing(foundRow) Then
                                            If val >= foundRow("Value1") Then score = "H"
                                            If val >= foundRow("Value2") And val < foundRow("Value1") Then score = "M"
                                            If val > 0 And val < foundRow("Value2") Then score = "L"
                                            sql = "UPDATE Item_Sales SET Score = LEFT(Score,2) + '" & score & "' WHERE Str_Id = '" & thisStore & "' " & _
                                                "AND Sku = '" & item & "' AND eDate = '" & thisEDate & "'"
                                            cmd = New SqlCommand(sql, con2)
                                            cmd.ExecuteNonQuery()
                                        End If
                                    End If
                                End If
                            End If
                        End While
                        con.Close()
                        con2.Close()
                    Else

                        con.Open()
                        con2.Open()
                        sql = "SELECT m." & typ & ", ISNULL(SUM(Sales_Cost),0) / ISNULL(SUM(Sold),0) AS Cost FROM Item_Sales d " & _
                            "JOIN Item_Master m ON m.Sku=d.Sku WHERE Str_Id = '" & thisStore & "' " & _
                            "AND eDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "'  AND Sold > 0 GROUP BY m." & typ & " "
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            dept = rdr(typ)
                            val = rdr("Cost")
                            foundRow = scoreTbl3.Rows.Find(typ)
                            If Not IsNothing(foundRow) Then
                                If val >= foundRow("Value1") Then score = "H"
                                If val >= foundRow("Value2") And val < foundRow("Value1") Then score = "M"
                                If val > 0 And val < foundRow("Value2") Then score = "L"
                                sql = "UPDATE " & tblName & " SET Score = LEFT(Score,2) + '" & score & "' WHERE Str_Id = '" & thisStore & "' " & _
                                    "AND ID = '" & dept & "' AND eDate = '" & thisEDate & "'"
                                cmd = New SqlCommand(sql, con2)
                                cmd.ExecuteNonQuery()
                            End If
                        End While
                    End If
                    con.Close()
                    con2.Close()
                Next
            Next
            Dim m As String = "Score Cost " & typ
            Call Update_Process_Log(1, m, "", "")
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            Console.WriteLine(ex.Message)
            Dim el As New ScoreItems.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "ScoreItems, Score Cost " & typ)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub
    Private Sub Verify_Setup()

    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        Stopwatch.Stop()
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "Scoring"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & CDate(Now) & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
End Module

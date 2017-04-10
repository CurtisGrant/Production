Imports System.Data.SqlClient
Imports System.Xml
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module Module1
    Private eDateTBL As DataTable
    Private con, con2, con3 As SqlConnection
    Private cmd As SqlCommand
    Public rdr, rdr2 As SqlDataReader
    Public client, sql, rcErrorPath, xmlPath, conString, conString2, constr As String
    Private oTest As Object
    Private stopWatch As Stopwatch
    Private totlQty, totlCost, totlRetail, totlMarkdown As Decimal
    Private records As Integer
    Private msg As String

    Sub Main()
        ''Try
        Dim fileSZ As Integer = 0
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim rcServer, rcConString, rcExePath, rcPassword As String
        Dim fld As String = ""
        Dim valu As String = ""
        Dim cust As String = "N"
        ''Dim oktoprocessCustomers As Boolean = False
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
        ''rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;User Id=sa;Password=" & rcPassword & ""
        rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True" & ""
        eDateTBL = New DataTable
        Dim column As New DataColumn
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "processEdate"
        eDateTBL.Columns.Add(column)
        Dim primarykey(1) As DataColumn
        primarykey(0) = eDateTBL.Columns("processEdate")
        eDateTBL.PrimaryKey = primarykey
        Dim server, dbase, thePath, sqlUserId, sqlPassword, exePath As String
        Dim cmd As SqlCommand
        Dim rdr As SqlDataReader
        Dim cnt As Int32 = 0
        Dim thisDate, thisEdate, thisSdate, clearDate, lastWeekseDate As Date
        Dim DayOfWeek As Integer

        Dim args As String() = Environment.GetCommandLineArgs()
        Dim txtArray() As String = args(1).Split(";")
        client = txtArray(0)
        server = txtArray(1)
        dbase = txtArray(2)
        xmlPath = txtArray(3)
        sqlUserId = txtArray(4)
        sqlPassword = txtArray(5)
        thisSdate = txtArray(6)
        thisEdate = txtArray(7)
        rcErrorPath = txtArray(8)
        cust = txtArray(9)
        eDateTBL.Rows.Add(thisEdate)
        lastWeekseDate = DateAdd(DateInterval.Day, -7, thisEdate)
        Console.WriteLine(dbase & " XMLUpdate")



        ''client = "TCM"
        '' ''server = "alrm6hn0ql.database.windows.net,1433"
        ''server = "LP-CURTIS"
        ''dbase = "PARGIF"
        ''xmlPath = "c:\RetailClarity\XMLs\TCM"
        '' ''xmlPath = "\\LPS2-2\RetailClarity\PARGIF\XMLs\Build"
        ''thisSdate = "4/2/2017"
        ''thisEdate = "4/8/2017"
        ''lastWeekseDate = DateAdd(DateInterval.Day, -7, thisEdate)
        ''sqlUserId = "sa"
        ''sqlPassword = "PGadm01!"
        ''rcErrorPath = "c:\RetailClarity\RCSystem\SYSFAIL"
        ''rcExePath = "c:\RetailClarity\EXEs"
        ''cust = "N"
        ' ''eDateTBL.Rows.Add("7/10/2016")
        ''eDateTBL.Rows.Add("4/8/2017")



        conString = "server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True" & ""
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        constr = conString

        thisDate = CDate(Date.Today)
        DayOfWeek = thisDate.DayOfWeek
        clearDate = DateAdd(DateInterval.Month, -6, thisEdate)
        '' If cust = "Y" Then oktoprocessCustomers = True
        stopWatch = New Stopwatch






        '' GoTo 100





        Console.WriteLine("Processing Calendar")
        thePath = xmlPath & "\Calendar.xml"
        If FileSize(thePath) > 0 Then Call Process_Calendar(thePath, con, con2, constr)

        Console.WriteLine("Processing " & client & " Item records")
        thePath = xmlPath & "\Items.xml"
        If FileSize(thePath) > 0 Then Call Process_Items(thePath, con, con2, constr)

        Dim rCon As New SqlConnection(rcConString)
        rCon.Open()
        sql = "UPDATE Client_Master SET Last_XML_Update = NULL WHERE Client_ID = '" & client & "'"
        cmd = New SqlCommand(sql, rCon)
        cmd.ExecuteNonQuery()
        rCon.Close()

        Console.WriteLine(" Purchase Request Header")
        thePath = xmlPath & "\Purchase_Request_Header.xml"
        If FileSize(thePath) > 0 Then Call Process_PREQ_Header(thePath, con, con2)

        Console.WriteLine(" Purchare Request Detail")
        thePath = xmlPath & "\Purchase_Request_Detail.xml"
        If FileSize(thePath) > 0 Then Call Process_PREQ_Detail(thePath, con, con2)

        Console.WriteLine("Processing " & client & " Inventory records")
        thePath = xmlPath & "\Inventory.xml"
        If FileSize(thePath) > 0 Then Call Process_Inventory(thePath, con, con2)

        Console.WriteLine("Processing " & client & " Purchase Order Header")
        thePath = xmlPath & "\POHeader.xml"
        If FileSize(thePath) > 0 Then Call Process_POHeader(thePath, con, con2)

        Console.WriteLine("Processing " & client & " Purchase Order Detail")
        thePath = xmlPath & "\PODetail.xml"
        If FileSize(thePath) > 0 Then Call Process_PODetail(thePath, con, con2)

        Console.WriteLine("Processing " & client & " Adjustment records")
        thePath = xmlPath & "\Adjustment.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "ADJ", con, con2)

        Console.WriteLine("Processing " & client & " Physical records")
        thePath = xmlPath & "\Physical.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "ADJ", con, con2)

        Console.WriteLine("Processing " & client & " Receiving records")
        thePath = xmlPath & "\Receipt.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "RECVD", con, con2)

        Console.WriteLine("Process " & client & " Return records")
        thePath = xmlPath & "\Return.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "RTV", con, con2)

        ''Console.WriteLine("Process " & client & " Unposted Sales")
        ''thePath = xmlPath & "\UnpostedSales.xml"
        ''If FileSize(thePath) > 0 Then Call Process_Data(thePath, "UnpostedSales", con, con2)

100:
        Console.WriteLine("Processing Sales for " & client)
        '' Console.ReadLine()

        thePath = xmlPath & "\Sales.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "Sold", con, con2)





        '' GoTo 200




        Call Merge_Tickets()
        Call Set_Max_OH(thisEdate)
        Call Update_Store_Table()

        Console.WriteLine("Processing " & client & " Transfer records")
        thePath = xmlPath & "\Transfer.xml"
        If FileSize(thePath) > 0 Then Call Process_Data(thePath, "XFER", con, con2)

        Console.WriteLine("Processing " & client & " Buyer records")
        thePath = xmlPath & "\Buyers.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Buyers", con, con2)

        Console.WriteLine("Processing " & client & " Class records")
        thePath = xmlPath & "\Classes.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Classes", con, con2)

        Console.WriteLine("Processing " & client & " Coupon_Code records")
        thePath = xmlPath & "\Coupon_Codes.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Coupon_Codes", con, con2)

        Console.WriteLine("Processing " & client & " Department records")
        thePath = xmlPath & "\Departments.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Departments", con, con2)

        Console.WriteLine("Processing " & client & " Seasons records") '
        thePath = xmlPath & "\Seasons.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Seasons", con, con2)

        Console.WriteLine("Processing " & client & " Store records")
        thePath = xmlPath & "\Stores.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Stores", con, con2)

        Console.WriteLine("Processing " & client & " Shipping Locations")
        thePath = xmlPath & "\Locations.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Locations", con, con2)

        Console.WriteLine("Processing " & client & " Product Line records")
        thePath = xmlPath & "\ProductLines.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "ProductLines", con, con2)

        Console.WriteLine("Processing " & client & " Vendor records")
        thePath = xmlPath & "\Vendors.xml"
        If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Vendors", con, con2)

        Try
            '    Create records for the next week when processing the last day of the week
            Dim nextSdate, nextEdate As Date
            Dim rdr2 As SqlDataReader
            Dim location, item, sku As String
            Dim onhand, committed, endoh As Int32
            Dim cost, retail As Decimal
            Dim dte As Date = Date.Today
            Dim newsku, newitem As String

            ''If thisDate = thisEdate Then
            If DayOfWeek = 0 Then                                              ' create new inventory records on Sunday
                Console.WriteLine("Adding Item_Inv records for next week.")
                ''                                        thisSdate and thisEdate were set when this thing first started
                con2.Open()
                nextSdate = DateAdd(DateInterval.Day, 7, thisSdate)
                nextEdate = DateAdd(DateInterval.Day, 7, thisEdate)
                cnt = 0
                sql = "UPDATE Item_Inv SET Sys_OH = End_OH WHERE eDate = '" & thisEdate & "'"
                cmd = New SqlCommand(sql, con2)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()

                sql = "SELECT Loc_Id, Sku, ISNULL(End_OH,0) End_OH, ISNULL(OnHand,0) OnHand, ISNULL(Committed,0) Committed, " & _
                    "ISNULL(Cost,0) AS Cost, ISNULL(Retail,0) AS Retail, Item FROM Item_Inv " & _
                    "WHERE ISNULL(End_OH,0) > 0 AND eDate = '" & lastWeekseDate & "' " & _
                    "ORDER BY Loc_ID, Sku"
                cmd = New SqlCommand(sql, con2)

                cmd.CommandTimeout = 120
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    location = rdr2("Loc_Id")
                    sku = rdr2("Sku")
                    newsku = Replace(sku, "'", "''")
                    endoh = rdr2("End_OH")
                    committed = rdr2("Committed")
                    onhand = rdr2("onhand")
                    cost = rdr2("Cost")
                    retail = rdr2("Retail")
                    item = rdr2("Item")
                    newitem = Replace(item, "'", "''")
                    cnt += 1
                    If cnt Mod 1000 = 0 Then Console.WriteLine(cnt & "  " & item)

                    con3.Open()
                    sql = "IF NOT EXISTS (SELECT Sku FROM Item_Inv WHERE Loc_Id = '" & location & "' AND Sku = '" & item & "' AND eDate = '" & nextEdate & "') " & _
                        "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Begin_OH, End_OH, OnHand, Committed, Max_OH, Cost, Retail, Item, YrWk) " & _
                        "SELECT '" & Trim(location) & "','" & Trim(newsku) & "','" & nextSdate & "','" & nextEdate & "'," &
                        onhand & "," & onhand & "," & onhand & "," & onhand & "," & committed & "," & cost & ", " & retail & ", '" & newitem & "', yrwk FROM Calendar " & _
                        "WHERE eDate = '" & nextEdate & "' AND Week_Id > 0 " & _
                        "ELSE " & _
                        "UPDATE Item_Inv SET Begin_OH = " & onhand & ", End_OH = " & onhand & " " & _
                        "WHERE Loc_Id = '" & location & "' AND Sku = '" & item & "' AND eDate = '" & nextEdate & "'"
                    cmd = New SqlCommand(sql, con3)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                    con3.Close()
                End While
                con2.Close()
                Dim m As String = "Created " & cnt & " Records"
                Call Update_Process_Log("1", "Create Inventory Records for Next Week", m, "")
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Create records for next week")
            MsgBox(ex.Message)

        End Try

        cnt = eDateTBL.Rows.Count
        If cnt = 0 Then
            con3.Open()
            Dim title As String = "XMLUpdate"
            Dim message As String = "Nothing to update in Client_Weekly_Summary"
            sql = "INSERT INTO Message (mDate, Title, Message) " & _
                "SELECT '" & Date.Now & "','" & title & "','" & message & "'"
            cmd = New SqlCommand(sql, con3)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            con3.Close()
        Else
            Dim dateStr As String = ""
            For Each rw In eDateTBL.Rows
                If dateStr = "" Then
                    dateStr = CDate(rw(0))
                Else
                    dateStr &= "," & CDate(rw(0))
                End If
            Next
            Console.WriteLine("Call Client_Weekly_Summary " & dateStr)
            Dim p As New ProcessStartInfo
            p.FileName = rcExePath & "\Client_Weekly_Summary.exe"
            p.Arguments = client & ";" & server & ";" & dbase & ";" & xmlPath & ";" & sqlUserId & ";" & sqlPassword & ";" & dateStr & ";" & rcErrorPath
            Console.WriteLine("Calling Client_Weekly_Summary " & p.Arguments)

            p.UseShellExecute = True
            p.WindowStyle = ProcessWindowStyle.Normal
            Dim proc As Process = Process.Start(p)
        End If

200:    con = New SqlConnection(rcConString)
        con.Open()
        sql = "UPDATE Client_Master SET Last_XML_Update = '" & Date.Now & "', Contact = " & cnt & " WHERE Client_Id = '" & client & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Process_Calendar(ByVal thePath, ByVal con, ByVal con2, ByVal constr)
        Try
            Dim ctbl As New DataTable
            Dim ds As New DataSet
            Dim year, period, week, yrprd, yrwks, prdwk As Integer
            Dim sdate, edate As Date
            Dim cnt As Integer = 0
            ctbl.Columns.Add("Year_Id")
            ctbl.Columns.Add("Prd_Id")
            ctbl.Columns.Add("Week_Id")
            ctbl.Columns.Add("sDate")
            ctbl.Columns.Add("eDate")
            ctbl.Columns.Add("YrPrd")
            ctbl.Columns.Add("YrWks")
            Dim xmlFile As XmlReader
            Dim rw As DataRow
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                ctbl = ds.Tables(0)
                For Each row In ctbl.Rows
                    cnt += 1
                    year = CInt(row("Year_Id"))
                    period = CInt(row("Prd_Id"))
                    week = CInt(row("Week_Id"))
                    sdate = CDate(row("sDate"))
                    edate = CDate(row("eDate"))
                    yrprd = CInt(row("YrPrd"))
                    yrwks = CInt(row("YrWks"))
                    prdwk = CInt(period & week)
                    If (period = 0 And week = 0) Or (period > 0 And week > 0) Then
                        sql = "IF NOT EXISTS (SELECT * FROM Calendar WHERE sDate = '" & sdate & "' AND Prd_Id = " & period & " AND Week_Id = " & week & ") " & _
                            "INSERT INTO Calendar (Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, YrWks, PrdWk) " & _
                            "SELECT " & year & "," & period & "," & week & ",'" & sdate & "','" & edate & "'," & yrprd & "," & yrwks & "," & prdwk
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                    End If
                Next
                con.close()
                Dim m As String = "Created " & cnt & " Records"
                Call Update_Process_Log("1", "Update Calendar", m, "")
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Update Calendar")
        End Try
    End Sub

    Private Sub Process_Items(ByVal thePath, ByVal con, ByVal con2, ByVal constr)
        Try
            stopWatch.Start()
            Console.WriteLine("Cleaning XML import file.")
            Dim cnt As Integer = 0
            Dim oTest As Object
            Dim isNumber As Boolean
            Dim item, descr, vid, vendor, vitem, note, buyer, dept, clss, season, pline, uom,
                sku, dim1, dim2, dim3, dim1_descr, dim2_descr, dim3_descr, status, sql, type, mktg As String
            Dim buyunit, sellunit As Int16
            Dim dim1_seq, dim2_seq, dim3_seq As Integer
            Dim cost, retail As Decimal
            Console.WriteLine("Cleaning XML import file")
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim dt2 As DataTable = New DataTable
            dt2.Columns.Add("Sku")
            dt2.Columns.Add("Description")
            dt2.Columns.Add("Vendor_Id")
            dt2.Columns.Add("Vendor")
            dt2.Columns.Add("Vend_Item_No")
            dt2.Columns.Add("Dept")
            dt2.Columns.Add("Buyer")
            dt2.Columns.Add("Class")
            dt2.Columns.Add("Curr_Cost")
            dt2.Columns.Add("Curr_Retail")
            dt2.Columns.Add("Season")
            dt2.Columns.Add("PLine")
            dt2.Columns.Add("Mktg")
            dt2.Columns.Add("UOM")
            dt2.Columns.Add("BuyUnit")
            dt2.Columns.Add("SellUnit")
            dt2.Columns.Add("Note")
            dt2.Columns.Add("Status")
            dt2.Columns.Add("Type")
            dt2.Columns.Add("Initial_Date", GetType(System.DateTime))
            dt2.Columns.Add("Last_Change_Date", GetType(System.DateTime))
            dt2.Columns.Add("Item")
            dt2.Columns.Add("DIM1")
            dt2.Columns.Add("DIM2")
            dt2.Columns.Add("DIM3")
            dt2.Columns.Add("DIM1_DESCR")
            dt2.Columns.Add("DIM2_DESCR")
            dt2.Columns.Add("DIM3_DESCR")
            dt2.Columns.Add("DIM1_SEQ_NO")
            dt2.Columns.Add("DIM2_SEQ_NO")
            dt2.Columns.Add("DIM3_SEQ_NO")
            Dim xmlFile As XmlReader
            Dim rw As DataRow
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    cnt += 1
                    If cnt Mod 1000 = 0 Then
                        Console.WriteLine("Processed " & cnt & " records.")
                    End If
                    sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                    If InStr(sku, "~") > 0 Then
                        Dim parts() As String = sku.Split("~"c)
                        item = parts(0)
                        dim1 = parts(1)
                        dim2 = parts(2)
                        dim3 = parts(3)
                    Else
                        item = sku
                        dim1 = Nothing
                        dim2 = Nothing
                        dim3 = Nothing
                    End If
                    dim1_descr = Trim(Microsoft.VisualBasic.Left(row("DIM1_DESCR"), 30))
                    dim2_descr = Trim(Microsoft.VisualBasic.Left(row("DIM2_DESCR"), 30))
                    dim3_descr = Trim(Microsoft.VisualBasic.Left(row("DIM3_DESCR"), 30))
                    oTest = row("DIM_1_SEQ_NO")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then dim1_seq = CInt(oTest) Else dim1_seq = 1
                    End If
                    oTest = row("DIM_2_SEQ_NO")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then dim2_seq = CInt(oTest) Else dim2_seq = 1
                    End If
                    oTest = row("DIM_3_SEQ_NO")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then dim3_seq = CInt(oTest) Else dim3_seq = 1
                    End If
                    descr = Trim(Microsoft.VisualBasic.Left(row("DESCRIPTION"), 40))
                    vid = Trim(Microsoft.VisualBasic.Left(row("VENDOR_ID"), 20))
                    vendor = Trim(Microsoft.VisualBasic.Left(row("VENDOR"), 40))
                    vitem = Trim(Microsoft.VisualBasic.Left(row("VEND_ITEM_NO"), 20))
                    buyer = Trim(Microsoft.VisualBasic.Left(row("BUYER"), 10))
                    If buyer = "" Then buyer = "NA"
                    dept = Trim(Microsoft.VisualBasic.Left(row("DEPT"), 10))
                    clss = Trim(Microsoft.VisualBasic.Left(row("CLASS"), 10))
                    season = Trim(Microsoft.VisualBasic.Left(row("SEASON"), 20))
                    pline = Trim(Microsoft.VisualBasic.Left(row("PLINE"), 20))
                    mktg = Trim(Microsoft.VisualBasic.Left(row("MKTG"), 20))
                    note = Trim(Microsoft.VisualBasic.Left(row("NOTE"), 20))
                    uom = Trim(Microsoft.VisualBasic.Left(row("UOM"), 10))
                    status = Trim(Microsoft.VisualBasic.Left(row("STATUS"), 10))
                    type = Trim(Microsoft.VisualBasic.Left(row("TYPE"), 1))
                    oTest = row("CURR_COST")
                    isNumber = IsNumeric(oTest)
                    If isNumber Then cost = CDec(oTest) Else cost = 0
                    oTest = row("CURR_RTL")
                    isNumber = IsNumeric(oTest)
                    If isNumber Then retail = CDec(oTest) Else retail = 0
                    uom = row("UOM")
                    If uom <> "" Then
                        oTest = row("BUYUNIT")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then buyunit = oTest
                        oTest = row("SELLUNIT")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then sellunit = oTest
                    Else
                        uom = "EA"
                        buyunit = 1
                        sellunit = 1
                    End If
                    rw = dt2.NewRow
                    oTest = sku
                    rw("Sku") = sku
                    rw("Description") = descr
                    rw("Vendor_Id") = vid
                    rw("Vendor") = vendor
                    rw("Vend_Item_No") = vitem
                    rw("Dept") = dept
                    rw("Buyer") = buyer
                    rw("Class") = clss
                    rw("Curr_Cost") = cost
                    rw("Curr_Retail") = retail
                    rw("Season") = season
                    rw("PLine") = pline
                    rw("Mktg") = mktg
                    rw("UOM") = uom
                    rw("BuyUnit") = buyunit
                    rw("SellUnit") = sellunit
                    rw("Note") = note
                    rw("Status") = status
                    rw("Type") = type
                    rw("Initial_Date") = Date.Today
                    rw("Last_Change_Date") = Date.Today
                    rw("Item") = item
                    rw("DIM1") = dim1
                    rw("DIM2") = dim2
                    rw("DIM3") = dim3
                    rw("DIM1_DESCR") = dim1_descr
                    rw("DIM2_DESCR") = dim2_descr
                    rw("DIM3_DESCR") = dim3_descr
                    ''rw("DIM1_SEQ_NO") = dim1_seq
                    ''rw("DIM2_SEQ_NO") = dim2_seq
                    ''rw("DIM3_SEQ_NO") = dim3_seq
                    dt2.Rows.Add(rw)
                Next
            End If

            Console.WriteLine("DELETE FROM ITEM_MASTER")
            con.Open()
            sql = "DELETE FROM Item_Master"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Console.WriteLine("Bulk Insert into Item_Master")
            Dim connection As SqlConnection = New SqlConnection(constr)
            Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
            connection.Open()
            bulkCopy.DestinationTableName = "dbo.Item_Master"
            bulkCopy.BulkCopyTimeout = 120
            bulkCopy.WriteToServer(dt2)
            connection.Close()

            Dim m As String = "Created " & cnt & " Records"
            Call Update_Process_Log("1", "Recreate Item_Master", m, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Items")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Process_Inventory(ByVal thePath, con, con2)
        Try
            stopWatch.Start()
            con.open()
            Dim sql As String
            sql = "IF OBJECT_ID('dbo._t1','U') IS NOT NULL DROP TABLE dbo._t1; "
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE [dbo].[_t1](" & _
               "[Loc_Id] [varchar](20) NOT NULL," & _
               "[Sku] [varchar](90) NOT NULL," & _
               "[sDate] [date] NOT NULL," & _
               "[eDate] [date] NOT NULL," & _
               "[OnHand] [decimal](18, 4) NULL," & _
               "[Cost] [decimal](10, 2) NULL," & _
               "[Retail] [decimal](10, 2) NULL," & _
               "[YrWk] [int] NULL," & _
               "[Begin_OH] [decimal](18, 4) NULL," & _
               "[End_OH] [decimal](18, 4) NULL," & _
               "[Committed] [decimal](18,4) NULL, " & _
               "[Sys_OH] [decimal](18, 4) NULL," & _
               "[Max_OH] [decimal](18, 4) NULL," & _
               "[Item] [varchar](30) NOT NULL," & _
               "[DIM1] [varchar](30) NULL," & _
               "[DIM2] [varchar](30) NULL," & _
               "[DIM3] [varchar](30) NULL) ON [PRIMARY]"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.close()

            Dim cnt As Int64 = 0
            Dim oTest As Object
            Dim onhand, avail, committed As Decimal
            Dim tqty As Decimal = 0
            Dim item, dim1, dim2, dim3, sku, loc As String
            Dim cost, retail, tcost, tretail As Decimal
            Dim dte As Date = Date.Today
            Dim dttime As DateTime
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim rdr As SqlDataReader
            Dim thisSdate, thisEdate As Date
            con.open()
            sql = "SELECT sDate, eDate FROM Calendar WHERE '" & Date.Today & "' BETWEEN sDate AND eDate AND Week_ID > 0"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader()
            While rdr.Read
                If Not IsDBNull(rdr("eDate")) And Not IsNothing(rdr("eDate")) Then thisEdate = CDate(rdr("eDate"))
                If Not IsDBNull(rdr("sDate")) And Not IsNothing(rdr("sDate")) Then thisSdate = CDate(rdr("sDate"))
            End While
            con.close()

            con.Open()
            Dim xmlFile As XmlReader
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)                                                                      '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    If row("Sku") <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                        If InStr(sku, "~") > 0 Then
                            Dim parts() As String = sku.Split("~"c)
                            item = parts(0)
                            dim1 = parts(1)
                            dim2 = parts(2)
                            dim3 = parts(3)
                        Else
                            item = sku
                            dim1 = Nothing
                            dim2 = Nothing
                            dim3 = Nothing
                        End If

                        loc = Trim(row("LOCATION"))
                        oTest = row("OnHand")
                        If IsDBNull(oTest) Then onhand = 0 Else onhand = CDec(oTest)
                        oTest = row("COMMITTED")
                        If IsDBNull(oTest) Then committed = 0 Else committed = CDec(oTest)
                        oTest = row("AVAIL")
                        If IsDBNull(oTest) Then avail = 0 Else avail = CDec(oTest)
                        oTest = row("COST")
                        If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                        oTest = row("RETAIL")
                        If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                        totlQty += onhand
                        totlCost += (cost * onhand)
                        totlRetail += (retail * onhand)
                        oTest = row("EXTRACT_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If oTest <> "" Then
                                dttime = oTest
                            Else : dttime = DateAndTime.Now
                            End If
                        Else : dttime = DateAndTime.Now
                        End If
                        sql = "INSERT INTO _t1 (Loc_Id, Sku, sDate, eDate, OnHand, End_OH, " & _
                             "Committed, Cost, Retail, YrWk, " & _
                             "Item, DIM1, DIM2, DIM3) " & _
                             "SELECT '" & loc & "','" & sku & "','" & thisSdate & "','" & thisEdate &
                             "'," & onhand & "," & avail & "," & committed & "," & cost & "," &
                             retail & ", YrWk ,'" & item & "','" & dim1 & "','" & dim2 & "','" &
                             dim3 & "' FROM Calendar WHERE eDate = '" & thisEdate & "' AND Prd_Id > 0 " & _
                             "AND Week_Id > 0"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()

                    Else
                        oTest = row("OnHand")
                        If IsNumeric(oTest) Then onhand = CDec(row("OnHand"))
                        oTest = row("AVAIL")
                        If IsNumeric(oTest) Then avail = CDec(oTest)
                        oTest = row("COST")
                        If IsNumeric(oTest) Then cost = CDec(row("COST"))
                        oTest = row("RETAIL")
                        If IsNumeric(oTest) Then retail = CDec(row("RETAIL"))
                        If onhand <> totlQty Then
                            msg = "Expected " & Format(avail, "###,###,##0.0000") & " Received " & Format(tqty, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Quantity Mismatch " & conString, "Process Inventory")
                        End If
                        If cost <> totlCost Then
                            msg = "Expected " & Format(cost, "$###,###,##0.00") & " Received " & Format(tcost, "$###,###,##0.00")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Cost Mismatch", "Process Inventory")
                        End If
                        If retail <> totlRetail Then
                            msg = "Expected " & Format(retail, "$###,###,##0.00") & " Received " & Format(tretail, "$###,###,##0.00")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Retail Mismatch", "Process Inventory")
                        End If
                    End If
                Next
            End If
            con.close()

            con.open()
            sql = "UPDATE Item_Inv SET End_OH = 0, Max_OH = 0, Committed = 0 WHERE eDate = '" & thisEdate & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "MERGE Item_Inv AS target USING _t1 AS source ON (source.Loc_Id = target.Loc_Id AND source.Sku = target.Sku " & _
                    "AND source.eDate = target.eDate) " & _
                "WHEN NOT MATCHED BY TARGET THEN " & _
                    "INSERT (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, OnHand, Begin_OH, End_OH, Committed, " & _
                        "Max_OH, Item, DIM1, DIM2, DIM3) " & _
                    "VALUES (source.Loc_Id, source.Sku, source.sDate, source.eDate, source.Cost, source.Retail, " & _
                        "source.YrWk, source.OnHand, source.Begin_OH, source.End_OH, source.Committed, source.Max_OH, " & _
                        "source.Item, source.DIM1, source.DIM2, source.DIM3) " & _
                "WHEN MATCHED THEN " & _
                    "UPDATE SET target.OnHand = source.OnHand, target.Begin_OH = source.Begin_OH, " & _
                    "target.End_OH = source.End_OH, target.Cost = source.Cost, target.Retail = source.retail, " & _
                    "target.Max_OH = source.Max_OH, target.Yrwk = source.YrWk, target.Committed = source.Committed;"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            sql = "IF EXISTS (SELECT * FROM sysobjects WHERE name = '_t1' AND xtype = 'U') DROP TABLE dbo._t1 "
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", "Update Inventory", m, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Inventory")
        End Try
    End Sub

    Private Sub Process_PREQ_Header(ByVal thePath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim preq, batch, vend_no, vend, loc, buyer, alloc, mrg As String
            Dim ord, del, can, extract As Date
            Dim amt, totlCost As Decimal
            totlCost = 0
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim xmlFile As XmlReader
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                con2.open()
                sql = "DELETE FROM Purchase_Request_Header"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("VEND_NO")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        oTest = row("PREQ_NO")
                        If Not IsDBNull(oTest) Then preq = CStr(oTest) Else preq = ""
                        oTest = row("BATCH")
                        If Not IsDBNull(oTest) Then batch = CStr(oTest)
                        oTest = row("VEND_NO")
                        If Not IsDBNull(oTest) Then vend_no = CStr(oTest)
                        oTest = row("VENDOR")
                        If Not IsDBNull(oTest) Then vend = CStr(oTest)
                        oTest = row("LOC_ID")
                        If IsDBNull(oTest) Then oTest = "UNKNOWN"
                        loc = CStr(oTest)
                        oTest = row("BUYER")
                        If IsDBNull(oTest) Then buyer = "UNKNOWN"
                        buyer = CStr(oTest)
                        oTest = row("ORDER_DATE")
                        If IsDBNull(oTest) Then oTest = 0
                        ord = CDate(oTest)
                        oTest = row("DELIVER_DATE")
                        If IsDBNull(oTest) Then oTest = 0
                        del = CDate(oTest)

                        oTest = row("CANCEL_DATE")
                        If IsDBNull(oTest) Then oTest = 0
                        can = CDate(oTest)
                        oTest = row("ORDER_TOTAL")
                        If IsDBNull(oTest) Then amt = 0 Else amt = CDec(oTest)
                        totlCost += amt
                        oTest = row("ISALLOCATED")
                        If IsDBNull(oTest) Then oTest = "N"
                        alloc = CStr(oTest)
                        oTest = row("MERGED")
                        If IsDBNull(oTest) Then oTest = "S"
                        mrg = CStr(oTest)
                        oTest = row("EXTRACT_DATE")
                        If IsDBNull(oTest) Then oTest = 0
                        extract = CDate(oTest)
                        sql = "INSERT INTO Purchase_Request_Header(PREQ_NO, Batch_Id, Loc_Id, " & _
                            "Vendor_Id, Buyer, Order_Date, Deliver_Date, Cancel_Date, Order_Total, " & _
                            "Allocated, Merged, Last_Update) " & _
                            "SELECT '" & preq & "','" & batch & "','" & loc & "','" & vend_no & "','" & _
                            buyer & "','" & ord & "','" & del & "','" & can & "'," & totlCost & ",'" & _
                            alloc & "','" & mrg & "','" & Date.Today & "'"
                        cmd = New SqlCommand(sql, con2)
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("ORDER_TOTAL")
                        If IsDBNull(oTest) Then amt = 0 Else amt = CDec(oTest)
                        If amt <> totlCost Then
                            msg = "Expected " & Format(amt, "###,###,##0.00") & " Received " & Format(totlCost, "###,###,##0.00")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Record Count Mismatch " & conString, "Process PREQ Header")
                        End If
                    End If
                Next
                con.close()
                con2.close()
            End If
            Dim m As String = "Created " & cnt & " records"
            Call Update_Process_Log("1", "Create Purchase Request Header", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process PREQ Header")
        End Try
    End Sub

    Private Sub Process_PREQ_Detail(ByVal thePath, ByVal con, byvalcon2)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim preq, item, desc, uom, dim1, dim2, dim3, sku As String
            Dim seq, numer, denom As Integer
            Dim cost, qty As Decimal
            Dim totlCost As Decimal = 0
            Dim totlQty As Decimal = 0
            Dim extract As Date
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim xmlfile As XmlReader
            xmlfile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlfile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                con2.Open()
                sql = "DELETE FROM Purchase_Request_Detail"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                dt = ds.Tables(0)
                For Each row As DataRow In dt.Rows
                    oTest = row("ITEM_NO")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        oTest = row("PREQ_NO")
                        If Not IsDBNull(oTest) Then preq = CStr(oTest) Else preq = ""
                        oTest = row("SEQ_NO")
                        If Not IsDBNull(oTest) Then seq = CInt(oTest) Else seq = 0
                        sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                        If InStr(sku, "~") > 0 Then
                            Dim parts() As String = sku.Split("~"c)
                            item = parts(0)
                            dim1 = parts(1)
                            dim2 = parts(2)
                            dim3 = parts(3)
                        Else
                            item = sku
                            dim1 = Nothing
                            dim2 = Nothing
                            dim3 = Nothing
                        End If
                        
                        oTest = row("DESC")
                        If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = ""
                        oTest = row("UOM")
                        If Not IsDBNull(oTest) Then uom = CStr(oTest) Else uom = "EACH"
                        oTest = row("NUMER")
                        If IsDBNull(oTest) Or Not IsNumeric(oTest) Then numer = 1 Else numer = CInt(oTest)
                        oTest = row("DENOM")
                        If IsDBNull(oTest) Or Not IsNothing(oTest) Then denom = 1 Else denom = CInt(oTest)
                        oTest = row("COST")
                        If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                        oTest = row("QTY")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        qty = CDec(oTest)
                        totlQty += qty
                        totlCost += (qty * cost)
                        sql = "INSERT INTO Purchase_Request_Detail(PREQ_NO, Seq_No, Sku, Description, " & _
                            "Stock_Units, Numer, Denom, Cost, Item, DIM1, DIM2, DIM3, Qty) " & _
                            "SELECT '" & preq & "'," & seq & ",'" & sku & "','" & desc & "','" & _
                            uom & "'," & numer & "," & denom & "," & cost & ",'" & item & "','" & _
                            dim1 & "','" & dim2 & "','" & dim3 & "'," & qty & " "
                        cmd = New SqlCommand(sql, con2)
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("COST")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        cost = CDec(oTest)
                        oTest = row("QTY")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        qty = CDec(oTest)
                        If cost <> totlCost Then
                            msg = "Expected " & Format(cost, "###,###,##0.00") & " Received " & Format(totlCost, "###,###,##0.00")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Record Count Mismatch " & conString, "Process PREQ Detail")
                        End If
                        If qty <> totlQty Then
                            msg = "Expected " & Format(qty, "###,###,##0.00") & " Received " & Format(totlQty, "###,###,##0.00")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Record Count Mismatch " & conString, "Process PREQ Detail")
                        End If
                    End If
                Next
                con.close()
                con2.Close()
                Dim m As String = "Created " & cnt & " Records"
                Call Update_Process_Log("1", "Create Purchase Request Detail", m, "")
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process PREQ Detail")
        End Try
    End Sub

    Private Sub Process_POHeader(ByVal thePath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim recvd, lines, open_lines As Integer
            Dim amt, recvd_cost, ord_qty, open_amt As Decimal
            Dim store, po, vendor, buyer, stat As String
            Dim ordDate, dueDate, canDate As Date
            Dim datetime As DateTime
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim xmlFile As XmlReader
            Dim row As DataRow
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                sql = "DELETE FROM PO_Header"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("PO")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        oTest = row("LOCATION")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = CStr(oTest) Else store = "1"
                        oTest = row("PO")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then po = CStr(oTest) Else po = "NA"
                        oTest = row("OrderDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then ordDate = CDate(oTest) Else ordDate = "1900-01-01"
                        oTest = row("DueDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then dueDate = CDate(oTest) Else dueDate = "1900-01-01"
                        oTest = row("CancelDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then canDate = CDate(oTest) Else canDate = "1900-01-01"
                        oTest = row("Vendor")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vendor = CStr(oTest) Else vendor = "NA"
                        oTest = row("BUYER")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = CStr(oTest) Else buyer = "NA"
                        oTest = row("STATUS")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then stat = CStr(oTest) Else stat = "X"
                        oTest = row("AMT")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = CDec(oTest) Else amt = 0
                        oTest = row("RECVD_COST")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then recvd_cost = CDec(oTest) Else recvd_cost = 0
                        oTest = row("LINES")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lines = CInt(oTest) Else lines = 0
                        oTest = row("ORD_QTY")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ord_qty = CDec(oTest) Else ord_qty = 0
                        oTest = row("OPEN_LINES")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then open_lines = CInt(oTest) Else open_lines = 0
                        oTest = row("OPEN_AMT")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then open_amt = CDec(oTest) Else open_amt = 0
                        oTest = row("EXTRACT_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then datetime = CDate(oTest) Else datetime = Nothing
                        sql = "INSERT INTO PO_Header (Loc_Id, PO_NO, Order_Date, Due_Date, Cancel_Date, Vendor_Id, Buyer, Status, " & _
                            "Amount, Recvd_Cost, Lines, Ord_Qty, Open_Lines, Open_Amt) " & _
                            "SELECT '" & store & "','" & po & "','" & ordDate & "','" & dueDate & "','" & canDate & "','" &
                            vendor & "','" & buyer & "','" & stat & "', " & amt & ", " & recvd_cost & ", " & lines & ", " &
                            ord_qty & ", " & open_lines & ", " & open_amt
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("ORDERDATE")
                        If IsNumeric(oTest) Then recvd = CInt(oTest)
                        If cnt <> recvd Then
                            msg = "Expected " & Format(cnt, "###,###,##0") & " Received " & Format(cnt, "###,###,##0")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Record Count Mismatch " & conString, "Process PO Header")
                        End If
                    End If
                Next
                con.close()
            End If
            Dim m As String = "Created " & cnt & " records"
            Call Update_Process_Log("1", "Create POHeader", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process POHeader")
        End Try
    End Sub


    Private Sub Process_PODetail(ByVal thepath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Int32 = 0
            Dim oTest As Object
            Dim po, item, loc, sql, lastDate As String
            Dim seq As Int32
            Dim cost, retail, ordQty, recvdQty, expQty, tqty, trecvd, tcan, tcost, tretail As Decimal
            Dim canqty As Decimal
            Dim extractDate As DateTime
            Dim cmd As SqlCommand
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim loDate As Date = "2012-01-01"
            Dim thisDate As Date = Date.Today
            Dim xmlFile As XmlReader
            Dim row As DataRow
            xmlFile = Xml.XmlReader.Create(thepath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                sql = "DELETE FROM PO_Detail"
                con.open()
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                dt = ds.Tables(0)
                tqty = 0
                For Each row In dt.Rows
                    oTest = row("Sku")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        loc = Trim(row("LOCATION").ToString)
                        po = row("PO_NO")
                        oTest = row("SEQ_NO").ToString
                        seq = CInt(oTest)
                        item = Trim(row("Sku").ToString)
                        oTest = row("ORD_QTY")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ordQty = CDec(oTest) Else ordQty = 0
                        oTest = row("QTY_RECVD")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then recvdQty = CDec(oTest) Else recvdQty = 0
                        oTest = row("EXP_QTY")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then expQty = CDec(oTest) Else expQty = 0
                        tqty += expQty
                        oTest = row("LAST_RECVD_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then lastDate = CDate(oTest) Else lastDate = Nothing
                        oTest = row("COST")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then cost = CDec(oTest) Else cost = 0
                        tcost += cost * expQty
                        oTest = row("RETAIL")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then retail = CDec(oTest) Else retail = 0
                        tretail += retail * expQty
                        'vendor = row("VEND_NO")
                        'buyer = row("BUYER")
                        'stat = row("STATUS")
                        extractDate = row("EXTRACT_DATE")
                        If lastDate <> "" Then
                            sql = "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, " & _
                                "Cost, Retail, Last_Recvd_Date, Item, DIM1, DIM2, DIM3) " & _
                                "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & ordQty & "," & recvdQty & "," &
                                expQty & "," & cost & "," & retail & ",'" & lastDate & "',Item, DIM1, DIM2, DIM3 FROM Item_Master " & _
                                "WHERE Sku = '" & item & "'"
                        Else
                            sql = "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, " & _
                                "Cost, Retail, Last_Recvd_Date, Item, DIM1, DIM2, DIM3) " & _
                                "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & ordQty & "," & recvdQty & "," &
                                expQty & "," & cost & "," & retail & ", NULL, Item, DIM1, DIM2, DIM3 FROM Item_Master " & _
                                "WHERE Sku = '" & item & "'"
                        End If
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                    Else
                        Dim oqty, rqty, cqty, ocost, oretail As Decimal
                        oTest = row("ORD_QTY")
                        If IsNumeric(oTest) Then oqty = CDec(oTest)
                        oTest = row("QTY_RECVD")
                        If IsNumeric(oTest) Then rqty = CDec(oTest)
                        oTest = row("EXP_QTY")
                        If IsNumeric(oTest) Then cqty = CDec(oTest)
                        oTest = row("COST")
                        If IsNumeric(oTest) Then ocost = CDec(oTest)
                        oTest = row("RETAIL")
                        If IsNumeric(oTest) Then oretail = CDec(oTest)
                        If oqty <> tqty Then
                            msg = "Expected " & Format(oqty, "###,###,##0.0000") & " Received " & Format(ordQty, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Order Quantity Mismatch " & conString, "Process PO Detail")
                        End If
                        If rqty <> trecvd Then
                            msg = "Expected " & Format(rqty, "###,###,##0.0000") & " Received " & Format(recvdQty, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Received Quantity Mismatch " & conString, "Process PO Detail")
                        End If
                        If cqty <> tqty Then
                            msg = "Expected " & Format(cqty, "###,###,##0.0000") & " Received " & Format(tqty, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Quantity Due Mismatch " & conString, "Process PO Detail")
                        End If
                        If ocost <> tcost Then
                            msg = "Expected " & Format(ocost, "$###,###,##0.0000") & " Received " & Format(cost, "$###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Ordered @ Cost Mismatch " & conString, "Process PO Detail")
                        End If
                        If oretail <> tretail Then
                            msg = "Expected " & Format(oretail, "$###,###,##0.0000") & " Received " & Format(retail, "####,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Ordered @ Retail Mismatch " & conString, "Process PO Detail")
                        End If
                    End If
                Next
                con.close()
            End If
            Dim m As String = "Created " & cnt & " records"
            Call Update_Process_Log("1", "Create PODetail", m, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process PODetail")
        End Try
    End Sub

    Private Sub Process_Data(ByVal thePath, ByVal theField, con, con2)
        Try
            stopWatch = New Stopwatch
            stopWatch.Start()
            Dim cnt As Int32 = 0
            Dim oTest As Object
            Dim isPhysical As Boolean = False
            Dim id, sku, item, dim1, dim2, dim3, store, loc, dept, buyer, clss, sql, cust, tkt, reason,
                coupon, typ, nam, cust_typ, ttype As String
            Dim seq As Int32
            Dim cost, retail, mkdn, tCost, tRetail, tMkdn, qty, tQty As Decimal
            Dim dte, eDate As Date
            Dim transDate As DateTime
            Dim extractDate, datetimenow As DateTime
            Dim cmd As SqlCommand
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim rdr As SqlDataReader
            Dim loDate As Date = "2012-01-01"
            Dim thisDate As Date = Date.Today
            Dim xmlFile As XmlReader
            Dim row As DataRow
            If theField = "UnpostedSales" Then
                con.open()
                sql = "DELETE FROM Unposted_Sales"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.close()
            End If
            If InStr(thePath, "Physical") > 0 Then isPhysical = True
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                totlQty = 0
                totlCost = 0
                totlRetail = 0
                totlMarkdown = 0
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("SKU")
                    If Not IsDBNull(oTest) AndAlso oTest <> "TOTALS" Then
                        qty = 0
                        cnt += 1
                        If cnt Mod 100 = 0 Then Console.WriteLine(cnt)
                        oTest = row("TRANS_ID")
                        If Not IsDBNull(oTest) Then id = oTest Else id = ""
                        If theField = "ADJ" Then
                            If isPhysical Then
                                id = "P" & id
                            Else : id = "A" & id
                            End If
                        End If
                        oTest = row("SEQ_NO").ToString
                        If Not IsDBNull(oTest) Then seq = CInt(oTest) Else seq = 0
                        oTest = row("SKU")
                        If Not IsDBNull(oTest) Then sku = Trim(Microsoft.VisualBasic.Left(oTest, 90)) Else sku = ""
                        If InStr(sku, "~") > 0 Then
                            Dim parts() As String = sku.Split("~"c)
                            item = parts(0)
                            dim1 = parts(1)
                            dim2 = parts(2)
                            dim3 = parts(3)
                        Else
                            item = sku
                            dim1 = Nothing
                            dim2 = Nothing
                            dim3 = Nothing
                        End If
                        oTest = Trim(row("LOCATION"))
                        If Not IsDBNull(oTest) Then loc = CStr(oTest)
                        oTest = row("QTY")
                        qty = 0
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then qty = oTest
                        End If
                        cost = 0
                        oTest = row("COST")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then cost = oTest
                        End If
                        retail = 0
                        oTest = row("RETAIL")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then retail = oTest
                        End If
                        totlQty += qty
                        totlCost += (qty * cost)
                        totlRetail += (qty * retail)
                        oTest = row("TRANS_DATE")
                        If oTest = "0" Then oTest = "1900-01-01 00:00:00"

                        If CDate(oTest) < loDate Then GoTo 10

                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then
                            transDate = DateTime.Parse(oTest)
                        Else : transDate = "1900-01-01 00:00:00"
                        End If

                        Select Case theField
                            Case "Sold"
                                oTest = row("STR_ID")
                                If IsDBNull(oTest) Then store = "UNKNOWN" Else store = CStr(oTest)
                                oTest = row("DEPT")
                                If IsDBNull(oTest) Then dept = "UNKNOWN else dept =cstr(otest)"
                                oTest = row("BUYER")
                                If IsDBNull(oTest) Then buyer = "UNKNOWN" Else buyer = CStr(oTest)
                                oTest = row("CLASS")
                                If IsDBNull(oTest) Then clss = "UNKNOWN" Else clss = CStr(oTest)
                                oTest = row("CUST_NO")
                                If Not IsDBNull(oTest) Then cust = CStr(oTest)
                                oTest = row("TKT_NO")
                                If Not IsDBNull(oTest) Then tkt = oTest
                                oTest = row("MARKDOWN")
                                If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then mkdn = CDec(oTest) Else mkdn = 0
                                totlMarkdown += mkdn
                                oTest = row("MKDN_REASON")
                                If Not IsDBNull(oTest) Then reason = CStr(oTest)
                                oTest = row("COUPON_CODE")
                                If Not IsDBNull(oTest) Then coupon = CStr(oTest)
                            Case Else
                                mkdn = 0
                                store = ""
                                dept = ""
                                buyer = ""
                                clss = ""
                                reason = ""
                                coupon = ""
                                cust = ""
                                tkt = ""
                        End Select

                        oTest = row("EXTRACT_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then extractDate = oTest

                        Dim itexists As Boolean = Check_Log(con2, id, seq, loc, sku, transDate, store)
                        If Not itexists Then
                            tQty += qty
                            tCost += (qty * cost)
                            tRetail += (qty * retail)
                            tMkdn += mkdn
                            con.open()
                            sql = "SELECT eDate FROM Calendar WHERE '" & CDate(transDate) & "' BETWEEN sDate AND eDate AND Week_ID > 0"
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                If Not IsDBNull(rdr(0)) And Not IsNothing(rdr(0)) Then eDate = rdr(0) Else eDate = "2012-01-01"
                            End While
                            con.close()

                            cust_typ = "UNKNOWN"
                            If theField = "Sold" Then
                                con.open()
                                sql = "SELECT [Type] FROM Customers WHERE Cust_No = '" & cust & "'"
                                cmd = New SqlCommand(sql, con)
                                rdr = cmd.ExecuteReader
                                While rdr.Read
                                    cust_typ = rdr("Type")
                                End While
                                con.close()
                            End If
                            row = eDateTBL.Rows.Find(eDate)              ' We need to send the eDate to Client_Weekly_Summary
                            If IsNothing(row) Then
                                eDateTBL.Rows.Add(eDate)                 ' Add date to eDateTBL if it isn't already there
                            End If
                            Call Update_Tables(store, loc, eDate, sku, qty, retail, cost, mkdn, theField, dept, buyer, clss, "", cust_typ, con2)

                            datetimenow = DateAndTime.Now
                            con.Open()
                            sql = "IF NOT EXISTS (SELECT TRANS_ID FROM Daily_Transaction_Log WHERE TRANS_ID = '" & id & "' " & _
                                "AND SEQUENCE_NO = '" & seq & "' AND STORE = '" & store & "' AND SKU = '" & sku & "' " & _
                                "AND LOCATION = '" & loc & "' " & " AND TRANS_DATE = '" & transDate & "') " & _
                                "INSERT INTO Daily_Transaction_Log (TRANS_ID, SEQUENCE_NO, STORE, SKU, LOCATION, QTY, COST, RETAIL, " & _
                                "MKDN, TRANS_DATE, [TYPE], POST_DATE, DEPT,BUYER, CLASS, EXTRACT_DATE, CUST_NO, TKT_NO, MKDN_REASON, " & _
                                "COUPON_CODE, ITEM, DIM1, DIM2, DIM3) " & _
                                "SELECT '" & id & "', " & seq & ", '" & store & "', '" & sku & "','" & loc & "'," &
                                qty & "," & cost & "," & retail & ", " & mkdn & ",'" & transDate & "','" & theField & "','" & datetimenow & "','" & _
                                dept & "','" & buyer & "','" & clss & "', '" & extractDate & "','" & cust & "','" & tkt & "','" & _
                                reason & "','" & coupon & "','" & item & "','" & dim1 & "','" & dim2 & "','" & dim3 & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()

                            con.Close()
                        End If
                    Else
                        qty = CDec(row("QTY"))
                        cost = CDec(row("COST"))
                        retail = CDec(row("RETAIL"))
                        If theField = "Sold" Then mkdn = CDec(row("MARKDOWN"))
                        If qty <> totlQty Then
                            msg = "Expected " & qty & " Received " & totlQty & " "
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Quantity Mismatch", theField)
                        End If
                        If cost <> totlCost Then
                            msg = "Expected " & cost & " Received " & totlCost & " Cost"
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Cost Mismatch", theField)
                        End If
                        If retail <> totlRetail Then
                            msg = "Expected " & retail & " Received " & totlRetail & " Retail"
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Retail Mismatch", theField)
                        End If
                        If theField = "Sold" Then
                            If mkdn <> totlMarkdown Then
                                msg = "Expected " & mkdn & " Received " & totlMarkdown & " Markdown"
                                Dim el As New XMLUpdate.ErrorLogger
                                el.WriteToErrorLog(msg, "Markdown Mismatch", theField)
                            End If
                        End If
                    End If
10:             Next
            End If

            If isPhysical Then theField = "PHYS"
            Dim message As String = "Updated " & cnt & " records - Qty = " & qty &
                " Total Cost = " & Format(tCost, "###,###,###.00") & " Total Retail = " & Format(tRetail, "###,###,###.00") &
                " Total Markdowns = " & Format(tMkdn, "###,###,###.00")
            Call Update_Process_Log("1", "Update " & theField & "", message, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Data " & theField)
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Function Check_Log(ByVal con2, ByVal id, ByVal seq, ByVal location, ByVal sku, ByVal dte, ByVal store)
        Try
            Dim reslt As Boolean = False
            If con2.State = ConnectionState.Open Then con2.Close()
            con2.open()
            Dim sql As String = "SELECT TRANS_ID FROM Daily_Transaction_Log WHERE TRANS_ID = '" & id & "' " & _
                "AND SEQUENCE_NO = " & seq & " AND LOCATION = '" & location & "' AND SKU = '" & sku & "' " & _
                "AND TRANS_DATE = '" & dte & "'"
            If store <> "" Then sql &= " AND STORE = '" & store & "'"
            Dim cmd As SqlCommand = New SqlCommand(sql, con2)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            While rdr.Read
                If Not IsNothing(rdr(0)) And Not IsDBNull(rdr(0)) Then
                    reslt = True
                    con2.close()
                    Return reslt
                End If
            End While
            con2.close()
            Return reslt
        Catch ex As Exception
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Check Log")
        End Try
    End Function

    Private Sub Update_Tables(ByVal store, ByVal location, ByVal dte, ByVal item, ByVal qty, ByVal retail,
                             ByVal cost, ByVal mkdn, ByVal theField, ByVal dept, ByVal buyer, ByVal clss,
                             ByVal dttime, ByVal Cust_typ, ByVal con2)
        Try
            Dim cmd As SqlCommand
            Dim itsadate As Date
            Dim thisEdate As Date = dte
            Dim thisSdate As Date = DateAdd(DateInterval.Day, -6, dte)
            Dim prevEdate As Date = DateAdd(DateInterval.Day, -7, thisEdate)
            Dim thisDate As Date = CDate(Date.Today)
            Dim thisDept As String = dept
            Dim thisBuyer As String = buyer
            Dim thisClass As String = clss
            Dim thistime As DateTime
            Dim sql As String = ""
            Dim maxQty, endOH As Decimal
            If con.State = ConnectionState.Closed Then con2.open()
            If DateTime.TryParse(dttime, itsadate) Then
                thistime = dttime
            Else : thistime = DateAndTime.Now
            End If

            If Cust_typ = "WHOLESALE" Then store &= "-WHSL"

            If theField = "Sold" Then
                sql = "IF NOT EXISTS (SELECT * FROM Item_Sales WHERE Str_Id = '" & Trim(store) & "' AND Loc_Id = '" & location & "' " & _
                    "AND Sku = '" & item & "' AND eDate = '" & thisEdate & "') " & _
                    "INSERT INTO Item_Sales (Str_Id, Loc_Id, Sku, sDate, eDate, Sold, Avg_Cost, Retail_Price, Markdown, " & _
                    "Sales_Cost, Sales_Retail, YrWk, Item, DIM1, DIM2, DIM3) " & _
                     "SELECT '" & Trim(store) & "', '" & location & "', '" & Trim(item) & "', '" & thisSdate & "', '" & thisEdate & "', " & _
                     qty & ", " & cost & ", Curr_Retail, " & mkdn & ", " & cost * qty & ", " & (retail * qty) & ", " & _
                     "(SELECT YrWk FROM Calendar WHERE eDate = '" & thisEdate & "' AND Prd_Id > 0 AND Week_Id > 0), " & _
                     "Item, DIM1, DIM2, DIM3 " & _
                     "FROM Item_Master WHERE Sku = '" & Trim(item) & "' " & _
                    "ELSE " & _
                    "UPDATE p SET Sold = ISNULL(Sold,0) + " & qty & ", Sales_Cost = ISNULL(Sales_Cost,0) + " & cost * qty & ", Sales_Retail = " & _
                    "ISNULL(Sales_Retail,0) + " & retail * qty & ", " & _
                    "Markdown = ISNULL(Markdown,0) + " & mkdn & " " & _
                    "FROM Item_Sales AS p " & _
                        "WHERE Str_Id = '" & store & "' AND Loc_Id = '" & location & "' AND Sku = '" & item & "' AND eDate = '" & thisEdate & "' "
            Else
                sql = "IF NOT EXISTS (SELECT * FROM Item_Inv d WHERE Loc_Id = '" & Trim(location) & "' AND Sku = '" & Trim(item) & "' " & _
                    "AND eDate = '" & thisEdate & "') " & _
                    "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, " & theField & ", Cost, Retail, YrWk, Item, DIM1, DIM2, DIM3) " & _
                        "SELECT '" & Trim(location) & "', '" & Trim(item) & "', '" & thisSdate & "', '" & thisEdate & "', " & _
                        qty & ", " & cost & ", " & retail & ", " & _
                        "(SELECT YrWk FROM Calendar WHERE eDate = '" & thisEdate & "' AND Prd_Id > 0 AND Week_Id > 0), " & _
                        "Item, DIM1, DIM2, DIM3 " & _
                        "FROM Item_Master WHERE Sku = '" & Trim(item) & "' " & _
                    "ELSE " & _
                    "UPDATE d SET " & theField & " = ISNULL(" & theField & ",0) + " & qty & " " & _
                    "FROM Item_Inv AS d " & _
                        "WHERE Loc_Id = '" & Trim(location) & "' AND Sku = '" & Trim(item) & "' AND eDate = '" & thisEdate & "'"
            End If
            cmd = New SqlCommand(sql, con2)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            con2.close()

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Update Tables " & theField)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Process_Other_Data(ByVal thePath, ByVal theTable, con, con2)
        Try
            stopWatch = New Stopwatch
            stopWatch.Start()
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim cnt As Int16 = 0
            Dim id, descr, dept As String
            Dim eDate As DateTime
            Dim xDate As Date
            Dim xmlFile As XmlReader
            Dim thisUser = System.Environment.UserName
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)

            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("ID")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 100 = 0 Then Console.WriteLine(cnt)
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            id = oTest
                            descr = row("DESCRIPTION")
                            oTest = row("EXTRACT_DATE")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                eDate = row("EXTRACT_DATE")
                                xDate = CDate(eDate)
                            End If
                            If theTable = "Classes" Then
                                oTest = row("Dept")
                                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest Else dept = ""
                                sql = "IF NOT EXISTS (SELECT ID FROM Classes WHERE ID = '" & id & "' AND Dept = '" & dept & "') " & _
                                    "INSERT INTO Classes (ID, Dept, Description, Orig_Date, Status) " & _
                                    "SELECT '" & id & "','" & dept & "','" & descr & "','" & CDate(Date.Today) & "','Active' " & _
                                    "ELSE " & _
                                    "UPDATE Classes SET Description = '" & descr & "', Dept = '" & dept & "', Extract_Date = '" & eDate & "' " & _
                                    "WHERE ID = '" & id & "' AND Dept = '" & dept & "'"
                            Else
                                sql = "IF NOT EXISTS (SELECT ID FROM " & theTable & " WHERE ID = '" & id & "') " & _
                                    "INSERT INTO " & theTable & " (ID, Description, Orig_Date, Status) " & _
                                    "SELECT '" & id & "','" & descr & "','" & CDate(Date.Today) & "','Active' " & _
                                    "ELSE " & _
                                    "UPDATE " & theTable & " SET Description = '" & descr & "', Extract_Date = '" & eDate & "' " & _
                                    "WHERE ID = '" & id & "' "
                            End If
                            ''Console.WriteLine(sql)
                            con.open()
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                            con.close()
                        End If
                    Else
                        records = CInt(row("DESCRIPTION"))
                        If records <> cnt Then
                            msg = "Expected " & records & " Received " & cnt
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Records Received Mismatch", "Process " & theTable)
                        End If
                    End If
                Next
            End If
            '                           Flag records in History and not in the other table
            ''If theTable <> "Customers" Then
            ''    sql = "UPDATE h SET Status = 'Omitted' FROM " & theTable & " h " & _
            ''        "WHERE Extract_Date <> '" & xDate & "'"
            ''    con.open()
            ''    cmd = New SqlCommand(sql, con)
            ''    cmd.CommandTimeout = 120
            ''    cmd.ExecuteNonQuery()
            ''    con.close()
            ''End If

            con.open()
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            con.close()

            '                           Insert a record for buyer OTHER if we don't already have one
            If theTable = "Buyers" Then
                con.open()
                sql = "IF NOT EXISTS (SELECT * FROM Buyers WHERE ID = 'OTHER') " & _
               "INSERT INTO Buyers (ID, Description, Orig_Date) " & _
               "SELECT 'OTHER','MISCELLANEOUS BUYER','" & CDate(Date.Today) & "'"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()
                con.close()
            End If

            Dim p As String = "Update " & theTable
            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", p, m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Process Other Data " & theTable)
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Merge_Tickets()
        Try
            Console.WriteLine("Updating Tickets table")
            con.Open()
            sql = "DECLARE @maxDate AS Date " & _
                "SET @maxDate = (SELECT MAX(Date) FROM Tickets) " & _
                "SELECT CONVERT(varchar(10),LOCATION) AS Str_Id, CONVERT(varchar(30),TRANS_ID) AS Ticket_No, " & _
                "CONVERT(varchar(30),CUST_NO) AS Cust_No, CONVERT(date,TRANS_DATE) AS 'Date', " & _
                "ISNULL(SUM(QTY * RETAIL),0) as Amt, COUNT(*) AS Items INTO #t1 FROM Daily_Transaction_Log " & _
                "WHERE TYPE = 'Sold' AND TRANS_DATE >= @maxDate GROUP BY LOCATION, TRANS_ID, Cust_No, TRANS_DATE " & _
                "MERGE Tickets AS t " & _
                "USING #t1 AS s " & _
                "ON (t.Str_Id = s.Str_Id AND t.Ticket_No = s.Ticket_No AND t.Cust_No = s.Cust_No AND t.Date = s.Date) " & _
                "WHEN NOT MATCHED BY Target " & _
                "THEN INSERT (Ticket_No, Str_Id, Cust_No, Date, Amt, Items) " & _
                "VALUES (s.Ticket_No, s.Str_Id, s.Cust_No, s.Date, s.Amt, s.Items);"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            Dim m As String = ""
            Call Update_Process_Log("1", "Merge Tickets", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Merge_Ticket")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Set_Max_OH(ByVal thisDate As Date)
        Try
            Dim stopwatch As New Stopwatch
            stopwatch.Start()
            Console.WriteLine("Updating Max_OH")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            If con3.State = ConnectionState.Open Then con3.Close()
            con2.Open()
            Dim maxx, onhand, sold As Decimal
            Dim location, sku As String
            Dim cnt As Integer = 0
            Dim tbl As New DataTable
            Dim edate As Date
            sql = "SELECT Loc_Id, i.Sku, eDate, OnHand INTO #t1 FROM Item_Inv i WHERE eDate = '" & thisDate & "' " & _
               "SELECT Loc_Id, Sku, eDate, SUM(ISNULL(Sold,0)) AS Sold INTO #t2 FROM Item_Sales " & _
               "WHERE eDate = '" & thisDate & "' GROUP BY Loc_Id, Sku, eDate " & _
               "SELECT t1. Loc_Id, t1.Sku, t1.eDate, t1.OnHand, ISNULL(t2.Sold,0) AS Sold FROM #t1 t1 " & _
               "LEFT JOIN #t2 t2 ON t2.Loc_Id = t1.Loc_Id AND t2.Sku = t1.Sku AND t1.eDate = t2.eDate"
            Using con As New SqlConnection(conString)
                con.Open()
                Using adptr As New SqlDataAdapter(sql, con)
                    adptr.Fill(tbl)
                End Using
                con.Close()
            End Using
            For Each row As DataRow In tbl.Rows
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = row("OnHand")
                If IsDBNull(oTest) Then onhand = 0 Else onhand = CDec(oTest)
                oTest = row("Sku")
                If IsDBNull(oTest) Then sku = "" Else sku = Replace(oTest, "'", "''")
                oTest = row("Loc_Id")
                If IsDBNull(oTest) Then location = "" Else location = Replace(oTest, "'", "''")
                oTest = row("eDate")
                If IsDBNull(oTest) Then edate = "1900-01-01" Else edate = CDate(oTest)
                oTest = row("Sold")
                If IsDBNull(oTest) Then sold = 0 Else sold = CDec(oTest)
                If onhand < 0 Then
                    maxx = sold
                Else
                    maxx = sold + onhand
                End If
                If maxx < 0 Then maxx = 0
                sql = "UPDATE Item_Inv SET Max_OH = " & maxx & " WHERE Loc_Id = '" & location & "' AND Sku = '" &
                    sku & "' AND eDate = '" & thisDate & "'"
                cmd = New SqlCommand(sql, con2)
                cmd.ExecuteNonQuery()
            Next
            con2.Close()

            Dim p As String = "Set Max Onhand"
            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", p, m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Set Max_OH")
        End Try
    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        stopWatch.Stop()
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "XMLUpdate"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & Date.Now & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Update_Store_Table()
        con.Open()
        con2.Open()
        sql = "SELECT DISTINCT Str_ID FROM Item_Sales"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Str_ID")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                sql = "IF NOT EXISTS (SELECT * FROM Stores WHERE ID = '" & CStr(oTest) & "') " & _
                    "INSERT INTO Stores(ID, Description, Status, Last_Change_Date, Last_Change_User) " & _
                    "SELECT '" & CStr(oTest) & "','WHOLESALE','Active','" & Date.Today & "','RC'"
                cmd = New SqlCommand(sql, con2)
                cmd.ExecuteNonQuery()
            End If
        End While
        con.Close()
        con2.Close()
        Dim m As String = ""
        Call Update_Process_Log("1", "Update Stores for Wholesale", m, "")
    End Sub

    Private Function FileSize(ByVal fileName As String)
        Dim length As Integer = 0
        If System.IO.File.Exists(fileName) Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(fileName)
            length = infoReader.Length
            If length = 0 Then
                Dim el As New XMLUpdate.ErrorLogger
                el.WriteToErrorLog(fileName, "HAS NO RECORDS", "Check File Exists")
            End If
        Else
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(fileName & " not found", "", "Check File Size")
        End If
        Return length
    End Function

    Private Function FixHyphen(val As String) As String
        Dim newval As String = Replace(val, "''", "'")
        Return newval
    End Function

    Public Function cleanData(ByVal input As String) As String
        Dim r As String = Regex.Replace(input, "[^A-Za-z0-9\-/]", "")
        Return r.Replace(input, [String].Empty)
    End Function
End Module

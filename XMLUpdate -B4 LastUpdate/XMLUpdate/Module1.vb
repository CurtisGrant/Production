Imports System.Data.SqlClient
Imports System.Xml
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module Module1
    Private eDateTBL As DataTable
    Private con, con2, con3 As SqlConnection
    Public rcCon As SqlConnection
    Private cmd, cmd2, cmd3 As SqlCommand
    Private rdr, rdr2 As SqlDataReader
    Public client, sql, rcErrorPath, xmlPath, conString, conString2, constr, rcConString As String
    Private oTest As Object
    Private stopWatch As Stopwatch
    Private totlQty, totlCost, totlRetail, totlMarkdown, diff As Decimal
    Private records As Integer
    Private msg, thisprocess, xmlProcess As String

    Sub Main()
        ''Try
        Dim fileSZ As Integer = 0
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim rcServer, rcExePath, RCUserId, rcPassword As String
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
            If fld = "USERID" Then RCUserId = valu
            If fld = "PD" Then rcPassword = valu
        End While
        ''rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;User Id=" & RCUserId & ";Password=" & rcPassword & ""
        rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True"
        rcCon = New SqlConnection(rcConString)
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
        Dim thisDate, thisEdate, thisSdate, lastWeekseDate As Date
        Dim DayOfWeek As Integer
        thisDate = CDate(Date.Today)
        DayOfWeek = thisDate.DayOfWeek

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
        xmlProcess = txtArray(10)




        ''eDateTBL.Rows.Add(thisEdate)
        ''lastWeekseDate = DateAdd(DateInterval.Day, -7, thisEdate)



        ''client = "TCM"
        ''''''server = "alrm6hn0ql.database.windows.net,1433"
        ''server = "LP-CURTIS"
        ''dbase = "TCM"
        ''xmlPath = "c:\RetailClarity\XMLs\TCM"
        ''''''xmlPath = "\\LPS2-2\RetailClarity\TCM\XMLs\Build"
        ''thisSdate = "10/29/2017"
        ''thisEdate = "11/4/2017"
        ''lastWeekseDate = DateAdd(DateInterval.Day, -7, thisEdate)
        ''sqlUserId = "sa"
        ''sqlPassword = "PGadm01!"
        ''rcErrorPath = "c:\RetailClarity\RCSystem\SYSFAIL"
        ''rcExePath = "c:\RetailClarity\EXEs"
        ''cust = "N"
        ''eDateTBL.Rows.Add("11/4/2017")
        ''xmlProcess = "ALL"




        conString = "server=" & server & ";Initial Catalog=" & dbase & ";User ID=" & sqlUserId & ";Password=" & sqlPassword & ""
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        constr = conString










        ''       xmlProcess = "SALES"









        stopWatch = New Stopwatch

        If xmlProcess = "ALL" Then
            Dim rCon As New SqlConnection(rcConString)
            rCon.Open()
            sql = "UPDATE Client_Master SET Last_XML_Update = NULL WHERE Client_ID = '" & client & "'"
            cmd = New SqlCommand(sql, rCon)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            rCon.Close()
            Console.WriteLine("Processing Calendar")
            thePath = xmlPath & "\Calendar.xml"
            If FileSize(thePath) > 0 Then Call Process_Calendar(thePath, con, con2, constr)
        End If

        If xmlProcess = "ITEMS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Item records")
            thePath = xmlPath & "\Items.xml"
            If FileSize(thePath) > 0 Then Call Process_Items(thePath, con, con2, constr)
            If xmlProcess = "ITEMS" Then Exit Sub
        End If

        If xmlProcess = "BARCODES" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing Barcodes")
            thePath = xmlPath & "\Barcodes.xml"
            If FileSize(thePath) > 0 Then Call Process_Barcodes(thePath, con, con2, constr)
            If xmlProcess = "BARCODES" Then Exit Sub
        End If

        If xmlProcess = "PREQHEADER" Or xmlProcess = "ALL" Then
            Console.WriteLine(" Purchase Request Header")
            thePath = xmlPath & "\Purchase_Request_Header.xml"
            If FileSize(thePath) > 0 Then Call Process_PREQ_Header(thePath, con, con2)
            If xmlProcess = "PREQHEADER" Then Exit Sub
        End If

        If xmlProcess = "PREQDETAIL" Or xmlProcess = "ALL" Then
            Console.WriteLine(" Purchase Request Detail")
            thePath = xmlPath & "\Purchase_Request_Detail.xml"
            If FileSize(thePath) > 0 Then Call Process_PREQ_Detail(thePath, con, con2)
            If xmlProcess = "PREQDETAIL" Then Exit Sub
        End If

        If xmlProcess = "INVENTORY" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Inventory records")
            thePath = xmlPath & "\Inventory.xml"
            If FileSize(thePath) > 0 Then Call Process_Inventory(thePath, con, con2)
            If xmlProcess = "INVENTORY" Then Exit Sub
        End If

        If xmlProcess = "POHEADER" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Purchase Order Header")
            thePath = xmlPath & "\POHeader.xml"
            If FileSize(thePath) > 0 Then Call Process_POHeader(thePath, con, con2)
            If xmlProcess = "POHEADER" Then Exit Sub
        End If

        If xmlProcess = "PODETAIL" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Purchase Order Detail")
            thePath = xmlPath & "\PODetail.xml"
            If FileSize(thePath) > 0 Then Call Process_PODetail(thePath, con, con2)
            If xmlProcess = "PODETAIL" Then Exit Sub
        End If

        If xmlProcess = "ADJUSTMENTS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing Adjustments")
            thePath = xmlPath & "\Adjustment.xml"
            If FileSize(thePath) > 0 Then Call Process_Data(thePath, "ADJ", con, con2)
            If xmlProcess = "ADJUSTMENTS" Then Exit Sub
        End If

        If xmlProcess = "PHYSICAL" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Physical records")
            thePath = xmlPath & "\Physical.xml"
            If FileSize(thePath) > 0 Then Call Process_Data(thePath, "ADJ", con, con2)
            If xmlProcess = "PHYSICAL" Then Exit Sub
        End If

        If xmlProcess = "RECEIPTS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Receiving records")
            thePath = xmlPath & "\Receipt.xml"
            If FileSize(thePath) > 0 Then Call Process_Data(thePath, "RECVD", con, con2)
            If xmlProcess = "RECEIPTS" Then Exit Sub
        End If

        If xmlProcess = "RETURNS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Process " & client & " Return records")
            thePath = xmlPath & "\Return.xml"
            If FileSize(thePath) > 0 Then Call Process_Data(thePath, "RTV", con, con2)
            If xmlProcess = "RETURNS" Then Exit Sub
        End If

        If xmlProcess = "SALES" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing Sales for " & client)
            Dim files As String() = Directory.GetFiles(xmlPath, "Sales*.xml")
            If files.Length > 0 Then
                For Each File As String In files
                    Call Process_Data(File, "Sold", con, con2)
                Next
            End If
            Call Merge_Tickets()
            If xmlProcess = "SALES" Then Exit Sub
        End If

        ''     Below code commented because Order data is already in Ticket History

        ''If xmlProcess = "ORDERS" Or xmlProcess = "ALL" Then
        ''    Console.WriteLine("Processing Orders for " & client)
        ''    thePath = xmlPath & "\Orders.xml"
        ''    If FileSize(thePath) > 0 Then Call Process_Data(thePath, "Order", con, con2)
        ''    If xmlProcess = "ORDERS" Then Exit Sub
        ''End If

        If xmlProcess = "TRANSFERS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Transfer records")
            thePath = xmlPath & "\Transfer.xml"
            If FileSize(thePath) > 0 Then Call Process_Data(thePath, "XFER", con, con2)
            If xmlProcess = "TRANSFERS" Then Exit Sub
        End If

        If xmlProcess = "ALL" Then
            Call Set_Max_OH(thisEdate)
            ''  Call Update_BeginOH
            Call Update_Store_Table()
        End If

        If xmlProcess = "BUYERS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Buyer records")
            thePath = xmlPath & "\Buyers.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Buyers", con, con2)
            If xmlProcess = "BUYERS" Then Exit Sub
        End If

        If xmlProcess = "CLASSES" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Class records")
            thePath = xmlPath & "\Classes.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Classes", con, con2)
            If xmlProcess = "CLASSES" Then Exit Sub
        End If

        If xmlProcess = "COUPONS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Coupon_Code records")
            thePath = xmlPath & "\Coupon_Codes.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Coupon_Codes", con, con2)
            Console.WriteLine("Process " & client & " Coupons")
            Call Process_Coupons()
            If xmlProcess = "COUPONS" Then Exit Sub
        End If

        If xmlProcess = "DEPARTMENTS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Department records")
            thePath = xmlPath & "\Departments.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Departments", con, con2)
            If xmlProcess = "DEPARTMENTS" Then Exit Sub
        End If

        If xmlProcess = "SEASONS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Seasons records") '
            thePath = xmlPath & "\Seasons.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Seasons", con, con2)
            If xmlProcess = "SEASONS" Then Exit Sub
        End If

        If xmlProcess = "STORES" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Store records")
            thePath = xmlPath & "\Stores.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Stores", con, con2)
            If xmlProcess = "STORES" Then Exit Sub
        End If

        If xmlProcess = "LOCATIONS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Shipping Locations")
            thePath = xmlPath & "\Locations.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Locations", con, con2)
            If xmlProcess = "LOCATIONS" Then Exit Sub
        End If

        If xmlProcess = "PLINES" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Product Line records")
            thePath = xmlPath & "\ProductLines.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "ProductLines", con, con2)
            If xmlProcess = "PLINES" Then Exit Sub
        End If

        If xmlProcess = "VENDORS" Or xmlProcess = "ALL" Then
            Console.WriteLine("Processing " & client & " Vendor records")
            thePath = xmlPath & "\Vendors.xml"
            If FileSize(thePath) > 0 Then Call Process_Other_Data(thePath, "Vendors", con, con2)
            If xmlProcess = "VENDORS" Then Exit Sub
        End If

        If xmlProcess <> "ALL" Then Exit Sub

100:    Try                 '    Create records for the next week when processing the last day of the week
            Dim nextSdate, nextEdate As Date
            If DayOfWeek = 0 Then                                              ' create new inventory records on Sunday
                Console.WriteLine("Adding Item_Inv records for next week.")
                ''                                        thisSdate and thisEdate were set when this thing first started
                con2.Open()
                thisprocess = "XMLUpdate - Set SysOH"
                Console.WriteLine("Setting Sys_OH = End_OH")
                sql = "UPDATE Item_Inv SET Sys_OH = End_OH WHERE eDate = '" & lastWeekseDate & "'"
                cmd = New SqlCommand(sql, con2)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()

                nextSdate = thisSdate
                nextEdate = thisEdate
                thisprocess = "XMLUpdate - Create Next Week's Inv Records"
                sql = "DECLARE @sDate date = '" & nextSdate & "' " &
                    "DECLARE @eDate date = '" & nextEdate & "' " &
                    "DECLARE @lasteDate date = '" & lastWeekseDate & "' " &
                    "IF OBJECT_ID('tempdb.dbo.#t1','U') IS NOT NULL DROP TABLE dbo.#t1; " &
                    "SELECT Loc_Id, Sku, @sDate sDate, @eDate eDate, Cost, Retail, c.YrWk, Begin_OH, End_OH, " &
                    "Item, DIM1, DIM2, DIM3 INTO #t1 FROM Item_Inv i " &
                    "JOIN Calendar c ON c.eDate = @eDate AND c.week_Id > 0 " &
                    "WHERE i.eDate = @lasteDate AND End_OH <> 0 " &
                    "MERGE Item_Inv AS t USING #t1 AS s ON s.Loc_Id=t.Loc_Id AND s.Sku=t.Sku AND s.eDate=t.eDate " &
                    "WHEN NOT MATCHED BY Target THEN " &
                    "INSERT(Loc_Id, sku, sDate, eDate, cost, retail, YrWk, Begin_OH, End_OH, item, dim1, dim2, dim3) " &
                    "VALUES(s.Loc_Id, s.Sku, s.sDate, s.eDate, s.Cost, s.Retail, s.YrWk, s.End_OH, s.End_OH, s.Item, s.DIM1, s.DIM2, s.DIM3) " &
                    "WHEN MATCHED THEN UPDATE SET t.Begin_OH = s.End_OH;"
                cmd = New SqlCommand(sql, con2)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()
                con2.Close()
                Dim m As String = "Created " & cnt & " Records"
                Call Update_Process_Log("1", "Create Inventory Records for Next Week", m, "")
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)

        End Try

        cnt = eDateTBL.Rows.Count
        If cnt = 0 Then
            con3.Open()
            Dim title As String = "XMLUpdate"
            Dim message As String = "Nothing to update in Client_Weekly_Summary"
            sql = "INSERT INTO Message (mDate, Title, Message) " &
                "SELECT '" & Date.Now & "','" & title & "','" & message & "'"
            cmd = New SqlCommand(sql, con3)
            cmd.CommandTimeout = 480
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
        cmd.CommandTimeout = 480
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
            thisprocess = "XMLUpdate - Process Calendar - Read XML File"
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                ctbl = ds.Tables(0)
                For Each row In ctbl.Rows
                    cnt += 1
                    year = CInt(row("Year_Id"))
                    'Console.WriteLine(year)
                    period = CInt(row("Prd_Id"))
                    ' Console.WriteLine(period)
                    week = CInt(row("Week_Id"))
                    ' Console.WriteLine(week)
                    sdate = CDate(row("sDate"))
                    ' Console.WriteLine(sdate)
                    edate = CDate(row("eDate"))
                    ' Console.WriteLine(edate)
                    yrprd = CInt(row("YrPrd"))
                    'Console.WriteLine(yrprd)
                    yrwks = CInt(row("YrWks"))
                    ' Console.WriteLine(yrwks)
                    prdwk = CInt(period & week)
                    ' Console.WriteLine(prdwk)
                    If (period = 0 And week = 0) Or (period > 0 And week > 0) Then
                        thisprocess = "XMLUpdate - Process Calendar - Insert into Calendar"
                        sql = "IF NOT EXISTS (SELECT * FROM Calendar WHERE sDate = '" & sdate & "' AND Prd_Id = " & period & " AND Week_Id = " & week & ") " &
                            "INSERT INTO Calendar (Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, YrWks, PrdWk) " &
                            "SELECT " & year & "," & period & "," & week & ",'" & sdate & "','" & edate & "'," & yrprd & "," & yrwks & "," & prdwk
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 480
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Process_Items(ByVal thePath, ByVal con, ByVal con2, ByVal constr)
        Try
            stopWatch.Start()
            Console.WriteLine("Cleaning XML import file.")
            Dim cnt As Integer = 0
            Dim oTest As Object
            Dim isNumber As Boolean
            Dim item, descr, vid, vendor, vitem, note, buyer, dept, clss, custom1, custom2, custom3, custom4, custom5, uom,
                sku, dim1, dim2, dim3, dim1_descr, dim2_descr, dim3_descr, status, sql, type As String
            Dim buyunit, sellunit As Int16
            Dim cost, retail As Decimal
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
            dt2.Columns.Add("Custom1")
            dt2.Columns.Add("Custom2")
            dt2.Columns.Add("Custom3")
            dt2.Columns.Add("Custom4")
            dt2.Columns.Add("Custom5")
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
            Dim xmlFile As XmlReader
            Dim rw As DataRow
            thisprocess = "XMLUpdate - Process Items - Read XML File"
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
                    If sku <> "TOTALS" Then
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
                        descr = Trim(Microsoft.VisualBasic.Left(row("DESCRIPTION"), 40))
                        vid = Trim(Microsoft.VisualBasic.Left(row("VENDOR_ID"), 20))
                        vendor = Trim(Microsoft.VisualBasic.Left(row("VENDOR"), 40))
                        vitem = Trim(Microsoft.VisualBasic.Left(row("VEND_ITEM_NO"), 20))
                        buyer = Trim(Microsoft.VisualBasic.Left(row("BUYER"), 10))
                        If buyer = "" Then buyer = "UNKNOWN"
                        dept = Trim(Microsoft.VisualBasic.Left(row("DEPT"), 10))
                        If dept = "" Then dept = "UNKNOWN"
                        clss = Trim(Microsoft.VisualBasic.Left(row("CLASS"), 10))
                        If clss = "" Then clss = "UNKNOWN"
                        custom1 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM_1"), 30))
                        custom2 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM_2"), 30))
                        custom3 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM_3"), 30))
                        custom4 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM_4"), 30))
                        custom5 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM_5"), 30))
                        note = Trim(Microsoft.VisualBasic.Left(row("NOTE"), 20))
                        uom = Trim(Microsoft.VisualBasic.Left(row("UOM"), 10))
                        status = Trim(Microsoft.VisualBasic.Left(row("STATUS"), 10))
                        oTest = Trim(Microsoft.VisualBasic.Left(row("TYPE"), 10))
                        If Not IsDBNull(oTest) Then type = CStr(oTest) Else type = "N"
                        oTest = row("CURR_COST")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero) Else cost = 0
                        oTest = row("CURR_RTL")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then retail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero) Else retail = 0
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
                        rw("Custom1") = custom1
                        rw("Custom2") = custom2
                        rw("Custom3") = custom3
                        rw("Custom4") = custom4
                        rw("Custom5") = custom5
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
                        dt2.Rows.Add(rw)
                    End If
                Next
            End If
            Console.WriteLine("DELETE FROM ITEM_MASTER")
            thisprocess = "XMLUpdate - Process Items - Delete from Item Master"
            con.Open()
            sql = "DELETE FROM Item_Master"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()

            Console.WriteLine("Bulk Insert into Item_Master")
            thisprocess = "XMLUpdate - Bulk Insert into Item Master"
            Dim connection As SqlConnection = New SqlConnection(constr)
            Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
            connection.Open()
            bulkCopy.DestinationTableName = "dbo.Item_Master"
            bulkCopy.BulkCopyTimeout = 960
            bulkCopy.WriteToServer(dt2)
            connection.Close()

            Dim m As String = "Created " & cnt & " Records"
            Call Update_Process_Log("1", "Recreate Item_Master", m, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Process_Barcodes(ByVal thePath, con, con2, constr)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim dte As Date
            Dim barcode, type, sku, item, dim1, dim2, dim3 As String
            Console.WriteLine("Loading Barcode XML file")

            con.open()                                          ' Clear out Item_Barcodes
            sql = "DELETE FROM Item_Barcodes"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.close()

            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            dt.Columns.Add("Sku")
            dt.Columns.Add("Item")
            dt.Columns.Add("DIM1")
            dt.Columns.Add("DIM2")
            dt.Columns.Add("DIM3")
            dt.Columns.Add("Barcode")
            dt.Columns.Add("Type")
            dt.Columns.Add("Date")
            Dim xmlFile As XmlReader
            Dim row As DataRow
            thisprocess = "XMLUpdate - Process Barcodes - Bulk Copy XML File"
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                ''con.open()
                dt = ds.Tables(0)
                ''For Each row In dt.Rows
                ''    cnt += 1
                ''    If cnt Mod 1000 = 0 Then
                ''        Console.WriteLine("Processed " & cnt & " records.")
                ''    End If
                ''    sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                ''    If InStr(sku, "~") > 0 Then
                ''        Dim parts() As String = sku.Split("~"c)
                ''        item = parts(0)
                ''        dim1 = parts(1)
                ''        dim2 = parts(2)
                ''        dim3 = parts(3)
                ''    Else
                ''        item = sku
                ''        dim1 = Nothing
                ''        dim2 = Nothing
                ''        dim3 = Nothing
                ''    End If
                ''    type = row("BARCOD_ID")
                ''    barcode = row("BARCOD")
                ''    dte = row("EXTRACT_DATE")
                ''    sql = "IF NOT EXISTS (SELECT Barcode FROM Item_Barcodes WHERE Sku = '" & sku & "' AND Barcode = '" & barcode & "' AND Type = '" & type & "') " & _
                ''        "INSERT INTO Item_Barcodes(Sku, Item, Dim1, Dim2, Dim3, Barcode, Type, Extract_Date) " & _
                ''        "SELECT '" & sku & "','" & item & "','" & dim1 & "','" & dim2 & "','" & dim3 & "','" & barcode & "','" & type & "','" & dte & "'"
                ''    cmd = New SqlCommand(sql, con)
                ''    cmd.ExecuteNonQuery()
                ''Next
                ''con.close()
                cnt = dt.Rows.Count
                Dim connection As SqlConnection = New SqlConnection(constr)
                Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
                connection.Open()
                bulkCopy.DestinationTableName = "dbo.Item_Barcodes"
                bulkCopy.BulkCopyTimeout = 480
                bulkCopy.WriteToServer(dt)
                connection.Close()
            End If
            Dim m As String = "Created " & cnt & " Records"
            Call Update_Process_Log("1", "Create Barcodes", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Process_Inventory(ByVal thePath, con, con2)
        Try
            stopWatch.Start()
            con2.open()
            thisprocess = "XMLUpdate - Process Inventory - Create temp Table"
            Dim sql As String
            sql = "IF OBJECT_ID('dbo._t1','U') IS NOT NULL DROP TABLE dbo._t1; "
            cmd = New SqlCommand(sql, con2)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE _t1(" &
                   "[Loc_Id] [varchar](20) NOT NULL," &
                   "[Sku] [varchar](90) NOT NULL," &
                   "[sDate] [date] NOT NULL," &
                   "[eDate] [date] NOT NULL," &
                   "[Cost] [decimal](18,4) NULL," &
                   "[Retail] [decimal](18,4) NULL," &
                   "[YrWk] [int] NULL," &
                   "[Begin_OH] [decimal](18, 4) NULL," &
                   "[End_OH] [decimal](18, 4) NULL," &
                   "[Sys_OH] [decimal](18, 4) NULL," &
                   "[Max_OH] [decimal](18, 4) NULL," &
                   "[Item] [varchar](30) NOT NULL," &
                   "[DIM1] [varchar](30) NULL," &
                   "[DIM2] [varchar](30) NULL," &
                   "[DIM3] [varchar](30) NULL) ON [PRIMARY]"
            cmd = New SqlCommand(sql, con2)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()

            thisprocess = "XMLUpdate - Process Inventory - Get sDate & eDate"
            Dim cnt As Int64 = 0
            Dim oTest As Object
            Dim onhand As Decimal
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

            thisprocess = "XMLUpdate - Process Inventory - Read XML File"
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
                        If IsDBNull(oTest) Then onhand = 0 Else onhand = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("COST")
                        If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("RETAIL")
                        If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
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
                        thisprocess = "XMLUpdate - Process Inventory - Insert into temp Table"
                        sql = "INSERT INTO _t1 (Loc_Id, Sku, sDate, eDate, End_OH, " &
                                 "Cost, Retail, YrWk, Item, DIM1, DIM2, DIM3) " &
                                 "SELECT '" & loc & "','" & sku & "','" & thisSdate & "','" & thisEdate &
                                 "'," & onhand & "," & cost & "," & retail & ", YrWk ,'" & item & "','" & dim1 & "','" & dim2 & "','" &
                                 dim3 & "' FROM Calendar c JOIN Item_Master m ON m.Sku = '" & sku & "' WHERE eDate = '" & thisEdate & "' " &
                                 "AND m.[Type] = 'I' AND Prd_Id > 0 AND Week_Id > 0"
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 960
                        cmd.ExecuteNonQuery()

                    Else
                        oTest = row("OnHand")
                        If IsNumeric(oTest) Then onhand = Decimal.Round(CDec(row("OnHand")), 4, MidpointRounding.AwayFromZero)
                        oTest = row("COST")
                        If IsNumeric(oTest) Then cost = Decimal.Round(CDec(row("COST")), 4, MidpointRounding.AwayFromZero)
                        oTest = row("RETAIL")
                        If IsNumeric(oTest) Then retail = Decimal.Round(CDec(row("RETAIL")), 4, MidpointRounding.AwayFromZero)
                        If onhand <> totlQty Then
                            msg = "Expected " & Format(onhand, "###,###,##0.0000") & " Received " & Format(tqty, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Quantity Mismatch " & conString, "Process Inventory")
                        End If
                        If cost <> totlCost Then
                            msg = "Expected " & Format(cost, "$###,###,##0.0000") & " Received " & Format(tcost, "$###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Cost Mismatch", "Process Inventory")
                        End If
                        If retail <> totlRetail Then
                            msg = "Expected " & Format(retail, "$###,###,##0.0000") & " Received " & Format(tretail, "$###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Retail Mismatch", "Process Inventory")
                        End If
                    End If
                Next
            End If
            con2.close()

            Console.WriteLine("Set End_OH, Max_OH and Committed to 0")
            con.open()
            thisprocess = "XMLUpdate - Process Inventory - Set Qty Fields to 0"
            sql = "UPDATE Item_Inv SET End_OH = 0, Max_OH = 0 WHERE eDate = '" & thisEdate & "'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            thisprocess = "XMLUpdate - Process Inventory - Merge"
            Console.WriteLine("Merge _t1 and Item_Inv")
            sql = "MERGE Item_Inv AS target USING _t1 AS source ON (source.Loc_Id = target.Loc_Id AND source.Sku = target.Sku " &
                      "AND source.eDate = target.eDate) " &
                      "WHEN NOT MATCHED BY TARGET THEN " &
                      "INSERT (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, Begin_OH, End_OH, Max_OH, Item, DIM1, DIM2, DIM3) " &
                      "VALUES (source.Loc_Id, source.Sku, source.sDate, source.eDate, source.Cost, source.Retail, " &
                          "source.YrWk, source.Begin_OH, source.End_OH, source.Max_OH, " &
                          "source.Item, source.DIM1, source.DIM2, source.DIM3) " &
                      "WHEN MATCHED THEN " &
                      "UPDATE SET target.Begin_OH = source.Begin_OH, target.End_OH = source.End_OH, target.Cost = source.Cost, " &
                      "target.Retail = source.retail, target.Max_OH = source.Max_OH, target.Yrwk = source.YrWk;"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            cmd.ExecuteNonQuery()

            thisprocess = "XMLUpdate - Process Inventory - Drop temp Table"
            Console.WriteLine("Drop _t1")
            sql = "IF OBJECT_ID('dbo._t1','U') IS NOT NULL DROP TABLE dbo. _t1; "
            cmd = New SqlCommand(sql, con)
            con.Close()

            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", "Update Inventory", m, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Update_BeginOH()
        Try
            con.Open()
            thisprocess = "XMLUpdate - Process Inventory - Update Begin_OH"
            Console.WriteLine("Update Begin_OH")
            sql = "UPDATE i SET i.Begin_OH = ISNULL(End_OH,0) " &
                    "-ISNULL(ADJ, 0) " &
                    "- ISNULL(RECVD,0) " &
                    "- ISNULL(XFER,0) " &
                    "+ ISNULL(RTV,0) " &
                    "+ ISNULL((SELECT ISNULL(SUM(Sold),0) FROM Item_Sales s " &
                    "WHERE s.Loc_Id = i.Loc_Id AND s.Sku = i.Sku AND s.eDate = i.eDate),0) " &
                    "FROM Item_Inv i WHERE i.Begin_OH IS NULL"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub
    Private Sub Process_PREQ_Header(ByVal thePath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim preq, batch, vend_no, vend, loc, buyer, alloc, mrg, terms, custom, custom2, custom3, custom4, custom5 As String
            Dim ord, del, can, extract As Date
            Dim amt, totlCost As Decimal
            totlCost = 0
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim xmlFile As XmlReader
            thisprocess = "XMLUpdate - Process PREQ Header - Read XML File"
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                con2.open()
                thisprocess = "XMLUpdate - Process PREQ Header - Delete from Purchase_Request_Header"
                sql = "DELETE FROM Purchase_Request_Header"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
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
                        oTest = row("TERMS")
                        If IsDBNull(oTest) Then oTest = Nothing
                        terms = oTest
                        oTest = row("CUSTOM_1")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom = CStr(oTest) Else custom = Nothing
                        oTest = row("CUSTOM_2")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom2 = CStr(oTest) Else custom2 = Nothing
                        oTest = row("CUSTOM_3")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom3 = CStr(oTest) Else custom3 = Nothing
                        oTest = row("CUSTOM_4")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom4 = CStr(oTest) Else custom4 = Nothing
                        oTest = row("CUSTOM_5")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom5 = CStr(oTest) Else custom5 = Nothing
                        thisprocess = "XMLUpdate - Process PREQ Header - Insert into Purchase_Request_Header"
                        sql = "INSERT INTO Purchase_Request_Header(PREQ_NO, Batch_Id, Loc_Id, " &
                            "Vendor_Id, Buyer, Order_Date, Deliver_Date, Cancel_Date, Order_Total, " &
                            "Allocated, Merged, Last_Update, Terms, Custom1, Custom2, Custom3, Custom4, Custom5) " &
                            "SELECT '" & preq & "','" & batch & "','" & loc & "','" & vend_no & "','" &
                            buyer & "','" & ord & "','" & del & "','" & can & "'," & totlCost & ",'" &
                            alloc & "','" & mrg & "','" & Date.Today & "','" & terms & "','" & custom & "','" &
                            custom2 & "','" & custom3 & "','" & custom4 & "','" & custom5 & "'"
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 480
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("ORDER_TOTAL")
                        If IsDBNull(oTest) Then amt = 0 Else amt = CDec(oTest)
                        If amt <> totlCost Then
                            msg = "Expected " & Format(amt, "###,###,##0.0000") & " Received " & Format(totlCost, "###,###,##0.0000")
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
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
            thisprocess = "XMLUpdate - Process PREQ Detail - Read XML File"
            Dim xmlfile As XmlReader
            xmlfile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlfile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                con2.Open()
                thisprocess = "XMLUpdate - Process PREQ Detail - Delete from Purchase_Request_Detail"
                sql = "DELETE FROM Purchase_Request_Detail"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
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
                        If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("QTY")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        qty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        totlQty += qty
                        totlCost += (qty * cost)
                        thisprocess = "XMLUpdate - Process PREQ Detail - Insert into Purchase_Request_Detail"
                        sql = "INSERT INTO Purchase_Request_Detail(PREQ_NO, Seq_No, Sku, Description, " &
                            "Stock_Units, Numer, Denom, Cost, Item, DIM1, DIM2, DIM3, Qty) " &
                            "SELECT '" & preq & "'," & seq & ",'" & sku & "','" & desc & "','" &
                            uom & "'," & numer & "," & denom & "," & cost & ",'" & item & "','" &
                            dim1 & "','" & dim2 & "','" & dim3 & "'," & qty & " "
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 480
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("COST")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("QTY")
                        If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                        qty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        If cost <> totlCost Then
                            msg = "Expected " & Format(cost, "###,###,##0.0000") & " Received " & Format(totlCost, "###,###,##0.0000")
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Record Count Mismatch " & conString, "Process PREQ Detail")
                        End If
                        If qty <> totlQty Then
                            msg = "Expected " & Format(qty, "###,###,##0.0000") & " Received " & Format(totlQty, "###,###,##0.0000")
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Process_POHeader(ByVal thePath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Integer = 0
            Dim recvd, lines, open_lines As Integer
            Dim amt, recvd_cost, ord_qty, open_amt As Decimal
            Dim store, po, vendor, buyer, stat, terms, custom, custom2, custom3, custom4, custom5 As String
            Dim ordDate, dueDate, canDate As Date
            Dim datetime As DateTime
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            thisprocess = "XMLUpdate - Process PO Header - Read XML File"
            Dim xmlFile As XmlReader
            Dim row As DataRow
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.open()
                thisprocess = "XMLUpdate - Process PO Detail - Delete from PO_Header"
                sql = "DELETE FROM PO_Header"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()
                Console.WriteLine("Deleted everything in PO_Header")
                thisprocess = "XMLUpdate - Process PO Header - Read XML File"
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("PO")
                    If oTest <> "TOTALS" Then
                        cnt += 1
                        If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                        oTest = row("LOCATION")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = CStr(oTest) Else store = "1"
                        oTest = row("PO")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then po = CStr(oTest) Else po = "UNKNOWN"
                        oTest = row("OrderDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then ordDate = CDate(oTest) Else ordDate = "1900-01-01"
                        oTest = row("DueDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then dueDate = CDate(oTest) Else dueDate = "1900-01-01"
                        oTest = row("CancelDate")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then canDate = CDate(oTest) Else canDate = "1900-01-01"
                        oTest = row("Vendor")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vendor = CStr(oTest) Else vendor = "UNKNOWN"
                        oTest = row("BUYER")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = CStr(oTest) Else buyer = "UNKNOWN"
                        oTest = row("STATUS")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then stat = CStr(oTest) Else stat = "X"
                        oTest = row("AMT")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            amt = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : amt = 0
                        End If
                        oTest = row("RECVD_COST")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            recvd_cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : recvd_cost = 0
                        End If
                        oTest = row("LINES")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lines = CInt(oTest) Else lines = 0
                        oTest = row("ORD_QTY")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            ord_qty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : ord_qty = 0
                        End If
                        oTest = row("OPEN_LINES")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then open_lines = CInt(oTest) Else open_lines = 0
                        oTest = row("OPEN_AMT")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            open_amt = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : open_amt = 0
                        End If
                        oTest = row("EXTRACT_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then datetime = CDate(oTest) Else datetime = Nothing
                        oTest = row("TERMS")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then terms = CStr(oTest) Else terms = Nothing
                        oTest = row("CUSTOM_1")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom = CStr(oTest) Else custom = Nothing
                        oTest = row("CUSTOM_2")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom2 = CStr(oTest) Else custom2 = Nothing
                        oTest = row("CUSTOM_3")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom3 = CStr(oTest) Else custom3 = Nothing
                        oTest = row("CUSTOM_4")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom4 = CStr(oTest) Else custom4 = Nothing
                        oTest = row("CUSTOM_5")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom5 = CStr(oTest) Else custom5 = Nothing
                        thisprocess = "XMLUpdate - Process PO Header - Insert into PO_Header"
                        sql = "IF NOT EXISTS (SELECT PO_NO FROM PO_Header WHERE PO_NO = '" & po & "') " &
                            "INSERT INTO PO_Header (Loc_Id, PO_NO, Order_Date, Due_Date, Cancel_Date, Vendor_Id, Buyer, Status, " &
                            "Amount, Recvd_Cost, Lines, Ord_Qty, Open_Lines, Open_Amt, Terms, Custom1, Custom2, Custom3, Custom4, Custom5) " &
                            "SELECT '" & store & "','" & po & "','" & ordDate & "','" & dueDate & "','" & canDate & "','" &
                            vendor & "','" & buyer & "','" & stat & "', " & amt & ", " & recvd_cost & ", " & lines & ", " &
                            ord_qty & ", " & open_lines & ", " & open_amt & ",'" & terms & "','" & custom & "','" & custom2 & "','" &
                            custom3 & "','" & custom4 & "','" & custom5 & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 480
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub


    Private Sub Process_PODetail(ByVal thepath, ByVal con, ByVal con2)
        Try
            stopWatch.Start()
            Dim cnt As Int32 = 0
            Dim oTest As Object
            Dim po, item, loc, sql, lastDate As String
            Dim seq As Int32
            Dim cost, retail, qty, tqty, trecvd, tcan, tcost, tretail As Decimal
            Dim oqty, rqty, eqty, cqty, ocost, oretail As Decimal
            Dim ordQty As Decimal = 0
            Dim recvdQty As Decimal = 0
            Dim expQty As Decimal = 0
            Dim extractDate As DateTime
            Dim cmd As SqlCommand
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim loDate As Date = "2012-01-01"
            Dim thisDate As Date = Date.Today
            Dim xmlFile As XmlReader
            thisprocess = "XMLUpdate - Process PO Detail - Read XML File"
            Dim row As DataRow
            xmlFile = Xml.XmlReader.Create(thepath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                thisprocess = "XMLUpdate - Process PO Detail - Delete from PO_Detail"
                sql = "DELETE FROM PO_Detail"
                con.open()
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
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
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            ''oqty = Decimal.Round(CDec(oTest), 4, MidpeointRounding.AwayFromZero)
                            oqty = CDec(oTest)
                        Else : oqty = 0
                        End If
                        ordQty += oqty
                        oTest = row("QTY_RECVD")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            rqty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : rqty = 0
                        End If
                        recvdQty += rqty
                        oTest = row("EXP_QTY")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            eqty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : eqty = 0
                        End If
                        expQty += eqty
                        oTest = row("LAST_RECVD_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then lastDate = CDate(oTest) Else lastDate = Nothing
                        oTest = row("COST")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : cost = 0
                        End If
                        tcost += (oqty * cost)
                        oTest = row("RETAIL")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            retail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        Else : retail = 0
                        End If
                        tretail += (oqty * retail)
                        'vendor = row("VEND_NO")
                        'buyer = row("BUYER")
                        'stat = row("STATUS")
                        extractDate = row("EXTRACT_DATE")
                        If lastDate <> "" Then
                            thisprocess = "XMLUpdate - Process PO Detail - Insert into PO_Detail"
                            sql = "IF NOT EXISTS (SELECT PO_NO FROM PO_Detail WHERE PO_NO = '" & po & "' " &
                                "AND Sku = '" & item & "' AND Seq_No = " & seq & ") " &
                                "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, " &
                                "Cost, Retail, Last_Recvd_Date, Item, DIM1, DIM2, DIM3) " &
                                "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & oqty & "," & rqty & "," &
                                eqty & "," & cost & "," & retail & ",'" & lastDate & "',Item, DIM1, DIM2, DIM3 FROM Item_Master " &
                                "WHERE Sku = '" & item & "'"
                        Else
                            sql = " IF NOT EXISTS (SELECT PO_NO FROM PO_Detail WHERE PO_NO = '" & po & "' " &
                                "AND Sku = '" & item & "' AND Seq_No = " & seq & ") " &
                                "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, " &
                                "Cost, Retail, Last_Recvd_Date, Item, DIM1, DIM2, DIM3) " &
                                "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & oqty & "," & rqty & "," &
                                eqty & "," & cost & "," & retail & ", NULL, Item, DIM1, DIM2, DIM3 FROM Item_Master " &
                                "WHERE Sku = '" & item & "'"
                        End If
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 480
                        cmd.ExecuteNonQuery()
                    Else
                        oTest = row("ORD_QTY")
                        If IsNumeric(oTest) Then oqty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("QTY_RECVD")
                        If IsNumeric(oTest) Then rqty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("EXP_QTY")
                        If IsNumeric(oTest) Then cqty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("COST")
                        If IsNumeric(oTest) Then ocost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        oTest = row("RETAIL")
                        If IsNumeric(oTest) Then oretail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        If oqty <> ordQty Then
                            diff = Math.Abs((oqty - ordQty) / oqty)
                            If diff > 0.02 Then
                                msg = "Expected " & Format(oqty, "###,###,##0.0000") & " Received " & Format(ordQty, "###,###,##0.0000")
                                Dim el As New XMLUpdate.ErrorLogger
                                el.WriteToErrorLog(msg, "Order Quantity Mismatch " & conString, "Process PO Detail")
                            End If
                        End If
                        If ocost <> tcost Then
                            diff = Math.Abs((ocost - tcost) / ocost)
                            If diff > 0.01 Then
                                msg = "Expected " & Format(ocost, "$###,###,##0.0000") & " Received " & Format(cost, "$###,###,##0.0000")
                                Dim el As New XMLUpdate.ErrorLogger
                                el.WriteToErrorLog(msg, "Ordered @ Cost Mismatch " & conString, "Process PO Detail")
                            End If
                        End If
                        If oretail <> tretail Then
                            diff = Math.Abs((oretail - tretail) / oretail)
                            If diff > 0.02 Then
                                msg = "Expected " & Format(oretail, "$###,###,##0.0000") & " Received " & Format(tretail, "####,###,##0.0000")
                                Dim el As New XMLUpdate.ErrorLogger
                                el.WriteToErrorLog(msg, "Ordered @ Retail Mismatch " & conString, "Process PO Detail")
                            End If
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Process_Data(ByVal thePath, ByVal theField, con, con2)
        Try
            stopWatch = New Stopwatch
            stopWatch.Start()
            Dim cnt As Int32 = 0
            Dim oTest As Object
            Dim isPhysical As Boolean = False
            Dim id, sku, item, dim1, dim2, dim3, store, loc, dept, buyer, clss, sql, cust, tkt, reason, transtype,
                coupon, typ, nam, ttype, ordtype, whsl, station, drawer, sales_rep, userid, sale_type As String
            Dim custom, custom2, custom3, custom4, custom5 As String
            Dim seq As Int32
            Dim cost, retail, mkdn, tCost, tRetail, tMkdn, qty, tQty As Decimal
            Dim dte, eDate As Date
            Dim transDate, tktDate As DateTime
            Dim extractDate, datetimenow, lastUpdate As DateTime
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
                thisprocess = "XMLUpdate - Process Date - Delete from Unposted_Sales"
                sql = "DELETE FROM Unposted_Sales"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()
                con.close()
            End If

            con.open
            If InStr(thePath, "Physical") > 0 Then isPhysical = True
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            thisprocess = "XMLUpdate - Process Data - Read " & theField & " XML File"
            ds.ReadXml(xmlFile)            '  bulk insert XML into dataset
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                totlQty = 0
                totlCost = 0
                totlRetail = 0
                totlMarkdown = 0
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    oTest = row("SKU")
                    If Not IsDBNull(oTest) AndAlso oTest <> "TOTALS" Then
                        loc = "UNKNOWN"
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
                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then loc = CStr(oTest)
                        qty = 0
                        oTest = row("QTY")
                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then qty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        End If
                        cost = 0
                        oTest = row("COST")
                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        End If
                        retail = 0
                        oTest = row("RETAIL")
                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then retail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                        End If
                        totlQty += qty
                        totlCost += Decimal.Round((qty * cost), 4, MidpointRounding.AwayFromZero)
                        totlRetail += Decimal.Round((qty * retail), 4, MidpointRounding.AwayFromZero)
                        oTest = row("TRANS_DATE")
                        If oTest = "0" Then oTest = "1900-01-01 00:00:00"

                        If CDate(oTest) < loDate Then GoTo 10

                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then
                            transDate = DateTime.Parse(oTest)
                        Else : transDate = "1900-01-01 00:00:00"
                        End If
                        ordtype = ""
                        whsl = "N"
                        mkdn = 0
                        store = Nothing
                        dept = Nothing
                        buyer = Nothing
                        clss = Nothing
                        cust = Nothing
                        tkt = Nothing
                        reason = Nothing
                        coupon = Nothing
                        ordtype = Nothing
                        sales_rep = Nothing
                        userid = Nothing
                        station = Nothing
                        drawer = Nothing
                        custom = Nothing
                        custom2 = Nothing
                        custom3 = Nothing
                        custom4 = Nothing
                        custom5 = Nothing



                        ''  Uncomment below code when new extract is running on TCM



                        ''lastUpdate = "1900-01-01 00:00:00"
                        ''If theField <> "Sold" And theField <> "Order" Then
                        ''    oTest = row("LAST_UPDATE")
                        ''    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lastUpdate = CDate(oTest)
                        ''End If
                        If theField = "Sold" Or theField = "Order" Then
                            oTest = row("STR_ID")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then store = CStr(oTest)
                            oTest = row("DEPT")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then dept = CStr(oTest)
                            oTest = row("BUYER")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then buyer = CStr(oTest)
                            oTest = row("CLASS")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then clss = CStr(oTest)
                            oTest = row("CUST_NO")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then cust = CStr(oTest)
                            oTest = row("TKT_NO")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then tkt = oTest
                            oTest = row("TICKET_DATE")
                            If Not IsDBNull(oTest) Then tktDate = CDate(oTest) Else tktDate = "1900-01-01 00:00:00"
                            oTest = row("MARKDOWN")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then mkdn = CDec(oTest)
                            totlMarkdown += mkdn
                            oTest = row("MKDN_REASON")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then reason = CStr(oTest)
                            oTest = row("COUPON_CODE")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then coupon = CStr(oTest)
                            oTest = row("ORD_TYPE")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then ordtype = CStr(oTest)
                            oTest = row("USER_ID")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then userid = CStr(oTest)
                            oTest = row("SALES_REP")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then sales_rep = CStr(oTest)
                            oTest = row("STATION")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then station = CStr(oTest)
                            oTest = row("DRAWER")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then drawer = CStr(oTest)
                            If theField = "Sold" Then sale_type = "T" Else sale_type = "O"
                            If ordtype = "WHOLESALE" Then whsl = "Y"
                            oTest = row("CUSTOM_1")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then custom = CStr(oTest)
                            oTest = row("CUSTOM_2")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then custom2 = CStr(oTest)
                            oTest = row("CUSTOM_3")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then custom3 = CStr(oTest)
                            oTest = row("CUSTOM_4")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then custom4 = CStr(oTest)
                            oTest = row("CUSTOM_5")
                            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then custom5 = CStr(oTest)
                        End If
                        oTest = row("TRANS_DATE")
                        If oTest = "0" Then oTest = "1900-01-01 00:00:00"

                        If CDate(oTest) < loDate Then GoTo 10

                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) And oTest <> "" Then
                            transDate = DateTime.Parse(oTest)
                        Else : transDate = "1900-01-01 00:00:00"
                        End If

                        oTest = row("EXTRACT_DATE")
                        If oTest = "0" Then oTest = "1900-01-01 00:00:00"
                        If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) And oTest <> "" Then
                            extractDate = DateTime.Parse(oTest)
                        Else : extractDate = "1900-01-01 00:00:00"
                        End If

                        Dim itexists As Boolean = Check_Log(con2, id, seq, loc, sku, transDate, store, theField)

                        If Not itexists Then
                            tQty += qty
                            tCost += (qty * cost)
                            tRetail += (qty * retail)
                            tMkdn += mkdn
                            thisprocess = "XMLUpdate - Process Data - Get eDate"
                            If con2.state = ConnectionState.Closed Then con2.open()
                            sql = "SELECT eDate FROM Calendar WHERE '" & CDate(transDate) & "' BETWEEN sDate AND eDate AND Week_ID > 0"
                            cmd = New SqlCommand(sql, con2)
                            cmd.CommandTimeout = 120
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                If Not IsDBNull(rdr(0)) AndAlso Not IsNothing(rdr(0)) Then eDate = rdr(0) Else eDate = "2012-01-01"
                            End While
                            If con2.state = ConnectionState.Open Then con2.close()

                            row = eDateTBL.Rows.Find(eDate)              ' We need to send the eDate to Client_Weekly_Summary
                            If IsNothing(row) Then
                                eDateTBL.Rows.Add(eDate)                 ' Add date to eDateTBL if it isn't already there
                            End If

                            If theField = "Order" Then transtype = "Sold" Else transtype = theField

                            datetimenow = DateAndTime.Now
                            If qty <> 0 Then                            ' Don't do anything unless qty <> 0
                                Call Update_Tables(store, loc, eDate, sku, qty, retail, cost, mkdn, theField, dept, buyer, clss, "", whsl, con2)
                                If transtype = "Sold" Then
                                    sql = "INSERT INTO Daily_Transaction_Log (TRANS_ID, SEQUENCE_NO, STORE, SKU, LOCATION, QTY, COST, RETAIL, " &
                                        "MKDN, TRANS_DATE, [TYPE], POST_DATE, DEPT, BUYER, CLASS, EXTRACT_DATE, CUST_NO, TKT_NO, MKDN_REASON, " &
                                        "COUPON_CODE, ITEM, DIM1, DIM2, DIM3, ORD_TYPE, TKT_DATE, SALES_REP, USER_ID, STATION, DRAWER, SALE_TYPE, " &
                                        "WHOLESALE, CUSTOM1, CUSTOM2, CUSTOM3, CUSTOM4, CUSTOM5) " &
                                        "SELECT '" & id & "', " & seq & ", '" & store & "', '" & sku & "','" & loc & "'," &
                                            qty & "," & cost & "," & retail & ", " & mkdn & ",'" & transDate & "','" & transtype & "','" & datetimenow & "','" &
                                            dept & "','" & buyer & "','" & clss & "', '" & extractDate & "','" & cust & "','" & tkt & "','" &
                                            reason & "','" & coupon & "','" & item & "','" & dim1 & "','" & dim2 & "','" & dim3 & "','" & ordtype & "','" &
                                            tktDate & "','" & sales_rep & "','" & userid & "','" & station & "','" & drawer & "','" & sale_type & "','" &
                                            whsl & "','" & custom & "','" & custom2 & "','" & custom3 & "','" & custom4 & "','" & custom5 & "'"
                                    thisprocess = "XMLUpdate - Process Data - Insert into Daily_Transaction_Log - " & sql & ""
                                    cmd = New SqlCommand(sql, con)
                                    cmd.CommandTimeout = 480
                                    cmd.ExecuteNonQuery()
                                Else
                                    sql = "INSERT INTO Inv_Transaction_Log (TRANS_ID, SEQUENCE_NO, SKU, LOCATION, TRANS_DATE, TYPE, QTY, COST, " &
                                        "RETAIL, EXTRACT_DATE, POST_DATE, ITEM, DIM1, DIM2, DIM3) " &
                                        "SELECT '" & id & "', " & seq & ", '" & sku & "', '" & loc & "', '" & transDate & "', '" & transtype & "', " &
                                        qty & ", " & cost & ", " & retail & ", '" & extractDate & "', '" & datetimenow & "', '" & item & "', '" &
                                        dim1 & "', '" & dim2 & "', '" & dim3 & "'"


                                    ''  uncomment below code when new extract is running on TCM


                                    ''sql = "INSERT INTO Inv_Transaction_Log (TRANS_ID, SEQUENCE_NO, SKU, LOCATION, TRANS_DATE, TYPE, QTY, COST, " &
                                    ''    "RETAIL, EXTRACT_DATE, POST_DATE, ITEM, DIM1, DIM2, DIM3, LAST_UPDATE) " &
                                    ''    "SELECT '" & id & "', " & seq & ", '" & sku & "', '" & loc & "', '" & transDate & "', '" & transtype & "', " &
                                    ''    qty & ", " & cost & ", " & retail & ", '" & extractDate & "', '" & datetimenow & "', '" & item & "', '" &
                                    ''    dim1 & "', '" & dim2 & "', '" & dim3 & "','" & lastUpdate & "'"
                                    cmd = New SqlCommand(sql, con)
                                    cmd.CommandTimeout = 480
                                    cmd.ExecuteNonQuery()
                                End If
                            End If
                        End If
                    Else
                        qty = Decimal.Round(CDec(row("QTY")), 4, MidpointRounding.AwayFromZero)
                        cost = Decimal.Round(CDec(row("COST")), 4, MidpointRounding.AwayFromZero)
                        retail = Decimal.Round(CDec(row("RETAIL")), 4, MidpointRounding.AwayFromZero)
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
                If con.state = ConnectionState.Open Then con.Close()
            End If

            If isPhysical Then theField = "PHYS"
            Dim message As String = "Updated " & cnt & " records - Qty = " & qty &
                " Total Cost = " & Format(tCost, "###,###,###.0000") & " Total Retail = " & Format(tRetail, "###,###,###.0000") &
                " Total Markdowns = " & Format(tMkdn, "###,###,###.0000")
            Call Update_Process_Log("1", "Update " & theField & "", message, "")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Function Check_Log(ByVal con2, ByVal id, ByVal seq, ByVal location, ByVal sku, ByVal dte, ByVal store, ByVal theField)
        Try
            Dim reslt As Boolean = False
            Dim sql As String
            If con2.State = ConnectionState.Open Then con2.Close()
            con2.open()
            If theField = "Sold" Or theField = "Order" Then
                sql = "SELECT TRANS_ID FROM Daily_Transaction_Log WHERE TRANS_ID = '" & id & "' " &
                "AND SEQUENCE_NO = " & seq & " AND LOCATION = '" & location & "' AND SKU = '" & sku & "' " &
                "AND TRANS_DATE = '" & dte & "'"
            Else
                sql = "SELECT TRANS_ID FROM Inv_Transaction_Log WHERE TRANS_ID = '" & id & "' " &
                "AND SEQUENCE_NO = " & seq & " AND LOCATION = '" & location & "' AND SKU = '" & sku & "' " &
                "AND TRANS_DATE = '" & dte & "'"
            End If

            If store <> "" Then sql &= " AND STORE = '" & store & "'"
            thisprocess = "XMLUpdate - Check_Log - " & sql & ""
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
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Function

    Private Sub Update_Tables(ByVal store, ByVal location, ByVal dte, ByVal item, ByVal qty, ByVal retail,
                             ByVal cost, ByVal mkdn, ByVal theField, ByVal dept, ByVal buyer, ByVal clss,
                             ByVal dttime, ByVal whsl, ByVal con2)
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
            If con2.State = ConnectionState.Closed Then con2.open()
            If DateTime.TryParse(dttime, itsadate) Then
                thistime = dttime
            Else : thistime = DateAndTime.Now
            End If

            If whsl = "Y" Then store &= "-WHSL"

            If theField = "Sold" Or theField = "Order" Then
                sql = "IF NOT EXISTS (SELECT * FROM Item_Sales WHERE Str_Id = '" & Trim(store) & "' AND Loc_Id = '" & location & "' " &
                    "AND Sku = '" & item & "' AND eDate = '" & thisEdate & "') " &
                    "INSERT INTO Item_Sales (Str_Id, Loc_Id, Sku, sDate, eDate, Sold, Avg_Cost, Retail_Price, Markdown, " &
                    "Sales_Cost, Sales_Retail, YrWk, Item, DIM1, DIM2, DIM3) " &
                     "SELECT '" & Trim(store) & "', '" & location & "', '" & Trim(item) & "', '" & thisSdate & "', '" & thisEdate & "', " &
                     qty & ", " & cost & ", Curr_Retail, " & mkdn & ", " & cost * qty & ", " & (retail * qty) & ", " &
                     "(SELECT YrWk FROM Calendar WHERE eDate = '" & thisEdate & "' AND Prd_Id > 0 AND Week_Id > 0), " &
                     "Item, DIM1, DIM2, DIM3 " &
                     "FROM Item_Master WHERE Sku = '" & Trim(item) & "' " &
                    "ELSE " &
                    "UPDATE p SET Sold = ISNULL(Sold,0) + " & qty & ", Sales_Cost = ISNULL(Sales_Cost,0) + " & cost * qty & ", Sales_Retail = " &
                    "ISNULL(Sales_Retail,0) + " & retail * qty & ", " &
                    "Markdown = ISNULL(Markdown,0) + " & mkdn & " " &
                    "FROM Item_Sales AS p " &
                        "WHERE Str_Id = '" & store & "' AND Loc_Id = '" & location & "' AND Sku = '" & item & "' AND eDate = '" & thisEdate & "' "
            Else
                sql = "IF NOT EXISTS (SELECT * FROM Item_Inv d WHERE Loc_Id = '" & Trim(location) & "' AND Sku = '" & Trim(item) & "' " &
                    "AND eDate = '" & thisEdate & "') " &
                    "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, " & theField & ", Cost, Retail, YrWk, Item, DIM1, DIM2, DIM3, " &
                    "Begin_OH, End_OH, Max_OH) " &
                    "SELECT '" & Trim(location) & "', '" & Trim(item) & "', '" & thisSdate & "', '" & thisEdate & "', " &
                        qty & ", " & cost & ", " & retail & ", " &
                        "(SELECT YrWk FROM Calendar WHERE eDate = '" & thisEdate & "' AND Prd_Id > 0 AND Week_Id > 0), " &
                        "Item, DIM1, DIM2, DIM3, 0, " & qty & ", " & qty & " " &
                        "FROM Item_Master WHERE Sku = '" & Trim(item) & "' " &
                    "ELSE " &
                    "UPDATE d SET " & theField & " = ISNULL(" & theField & ",0) + " & qty & " " &
                    "FROM Item_Inv AS d " &
                        "WHERE Loc_Id = '" & Trim(location) & "' AND Sku = '" & Trim(item) & "' AND eDate = '" & thisEdate & "'"
            End If
            thisprocess = "XMLUpdate - Update Tables - " & sql & ""
            cmd = New SqlCommand(sql, con2)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con2.close()
            ''
            ''  if trans_date in a previous week
            ''  write a record or something and run "ForwardFill" to correct Item_Inv
            ''
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
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
            thisprocess = "XMLUpdate - Process Other Data - Read XML File " & thePath & ""
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
                                sql = "IF NOT EXISTS (SELECT ID FROM Classes WHERE ID = '" & id & "' AND Dept = '" & dept & "') " &
                                    "INSERT INTO Classes (ID, Dept, Description, Orig_Date, Status) " &
                                    "SELECT '" & id & "','" & dept & "','" & descr & "','" & CDate(Date.Today) & "','Active' " &
                                    "ELSE " &
                                    "UPDATE Classes SET Description = '" & descr & "', Dept = '" & dept & "', Extract_Date = '" & eDate & "' " &
                                    "WHERE ID = '" & id & "' AND Dept = '" & dept & "'"
                            Else
                                sql = "IF NOT EXISTS (SELECT ID FROM " & theTable & " WHERE ID = '" & id & "') " &
                                    "INSERT INTO " & theTable & " (ID, Description, Orig_Date, Status) " &
                                    "SELECT '" & id & "','" & descr & "','" & CDate(Date.Today) & "','Active' " &
                                    "ELSE " &
                                    "UPDATE " & theTable & " SET Description = '" & descr & "', Extract_Date = '" & eDate & "' " &
                                    "WHERE ID = '" & id & "' "
                            End If
                            thisprocess = "XMLUpdate - Process Other Data - " & sql & ""
                            con.open()
                            cmd = New SqlCommand(sql, con)
                            cmd.CommandTimeout = 480
                            cmd.ExecuteNonQuery()
                            con.close()
                        End If
                    Else
                        records = CInt(row("DESCRIPTION"))
                        If records <> cnt Then
                            msg = "Expected " & records & " Received " & cnt
                            Dim el As New XMLUpdate.ErrorLogger
                            el.WriteToErrorLog(msg, "Records Received Mismatch", thisprocess)
                        End If
                    End If
                Next
            End If

            con.open()
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.close()

            '                           Insert a record for buyer OTHER if we don't already have one
            ''If theTable = "Buyers" Then
            thisprocess = "XMLUpdate - Process Other Data - Fix Buyers"
            con.open()
            sql = "IF NOT EXISTS (SELECT * FROM Buyers WHERE ID = 'OTHER') " &
           "INSERT INTO Buyers (ID, Description, Orig_Date) " &
           "SELECT 'OTHER','MISCELLANEOUS BUYER','" & CDate(Date.Today) & "'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.close()
            ''End If

            Dim p As String = "Update " & theTable
            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", p, m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Process_Coupons()
        Try
            stopWatch = New Stopwatch
            stopWatch.Start()
            Dim ds As DataSet = New DataSet
            Dim dt As DataTable = New DataTable
            Dim cnt As Int16 = 0
            Dim id, code, thePath As String
            Dim seq As Integer
            thePath = xmlPath & "\Coupons.xml"
            Dim xmlFile As XmlReader
            thisprocess = "XMLUpdate - Process Coupons - Read XML File"
            xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
            ds.ReadXml(xmlFile)
            If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                con.Open()
                dt = ds.Tables(0)
                For Each row In dt.Rows
                    id = ""
                    code = ""
                    oTest = row("TRANS_ID")
                    If Not IsNothing(oTest) And oTest <> "" Then id = CStr(oTest)
                    oTest = row("CODE")
                    If Not IsNothing(oTest) And oTest <> "" Then code = CStr(oTest)
                    oTest = row("SEQ_NO")
                    If IsNumeric(oTest) Then seq = CInt(oTest) Else seq = 0
                    sql = "UPDATE Daily_Transaction_Log SET COUPON_CODE = '" & code & "' " &
                        "WHERE TRANS_ID = '" & id & "' AND SEQUENCE_NO = " & seq & ""

                    thisprocess = "XMLUpdate - Process Coupons - " & sql & ""
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 480
                    cmd.ExecuteNonQuery()
                Next
                con.Close()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Merge_Tickets()
        Try
            Console.WriteLine("Updating Tickets table")
            thisprocess = "XMLUpdate - Merge Tickets"
            con.Open()
            sql = "DECLARE @maxDate AS Date " &
                "SET @maxDate = (SELECT MAX(Date) FROM Tickets) " &
                "SET @maxDate = CASE WHEN @maxDate IS NULL THEN '1900-01-01' END " &
                "SELECT CONVERT(varchar(10),STORE) AS Str_Id, CONVERT(varchar(30),TRANS_ID) AS Ticket_No, " &
                "CONVERT(varchar(30),CUST_NO) AS Cust_No, CONVERT(date,TRANS_DATE) AS 'Date', " &
                "ISNULL(SUM(QTY * RETAIL),0) as Amt, COUNT(*) AS Items INTO #t1 FROM Daily_Transaction_Log " &
                "WHERE TYPE = 'Sold' AND TRANS_DATE >= @maxDate GROUP BY STORE, TRANS_ID, Cust_No, TRANS_DATE " &
                "MERGE Tickets AS t " &
                "USING #t1 AS s " &
                "ON (t.Str_Id = s.Str_Id AND t.Ticket_No = s.Ticket_No AND t.Cust_No = s.Cust_No AND t.Date = s.Date) " &
                "WHEN NOT MATCHED BY Target " &
                "THEN INSERT (Ticket_No, Str_Id, Cust_No, Date, Amt, Items) " &
                "VALUES (s.Ticket_No, s.Str_Id, s.Cust_No, s.Date, s.Amt, s.Items);"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
            Dim m As String = ""
            Call Update_Process_Log("1", "Merge Tickets", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Set_Max_OH(ByVal thisDate As Date)
        Try
            Dim stopwatch As New Stopwatch
            stopwatch.Start()
            Console.WriteLine("Updating Max_OH")
            thisprocess = "XMLUpdate - Set Max On Hand - Select Items"
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            If con3.State = ConnectionState.Open Then con3.Close()
            Dim maxx, end_OH, sold As Decimal
            Dim location, sku As String
            Dim cnt As Integer = 0
            Dim tbl As New DataTable
            Dim edate As Date
            con.Open()
            con2.Open()
            sql = "SELECT Loc_Id, i.Sku, eDate, End_OH INTO #t1 FROM Item_Inv i WHERE eDate = '" & thisDate & "' " &
               "SELECT Loc_Id, Sku, eDate, SUM(ISNULL(Sold,0)) AS Sold INTO #t2 FROM Item_Sales " &
               "WHERE eDate = '" & thisDate & "' GROUP BY Loc_Id, Sku, eDate " &
               "SELECT t1. Loc_Id, t1.Sku, t1.eDate, t1.End_OH, ISNULL(t2.Sold,0) AS Sold FROM #t1 t1 " &
               "LEFT JOIN #t2 t2 ON t2.Loc_Id = t1.Loc_Id AND t2.Sku = t1.Sku AND t1.eDate = t2.eDate"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("End_OH")
                If IsDBNull(oTest) Then end_OH = 0 Else end_OH = CDec(oTest)
                oTest = rdr("Sku")
                If IsDBNull(oTest) Then sku = "" Else sku = Replace(oTest, "'", "''")
                oTest = rdr("Loc_Id")
                If IsDBNull(oTest) Then location = "" Else location = Replace(oTest, "'", "''")
                oTest = rdr("eDate")
                If IsDBNull(oTest) Then edate = "1900-01-01" Else edate = CDate(oTest)
                oTest = rdr("Sold")
                If IsDBNull(oTest) Then sold = 0 Else sold = CDec(oTest)
                If end_OH < 0 Then
                    maxx = sold
                Else
                    maxx = sold + end_OH
                End If
                If maxx < 0 Then maxx = 0
                sql = "UPDATE Item_Inv SET Max_OH = " & maxx & " WHERE Loc_Id = '" & location & "' AND Sku = '" &
                    sku & "' AND eDate = '" & thisDate & "'"
                thisprocess = "XMLUpdate - Set Max On Hand - " & sql & ""
                cmd = New SqlCommand(sql, con2)
                cmd.CommandTimeout = 480
                cmd.ExecuteNonQuery()
            End While
            con.Close()
            con2.Close()
            Dim p As String = "Set Max Onhand"
            Dim m As String = "Updated " & cnt & " records"
            Call Update_Process_Log("1", p, m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        Try
            stopWatch.Stop()
            Dim ts As TimeSpan = stopWatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim pgm As String = "XMLUpdate"
            con.Open()
            sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " &
                "SELECT '" & Date.Now & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
            thisprocess = "XMLUpdate - Update Process Log - " & sql & ""
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Sub Update_Store_Table()
        Try
            con.Open()
            con2.Open()
            sql = "SELECT DISTINCT Str_ID FROM Item_Sales"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Str_ID")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    sql = "IF NOT EXISTS (SELECT * FROM Stores WHERE ID = '" & CStr(oTest) & "') " &
                        "INSERT INTO Stores(ID, Description, Status, Last_Change_Date, Last_Change_User) " &
                        "SELECT '" & CStr(oTest) & "','" & CStr(oTest) & "','Active','" & Date.Today & "','RC'"
                    thisprocess = "XMLUpdate - Update Store Table - " & sql & ""
                    cmd = New SqlCommand(sql, con2)
                    cmd.CommandTimeout = 480
                    cmd.ExecuteNonQuery()
                End If
            End While
            con.Close()
            con2.Close()
            Dim m As String = ""
            Call Update_Process_Log("1", "Update Stores for Wholesale", m, "")
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim el As New XMLUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, thisprocess)
        End Try
    End Sub

    Private Function FileSize(ByVal fileName As String)
        Dim length As Integer = 0
        If System.IO.File.Exists(fileName) Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(fileName)
            length = infoReader.Length
            If length = 0 Then
                Dim el As New XMLUpdate.ErrorLogger
                el.WriteToErrorLog(fileName, "HAS NO RECORDS", "Check File Size")
            End If
            Dim fileDate As DateTime = File.GetLastWriteTime(fileName)
            Dim hoursDiff = DateDiff(DateInterval.Hour, fileDate, Date.Now)
            If hoursDiff > 26 Then
                Dim el As New XMLUpdate.ErrorLogger
                If fileName <> xmlPath & "\ErrLog.txt" Then
                    el.WriteToErrorLog(fileName & " IS OLDER THAN 24 HOURS", "", "Check File Datetime")
                End If
            End If
        Else
            If fileName <> xmlPath & "\ErrLog.txt" Then
                Dim el As New XMLUpdate.ErrorLogger
                el.WriteToErrorLog(fileName & " NOT FOUND", "", "Check File Exists")
            End If
        End If
        Return length
    End Function

    Private Function FixHyphen(val As String) As String
        Dim newval As String = Replace(val, "''", "'")
        Return newval
    End Function

    Public Function cleanData(ByVal input As String) As String
        Dim r As String = Regex.Replace(input, "[~A-Za-z0-9\-/]", "")
        Return r.Replace(input, [String].Empty)
    End Function
End Module

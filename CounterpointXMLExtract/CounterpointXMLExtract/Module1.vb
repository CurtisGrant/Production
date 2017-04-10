Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Public errorPath As String
    Private con, cpCon As SqlConnection
    Private StopWatch As Stopwatch
    Private sql As String
    Private cmd As SqlCommand
    Private rdr As SqlDataReader
    Private qty, cost, retail, discountAmt, discount As Decimal

    Sub Main()

        StopWatch = New Stopwatch
        stopWatch.Start()
        Dim thisDate As Date = CDate(Date.Today)
        Dim minADJdate, minRCPTdate, minRTNdate, minSALdate, minXFERdate, minPOdate, thisEdate, thisSdate As Date
        Dim todaysDate As Date = Date.Today
        Dim DateTimeString As String = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
        Dim ADJwks As Integer = 2
        Dim RCPTwks As Integer = 2
        Dim RTNwks As Integer = 2
        Dim SALwks As Integer = 2
        Dim XFERwks As Integer = 2
        Dim POwks As Integer = 2
        Dim cnt As Int32 = 0
        Dim files As Integer = 1
        Dim totlQty As Decimal = 0
        Dim totlCost As Decimal = 0
        Dim totlRetail As Decimal = 0
        Dim totlMarkdown As Decimal = 0
        Dim oTest As Object
        Dim fld, valu, server, database, client, user, password, sql, rcServer, xmlPath, conString, path, cpString As String
        Dim cpServer, cpDatabase, cpUserID, cpPassword As String
        Dim logPath As String = ""
        Dim xmlWriter As XmlTextWriter
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCExtract.xml")
        Dim rcConString, rcExePath, rcPassWord As String
        While xmlReader.Read
            Select Case xmlReader.NodeType
                Case XmlNodeType.Element
                    fld = xmlReader.Name
                Case XmlNodeType.Text
                    valu = xmlReader.Value
                Case XmlNodeType.EndElement
                    'Console.WriteLine("</" & xmlReader.Name)
            End Select
            If fld = "COUNTERPOINTSERVER" Then cpServer = valu
            If fld = "COUNTERPOINTDATABASE" Then cpDatabase = valu
            If fld = "COUNTERPOINTUSERID" Then cpUserID = valu
            If fld = "COUNTERPOINTPASSWORD" Then cpPassword = valu
            If fld = "RCSERVER" Then server = valu
            If fld = "RCDATABASE" Then database = valu
            If fld = "RCUSERID" Then user = valu
            If fld = "RCPASSWORD" Then password = valu
            If fld = "ERRORPATH" Then errorPath = valu
            If fld = "EXEPATH" Then rcExePath = valu
            If fld = "XMLPATH" Then xmlPath = valu
            If fld = "CLIENTID" Then client = valu
            If fld = "ADJWKS" And IsNumeric(valu) Then ADJwks = CInt(valu)
            If fld = "POWKS" And IsNumeric(valu) Then POwks = CInt(valu)
            If fld = "RCPTWKS" And IsNumeric(valu) Then RCPTwks = CInt(valu)
            If fld = "RTNWKS" And IsNumeric(valu) Then RTNwks = CInt(valu)
            If fld = "SALWKS" And IsNumeric(valu) Then SALwks = CInt(valu)
            If fld = "XFERWKS" And IsNumeric(valu) Then XFERwks = CInt(valu)
        End While

        thisSdate = Date.Today.AddDays(0 - Date.Today.DayOfWeek)
        thisEdate = Date.Today.AddDays(6 - Date.Today.DayOfWeek)

        minADJdate = DateAdd(DateInterval.Day, ADJwks * -14, thisSdate)
        minRCPTdate = DateAdd(DateInterval.Day, RCPTwks * -14, thisSdate)
        minRTNdate = DateAdd(DateInterval.Day, RTNwks * -14, thisSdate)
        minSALdate = DateAdd(DateInterval.Day, SALwks * -14, thisSdate)
        minXFERdate = DateAdd(DateInterval.Day, XFERwks * -14, thisSdate)
        minPOdate = DateAdd(DateInterval.Day, POwks * -185, thisDate)

        conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & "; Integrated Security=True"
        cpCon = New SqlConnection(conString)
       
        errorPath = xmlPath                            ' force the error log to write to the same folder as the xmls



        ''  GoTo 100


        '               Delete errlog first thing
        Try
            Dim thisDay As Integer = Date.Now.DayOfWeek
            path = xmlPath & "\errlog.txt"
            If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Delete errlog")
        End Try
        Try
            cnt = 0
            If (Not System.IO.Directory.Exists(xmlPath)) Then
                System.IO.Directory.CreateDirectory(xmlPath)
            End If
            Console.WriteLine("Extracting Buyers")
            Dim id, desc As String
            path = xmlPath & "\Buyers.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Buyers")
            Dim constr2 As String = ""
            Dim stat As String = ""
            cpCon.Open()
            sql = "SELECT DISTINCT CATEG_COD AS ID, s.NAM_UPR AS DESCR INTO #t1 FROM PO_VEND p " & _
                "JOIN SY_USR s ON CATEG_COD = USR_ID " & _
                "WHERE CATEG_COD IS NOT NULL " & _
                "SELECT DISTINCT BUYER AS ID, '' AS DESCR INTO #t2 FROM PO_ORD_HDR WHERE BUYER IS NOT NULL; " & _
                "MERGE #t1 AS t USING #t2 AS s ON (s.ID = t.ID) " & _
                "WHEN NOT MATCHED BY TARGET THEN INSERT(ID, DESCR) VALUES (s.ID, s.ID); " & _
                "SELECT ID, DESCR FROM #t1"
            cmd = New SqlCommand(sql, cpCon)

            If Err.Number <> 0 Then Console.WriteLine(Err.Number)

            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = rdr("DESCR")
                    If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                    xmlWriter.WriteStartElement("Buyer")
                    xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                    xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                    xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                    xmlWriter.WriteEndElement()
                End If
            End While
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Buyer")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Buyers", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Buyers")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Calendar")
            path = xmlPath & "\Calendar.xml"
            cpCon.Open()
            sql = "SELECT * FROM SY_CALNDR WHERE CALNDR_ID = 2016"
            Dim cmd2 As New SqlCommand(sql, cpCon)
            Dim rdr2 As SqlDataReader = cmd2.ExecuteReader

            Dim row As DataRow
            Dim prdcnt As Int16 = 0
            Dim theLasteDate As Date
            Dim theYear As Integer = 2016
            Dim tbl As New DataTable
            tbl.Columns.Add("Year_Id")
            tbl.Columns.Add("Prd_Id")
            tbl.Columns.Add("Week_Id")
            tbl.Columns.Add("sDate")
            tbl.Columns.Add("eDate")
            tbl.Columns.Add("YrPrd")
            tbl.Columns.Add("YrWks")
            While rdr2.Read
                ' Add records for the year range
                theLasteDate = rdr2.Item(1)
                row = tbl.NewRow
                row.Item("Year_Id") = theYear                                                         ' Year
                row.Item("Prd_Id") = 0                                                               ' Period_Id
                row.Item("Week_Id") = 0                                                               ' Week_Id
                row.Item("sDate") = CDate(theLasteDate)                                                    ' sDate
                row.Item("eDate") = CDate(rdr2(2))                                                         ' edate
                row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrPrd
                row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
                tbl.Rows.Add(row)
                ' Add records for period ranges (called Months in CounterPoint)
                For i = 2 To 28 Step 2
                    oTest = rdr2.Item(i + 16)
                    If Not IsDBNull(oTest) Then
                        prdcnt += 1
                        row = tbl.NewRow
                        row.Item("Year_Id") = theYear                                                         ' Year
                        row.Item("Prd_Id") = prdcnt                                                          ' Period_Id
                        row.Item("Week_Id") = 0                                                               ' Week_Id
                        row.Item("sDate") = CDate(theLasteDate)                                                    ' sDate
                        row.Item("eDate") = CDate(rdr2(i + 16))                                                    ' edate
                        row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & Format(prdcnt, "00")  ' YrPrd
                        row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
                        theLasteDate = rdr2.Item(i + 16)
                        theLasteDate = DateAdd(DateInterval.Day, 1, theLasteDate)
                        tbl.Rows.Add(row)
                    End If
                Next
                prdcnt = 0
                theLasteDate = rdr2.Item(1)
                ' Add records for week ranges
                For i = 2 To 108 Step 2
                    oTest = rdr2.Item(i + 44)
                    If Not IsDBNull(oTest) Then
                        prdcnt += 1
                        row = tbl.NewRow
                        row.Item("Year_Id") = theYear                                                         ' Year
                        row.Item("Prd_Id") = 0                                                               ' Period_Id
                        row.Item("Week_Id") = prdcnt                                                          ' Week_Id
                        row.Item("sDate") = CDate(theLasteDate)                                                    ' sDate
                        row.Item("eDate") = CDate(rdr2(i + 44))                                                    ' edate
                        row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrPrd
                        row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & Format(prdcnt, "00")  ' YrWks
                        theLasteDate = rdr2.Item(i + 44)
                        theLasteDate = DateAdd(DateInterval.Day, 1, theLasteDate)
                        tbl.Rows.Add(row)
                    End If
                Next
            End While
            cpCon.Close()
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Calendar")
            For Each row In tbl.Rows
                oTest = row("Year_Id")
                xmlWriter.WriteStartElement("Date")
                xmlWriter.WriteElementString("Year_Id", oTest)
                oTest = row("Prd_Id")
                xmlWriter.WriteElementString("Prd_Id", oTest)
                oTest = row("Week_Id")
                xmlWriter.WriteElementString("Week_Id", oTest)
                oTest = row("sDate")
                xmlWriter.WriteElementString("sDate", oTest)
                oTest = row("eDate")
                xmlWriter.WriteElementString("eDate", oTest)
                oTest = row("YrPrd")
                xmlWriter.WriteElementString("YrPrd", oTest)
                oTest = row("YrWks")
                xmlWriter.WriteElementString("YrWks", oTest)
                xmlWriter.WriteEndElement()
            Next
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Calendar")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Classes")
            path = xmlPath & "\Classes.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Classes")
            Dim dept, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT SUBCAT_COD AS ID, CATEG_COD AS DEPT, DESCR FROM IM_SUBCAT_COD WHERE SUBCAT_COD IS NOT NULL"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                dept = rdr("Dept")
                desc = rdr("DESCR")
                xmlWriter.WriteStartElement("Class")
                xmlWriter.WriteElementString("ID", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("DEPT", Replace(dept, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Class")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Classes", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Classes")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Departments")
            path = xmlPath & "\Departments.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Departments")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT CATEG_COD AS ID, DESCR FROM IM_CATEG_COD WHERE CATEG_COD IS NOT NULL"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA'"
                xmlWriter.WriteStartElement("Department")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Department")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Departments", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Departments")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Product Lines")
            path = xmlPath & "\ProductLines.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("ProductLines")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT PROF_COD AS ID, DESCR FROM IM_ITEM_PROF_COD WHERE PROF_COD IS NOT NULL " & _
                "AND VAL_3 ='Y'"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                xmlWriter.WriteStartElement("ProductLine")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("ProductLine")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Product Lines", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Product Lines")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Seasons")
            path = xmlPath & "\Seasons.xml"
            Dim id, desc As String
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Seasons")
            cpCon.Open()
            sql = "SELECT DISTINCT PROF_COD AS ID, DESCR FROM IM_ITEM_PROF_COD WHERE PROF_COD IS NOT NULL " & _
                "AND VAL_1 ='Y'"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                xmlWriter.WriteStartElement("Season")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Season")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Seasons", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Product Lines")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Locations")
            path = xmlPath & "\Locations.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Location")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT LOC_ID AS ID, DESCR_UPR FROM IM_LOC WHERE LOC_ID IS NOT NULL "
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR_UPR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                xmlWriter.WriteStartElement("Location")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Location")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Locations", m, "")
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Locations")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Stores")
            path = xmlPath & "\Stores.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Stores")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT s.STR_ID AS ID, DESCR_UPR, STK_LOC_ID FROM PS_STR s " & _
                "JOIN PS_STR_CFG_PS c ON c.STR_ID = s.STR_ID WHERE STK_LOC_ID IS NOT NULL "
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR_UPR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                xmlWriter.WriteStartElement("Store")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LOCATION", Replace(rdr("STK_LOC_ID"), "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Store")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Stores", m, "")
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Stores")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Vendors")
            path = xmlPath & "\Vendors.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Vendors")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT DISTINCT VEND_NO AS ID, NAM_UPR FROM PO_VEND WHERE VEND_NO IS NOT NULL "
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("NAM_UPR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                xmlWriter.WriteStartElement("Vendor")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Vendor")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Vendors", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extract Vendors")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Coupon Codes")
            path = xmlPath & "\Coupon_Codes.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Coupon_Code")
            Dim id, desc As String
            cpCon.Open()
            sql = "SELECT REAS_COD, DESCR FROM PS_REAS_COD ORDER BY REAS_COD"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("Code")
                oTest = rdr("REAS_COD")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    cnt += 1
                    xmlWriter.WriteElementString("ID", Replace(oTest, "'", "''"))
                    oTest = rdr("DESCR")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then desc = oTest Else desc = ""
                    xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                    xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                    xmlWriter.WriteEndElement()
                End If
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Code")
            xmlWriter.WriteElementString("ID", "TOTALS")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Coupon Codes", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extract Coupon Codes")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Customers")
            path = xmlPath & "\Customers.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Customers")
            Dim flipIt As String
            cpCon.Open()
            sql = "SELECT CUST_NO, NAM_UPR AS NAME, FST_NAM_UPR AS FIRST_NAME, LST_NAM_UPR AS LAST_NAME, ADRS_1 AS ADDRESS_1, " & _
                "ADRS_2 AS ADDRESS_2, ADRS_3 AS ADDRESS_3, CITY, STATE, ZIP_COD AS ZIP, CONTCT_2 AS SPOUSE, COMMNT AS BIRTHDAY, " & _
                "PHONE_1, PHONE_2, EMAIL_ADRS_1 AS EMAIL, CATEG_COD AS TYPE, BAL AS BALANCE, " & _
                "FST_SAL_DAT AS FIRST_SALE_DATE, LST_SAL_DAT AS LAST_SALE_DATE, LST_MAINT_DT AS LAST_UPDATE, " & _
                "LOY_PTS_BAL AS LOYALTY_POINTS, INCLUDE_IN_MARKETING_MAILOUTS AS OK_TO_EMAIL, PROF_ALPHA_2 AS OK_TO_MAIL, " & _
                "MBL_PHONE_1 AS CELL_1, MBL_PHONE_2 AS CELL_2 FROM AR_CUST"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            cnt = 0
            While rdr.Read
                xmlWriter.WriteStartElement("Customer")
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("CUST_NO")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("CUST_NO", oTest)
                oTest = rdr("NAME")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("NAME", oTest)
                oTest = rdr("FIRST_NAME")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("FIRST_NAME", oTest)
                oTest = rdr("LAST_NAME")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_NAME", oTest)
                oTest = rdr("ADDRESS_1")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("ADDRESS_1", oTest)
                oTest = rdr("ADDRESS_2")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("ADDRESS_2", oTest)
                oTest = rdr("ADDRESS_3")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("ADDRESS_3", oTest)
                oTest = rdr("CITY")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("CITY", oTest)
                oTest = rdr("STATE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("STATE", oTest)
                oTest = rdr("ZIP")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("ZIP", oTest)
                oTest = rdr("SPOUSE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("SPOUSE", oTest)
                oTest = rdr("BIRTHDAY")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("BIRTHDAY", oTest)
                oTest = rdr("PHONE_1")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("PHONE_1", oTest)
                oTest = rdr("PHONE_2")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("PHONE_2", oTest)
                oTest = rdr("EMAIL")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("EMAIL", oTest)
                oTest = rdr("TYPE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("TYPE", oTest)
                oTest = rdr("BALANCE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("BALANCE", oTest)
                oTest = rdr("FIRST_SALE_DATE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("FIRST_SALE_DATE", oTest)
                oTest = rdr("LAST_SALE_DATE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_SALE_DATE", oTest)
                oTest = rdr("LAST_UPDATE")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_UPDATE", oTest)
                oTest = rdr("LOYALTY_POINTS")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("LOYALTY_POINTS", oTest)
                oTest = rdr("OK_TO_EMAIL")
                If IsDBNull(oTest) Then oTest = Nothing
                If oTest = "N" Then
                    flipIt = "Y"
                Else : flipIt = "N"
                End If
                xmlWriter.WriteElementString("OK_TO_EMAIL", flipIt)
                oTest = rdr("OK_TO_MAIL")
                If IsDBNull(oTest) Then oTest = Nothing
                If IsNothing(oTest) Then oTest = "Y" Else oTest = "N"
                xmlWriter.WriteElementString("OK_TO_MAIL", oTest)
                oTest = rdr("CELL_1")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("CELL_1", oTest)
                oTest = rdr("CELL_2")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("CELL_2", oTest)
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            '
            '                        Write a total record
            '
            xmlWriter.WriteStartElement("Customer")
            xmlWriter.WriteElementString("CUST_NO", "TOTALS")
            xmlWriter.WriteElementString("NAME", cnt)
            xmlWriter.WriteElementString("FIRST_NAME", "")
            xmlWriter.WriteElementString("LAST_NAME", "")
            xmlWriter.WriteElementString("ADDRESS_1", "")
            xmlWriter.WriteElementString("ADDRESS_2", "")
            xmlWriter.WriteElementString("ADDRESS_3", "")
            xmlWriter.WriteElementString("CITY", "")
            xmlWriter.WriteElementString("STATE", "")
            xmlWriter.WriteElementString("ZIP", "")
            xmlWriter.WriteElementString("SPOUSE", "")
            xmlWriter.WriteElementString("BIRTHDAY", "")
            xmlWriter.WriteElementString("PHONE_1", "")
            xmlWriter.WriteElementString("PHONE_2", "")
            xmlWriter.WriteElementString("EMAIL", "")
            xmlWriter.WriteElementString("TYPE", "")
            xmlWriter.WriteElementString("BALANCE", "")
            xmlWriter.WriteElementString("FIRST_SALE_DATE", "")
            xmlWriter.WriteElementString("LAST_SALE_DATE", "")
            xmlWriter.WriteElementString("LAST_UPDATE", "")
            xmlWriter.WriteElementString("LOYALTY_POINTS", "")
            xmlWriter.WriteElementString("OK_TO_EMAIL", "")
            xmlWriter.WriteElementString("OK_TO_MAIL", "")
            xmlWriter.WriteElementString("CELL_1", "")
            xmlWriter.WriteElementString("CELL_2", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("3", "Extract Customers", m, "")

        Catch ex As Exception

            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Customer Data")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Item data")
            path = xmlPath & "\Items.xml"
            Dim sku As String
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Items")
            cpCon.Open()
            sql = "SELECT i.ITEM_NO, DIM_1_UPR AS DIM1, DIM_2_UPR AS DIM2, DIM_3_UPR AS DIM3, " & _
                "GRID_DIM_1_TAG AS DIM1_DESCR, GRID_DIM_2_TAG AS DIM2_DESCR, GRID_DIM_3_TAG AS DIM3_DESCR, " & _
                "DIM_1_SEQ_NO, DIM_2_SEQ_NO, DIM_3_SEQ_NO, " & _
                "DESCR AS DESCRIPTION, i.VEND_ITEM_NO, pv.VEND_ITEM_NO AS povItem, i.ITEM_VEND_NO AS VENDOR_ID, " & _
                "CASE WHEN LEN(i.ITEM_VEND_NO) > 0 THEN i.ITEM_NO ELSE pv.VEND_NO END AS VENDOR_ID, p.NAM AS VENDOR, " & _
                "CASE WHEN LEN(pv.LST_ORD_BUYER) > 0 THEN pv.LST_ORD_BUYER ELSE p.CATEG_COD END AS BUYER, " & _
                "i.CATEG_COD AS DEPT, i.SUBCAT_COD AS CLASS, i.PRC_1 AS CURR_RTL, pv.UNIT_COST / pv. PUR_NUMER AS CURR_COST, " & _
                "i.PROF_COD_1 AS SEASON, i.PROF_COD_3 AS PLINE, i.PROF_COD_5 AS MKTG, i.ALT_1_UNIT AS UOM, i.ALT_1_NUMER AS BUYUNIT, " & _
                "i.ALT_1_DENOM AS SELLUNIT, imn.NOTE_TXT, i.STAT AS STATUS, i.ITEM_TYP AS TYPE " & _
                "FROM IM_ITEM AS i " & _
                "LEFT JOIN IM_GRID_DIM_1 g1 ON g1.ITEM_NO = i.ITEM_NO " & _
                "LEFT JOIN IM_GRID_DIM_2 g2 ON g2.ITEM_NO = i.ITEM_NO " & _
                "LEFT JOIN IM_GRID_DIM_3 g3 ON g3.ITEM_NO = i.ITEM_NO " & _
                "LEFT JOIN PO_VEND AS p ON p.VEND_NO = i. ITEM_VEND_NO " & _
                "LEFT JOIN PO_VEND_ITEM AS pv ON pv.ITEM_NO = i.ITEM_NO AND i.ITEM_VEND_NO = pv.VEND_NO " & _
                "LEFT JOIN IM_ITEM_NOTE AS imn ON imn.ITEM_NO =i.ITEM_NO AND NOTE_ID = 'PO'"

            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("ITEM")
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("ITEM_NO")
                If IsDBNull(oTest) Then sku = "NA" Else sku = oTest
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then
                    sku &= "~" & UCase(oTest) & "~" & UCase(rdr("DIM2")) & "~" & UCase(rdr("DIM3"))
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))

                oTest = rdr("DIM1_DESCR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM1_DESCR", Replace(oTest, "'", "''"))

                oTest = rdr("DIM2_DESCR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM2_DESCR", Replace(oTest, "'", "''"))

                oTest = rdr("DIM3_DESCR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM3_DESCR", Replace(oTest, "'", "''"))

                oTest = rdr("DIM_1_SEQ_NO")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM_1_SEQ_NO", oTest)

                oTest = rdr("DIM_2_SEQ_NO")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM_2_SEQ_NO", oTest)

                oTest = rdr("DIM_3_SEQ_NO")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM_3_SEQ_NO", oTest)

                oTest = rdr("DESCRIPTION")
                If InStr(oTest, "") > 0 Then oTest = oTest.Replace(Chr(34), Chr(34) & Chr(34))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(oTest, "'", "''"))

                Dim vitem As String = ""
                oTest = rdr("VEND_ITEM_NO")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    vitem = oTest
                Else
                    oTest = rdr("povItem")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vitem = rdr("povItem")
                End If
                xmlWriter.WriteElementString("VEND_ITEM_NO", Replace(vitem, "'", "''"))

                oTest = rdr("VENDOR_ID")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("VENDOR_ID", Replace(oTest, "'", "''"))

                oTest = rdr("VENDOR")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("VENDOR", Replace(oTest, "'", "''"))

                oTest = rdr("BUYER")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("BUYER", Replace(oTest, "'", "''"))

                oTest = rdr("DEPT")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DEPT", Replace(oTest, "'", "''"))

                oTest = rdr("CLASS")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CLASS", Replace(oTest, "'", "''"))

                oTest = rdr("CURR_Rtl")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("CURR_RTL", Replace(oTest, "'", "''"))

                oTest = rdr("CURR_COST")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("CURR_COST", Replace(oTest, "'", "''"))

                oTest = rdr("SEASON")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("SEASON", Replace(oTest, "'", "''"))

                oTest = rdr("PLINE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("PLINE", Replace(oTest, "'", "''"))

                oTest = rdr("MKTG")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("MKTG", Replace(oTest, "'", "''"))

                oTest = rdr("UOM")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("UOM", Replace(oTest, "'", "''"))

                oTest = rdr("BUYUNIT")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("BUYUNIT", Replace(oTest, "'", "''"))

                oTest = rdr("SELLUNIT")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("SELLUNIT", Replace(oTest, "'", "''"))

                oTest = rdr("STATUS")
                If IsDBNull(oTest) Then oTest = ""
                If oTest = "A" Then oTest = "Active" Else oTest = "Inactive" '              items are either Active or Inactive
                xmlWriter.WriteElementString("STATUS", Replace(oTest, "'", "''"))

                oTest = rdr("NOTE_TXT")
                If IsDBNull(oTest) Then oTest = "" Else oTest = Trim(Left(oTest, 20))
                xmlWriter.WriteElementString("NOTE", Replace(oTest, "'", "''"))

                oTest = rdr("TYPE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("TYPE", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteStartElement("SKU")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("DIM1_DESCR", "")
            xmlWriter.WriteElementString("DIM2_DESCR", "")
            xmlWriter.WriteElementString("DIM3_DESCR", "")
            xmlWriter.WriteElementString("DESCRIPTION", cnt)
            xmlWriter.WriteElementString("VEND_ITEM_NO", "")
            xmlWriter.WriteElementString("VENDOR_ID", "")
            xmlWriter.WriteElementString("VENDOR", "")
            xmlWriter.WriteElementString("BUYER", "")
            xmlWriter.WriteElementString("DEPT", "")
            xmlWriter.WriteElementString("CLASS", "")
            xmlWriter.WriteElementString("CURR_RTL", "")
            xmlWriter.WriteElementString("CURR_COST", "")
            xmlWriter.WriteElementString("SEASON", "")
            xmlWriter.WriteElementString("PLINE", "")
            xmlWriter.WriteElementString("MKTG", "")
            xmlWriter.WriteElementString("UOM", "")
            xmlWriter.WriteElementString("BUYUNIT", "")
            xmlWriter.WriteElementString("SELLUNIT", "")
            xmlWriter.WriteElementString("STATUS", "")
            xmlWriter.WriteElementString("NOTE", "")
            xmlWriter.WriteElementString("TYPE", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Items", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Item Data")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Dim sku, item, barcod_id, barcod As String
            Console.WriteLine("Extracting Barcode data")
            path = xmlPath & "\Barcodes.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Barcodes")
            cpCon.Open()
            sql = "SELECT ITEM_NO, BARCOD_ID, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, BARCOD FROM IM_BARCOD"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("BARCODE")
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("ITEM_NO")
                If IsDBNull(oTest) Then sku = "NA" Else sku = oTest
                item = sku
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    sku &= "~" & UCase(oTest) & "~" & UCase(rdr("DIM_2_UPR")) & "~" & UCase(rdr("DIM_3_UPR"))
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                xmlWriter.WriteElementString("ITEM", Replace(item, "'", "''"))

                oTest = rdr("DIM_1_UPR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM1", Replace(oTest, "'", "''"))

                oTest = rdr("DIM_2_UPR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM2", Replace(oTest, "'", "''"))

                oTest = rdr("DIM_3_UPR")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DIM3", Replace(oTest, "'", "''"))

                oTest = rdr("BARCOD_ID")
                If IsDBNull(oTest) Then barcod_id = "NA" Else barcod_id = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("BARCOD_ID", barcod_id)

                oTest = rdr("BARCOD")
                If IsDBNull(oTest) Then barcod = "NA" Else barcod = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("BARCOD", barcod)

                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteStartElement("SKU")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("ITEM", "")
            xmlWriter.WriteElementString("DIM1", "")
            xmlWriter.WriteElementString("DIM2", "")
            xmlWriter.WriteElementString("DIM3", "")
            xmlWriter.WriteElementString("BARCOD_ID", "")
            xmlWriter.WriteElementString("BARCOD", cnt)
            xmlWriter.WriteElementString("EXTRACT_DATE", "")
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Barcode Data")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Inventory data")
            cpCon.Open()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Dim avail As Decimal
            Dim commited As Decimal = 0
            Dim totlcommited As Decimal = 0
            path = xmlPath & "\Inventory.xml"
            Dim sku As String
            Dim path2 As String = xmlPath & "Inventory_" & Today & ".xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Inventory")
            sql = "SELECT ii.LOC_ID AS LOCATION, ii.ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, " & _
                "CASE WHEN c.QTY_ON_HND IS NULL THEN ii.QTY_ON_HND ELSE c.QTY_ON_HND END AS ONHAND, " & _
                "CASE WHEN c.QTY_AVAIL IS NOT NULL THEN c.QTY_AVAIL ELSE ii.QTY_AVAIL END AS AVAIL, " & _
                "CASE WHEN c.QTY_COMMIT IS NOT NULL THEN c.QTY_COMMIT ELSE ii.QTY_COMMIT END AS COMMITED, " & _
                "CASE WHEN ii.AVG_COST = 0 THEN ii.STD_COST ELSE ii.AVG_COST END AS COST, i.PRC_1 AS RETAIL " & _
                "FROM IM_INV AS ii  " & _
                "LEFT JOIN IM_INV_CELL c ON c.ITEM_NO = ii.ITEM_NO AND c.LOC_ID = ii.LOC_ID " & _
                "LEFT JOIN IM_ITEM AS i ON i.ITEM_NO = ii.ITEM_NO " & _
                "WHERE ii.Qty_Avail <> 0"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("ONHAND")
                If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                oTest = rdr("COMMITED")
                If IsDBNull(oTest) Then commited = 0 Else commited = CDec(oTest)
                oTest = rdr("AVAIL")
                If IsDBNull(oTest) Then avail = 0 Else avail = CDec(oTest)
                xmlWriter.WriteStartElement("ITEM")
                totlQty += qty
                totlcommited += commited
                oTest = rdr("LOCATION")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then oTest = FixData(oTest) Else oTest = "NA"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("ITEM_NO")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then oTest = FixData(oTest) Else oTest = "NA"
                sku = oTest
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                xmlWriter.WriteElementString("ONHAND", qty)
                xmlWriter.WriteElementString("COMMITTED", commited)
                xmlWriter.WriteElementString("AVAIL", avail)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()

            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("ITEM")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("ONHAND", totlQty)
            xmlWriter.WriteElementString("COMMITTED", totlcommited)
            xmlWriter.WriteElementString("AVAIL", "")
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Inventory", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Inventory Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try


        ''Stop


        Try
            cnt = 0
            totlRetail = 0
            Dim amt As Decimal
            Dim preq, batch, vend_id, vend, loc, buyer, alloc, mrg As String
            Dim ord, del, can As Date
            path = xmlPath & "\Purchase_Request_Header.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseRequestHeader")
            Console.WriteLine("extracting Purchase Requests Header")
            cpCon.Open()
            sql = "SELECT PREQ_NO, BAT_ID, VEND_NO, VEND_NAM, LOC_ID, BUYER, ORD_DAT, DELIV_DAT, CANCEL_DAT, ORD_TOT, " & _
                "IS_ALLOC, ALLOC_SEP_OR_MERGED FROM PO_PREQ_HDR"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                preq = Nothing
                batch = Nothing
                vend_id = Nothing
                vend = Nothing
                loc = Nothing
                buyer = Nothing
                alloc = Nothing
                mrg = Nothing
                ord = Nothing
                del = Nothing
                can = Nothing
                amt = 0
                oTest = rdr("PREQ_NO")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then preq = CStr(oTest)
                oTest = rdr("BAT_ID")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then batch = CStr(oTest)
                oTest = rdr("VEND_NO")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vend_id = CStr(oTest)
                oTest = rdr("VEND_NAM")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vend = CStr(oTest)
                oTest = rdr("LOC_ID")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then loc = CStr(oTest)
                oTest = rdr("BUYER")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = CStr(oTest)
                oTest = rdr("ORD_DAT")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ord = CDate(oTest)
                oTest = rdr("DELIV_DAT")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then del = CDate(oTest)
                oTest = rdr("CANCEL_DAT")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then can = CDate(oTest)
                oTest = rdr("ORD_TOT")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    totlRetail += CDec(oTest)
                    amt = CDec(oTest)
                End If
                oTest = rdr("IS_ALLOC")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then alloc = CStr(oTest)
                oTest = rdr("ALLOC_SEP_OR_MERGED")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then mrg = CStr(oTest)
                xmlWriter.WriteStartElement("PREQ")
                xmlWriter.WriteElementString("PREQ_NO", preq)
                xmlWriter.WriteElementString("BATCH", batch)
                xmlWriter.WriteElementString("VEND_NO", vend_id)
                xmlWriter.WriteElementString("VENDOR", vend)
                xmlWriter.WriteElementString("LOC_ID", loc)
                xmlWriter.WriteElementString("BUYER", buyer)
                xmlWriter.WriteElementString("ORDER_DATE", ord)
                xmlWriter.WriteElementString("DELIVER_DATE", del)
                xmlWriter.WriteElementString("CANCEL_DATE", can)
                xmlWriter.WriteElementString("ORDER_TOTAL", amt)
                xmlWriter.WriteElementString("ISALLOCATED", alloc)
                xmlWriter.WriteElementString("MERGED", mrg)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            xmlWriter.WriteStartElement("PREQ")
            xmlWriter.WriteElementString("PREQ_NO", "")
            xmlWriter.WriteElementString("BATCH", "")
            xmlWriter.WriteElementString("VEND_NO", "TOTALS")
            xmlWriter.WriteElementString("VENDOR", "")
            xmlWriter.WriteElementString("LOC_ID", "")
            xmlWriter.WriteElementString("BUYER", "")
            xmlWriter.WriteElementString("ORDER_DATE", "")
            xmlWriter.WriteElementString("DELIVER_DATE", "")
            xmlWriter.WriteElementString("CANCEL_DATE", "")
            xmlWriter.WriteElementString("ORDER_TOTAL", totlRetail)
            xmlWriter.WriteElementString("ISALLOCATED", "")
            xmlWriter.WriteElementString("MERGED", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Purchase Request Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            Console.WriteLine("Extracting Purchase Request Details")
            cnt = 0
            totlCost = 0
            totlQty = 0
            Dim sku As String = ""
            path = xmlPath & "\Purchase_Request_Detail.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseRequestDetail")
            cpCon.Open()
            sql = "SELECT l.PREQ_NO, l.SEQ_NO, ITEM_NO, DESCR_UPR, ORD_UNIT, ORD_QTY_NUMER, ORD_QTY_DENOM, " & _
                "ORD_COST, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, C.ORD_QTY FROM PO_PREQ_LIN l " & _
                "LEFT JOIN PO_PREQ_CELL C ON c.PREQ_NO = l.PREQ_NO AND c.SEQ_NO = l.SEQ_NO"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("PREQ_NO")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteStartElement("PREQ")
                xmlWriter.WriteElementString("PREQ_NO", oTest)
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("ITEM_NO")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                sku = oTest
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    sku &= "~" & UCase(oTest) & "~" & UCase(rdr("DIM_2_UPR")) & "~" & UCase(rdr("DIM_3_UPR"))
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("ITEM_NO")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("ITEM_NO", oTest)
                oTest = rdr("DESCR_UPR")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("DESC", oTest)
                oTest = rdr("ORD_UNIT")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("UOM", oTest)
                oTest = rdr("ORD_QTY_NUMER")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("NUMER", oTest)
                oTest = rdr("ORD_QTY_DENOM")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("DENOM", oTest)
                oTest = rdr("ORD_COST")
                If IsDBNull(oTest) Then oTest = 0
                cost = CDec(oTest)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("DIM_1_UPR")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("DIM1", oTest)
                oTest = rdr("DIM_2_UPR")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("DIM2", oTest)
                oTest = rdr("DIM_3_UPR")
                If IsDBNull(oTest) Then oTest = ""
                oTest = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("DIM3", oTest)
                oTest = rdr("ORD_QTY")
                If IsDBNull(oTest) Then oTest = 0
                qty = CDec(oTest)
                totlQty += qty
                totlCost += (cost * qty)
                xmlWriter.WriteElementString("QTY", oTest)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            xmlWriter.WriteStartElement("PREQ")
            xmlWriter.WriteElementString("PREQ_NO", "")
            xmlWriter.WriteElementString("SEQ_NO", "")
            xmlWriter.WriteElementString("SKU", "")
            xmlWriter.WriteElementString("ITEM_NO", "TOTALS")
            xmlWriter.WriteElementString("DESC", "")
            xmlWriter.WriteElementString("UOM", "")
            xmlWriter.WriteElementString("NUMER", "")
            xmlWriter.WriteElementString("DENOM", "")
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("DIM1", "")
            xmlWriter.WriteElementString("DIM2", "")
            xmlWriter.WriteElementString("DIM3", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Purchase Request Detail Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Purchase Orders")
            cpCon.Open()
            cnt = 0
            totlCost = 0
            totlRetail = 0
            totlQty = 0
            Dim pos As Integer = 0
            Dim ordQty, recvQty, canQty, expQty As Decimal
            Dim tbl As New DataTable
            Dim row, foundRow As DataRow
            Dim column As New DataColumn
            Dim po, store, vendor, buyer, orddate, duedate, candate, sku, dim1, dim2, dim3 As String
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "PO"
            tbl.Columns.Add(column)
            Dim primaryKey(1) As DataColumn
            primaryKey(0) = tbl.Columns("PO")
            tbl.PrimaryKey = primaryKey
            tbl.Columns.Add("Str_id", GetType(System.String))
            tbl.Columns.Add("OrdDate", GetType(System.String))
            tbl.Columns.Add("DueDate", GetType(System.String))
            tbl.Columns.Add("CanDate", GetType(System.String))
            tbl.Columns.Add("Vendor", GetType(System.String))
            tbl.Columns.Add("Buyer", GetType(System.String))
            tbl.Columns.Add("Status", GetType(System.String))
            tbl.Columns.Add("Amt", GetType(System.String))
            tbl.Columns.Add("Recvd_Cost", GetType(System.String))
            tbl.Columns.Add("Lines", GetType(System.String))
            tbl.Columns.Add("STK_Qty", GetType(System.String))
            tbl.Columns.Add("Open_Lines", GetType(System.String))
            tbl.Columns.Add("Open_Amt", GetType(System.String))
            path = xmlPath & "\PODetail.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseOrder")
            sql = "SELECT h.LOC_ID, h.PO_NO, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " & _
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.ORD_QTY ELSE l.ORD_QTY * ORD_QTY_NUMER END AS ORD_QTY, " & _
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.QTY_RECVD ELSe l.QTY_RECVD * ORD_QTY_NUMER END AS QTY_RECVD, " & _
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.QTY_EXPECTD ELSE l.QTY_EXPECTD * ORD_QTY_NUMER END AS QTY_EXPECTED, " & _
                "h.ORD_DAT, h.DELIV_DAT, h.CANCEL_DAT, ORD_COST, PRC_1, h.VEND_NO, BUYER, PO_STAT, " & _
                "(SELECT MAX(RECVR_DAT) FROM PO_RECVR_HIST_LIN r WHERE r.PO_NO = h.PO_NO " & _
                "AND r.RECVR_LOC_ID = h.LOC_ID AND r.ITEM_NO = l.ITEM_NO) AS LAST_RECVD_DATE, " & _
                "h.ORD_SUB_TOT AS AMT, h.RECVD_TOT_COST AS RECVD_COST, h.LIN_CNT AS LINES, h.ORD_QTY_IN_STK_UNITS AS STK_QTY, " & _
                "OPN_LIN_CNT AS OPEN_LINES, h.OPN_PO_TOT AS OPEN_AMT " & _
                "FROM PO_ORD_HDR h " & _
                "JOIN PO_ORD_LIN l ON l.PO_NO = h.PO_NO " & _
                "LEFT JOIN PO_ORD_CELL c ON c.PO_NO = l.PO_NO AND c.SEQ_NO = l.SEQ_NO " & _
                "JOIN IM_ITEM i ON i.ITEM_NO = l.ITEM_NO " & _
                "WHERE ORD_DAT >= '" & minPOdate & "' " & _
                "UNION " & _
                "SELECT RECVR_LOC_ID AS Loc_Id, 'X'+l.RECVR_NO AS PO_NO, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, 0 AS ORD_QTY, " & _
                "CASE WHEN c.RECVR_NO IS NOT NULL THEN c.QTY_RECVD ELSE l.QTY_RECVD * QTY_RECVD_NUMER END AS QTY_RECVD, " & _
                "l.NEW_QTY_EXPECTD * QTY_RECVD_NUMER, RECVR_DAT AS ORD_DAT, RECVR_DAT AS DELIV_DAT, " & _
                "RECVR_DAT AS CANCEL_DAT, RECVD_COST / QTY_RECVD_NUMER AS ORD_COST, UNIT_RETL_VAL / QTY_RECVD_NUMER AS PRC_1, " & _
                "l.ITEM_VEND_NO AS VEND_NO, '' AS BUYER, '' AS PO_STAT, RECVR_DAT AS LAST_RECVD_DAT, " & _
                "NULL AS AMT, NULL AS RECVD_COST, NULL AS LINES, NULL AS QRD_QTY, NULL AS OPEN_LINES, NULL AS OPEN_AMT " & _
                "FROM PO_RECVR_HIST_LIN AS l " & _
                "LEFT JOIN PO_RECVR_HIST_CELL c ON c.RECVR_NO = l.RECVR_NO AND c.SEQ_NO = l.SEQ_NO " & _
                "WHERE l.PO_NO IS NULL AND l.PO_SEQ_NO IS NULL AND RECVR_DAT >= '" & minPOdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("ONORDER")
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("LOC_ID")
                If IsDBNull(oTest) Then oTest = "1"
                oTest = Replace(oTest, "'", "''")
                store = oTest
                xmlWriter.WriteElementString("LOCATION", oTest)
                oTest = rdr("PO_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                po = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("PO_NO", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("ITEM_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                sku = Replace(oTest, "'", "")
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    dim1 = CStr(oTest)
                    oTest = Replace(rdr("DIM_2_UPR"), "'", "''")
                    If Not IsDBNull(oTest) Then dim2 = CStr(oTest) Else dim2 = Nothing
                    oTest = Replace(rdr("DIM_3_UPR"), "'", "''")
                    If Not IsDBNull(oTest) Then dim3 = CStr(oTest) Else dim3 = Nothing
                    sku &= "~" & dim1 & "~" & dim2 & "~" & dim3
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("ORD_QTY")
                If IsDBNull(oTest) Then oTest = 0
                ordQty = CDec(oTest)
                xmlWriter.WriteElementString("ORD_QTY", oTest)
                oTest = rdr("QTY_RECVD")
                If IsDBNull(oTest) Then oTest = 0
                recvQty = CDec(oTest)
                xmlWriter.WriteElementString("QTY_RECVD", oTest)
                oTest = rdr("QTY_EXPECTED")
                If IsDBNull(oTest) Then oTest = 0
                expQty = CDec(oTest)
                xmlWriter.WriteElementString("EXP_QTY", expQty)
                totlQty += expQty
                oTest = rdr("ORD_COST")
                If IsDBNull(oTest) Then oTest = 0
                totlCost += CDec(oTest) * expQty
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("PRC_1")
                If IsDBNull(oTest) Then oTest = 0
                totlRetail += CDec(oTest) * expQty
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("LAST_RECVD_DATE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("LAST_RECVD_DATE", oTest)
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
                foundRow = tbl.Rows.Find(po)
                oTest = foundRow
                If IsNothing(foundRow) Then
                    pos += 1
                    row = tbl.NewRow
                    If Not IsDBNull(po) Then row("PO") = po Else row("PO") = "NA"
                    If Not IsDBNull(store) Then row("Str_Id") = store Else store = "NA"
                    oTest = rdr("ORD_DAT")
                    If IsDBNull(oTest) Then oTest = ""
                    row("OrdDate") = CStr(oTest)
                    oTest = rdr("DELIV_DAT")
                    If IsDBNull(oTest) Then oTest = ""
                    row("DueDate") = CStr(oTest)
                    oTest = rdr("CANCEL_DAT")
                    If IsDBNull(oTest) Then oTest = ""
                    row("CanDate") = CStr(oTest)
                    oTest = rdr("VEND_NO")
                    If IsDBNull(oTest) Then oTest = ""
                    row("Vendor") = CStr(Replace(oTest, "'", "''"))
                    oTest = rdr("BUYER")
                    If IsDBNull(oTest) Then oTest = ""
                    row("Buyer") = CStr(Replace(oTest, "'", "''"))
                    oTest = rdr("PO_STAT")
                    If IsDBNull(oTest) Then oTest = ""
                    row("Status") = CStr(Replace(oTest, "'", "''"))
                    oTest = rdr("AMT")
                    If IsDBNull(oTest) Then oTest = 0
                    row("AMT") = CDec(oTest)
                    oTest = rdr("RECVD_COST")
                    If IsDBNull(oTest) Then oTest = 0
                    row("RECVD_COST") = CDec(oTest)
                    oTest = rdr("LINES")
                    If IsDBNull(oTest) Then oTest = 0
                    row("LINES") = CInt(oTest)
                    oTest = rdr("STK_QTY")
                    If IsDBNull(oTest) Then oTest = 0
                    row("STK_QTY") = CDec(oTest)
                    oTest = rdr("OPEN_LINES")
                    If IsDBNull(oTest) Then oTest = 0
                    row("OPEN_LINES") = CInt(oTest)
                    oTest = rdr("OPEN_AMT")
                    If IsDBNull(oTest) Then oTest = 0
                    row("OPEN_AMT") = CDec(oTest)
                    tbl.Rows.Add(row)
                End If
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("ONORDER")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("PO", "")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("ORD_QTY", totlQty)                    ' only if still due ie expqty
            xmlWriter.WriteElementString("QTY_RECVD", "")
            xmlWriter.WriteElementString("EXP_QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)                      ' extended cost of what is still due
            xmlWriter.WriteElementString("RETAIL", totlRetail)                  ' extended retail of what is still due
            xmlWriter.WriteElementString("LAST_RECVD_DATE", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            If tbl.Rows.Count > 0 Then
                path = xmlPath & "\POHeader.xml"
                xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
                xmlWriter.WriteStartDocument(True)
                xmlWriter.Formatting = Formatting.Indented
                xmlWriter.Indentation = 5
                xmlWriter.WriteStartElement("PurchaseOrder")
                Dim c As Integer = tbl.Rows.Count
                For Each row In tbl.Rows
                    xmlWriter.WriteStartElement("ONORDER")
                    If IsDBNull(row("Str_Id")) Then row("Str_Id") = "NA"
                    xmlWriter.WriteElementString("LOCATION", row("Str_Id"))
                    If IsDBNull(row("PO")) Then row("PO") = "NA"
                    xmlWriter.WriteElementString("PO", row("PO"))
                    If IsDBNull(row("OrdDate")) Then row("OrdDate") = "NA"
                    xmlWriter.WriteElementString("ORDERDATE", row("OrdDate"))
                    If IsDBNull(row("DueDate")) Then row("DueDate") = "NA"
                    xmlWriter.WriteElementString("DUEDATE", row("DueDate"))
                    If IsDBNull(row("CanDate")) Then row("CanDate") = "NA"
                    xmlWriter.WriteElementString("CANCELDATE", row("CanDate"))
                    If IsDBNull(row("Vendor")) Then row("Vendor") = "NA"
                    xmlWriter.WriteElementString("VENDOR", row("Vendor"))
                    If IsDBNull(row("Buyer")) Then row("Buyer") = "NA"
                    xmlWriter.WriteElementString("BUYER", row("Buyer"))
                    If IsDBNull(row("Status")) Then row("Status") = "NA"
                    xmlWriter.WriteElementString("STATUS", row("Status"))
                    If IsDBNull(row("Amt")) Then row("Amt") = 0
                    xmlWriter.WriteElementString("AMT", row("AMT"))
                    If IsDBNull(row("RECVD_COST")) Then row("RECVD_COST") = 0
                    xmlWriter.WriteElementString("RECVD_COST", row("RECVD_COST"))
                    If IsDBNull(row("Lines")) Then row("Lines") = 0
                    xmlWriter.WriteElementString("LINES", row("LINES"))
                    If IsDBNull(row("STK_QTY")) Then row("Stk_QTY") = 0
                    xmlWriter.WriteElementString("ORD_QTY", row("STK_QTY"))
                    If IsDBNull(row("OPEN_LINES")) Then row("OPEN_LINES") = 0
                    xmlWriter.WriteElementString("OPEN_LINES", row("OPEN_LINES"))
                    If IsDBNull(row("OPEN_AMT")) Then row("OPEN_AMT") = 0
                    xmlWriter.WriteElementString("OPEN_AMT", row("OPEN_AMT"))
                    xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                    xmlWriter.WriteEndElement()
                Next
                '
                '               Write a row for totals
                '
                xmlWriter.WriteStartElement("ONORDER")
                xmlWriter.WriteElementString("LOCATION", "")
                xmlWriter.WriteElementString("PO", "TOTALS")
                xmlWriter.WriteElementString("ORDERDATE", c)
                xmlWriter.WriteElementString("DUEDATE", "")
                xmlWriter.WriteElementString("CANCELDATE", "")
                xmlWriter.WriteElementString("VENDOR", "")
                xmlWriter.WriteElementString("BUYER", "")
                xmlWriter.WriteElementString("STATUS", "")
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
                xmlWriter.WriteEndElement()
                xmlWriter.WriteEndDocument()
                xmlWriter.Close()
            End If
            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Purchase Orders", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            MsgBox(ex.Message)
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Purchase Order Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Adjustment data")
            Dim sku As String
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            path = xmlPath & "\Adjustment.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Adjustments")
            cpCon.Open()
            sql = "SELECT h.EVENT_NO AS TRANS_ID, h.SEQ_NO, h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.LOC_ID AS LOCATION, " & _
                "CASE WHEN c.QTY IS NOT NULL THEN c.QTY ELSE h.QTY * QTY_NUMER END AS QTY, COST, UNIT_RETL_VAL AS RETAIL, h.TRX_DAT AS DATE " & _
                "FROM IM_ADJ_HIST h " & _
                "LEFT JOIN IM_ADJ_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.BAT_ID = h.BAT_ID AND c.ITEM_NO = h.ITEM_NO " & _
                "AND c.LOC_ID = h.LOC_ID AND c.SEQ_NO = h.SEQ_NO " & _
                "WHERE h.LST_MAINT_DT >= '" & minADJdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                xmlWriter.WriteStartElement("ADJUSTMENT")
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY")
                If IsNumeric(oTest) Then qty = CDec(oTest)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", oTest)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("ADJUSTMENT")
            xmlWriter.WriteElementString("TRANS_ID", "ADJUSTMENT")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Adjustments", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Adjustment Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Physical Count data")
            cpCon.Open()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            path = xmlPath & "\Physical.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Physical")
            sql = "SELECT h.EVENT_NO AS TRANS_ID, 0 AS SEQ_NO, h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " & _
                "h.LOC_ID AS LOCATION, AVG_COST AS COST, UNIT_RETL_VAL AS RETAIL, POST_DAT AS DATE, " & _
                "c.CNT_QTY_1, c.FRZ_QTY_ON_HND, " & _
                "CASE WHEN c.QTY_CNTD IS NOT NULL THEN c.QTY_CNTD - c.FRZ_QTY_ON_HND ELSE QTY_ADJ END AS QTY " & _
                "FROM IM_CNT_HIST h " & _
                "LEFT JOIN IM_CNT_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO AND c.LOC_ID = h.LOC_ID " & _
                "WHERE ITEM_TYP = 'I' AND QTY_ADJ <> 0 AND POST_DAT >= '" & minADJdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            Dim sku As String
            Dim qty As Decimal = 0
            While rdr.Read
                cnt += 1
                xmlWriter.WriteStartElement("PHYSICAL")
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"

                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("SEQ_NO", 0)
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY")
                If IsNumeric(oTest) Then qty = CDec(oTest) Else qty = 0
                totlQty += qty
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                cost = 0
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                retail = 0
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = "1/1/1900"
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("PHYSICAL")
            xmlWriter.WriteElementString("TRANS_ID", "PHYSICAL")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Physical Adjustments", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collecting Physical Count Records " & cnt)
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Receipt data")
            cpCon.Open()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Dim sku As String
            path = xmlPath & "\Receipt.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Receipt")
            sql = "SELECT l.RECVR_NO AS TRANS_ID, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " & _
                 "l.PO_NO, l.PO_SEQ_NO, l.RECVR_LOC_ID AS LOCATION, ORD_DAT AS ORDER_DATE,DELIV_DAT AS EXPECTED_DATE, " & _
                 "l.RECVD_COST / l.QTY_RECVD_NUMER AS COST, l.UNIT_RETL_VAL / l.QTY_RECVD_NUMER AS RETAIL, l.RECVR_DAT AS DATE,  " & _
                 "CASE WHEN c.QTY_RECVD IS NOT NULL THEN c.QTY_RECVD ELSE l.QTY_RECVD * l.QTY_RECVD_NUMER END AS QTY " & _
                 "FROM PO_RECVR_HIST_LIN AS l " & _
                 "LEFT JOIN PO_RECVR_HIST_CELL c ON c.RECVR_NO = l.RECVR_NO AND c.SEQ_NO = l.SEQ_NO " & _
                 "INNER JOIN PO_RECVR_HIST AS h ON h.RECVR_NO = l.RECVR_NO " & _
                 "LEFT JOIN PO_ORD_HDR p ON p.PO_NO = l.PO_NO " & _
                 "WHERE l.QTY_RECVD > 0 AND ITEM_TYP = 'I' AND l.RECVR_DAT >= '" & minRCPTdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                xmlWriter.WriteStartElement("RECEIPT")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)

                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY")
                If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", oTest)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("RECEIPT")
            xmlWriter.WriteElementString("TRANS_ID", "RECEIPT")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Receipts", m, "")

        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Receipt Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Return data")
            cpCon.Open()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Dim sku As String
            path = xmlPath & "\Return.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Return")
            sql = "SELECT l.RTV_NO AS TRANS_ID, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, l.RTV_LOC_ID AS LOCATION, " & _
                "COST / ISNULL(l.QTY_RETD_NUMER,1) AS COST, " & _
                "(UNIT_RETL_VAL / ISNULL(QTY_RETD_NUMER,1)) AS RETAIL, l.RTV_DAT AS DATE, " & _
                "CASE WHEN c.QTY_RETD IS NOT NULL THEN c.QTY_RETD ELSE l.QTY_RETD END AS QTY " & _
                "FROM PO_RTV_HIST_LIN as l " & _
                "LEFT JOIN PO_RTV_HIST_CELL c ON c.RTV_NO = l.RTV_NO AND c.SEQ_NO = l.SEQ_NO " & _
                "INNER JOIN PO_RTV_HIST AS h ON h.RTV_NO = l.RTV_NO " & _
                "WHERE ITEM_TYP = 'I' AND l.LST_MAINT_DT >= '" & minRTNdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("RETURN")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("QTY")
                If IsNumeric(oTest) Then qty = CDec(oTest) Else qty = 0
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                        oTest = rdr("QTY")
                        If IsNumeric(oTest) Then qty = CDec(oTest) Else qty = 0
                        totlQty += qty
                    Else
                        totlQty += qty
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                '' xmlWriter.WriteElementString("OVERRIDE", "")
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("RETURN")
            xmlWriter.WriteElementString("TRANS_ID", "RETURN")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records "
            Call Update_Process_Log("1", "Extract Returns", m, "")
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Return Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Console.WriteLine("Extracting Sales data")
            cpCon.Open()
            path = xmlPath & "\Sales.xml"
            Dim item, itemString, dim1, dim2, dim3, store, drawer, location As String
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Sales")
            sql = "CREATE TABLE #tt1 (TRANS_ID varchar(30), SEQ_NO int, STORE varchar(10), ITEM varchar(90), DIM1 varchar(30) NULL, " & _
            "DIM2 varchar(30) NULL, DIM3 varchar(30) NULL, LOCATION varchar(10), DRAWER varchar(10) NULL,  " & _
            "QTY decimal(18,4) NULL, COST decimal(18,2) NULL, RETAIL decimal(18,2) NULL, DATE datetime, " & _
            "MARKDOWN decimal(18,2) NULL, MKDN_REASON varchar(30) NULL,COUPON_CODE varchar(30) NULL, DEPT varchar(10) NULL, " & _
            "CLASS varchar(10) NULL, BUYER varchar(10) NULL, CUST_NO varchar(30) NULL,TKT_NO varchar(30) NULL) " & _
            "INSERT INTO #tt1 (TRANS_ID,SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, " & _
           "QTY, COST, RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO) " & _
           "SELECT ISNULL(h.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, ISNULL(l.STR_ID,'NA') STORE, " & _
            "ISNULL(l.ITEM_NO,'NA') AS ITEM, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, l.STK_LOC_ID AS LOCATION, " & _
            "DRW_ID AS DRAWER,  CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) ELSE c.QTY_SOLD END AS QTY, " & _
            "ISNULL(CONVERT(DECIMAL(10,2),l.UNIT_COST),0) AS COST, ISNULL(CONVERT(DECIMAL(10,2),l.PRC),0) AS RETAIL,  " & _
            "ISNULL(h.TKT_DT,'1/1/1900') AS DATE, " & _
            "ISNULL(CONVERT(DECIMAL(10,2),l.PRC_1 - l.PRC),0) AS MARKDOWN, " & _
            " l.PRC_OVRD_REAS, NULL, ISNULL(l.CATEG_COD,'NA') AS DEPT, " & _
            "ISNULL(l.SUBCAT_COD,'NA') AS CLASS, ISNULL(pv.LST_ORD_BUYER,'OTHER') AS BUYER, " & _
            "CUST_NO, h.TKT_NO FROM PS_TKT_HIST_LIN AS l " & _
            "JOIN PS_TKT_HIST h ON h.DOC_ID = l.DOC_ID AND h.BUS_DAT = l.BUS_DAT " & _
            "LEFT JOIN PS_TKT_HIST_LIN_CELL c ON c.DOC_ID = l.DOC_ID and c.LIN_SEQ_NO = l.LIN_SEQ_NO " & _
            "LEFT JOIN PO_VEND_ITEM pv ON pv.VEND_NO = l.ITEM_VEND_NO AND pv.ITEM_NO = l.ITEM_NO " & _
            "WHERE h.TKT_TYP = 'T' AND l.LIN_TYP IN ('S','R') AND c.BUS_DAT >= '" & minSALdate & "' " & _
           "SELECT DOC_ID, LIN_SEQ_NO, PROMPT_ALPHA INTO #t2 FROM PS_TKT_HIST_LIN_PROMPT " & _
           "WHERE BUS_DAT >= '" & minSALdate & "' AND PROMPT_COD = 'MKTG_CODE' " & _
           "UPDATE t1 SET t1.COUPON_CODE = PROMPT_ALPHA FROM #tt1 t1 JOIN #t2 t2 ON t2.DOC_ID = t1.TRANS_ID AND t2.LIN_SEQ_NO = SEQ_NO"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
                '
                '    changed from h.TKT_DT to h.BUS_DAT because Counterpoint uses BUS_DAT in Flas Sales report
                '   
            sql = "SELECT TRANS_ID,SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, STORE, QTY, COST, " & _
                "RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO FROM #tt1"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt = 300000 Then
                    xmlWriter.WriteEndElement()
                    xmlWriter.WriteEndDocument()
                    xmlWriter.Close()
                    path = xmlPath & "\Sales" & files & ".xml"
                    xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
                    xmlWriter.WriteStartDocument(True)
                    xmlWriter.Formatting = Formatting.Indented
                    xmlWriter.Indentation = 2
                    xmlWriter.WriteStartElement("Sales")
                    cnt = 0
                    files += 1
                End If
                item = ""
                itemString = ""
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("SALE")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("STORE")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("STR_ID", Replace(oTest, "'", "''"))
                oTest = rdr("ITEM")
                If IsDBNull(oTest) Then oTest = "NA"
                itemString = oTest
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        itemString &= "~" & oTest & "~" & rdr("DIM2") & "~" & rdr("DIM3")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(itemString, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then location = "1" Else location = oTest
                xmlWriter.WriteElementString("LOCATION", Replace(location, "'", "''"))
                oTest = rdr("DRAWER")
                If IsDBNull(oTest) Then drawer = "1" Else drawer = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("DRAWER", drawer)
                oTest = rdr("QTY")
                If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                ''If cost = 0 Then
                ''    oTest = rdr("AVG_COST")
                ''    If Not IsDBNull(oTest) Then cost = CDec(oTest)
                ''End If
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("MARKDOWN")
                If IsDBNull(oTest) Then discount = 0 Else discount = CDec(oTest)
                totlMarkdown += (discount * qty)
                discountAmt = discount * qty
                xmlWriter.WriteElementString("MARKDOWN", discountAmt)
                oTest = rdr("DEPT")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("DEPT", Replace(oTest, "'", "''"))
                oTest = rdr("CLASS")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("CLASS", Replace(oTest, "'", "''"))
                oTest = rdr("BUYER")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("BUYER", Replace(oTest, "'", "''"))
                oTest = rdr("CUST_NO")
                If IsDBNull(oTest) Then oTest = "U"
                xmlWriter.WriteElementString("CUST_NO", Replace(oTest, "'", "''"))
                oTest = rdr("TKT_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TKT_NO", oTest)
                oTest = rdr("MKDN_REASON")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("MKDN_REASON", Replace(oTest, "'", "''"))
                oTest = rdr("COUPON_CODE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("COUPON_CODE", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
                '
                '               Write a row for totals
                '
            xmlWriter.WriteStartElement("SALE")
            xmlWriter.WriteElementString("TRANS_ID", "SALES")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("STORE", "")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("DRAWER", "")
            xmlWriter.WriteElementString("STORE", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteElementString("MARKDOWN", totlMarkdown)
            xmlWriter.WriteElementString("DEPT", "")
            xmlWriter.WriteElementString("CLASS", "")
            xmlWriter.WriteElementString("BUYER", "")
            xmlWriter.WriteElementString("CUST_NO", "")
            xmlWriter.WriteElementString("TKT_NO", "")
            xmlWriter.WriteElementString("MKDN_REASON", "")
            xmlWriter.WriteElementString("COUPON_CODE", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Sales", m, "")
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            Console.WriteLine(ex.Message & " " & ex.StackTrace)
            Console.ReadLine()

            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Sale Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

100:    Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            files = 1
            Console.WriteLine("Extracting Transfer data")
            cpCon.Open()
            path = xmlPath & "\Transfer.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Transfer")
            sql = "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, " & _
                "c.DIM_2_UPR, c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " & _
                "CASE WHEN c.XFER_QTY IS NULL THEN l.XFER_QTY_NUMER * l.XFER_QTY ELSE c.XFER_QTY END AS QTY_OUT, l.xfer_qty_numer, " & _
                "CAST(l.FROM_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2)) AS RETAIL, " & _
                "CAST(l.FROM_UNIT_COST / l.XFER_QTY_NUMER AS DECIMAL(8,2)) AS COST," & _
                " h.SHIP_DAT AS Date_Out FROM IM_XFER_LIN AS l WITH (NOLOCK) " & _
                "INNER JOIN IM_XFER_HDR as h ON h.XFER_NO = l.XFER_NO " & _
                "LEFT JOIN IM_XFER_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_LIN_SEQ_NO = l.XFER_LIN_SEQ_NO " & _
                "WHERE h.LST_MAINT_DT >= '" & minXFERdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Dim sku As String
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("TRANSFER")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("FROM_LOC_ID")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY_OUT")
                If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest) * -1
                totlQty += qty
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(oTest)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("DATE_OUT")
                If IsDBNull(oTest) Then
                    If IsDBNull(oTest) Then oTest = CDate("1900-01-01")
                End If
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()

            cpCon.Open()
            sql = "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_IN_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, " & _
                "c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " & _
                "CASE WHEN c.QTY_RECVD IS NULL THEN l.XFER_QTY_NUMER * l.QTY_RECVD ELSE c.QTY_RECVD END AS QTY, l.xfer_qty_numer, " & _
                "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2)) " & _
                "ELSE CAST(TO_UNIT_RETL_VAL / l.XFER_QTY_NUMER AS DECIMAL) END AS RETAIL, " &
                "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_COST / (l.XFER_QTY_NUMER) AS DECIMAL(8,2)) " & _
                "ELSE CAST(TO_EXT_COST / (l.QTY_RECVD * l.XFER_QTY_NUMER) AS DECIMAL(8,2)) END AS COST, " & _
                "RECVD_DAT AS Date_In FROM IM_XFER_IN_HIST_LIN AS l WITH (NOLOCK) " & _
                "INNER JOIN IM_XFER_IN_HIST as h ON h.XFER_NO = l.XFER_NO AND h.EVENT_NO = l.EVENT_NO " & _
                "LEFT JOIN IM_XFER_IN_HIST_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_IN_LIN_SEQ_NO = l.XFER_IN_LIN_SEQ_NO " & _
                "WHERE h.LST_MAINT_DT >= '" & minXFERdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("TRANSFER")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("TO_LOC_ID")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY")
                If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = CDec(oTest)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = CDec(retail)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("DATE_IN")
                If IsDBNull(oTest) Then oTest = CDate("1900-01-01")
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("TRANSFER")
            xmlWriter.WriteElementString("TRANS_ID", "TRANSFER")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", Date.Now)
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Transfers", m, "")
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Transfer Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try
    End Sub

    Private Sub Update_Process_Log(ByVal modul As String, ByVal process As String, ByVal m As String, ByVal stat As String)
        ''     We're not connected to a local Retail Clarity database
        Exit Sub
        Try
            con.Open()
            StopWatch.Stop()
            Dim thisDateTime As DateTime = CDate(Now)
            Dim ts As TimeSpan = StopWatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim pgm As String = "Daily_XML_Extract"
            sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
               "SELECT '" & thisDateTime & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            Dim el As New CounterpointXMLExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Update Message (Collected Inventory Records)")
            If cpCon.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Function FixData(ByVal strng As String) As String
        If String.IsNullOrEmpty(strng.ToString) Then strng = "NA"
        Dim result As String
        result = Microsoft.VisualBasic.Replace(strng, ",", "|")
        FixData = result
    End Function
End Module

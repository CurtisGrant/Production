﻿Imports System
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
        StopWatch.Start()
        Dim thisDate As Date = CDate(Date.Today)
        Dim minADJdate, minORDdate, minRCPTdate, minRTNdate, minSALdate, minXFERdate, minPOdate, thisEdate, thisSdate As Date
        Dim todaysDate As Date = Date.Today
        Dim DateTimeString As String = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
        Dim ADJwks As Integer = 2
        Dim ORDwks As Integer = 2
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
            If fld = "ORDWKS" And IsNumeric(valu) Then ORDwks = CInt(valu)
        End While

        thisSdate = Date.Today.AddDays(0 - Date.Today.DayOfWeek)
        thisEdate = Date.Today.AddDays(6 - Date.Today.DayOfWeek)

        minADJdate = DateAdd(DateInterval.WeekOfYear, ADJwks * -1, thisSdate)
        minORDdate = DateAdd(DateInterval.WeekOfYear, ORDwks * -1, thisSdate)
        minRCPTdate = DateAdd(DateInterval.WeekOfYear, RCPTwks * -1, thisSdate)
        minRTNdate = DateAdd(DateInterval.WeekOfYear, RTNwks * -1, thisSdate)
        minSALdate = DateAdd(DateInterval.WeekOfYear, SALwks * -1, thisSdate)
        minXFERdate = DateAdd(DateInterval.WeekOfYear, XFERwks * -1, thisSdate)
        minPOdate = DateAdd(DateInterval.WeekOfYear, POwks * -1, thisDate)

        conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & "; Integrated Security=True"
        cpCon = New SqlConnection(conString)

        errorPath = xmlPath                            ' force the error log to write to the same folder as the xmls











        ''    GoTo 100








        '               Delete errlog first thing
        Try
            Dim thisDay As Integer = Date.Now.DayOfWeek
            path = xmlPath & "\errlog.txt"
            If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Delete errlog")
        End Try
        Try
            cnt = 0
            If (Not System.IO.Directory.Exists(xmlPath)) Then
                System.IO.Directory.CreateDirectory(xmlPath)
            End If
            Console.WriteLine("Extracting Buyers")
            Dim id, desc As String
            Dim lastUpdate As DateTime
            path = xmlPath & "\Buyers.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Buyers")
            Dim constr2 As String = ""
            Dim stat As String = ""
            cpCon.Open()
            sql = "SELECT USR_ID, NAM_UPR, LST_MAINT_DT LAST_UPDATE, LST_MAINT_USR_ID LAST_USER FROM SY_USR WHERE IS_BUYER = 'Y'"
            cmd = New SqlCommand(sql, cpCon)

            If Err.Number <> 0 Then Console.WriteLine(Err.Number)

            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                oTest = rdr("USR_ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = rdr("NAM_UPR")
                    If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                    oTest = rdr("LAST_UPDATE")
                    If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                    xmlWriter.WriteStartElement("Buyer")
                    xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                    xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                    xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                    oTest = rdr("LAST_USER")
                    If IsDBNull(oTest) Then oTest = Nothing
                    xmlWriter.WriteElementString("LAST_USER", oTest)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Buyers")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Calendar")
            path = xmlPath & "\Calendar.xml"
            cpCon.Open()
            sql = "SELECT * FROM SY_CALNDR WHERE CALNDR_ID = 2017"
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, dept, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT SUBCAT_COD AS ID, CATEG_COD AS DEPT, DESCR, LST_MAINT_DT LAST_UPDATE, " &
            "LST_MAINT_USR_ID LAST_USER FROM IM_SUBCAT_COD " &
            "WHERE SUBCAT_COD Is Not NULL"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = Nothing
                dept = rdr("Dept")
                desc = rdr("DESCR")
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUpdate = Nothing
                xmlWriter.WriteStartElement("Class")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DEPT", Replace(dept, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT CATEG_COD AS ID, DESCR, LST_MAINT_DT LAST_UPDATE, " &
                "LST_MAINT_USR_ID LAST_USER FROM IM_CATEG_COD " &
                "WHERE CATEG_COD Is Not NULL"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA'"
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteStartElement("Department")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT PROF_COD AS ID, DESCR, LST_MAINT_DT LAST_UPDATE, LST_MAINT_USR_ID LAST_USER FROM IM_ITEM_PROF_COD " &
            "WHERE PROF_COD Is Not NULL " &
                "And VAL_3 ='Y'"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteStartElement("ProductLine")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Product Lines")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Seasons")
            path = xmlPath & "\Seasons.xml"
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Seasons")
            cpCon.Open()
            sql = "SELECT DISTINCT PROF_COD AS ID, DESCR, LST_MAINT_DT LAST_UPDATE, LST_MAINT_USR_ID LAST_USER FROM IM_ITEM_PROF_COD " &
            "WHERE PROF_COD Is Not NULL " &
                "And VAL_1 ='Y'"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteStartElement("Season")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT LOC_ID AS ID, DESCR_UPR, LST_MAINT_DT LAST_UPDATE, LST_MAINT_USR_ID LAST_USER FROM IM_LOC " &
            "WHERE LOC_ID Is Not NULL "
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader

            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("DESCR_UPR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteStartElement("Location")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT s.STR_ID AS ID, DESCR_UPR, STK_LOC_ID, s.LST_MAINT_DT LAST_UPDATE, " &
                "s.LST_MAINT_USR_ID LAST_USER FROM PS_STR s " &
                "JOIN PS_STR_CFG_PS c ON c.STR_ID = s.STR_ID WHERE STK_LOC_ID Is Not NULL "
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
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT DISTINCT VEND_NO AS ID, NAM_UPR, LST_MAINT_DT LAST_UPDATE, " &
                "LST_MAINT_USR_ID LAST_USER FROM VI_PO_VEND_WITH_ADDRESS " &
                "WHERE VEND_NO Is Not NULL "
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                oTest = rdr("ID")
                If Not IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                oTest = rdr("NAM_UPR")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                oTest = rdr("LAST_USER")
                If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                xmlWriter.WriteStartElement("Vendor")
                xmlWriter.WriteElementString("ID", Replace(id, "'", "''"))
                xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim id, desc, lastUser As String
            Dim lastUpdate As DateTime
            cpCon.Open()
            sql = "SELECT REAS_COD, DESCR, LST_MAINT_DT LAST_UPDATE, LST_MAINT_USR_ID LAST_USER FROM PS_REAS_COD ORDER BY REAS_COD"
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
                    oTest = rdr("LAST_UPDATE")
                    If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                    oTest = rdr("LAST_USER")
                    If Not IsDBNull(oTest) Then lastUser = CStr(oTest) Else lastUser = Nothing
                    xmlWriter.WriteElementString("DESCRIPTION", Replace(desc, "'", "''"))
                    xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                    xmlWriter.WriteElementString("LAST_USER", lastUser)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim custom As String = Nothing
            Dim custom2 As String = Nothing
            Dim custom3 As String = Nothing
            Dim custom4 As String = Nothing
            Dim custom5 As String = Nothing
            cpCon.Open()
            sql = "SELECT CUST_NO, NAM_UPR AS NAME, FST_NAM_UPR AS FIRST_NAME, LST_NAM_UPR AS LAST_NAME, ADRS_1 AS ADDRESS_1, " &
                "ADRS_2 AS ADDRESS_2, ADRS_3 AS ADDRESS_3, CITY, STATE, ZIP_COD AS ZIP, CONTCT_2 AS SPOUSE, COMMNT AS BIRTHDAY, " &
                "PHONE_1, PHONE_2, EMAIL_ADRS_1 AS EMAIL, CATEG_COD AS TYPE, BAL AS BALANCE, " &
                "FST_SAL_DAT AS FIRST_SALE_DATE, LST_SAL_DAT AS LAST_SALE_DATE, LST_MAINT_DT AS LAST_UPDATE, LST_MAINT_USR_ID LAST_USER, " &
                "LOY_PTS_BAL AS LOYALTY_POINTS, INCLUDE_IN_MARKETING_MAILOUTS AS OK_TO_EMAIL, PROF_ALPHA_2 AS OK_TO_MAIL, " &
                "MBL_PHONE_1 AS CELL_1, MBL_PHONE_2 AS CELL_2 FROM VI_AR_CUST_WITH_ADDRESS"
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
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
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
                xmlWriter.WriteElementString("CUSTOM_1", custom)
                xmlWriter.WriteElementString("CUSTOM_2", custom2)
                xmlWriter.WriteElementString("CUSTOM_3", custom3)
                xmlWriter.WriteElementString("CUSTOM_4", custom4)
                xmlWriter.WriteElementString("CUSTOM_5", custom4)
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

            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Customer Data")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Console.WriteLine("Extracting Item data")
            path = xmlPath & "\Items.xml"
            Dim sku As String
            Dim lastUpdate As DateTime
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Items")
            cpCon.Open()
            sql = "SELECT i.ITEM_NO, g.DIM_1_UPR AS DIM1, g.DIM_2_UPR AS DIM2, g.DIM_3_UPR AS DIM3, " &
                "DisplayLabel1 AS DIM1_DESCR, DisplayLabel2 AS DIM2_DESCR, DisplayLabel3 AS DIM3_DESCR, i.DESCR AS DESCRIPTION, " &
                "i.VEND_ITEM_NO, pv.VEND_ITEM_NO AS povItem, i.ITEM_VEND_NO AS VENDOR_ID, " &
                "CASE WHEN LEN(i.ITEM_VEND_NO) > 0 THEN i.ITEM_NO ELSE pv.VEND_NO END AS VENDOR_ID, p.NAM AS VENDOR, " &
                "CASE WHEN LEN(pv.LST_ORD_BUYER) > 0 THEN pv.LST_ORD_BUYER ELSE p.CATEG_COD END AS BUYER, " &
                "i.CATEG_COD AS DEPT, i.SUBCAT_COD AS CLASS, i.PRC_1 AS CURR_RTL, CONVERT(Decimal(18,4),pv.UNIT_COST / pv.PUR_NUMER) AS CURR_COST, " &
                "CASE WHEN i.USR_LB_MERCHANT_SKU IS NOT NULL THEN i.USR_LB_MERCHANT_SKU ELSE gv.LB_MERCHANT_SKU END AS CUSTOM1, " &
                "i.PROF_COD_3 AS CUSTOM3, i.PROF_COD_5 AS CUSTOM5, i.ALT_1_UNIT AS UOM, i.ALT_1_NUMER AS BUYUNIT, " &
                "i.ALT_1_DENOM AS SELLUNIT, imn.NOTE_TXT, i.STAT AS STATUS, i.ITEM_TYP AS TYPE, i.STK_UNIT, pv.MIN_ORD_QTY, " &
                "ORD_MULT, i.LST_MAINT_DT LAST_UPDATE, i.LST_MAINT_USR_ID LAST_USER FROM VI_IM_ITEM_WITH_INV_TOTS AS i " &
                "LEFT JOIN VI_IM_ITEM_GRID_DIMS g ON g.ITEM_NO = i.ITEM_NO " &
                "LEFT JOIN USR_VI_MAG_ITEM_GRID_VAL gv ON i.ITEM_NO = gv.ITEM_NO And g.DIM_1_UPR = gv.DIM_1_UPR " &
                "And g.DIM_2_UPR = gv.DIM_2_UPR And g.DIM_3_UPR = gv.DIM_3_UPR " &
                "LEFT JOIN VI_PO_VEND_WITH_ADDRESS AS p ON p.VEND_NO = i. ITEM_VEND_NO " &
                "LEFT JOIN PO_VEND_ITEM AS pv ON pv.ITEM_NO = i.ITEM_NO And i.ITEM_VEND_NO = pv.VEND_NO " &
                "LEFT JOIN IM_ITEM_NOTE AS imn ON imn.ITEM_NO =i.ITEM_NO And NOTE_ID = 'PO'"

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
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("BUYER", Replace(oTest, "'", "''"))

                oTest = rdr("DEPT")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("DEPT", Replace(oTest, "'", "''"))

                oTest = rdr("CLASS")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("CLASS", Replace(oTest, "'", "''"))

                oTest = rdr("CURR_RTL")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("CURR_RTL", Replace(oTest, "'", "''"))

                oTest = rdr("CURR_COST")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("CURR_COST", Replace(oTest, "'", "''"))

                oTest = rdr("CUSTOM1")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CUSTOM_1", Replace(oTest, "'", "''"))

                xmlWriter.WriteElementString("CUSTOM_2", "")

                oTest = rdr("CUSTOM3")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CUSTOM_3", Replace(oTest, "'", "''"))

                xmlWriter.WriteElementString("CUSTOM_4", "")

                oTest = rdr("CUSTOM5")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CUSTOM_5", Replace(oTest, "'", "''"))

                oTest = rdr("UOM")
                If IsDBNull(oTest) Then oTest = "EA"
                xmlWriter.WriteElementString("UOM", Replace(oTest, "'", "''"))

                oTest = rdr("BUYUNIT")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("BUYUNIT", oTest)

                oTest = rdr("SELLUNIT")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("SELLUNIT", oTest)

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

                oTest = rdr("STK_UNIT")
                If IsDBNull(oTest) Then oTest = "EA"
                xmlWriter.WriteElementString("STOCK_UNIT", Replace(oTest, "'", "''"))

                oTest = rdr("MIN_ORD_QTY")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("MIN_ORDER_QTY", oTest)

                oTest = rdr("ORD_MULT")
                If IsDBNull(oTest) Then oTest = 1
                xmlWriter.WriteElementString("ORD_MULT", oTest)

                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteStartElement("ITEM")
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
            xmlWriter.WriteElementString("CUSTOM_1", "")
            xmlWriter.WriteElementString("CUSTOM_2", "")
            xmlWriter.WriteElementString("CUSTOM_3", "")
            xmlWriter.WriteElementString("CUSTOM_4", "")
            xmlWriter.WriteElementString("CUSTOM_5", "")
            xmlWriter.WriteElementString("UOM", "")
            xmlWriter.WriteElementString("BUYUNIT", "")
            xmlWriter.WriteElementString("SELLUNIT", "")
            xmlWriter.WriteElementString("STATUS", "")
            xmlWriter.WriteElementString("NOTE", "")
            xmlWriter.WriteElementString("TYPE", "")
            xmlWriter.WriteElementString("STOCK_UNIT", "")
            xmlWriter.WriteElementString("MIN_ORDER_QTY", "")
            xmlWriter.WriteElementString("ORD_MULT", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Items", m, "")

        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Item Data")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch.Start()
            cnt = 0
            Dim sku, item, barcod_id, barcod As String
            Dim lastUpdate As DateTime
            Console.WriteLine("Extracting Barcode data")
            path = xmlPath & "\Barcodes.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Barcodes")
            cpCon.Open()
            sql = "SELECT ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, BARCOD, BARCOD_ID, LST_MAINT_DT LAST_UPDATE FROM IM_BARCOD"
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

                oTest = rdr("BARCOD")
                If IsDBNull(oTest) Then barcod = "NA" Else barcod = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("BARCOD", barcod)

                oTest = rdr("BARCOD_ID")
                If IsDBNull(oTest) Then barcod_id = "NA" Else barcod_id = Replace(oTest, "'", "''")
                xmlWriter.WriteElementString("BARCOD_ID", barcod_id)

                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteStartElement("SKU")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("ITEM", "")
            xmlWriter.WriteElementString("DIM1", "")
            xmlWriter.WriteElementString("DIM2", "")
            xmlWriter.WriteElementString("DIM3", "")
            xmlWriter.WriteElementString("BARCOD", cnt)
            xmlWriter.WriteElementString("BARCOD_ID", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", "")
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            path = xmlPath & "\Inventory.xml"
            Dim sku As String
            Dim lastUpdate As DateTime
            Dim path2 As String = xmlPath & "Inventory_" & Today & ".xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Inventory")
            ''sql = "SELECT i.LOC_ID AS LOCATION, i.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " & _
            ''    "CASE WHEN c.QTY_ON_HND IS NULL THEN i.QTY_ON_HND ELSE c.QTY_ON_HND END AS ONHAND, " & _
            ''    "CASE WHEN c.QTY_AVAIL IS NULL THEN i.QTY_AVAIL ELSE c.QTY_AVAIL END AS AVAIL, " & _
            ''    "CASE WHEN c.QTY_COMMIT IS NULL THEN i.QTY_COMMIT ELSE c.QTY_COMMIT END AS COMMITED, " & _
            ''    "AVG_COST AS COST, PRC_1 AS RETAIL INTO #t1 FROM IM_INV i " & _
            ''    "LEFT JOIN IM_INV_CELL c ON c.ITEM_NO = i.ITEM_NO AND c.LOC_ID = i.LOC_ID " & _
            ''    "LEFT JOIN IM_ITEM im ON IM.ITEM_NO=i.ITEM_NO " & _
            ''    "WHERE i.QTY_AVAIL <> 0 " & _
            ''    "DELETE FROM #t1 WHERE AVAIL = 0 " & _
            ''    "SELECT LOCATION, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, ONHAND, AVAIL, COMMITED, COST, RETAIL from #t1"
            sql = "SELECT i.LOC_ID AS LOCATION, i.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " &
                 "c.QTY_ON_HND ONHAND, AVG_COST AS COST, PRC_1 AS RETAIL, COALESCE(c.LST_MAINT_DT, " &
                "i.LST_MAINT_DT) LAST_UPDATE, c.LST_MAINT_USR_ID LAST_USER INTO #t1 FROM IM_INV i " &
                 "LEFT JOIN IM_INV_CELL c ON c.ITEM_NO = i.ITEM_NO And c.LOC_ID = i.LOC_ID " &
                 "LEFT JOIN IM_ITEM im ON IM.ITEM_NO=i.ITEM_NO " &
                 "WHERE im.TRK_METH = 'G' AND c.QTY_ON_HND <> 0 " &
                 "INSERT INTO #t1(LOCATION, ITEM_NO, ONHAND, COST, RETAIL, LAST_UPDATE, LAST_USER) " &
                 "SELECT i.LOC_ID, i.ITEM_NO, i.QTY_ON_HND, AVG_COST, PRC_1, i.LST_MAINT_DT, i.LST_MAINT_USR_ID FROM IM_INV i " &
                 "LEFT JOIN IM_ITEM im ON im.ITEM_NO = i.ITEM_NO " &
                 "WHERE im.TRK_METH <> 'G' AND i.QTY_ON_HND <> 0 " &
                 "SELECT LOCATION, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, ONHAND, COST, RETAIL, LAST_UPDATE, LAST_USER from #t1"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("ONHAND")
                If IsDBNull(oTest) Then qty = 0 Else qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                xmlWriter.WriteStartElement("ITEM")
                totlQty += qty
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
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Inventory Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try


        ''Stop


        Try
            cnt = 0
            totlRetail = 0
            Dim amt As Decimal
            Dim preq, batch, vend_id, vend, loc, buyer, alloc, mrg, terms, custom, custom2, custom3, custom4, custom5 As String
            Dim ord, del, can As Date
            Dim lastUpdate As DateTime
            path = xmlPath & "\Purchase_Request_Header.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseRequestHeader")
            Console.WriteLine("extracting Purchase Requests Header")
            cpCon.Open()
            sql = "SELECT PREQ_NO, BAT_ID, VEND_NO, VEND_NAM, LOC_ID, BUYER, ORD_DAT, DELIV_DAT, CANCEL_DAT, ORD_TOT, " &
                "IS_ALLOC, ALLOC_SEP_OR_MERGED, TERMS_COD, COMMNT_1, LST_MAINT_DT LAST_UPDATE FROM PO_PREQ_HDR"
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
                custom = Nothing
                custom2 = Nothing
                custom3 = Nothing
                custom4 = Nothing
                custom5 = Nothing
                amt = 0
                terms = Nothing
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
                oTest = rdr("TERMS_COD")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then terms = CStr(oTest)
                oTest = rdr("COMMNT_1")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then custom = CStr(Replace(oTest, "'", "''"))
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
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
                xmlWriter.WriteElementString("TERMS", terms)
                xmlWriter.WriteElementString("CUSTOM_1", custom)
                xmlWriter.WriteElementString("CUSTOM_2", custom2)
                xmlWriter.WriteElementString("CUSTOM_3", custom3)
                xmlWriter.WriteElementString("CUSTOM_4", custom4)
                xmlWriter.WriteElementString("CUSTOM_5", custom5)
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
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
            xmlWriter.WriteElementString("TERMS", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Purchase Request Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            Console.WriteLine("Extracting Purchase Request Details")
            cnt = 0
            totlCost = 0
            totlQty = 0
            Dim sku As String = ""
            Dim lastUpdate As DateTime
            path = xmlPath & "\Purchase_Request_Detail.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseRequestDetail")
            cpCon.Open()
            sql = "SELECT l.PREQ_NO, l.SEQ_NO, ITEM_NO, DESCR_UPR, ORD_UNIT, ORD_QTY_NUMER, ORD_QTY_DENOM, " &
                "ORD_COST, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, C.ORD_QTY, " &
                "COALESCE(c.LST_MAINT_DT,l.LST_MAINT_DT) AS LAST_UPDATE FROM PO_PREQ_LIN l " &
                "LEFT JOIN PO_PREQ_CELL C ON c.PREQ_NO = l.PREQ_NO And c.SEQ_NO = l.SEQ_NO"
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
                cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
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
                qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlQty += qty
                totlCost += (cost * qty)
                xmlWriter.WriteElementString("QTY", oTest)
                oTest = rdr("LAST_UPDATE")
                If Not IsDBNull(oTest) Then lastUpdate = CDate(oTest) Else lastUpdate = "1900-01-01 00:00:00"
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim po, store, sku, dim1, dim2, dim3, terms, comment As String
            Dim lastUpdate As DateTime
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
            tbl.Columns.Add("Terms", GetType(System.String))
            tbl.Columns.Add("Custom_1", GetType(System.String))
            tbl.Columns.Add("Custom_2", GetType(System.String))
            tbl.Columns.Add("Custom_3", GetType(System.String))
            tbl.Columns.Add("Custom_4", GetType(System.String))
            tbl.Columns.Add("Custom_5", GetType(System.String))
            tbl.Columns.Add("LAST_UPDATE", GetType(System.String))
            path = xmlPath & "\PODetail.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("PurchaseOrder")
            sql = "SELECT h.LOC_ID, h.PO_NO, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " &
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.ORD_QTY ELSE l.ORD_QTY * ORD_QTY_NUMER END AS ORD_QTY, " &
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.QTY_RECVD ELSe l.QTY_RECVD * ORD_QTY_NUMER END AS QTY_RECVD, " &
                "CASE WHEN c.PO_NO IS NOT NULL THEN c.QTY_EXPECTD ELSE l.QTY_EXPECTD * ORD_QTY_NUMER END AS QTY_EXPECTED, " &
                "h.ORD_DAT, h.DELIV_DAT, h.CANCEL_DAT, ORD_COST, PRC_1, h.VEND_NO, BUYER, PO_STAT, " &
                "(SELECT MAX(RECVR_DAT) FROM PO_RECVR_HIST_LIN r WHERE r.PO_NO = h.PO_NO " &
                "AND r.RECVR_LOC_ID = h.LOC_ID AND r.ITEM_NO = l.ITEM_NO AND QTY_RECVD > 0) AS LAST_RECVD_DATE, " &
                "h.ORD_SUB_TOT AS AMT, h.RECVD_TOT_COST AS RECVD_COST, h.LIN_CNT AS LINES, h.ORD_QTY_IN_STK_UNITS AS STK_QTY, " &
                "OPN_LIN_CNT AS OPEN_LINES, h.OPN_PO_TOT AS OPEN_AMT, TERMS_COD, h.COMMNT_1, " &
                "COALESCE(c.LST_MAINT_DT,l.LST_MAINT_DT) AS LAST_UPDATE FROM PO_ORD_HDR h " &
                "JOIN PO_ORD_LIN l ON l.PO_NO = h.PO_NO " &
                "LEFT JOIN PO_ORD_CELL c ON c.PO_NO = l.PO_NO AND c.SEQ_NO = l.SEQ_NO " &
                "JOIN IM_ITEM i ON i.ITEM_NO = l.ITEM_NO " &
                "WHERE ORD_DAT >= '" & minPOdate & "' "
            ''"UNION " & _
            ''"SELECT RECVR_LOC_ID AS Loc_Id, 'X'+l.RECVR_NO AS PO_NO, l.SEQ_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, 0 AS ORD_QTY, " & _
            ''"CASE WHEN c.RECVR_NO IS NOT NULL THEN c.QTY_RECVD ELSE l.QTY_RECVD * QTY_RECVD_NUMER END AS QTY_RECVD, " & _
            ''"l.NEW_QTY_EXPECTD * QTY_RECVD_NUMER, RECVR_DAT AS ORD_DAT, RECVR_DAT AS DELIV_DAT, " & _
            ''"RECVR_DAT AS CANCEL_DAT, RECVD_COST / QTY_RECVD_NUMER AS ORD_COST, UNIT_RETL_VAL / QTY_RECVD_NUMER AS PRC_1, " & _
            ''"l.ITEM_VEND_NO AS VEND_NO, '' AS BUYER, '' AS PO_STAT, RECVR_DAT AS LAST_RECVD_DAT, " & _
            ''"NULL AS AMT, NULL AS RECVD_COST, NULL AS LINES, NULL AS QRD_QTY, NULL AS OPEN_LINES, NULL AS OPEN_AMT " & _
            ''"FROM PO_RECVR_HIST_LIN AS l " & _
            ''"LEFT JOIN PO_RECVR_HIST_CELL c ON c.RECVR_NO = l.RECVR_NO AND c.SEQ_NO = l.SEQ_NO " & _
            ''"WHERE l.PO_NO IS NULL AND l.PO_SEQ_NO IS NULL AND RECVR_DAT >= '" & minPOdate & "'"
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
                xmlWriter.WriteElementString("LOCATION", store)
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
                ordQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlQty += ordQty
                xmlWriter.WriteElementString("ORD_QTY", ordQty)
                oTest = rdr("QTY_RECVD")
                If IsDBNull(oTest) Then oTest = 0
                recvQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                xmlWriter.WriteElementString("QTY_RECVD", recvQty)
                oTest = rdr("QTY_EXPECTED")
                If IsDBNull(oTest) Then oTest = 0
                expQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                xmlWriter.WriteElementString("EXP_QTY", expQty)
                oTest = rdr("ORD_COST")
                If IsDBNull(oTest) Then oTest = 0
                totlCost += Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero) * ordQty
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("PRC_1")
                If IsDBNull(oTest) Then oTest = 0
                totlRetail += Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero) * ordQty
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("LAST_RECVD_DATE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("LAST_RECVD_DATE", oTest)
                oTest = rdr("LAST_UPDATE")
                If IsDBNull(oTest) Then lastUpdate = "1900-01-01 00:00:00" Else lastUpdate = CDate(oTest)
                xmlWriter.WriteElementString("LAST_UPDATE", Format(lastUpdate, "yyyy-MM-dd HH:mm:ss"))
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
                    oTest = rdr("TERMS_COD")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row("Terms") = oTest
                    oTest = rdr("COMMNT_1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("Custom_1") = Replace(oTest, "'", "''")
                    Else : row("Custom_1") = Nothing
                    End If
                    row("Custom_2") = Nothing
                    row("Custom_3") = Nothing
                    row("Custom_4") = Nothing
                    row("Custom_5") = Nothing
                    row("LAST_UPDATE") = lastUpdate
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
            xmlWriter.WriteElementString("EXP_QTY", "")
            xmlWriter.WriteElementString("COST", totlCost)                      ' extended cost of what is still due
            xmlWriter.WriteElementString("RETAIL", totlRetail)                  ' extended retail of what is still due
            xmlWriter.WriteElementString("LAST_RECVD_DATE", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteElementString("TERMS", "")
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
                    If IsDBNull(row("Terms")) Then row("Terms") = ""
                    xmlWriter.WriteElementString("TERMS", row("Terms"))
                    If IsDBNull(row("CUSTOM_1")) Then row("Custom_1") = ""
                    If IsDBNull(row("CUSTOM_2")) Then row("Custom_2") = ""
                    If IsDBNull(row("CUSTOM_3")) Then row("Custom_3") = ""
                    If IsDBNull(row("CUSTOM_4")) Then row("Custom_4") = ""
                    If IsDBNull(row("CUSTOM_5")) Then row("Custom_5") = ""
                    xmlWriter.WriteElementString("CUSTOM_1", row("Custom_1"))
                    xmlWriter.WriteElementString("CUSTOM_2", row("Custom_2"))
                    xmlWriter.WriteElementString("CUSTOM_3", row("Custom_3"))
                    xmlWriter.WriteElementString("CUSTOM_4", row("Custom_4"))
                    xmlWriter.WriteElementString("CUSTOM_5", row("Custom_5"))
                    xmlWriter.WriteElementString("LAST_UPDATE", row("LAST_UPDATE"))
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            ''sql = "SELECT h.EVENT_NO AS TRANS_ID, h.SEQ_NO, h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.LOC_ID As LOCATION, " &
            ''    "CASE WHEN c.QTY Is Not NULL THEN c.QTY ELSE h.QTY * QTY_NUMER END AS QTY, cost, UNIT_RETL_VAL As RETAIL, h.TRX_DAT AS DATE " &
            ''    "FROM IM_ADJ_HIST h " &
            ''    "LEFT JOIN IM_ADJ_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO And c.BAT_ID = h.BAT_ID And c.ITEM_NO = h.ITEM_NO " &
            ''    "And c.LOC_ID = h.LOC_ID And c.SEQ_NO = h.SEQ_NO " &
            ''    "WHERE h.LST_MAINT_DT >= '" & minADJdate & "'"
            sql = "create table #t1(TRANS_ID varchar(30), SEQ_NO integer, ITEM_NO varchar(30), DIM_1_UPR varchar(30), " &
                "   DIM_2_UPR varchar(30),DIM_3_UPR varchar(30), LOCATION varchar(30), QTY decimal(18,4), DATE date, COST decimal(18,4), " &
                "   RETAIL decimal(18,4), LAST_UPDATE datetime, LAST_USER varchar(40)) " &
                "insert into #t1(TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, QTY, DATE, COST, RETAIL, " &
                "   LAST_UPDATE, LAST_USER) " &
                "SELECT H.EVENT_NO TRANS_ID, H.SEQ_NO, H.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "   coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, " &
                "   H.LOC_ID LOCATION, coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, H.TRX_DAT DATE, " &
                "   case when H.QTY = 0 then 0 else cast(coalesce(H.EXT_COST / H.QTY, H.EXT_COST) as decimal(20,8)) end as COST, " &
                "   cast(H.UNIT_RETL_VAL * (H.QTY_DENOM / H.QTY_NUMER) as decimal(20,8)) as RETAIL, " &
                "   coalesce(c.LST_MAINT_DT, H.LST_MAINT_DT) as LAST_UPDATE, " &
                "   coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) as LAST_USER from IM_ADJ_HIST H " &
                "left join IM_ADJ_HIST_CELL C on H.EVENT_NO = C.EVENT_NO And H.BAT_ID = C.BAT_ID And H.ITEM_NO = C.ITEM_NO " &
                "   And H.LOC_ID = C.LOC_ID And H.SEQ_NO = C.SEQ_NO " &
                "WHERE H.LST_MAINT_DT >= '" & minADJdate & "' " &
                "select TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, sum(QTY) QTY, DATE, avg(COST) COST, avg(RETAIL) RETAIL, " &
                "   LAST_UPDATE, LAST_USER from #t1 " &
                "group by TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, DATE, LAST_UPDATE, LAST_USER"
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
                If IsNumeric(oTest) Then qty = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", oTest)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(CDec(oTest), 4, MidpointRounding.AwayFromZero)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_UPDATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("LAST_UPDATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Collected Adjustment Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Console.WriteLine("Extracting Coupons")
            cpCon.Open()
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            path = xmlPath & "\Coupons.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("COUPON")
            sql = "SELECT DOC_ID, LIN_SEQ_NO, BUS_DAT, PROMPT_ALPHA INTO #t2 FROM PS_TKT_HIST_LIN_PROMPT " &
           "WHERE BUS_DAT >= '" & minSALdate & "' AND PROMPT_COD = 'MKTG_CODE' "
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                xmlWriter.WriteStartElement("COUPON")
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("DOC_ID", oTest)
                oTest = rdr("LIN_SEQ_NO")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteStartElement("SEQ_NO", oTest)
                oTest = rdr("BUS_DAT")
                If IsDBNull(oTest) Then oTest = "1900-01-01"
                xmlWriter.WriteStartElement("BUS_DATE", oTest)
                oTest = rdr("PROMPT_ALPHA")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteStartElement("CODE", Replace(oTest, "'", "''"))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()
            xmlWriter.WriteStartElement("COUPON")
            xmlWriter.WriteElementString("TRANS_ID", "COUPON")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("BUS_DATE", "")
            xmlWriter.WriteElementString("CODE", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Coupon Records " & cnt)
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
            sql = "create table #t1(TRANS_ID varchar(30), SEQ_NO integer, ITEM_NO varchar(30), DIM_1_UPR varchar(30), " &
                "DIM_2_UPR varchar(30), DIM_3_UPR varchar(30), LOCATION varchar(30), QTY decimal(18,4), COST decimal(18,4), " &
                "RETAIL decimal(18,4), DATE date, LAST_UPDATE datetime, LAST_USER varchar(40)) " &
                "insert into #t1(TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, QTY, COST, RETAIL, DATE, " &
                "LAST_UPDATE, LAST_USER) " &
                "select  H.EVENT_NO TRANS_ID,  0 SEQ_NO, H.ITEM_NO,  '*' as DIM_1_UPR, '*' as DIM_2_UPR,  '*' as DIM_3_UPR, " &
                "  H.LOC_ID LOCATION, H.QTY_ADJ As QTY, H.AVG_COST COST, H.UNIT_RETL_AT_POST as RETAIL, H.POST_DAT DATE, " &
                "  H.LST_MAINT_DT, H.LST_MAINT_USR_ID from IM_CNT_HIST H " &
                "left join IM_CNT_HIST_CELL C on H.EVENT_NO = C.EVENT_NO And H.ITEM_NO = C.ITEM_NO And H.LOC_ID = C.LOC_ID " &
                " where c.ITEM_NO Is NULL And H.QTY_ADJ <> 0 And H.POST_DAT >= '" & minADJdate & "' " &
                "union all " &
                "select  C.EVENT_NO TRANS_ID,  0 SEQ_NO, C.ITEM_NO,  DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, C.LOC_ID LOCATION,  " &
                "  C.QTY_CNTD - C.QTY_ON_HND_BEFORE As QTY, H.AVG_COST COST, " &
                "  H.UNIT_RETL_AT_POST as RETAIL, H.POST_DAT DATE, C.LST_MAINT_DT, c.LST_MAINT_USR_ID from IM_CNT_HIST H " &
                "left join IM_CNT_HIST_CELL C on H.EVENT_NO = C.EVENT_NO And H.ITEM_NO = C.ITEM_NO And H.LOC_ID = C.LOC_ID " &
                " where c.QTY_CNTD - C.QTY_ON_HND_BEFORE <> 0 AND H.POST_DAT >= '" & minADJdate & "' " &
                "select TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, sum(QTY) QTY, avg(COST) COST, " &
                "avg(RETAIL) RETAIL, DATE, LAST_UPDATE, LAST_USER from #t1 " &
                "where qty <> 0 " &
                "group by TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, DATE, LAST_UPDATE, LAST_USER"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            Dim sku As String
            Dim id As String
            Dim qty As Decimal = 0
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                oTest = rdr("QTY")
                If IsDBNull(oTest) Or IsNothing(oTest) Then qty = 0
                If IsNumeric(oTest) Then qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero) Else qty = 0
                If qty <> 0 Then
                    xmlWriter.WriteStartElement("PHYSICAL")
                    oTest = rdr("TRANS_ID")
                    If IsDBNull(oTest) Then id = CStr(oTest) Else id = "NA"
                    totlQty += qty
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

                    xmlWriter.WriteElementString("QTY", qty)
                    oTest = rdr("COST")
                    cost = 0
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    totlCost += (qty * cost)
                    xmlWriter.WriteElementString("COST", cost)
                    oTest = rdr("RETAIL")
                    retail = 0
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    totlRetail += (qty * retail)
                    xmlWriter.WriteElementString("RETAIL", retail)
                    oTest = rdr("DATE")
                    If IsDBNull(oTest) Then oTest = "1/1/1900"
                    xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                    oTest = rdr("LAST_UPDATE")
                    If IsDBNull(oTest) Then oTest = "1/1/1900"
                    xmlWriter.WriteElementString("LAST_UPDATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                    oTest = rdr("LAST_USER")
                    If IsDBNull(oTest) Then oTest = Nothing
                    xmlWriter.WriteElementString("LAST_USER", oTest)
                    xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                    xmlWriter.WriteEndElement()
                End If
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            sql = "create table #t1(TRANS_ID varchar(30), SEQ_NO integer, ITEM_NO varchar(30), DIM_1_UPR varchar(30), " &
                "DIM_2_UPR varchar(30), DIM_3_UPR varchar(30), LOCATION varchar(30), QTY decimal(18,4), COST decimal(18,4), " &
                "RETAIL decimal(18,4), DATE date, LAST_UPDATE datetime, LAST_USER varchar(40)) " &
                "insert into #t1(TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, QTY, COST, RETAIL, DATE, " &
                "LAST_UPDATE, LAST_USER) " &
                "select  H.EVENT_NO TRANS_ID,  H.SEQ_NO, H.ITEM_NO,  coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR,  coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, " &
                "  H.LOC_ID LOCATION,  coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.EXT_COST / H.QTY, H.EXT_COST) as decimal(20,8)) end as COST, " &
                "  H.UNIT_RETL_VAL As RETAIL, H.TRX_DAT Date, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT) LAST_UPDATE," &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from PO_QRECV_HIST H " &
                "left join PO_QRECV_HIST_CELL C on H.EVENT_NO = C.EVENT_NO  And H.BAT_ID = C.BAT_ID " &
                " And H.ITEM_NO = C.ITEM_NO And H.LOC_ID = C.LOC_ID And H.TRX_DAT = C.TRX_DAT And H.SEQ_NO = C.SEQ_NO " &
                " where H.LST_MAINT_DT >= '" & minRCPTdate & "' " &
                " union all " &
                "Select  H.EVENT_NO TRANS_ID,  H.SEQ_NO, H.ITEM_NO,  coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR,  coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, " &
                "  H.RECVR_LOC_ID LOCATION,  coalesce(C.QTY_RECVD, H.QTY_RECVD) * (H.QTY_RECVD_NUMER / H.QTY_RECVD_DENOM) as QTY, " &
                "  case when H.QTY_RECVD = 0 then 0 else cast(H.EXT_COST / H.QTY_RECVD as decimal(20,8)) end as EXT_COST, " &
                "  H.UNIT_RETL_VAL as RETAIL, H.RECVR_DAT DATE, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT) LAST_UPDATE, " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from PO_RECVR_HIST_LIN H " &
                "join PO_RECVR_HIST " &
                "  on PO_RECVR_HIST.RECVR_NO = H.RECVR_NO " &
                "left join PO_RECVR_HIST_CELL C on H.RECVR_NO = C.RECVR_NO And H.SEQ_NO = C.SEQ_NO " &
                "WHERE PO_RECVR_HIST.IS_DROPSHIP_RECVR = 'N' and H.LST_MAINT_DT >= '" & minRCPTdate & "' " &
                "select TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, sum(QTY) QTY, sum(COST) COST, sum(RETAIL) RETAIL, " &
                "DATE, LAST_UPDATE, LAST_USER from #t1 " &
                "group by TRANS_ID, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOCATION, DATE, LAST_UPDATE, LAST_USER"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 960
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                sku = rdr("ITEM_NO")
                If IsDBNull(sku) Then sku = "NA"
                oTest = rdr("DIM_1_UPR")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                    End If
                End If
                xmlWriter.WriteStartElement("RECEIPT")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "1"
                xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                oTest = rdr("QTY")
                If IsDBNull(oTest) Then qty = 0 Else qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", oTest)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_UPDATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("LAST_UPDATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            sql = "select  H.EVENT_NO TRANS_ID,  H.SEQ_NO, H.RTV_LOC_ID LOCATION, H.ITEM_NO,  coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, " &
                "  coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, coalesce(C.QTY_RETD, H.QTY_RETD) * (H.QTY_RETD_NUMER / H.QTY_RETD_DENOM) as QTY, " &
                "  case when H.QTY_RETD = 0 then 0 else cast(coalesce(-H.EXT_COST / H.QTY_RETD, -H.EXT_COST) as decimal(20,8)) end as COST, " &
                "  H.UNIT_RETL_VAL * (H.QTY_RETD_DENOM / H.QTY_RETD_NUMER) as RETAIL, H.RTV_DAT DATE, " &
                "  coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT) LAST_UPDATE, " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from PO_RTV_HIST_LIN H " &
                "left join PO_RTV_HIST_CELL C " &
                "  on H.SEQ_NO = C.SEQ_NO " &
                " And H.RTV_NO = C.RTV_NO " &
                "where H.LST_MAINT_DT >= '" & minRTNdate & "'"
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
                If IsNumeric(oTest) Then qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero) Else qty = 0
                totlQty += qty
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
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                '' xmlWriter.WriteElementString("OVERRIDE", "")
                oTest = rdr("LAST_UPDATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("LAST_UPDATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                oTest = rdr("LAST_USER")
                If IsDBNull(oTest) Then oTest = Nothing
                xmlWriter.WriteElementString("LAST_USER", oTest)
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
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            Dim item, itemString, dim1, dim2, dim3, store, drawer, location, custom, custom2, custom3, custom4, custom5 As String
            Dim grandQty As Decimal = 0
            Dim grandCost As Decimal = 0
            Dim grandRetail As Decimal = 0
            Dim grandMarkdown As Decimal = 0
            custom = Nothing
            custom2 = Nothing
            custom3 = Nothing
            custom4 = Nothing
            custom5 = Nothing
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Sales")
            sql = "CREATE TABLE #tt1 (BUS_DATE date, TRANS_ID varchar(30), SEQ_NO int, STORE varchar(10), ITEM varchar(90), " &
                "DIM1 varchar(30) NULL, DIM2 varchar(30) NULL, DIM3 varchar(30) NULL, LOCATION varchar(10), DRAWER varchar(10) NULL,  " &
                "QTY decimal(18,4) NULL, COST decimal(18,4) NULL, RETAIL decimal(18,4) NULL, TKT_DT datetime, MARKDOWN decimal(18,4) NULL, " &
                "MKDN_REASON varchar(30) NULL,COUPON_CODE varchar(30) NULL, DEPT varchar(10) NULL, CLASS varchar(10) NULL, " &
                "BUYER varchar(10) NULL, CUST_NO varchar(30) NULL,TKT_NO varchar(30) NULL, USER_ID varchar(10) NULL, " &
                "SALES_REP varchar(10) NULL, STATION varchar(10) NULL) " &
                "INSERT INTO #tt1 (BUS_DATE, TRANS_ID,SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, QTY, COST, RETAIL, TKT_DT, " &
                "MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO, USER_ID, SALES_REP, STATION) " &
                "SELECT ISNULL(h.BUS_DAT,'1/1/1900'), ISNULL(h.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, " &
                "ISNULL(l.STR_ID,'NA') STORE, ISNULL(l.ITEM_NO,'NA') AS ITEM, c.DIM_1_UPR, " &
                "c.DIM_2_UPR, c.DIM_3_UPR, l.STK_LOC_ID AS LOCATION, DRW_ID AS DRAWER,  " &
                "CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) ELSE c.QTY_SOLD END AS QTY, ISNULL(l.COST,0) AS COST, " &
                "ISNULL(l.PRC,0) - ISNULL(l.LIN_DISC_AMT,0) AS RETAIL,  h.TKT_DT, " &
                "(ISNULL(l.PRC_1,0) - ISNULL(l.PRC,0) + ISNULL(l.LIN_DISC_AMT,0)) AS MARKDOWN, " &
                "l.PRC_OVRD_REAS, NULL, ISNULL(l.CATEG_COD,'NA') AS DEPT, ISNULL(l.SUBCAT_COD,'NA') AS CLASS, " &
                "ISNULL(pv.LST_ORD_BUYER,'OTHER') AS BUYER, CUST_NO, h.TKT_NO, USR_ID, l.SLS_REP, l.STA_ID FROM VI_PS_TKT_HIST_LIN AS l " &
                "JOIN VI_PS_TKT_HIST h ON h.DOC_ID = l.DOC_ID AND h.BUS_DAT = l.BUS_DAT " &
                "LEFT JOIN VI_PS_TKT_HIST_LIN_CELL c ON c.DOC_ID = h.DOC_ID AND c.BUS_DAT = h.BUS_DAT AND c.LIN_SEQ_NO = l.LIN_SEQ_NO " &
                "LEFT JOIN PO_VEND_ITEM pv ON pv.VEND_NO = l.ITEM_VEND_NO AND pv.ITEM_NO = l.ITEM_NO " &
                "WHERE l.LIN_TYP IN ('S','R') AND h.BUS_DAT >= '" & minSALdate & "'"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 48
            cmd.ExecuteNonQuery()
            sql = "SELECT BUS_DATE, TRANS_ID,SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, STORE, QTY, COST, RETAIL, TKT_DT, " &
                "MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, t.CUST_NO, TKT_NO, CATEG_COD, USER_ID, SALES_REP, STATION FROM #tt1 t " &
                "LEFT JOIN AR_CUST c ON c.CUST_NO = t.CUST_NO"

            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                If cnt = 30000 Then
                    xmlWriter.WriteStartElement("SALE")
                    xmlWriter.WriteElementString("TRANS_DATE", "")
                    xmlWriter.WriteElementString("TRANS_ID", "SALES")
                    xmlWriter.WriteElementString("SEQ_NO", 0)
                    xmlWriter.WriteElementString("STR_ID", "")
                    xmlWriter.WriteElementString("SKU", "TOTALS")
                    xmlWriter.WriteElementString("LOCATION", "")
                    xmlWriter.WriteElementString("DRAWER", "")
                    xmlWriter.WriteElementString("QTY", totlQty)
                    xmlWriter.WriteElementString("COST", totlCost)
                    xmlWriter.WriteElementString("RETAIL", totlRetail)
                    xmlWriter.WriteElementString("TICKET_DATE", "")
                    xmlWriter.WriteElementString("MARKDOWN", totlMarkdown)
                    xmlWriter.WriteElementString("DEPT", "")
                    xmlWriter.WriteElementString("CLASS", "")
                    xmlWriter.WriteElementString("BUYER", "")
                    xmlWriter.WriteElementString("CUST_NO", "")
                    xmlWriter.WriteElementString("TKT_NO", "")
                    xmlWriter.WriteElementString("TKT_DATE", "")
                    xmlWriter.WriteElementString("MKDN_REASON", "")
                    xmlWriter.WriteElementString("COUPON_CODE", "")
                    xmlWriter.WriteElementString("ORD_TYPE", "")
                    xmlWriter.WriteElementString("WHOLESALE", "")
                    xmlWriter.WriteElementString("USER_ID", "")
                    xmlWriter.WriteElementString("SALES_REP", "")
                    xmlWriter.WriteElementString("STATION", "")
                    xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                    xmlWriter.WriteEndElement()
                    xmlWriter.WriteEndElement()
                    xmlWriter.WriteEndDocument()
                    xmlWriter.Close()
                    path = xmlPath & "\Sales" & files & ".xml"
                    xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
                    xmlWriter.WriteStartDocument(True)
                    xmlWriter.Formatting = Formatting.Indented
                    xmlWriter.Indentation = 2
                    xmlWriter.WriteStartElement("Sales")
                    grandQty += totlQty
                    grandCost += totlCost
                    grandRetail += totlRetail
                    grandMarkdown += totlMarkdown
                    cnt = 0
                    totlQty = 0
                    totlCost = 0
                    totlRetail = 0
                    totlMarkdown = 0
                    files += 1
                End If
                item = ""
                itemString = ""
                xmlWriter.WriteStartElement("SALE")
                oTest = rdr("BUS_DATE")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd"))
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
                If IsDBNull(oTest) Then qty = 0 Else qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlQty += qty
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlCost += (qty * cost)
                xmlWriter.WriteElementString("COST", cost)
                oTest = rdr("RETAIL")
                If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("RETAIL", retail)
                oTest = rdr("TKT_DT")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TICKET_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
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
                oTest = rdr("CATEG_COD")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("ORD_TYPE", Replace(oTest, "'", "''"))
                ''If oTest = "WHOLESALE" Then
                ''    xmlWriter.WriteElementString("WHOLESALE", "Y")
                ''Else
                ''    xmlWriter.WriteElementString("WHOLESALE", "N")
                ''End If
                oTest = rdr("USER_ID")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("USER_ID", Replace(oTest, "'", "''"))
                oTest = rdr("SALES_REP")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("SALES_REP", Replace(oTest, "'", "''"))
                oTest = rdr("STATION")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("STATION", oTest)
                xmlWriter.WriteElementString("CUSTOM_1", custom)
                xmlWriter.WriteElementString("CUSTOM_2", custom2)
                xmlWriter.WriteElementString("CUSTOM_3", custom3)
                xmlWriter.WriteElementString("CUSTOM_4", custom4)
                xmlWriter.WriteElementString("CUSTOM_5", custom5)
                xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                xmlWriter.WriteEndElement()
            End While
            '
            '               Write a row for totals
            '
            xmlWriter.WriteStartElement("SALE")
            xmlWriter.WriteElementString("TRANS_DATE", "")
            xmlWriter.WriteElementString("TRANS_ID", "SALES")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("STR_ID", "")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("DRAWER", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TICKET_DATE", "")
            xmlWriter.WriteElementString("MARKDOWN", totlMarkdown)
            xmlWriter.WriteElementString("DEPT", "")
            xmlWriter.WriteElementString("CLASS", "")
            xmlWriter.WriteElementString("BUYER", "")
            xmlWriter.WriteElementString("CUST_NO", "")
            xmlWriter.WriteElementString("TKT_NO", grandQty)
            xmlWriter.WriteElementString("TKT_DATE", "")
            xmlWriter.WriteElementString("MKDN_REASON", grandCost)
            xmlWriter.WriteElementString("COUPON_CODE", grandRetail)
            xmlWriter.WriteElementString("ORD_TYPE", grandMarkdown)
            xmlWriter.WriteElementString("WHOLESALE", "")
            xmlWriter.WriteElementString("USER_ID", "")
            xmlWriter.WriteElementString("SALES_REP", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Sales", m, "")
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            Console.WriteLine(ex.Message & " " & ex.StackTrace)
            Console.ReadLine()

            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Sale Records")
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
            Console.WriteLine("Extracting Order data")
            cpCon.Open()
            path = xmlPath & "\Orders.xml"
            Dim item, itemString, dim1, dim2, dim3, store, drawer, location, custom, custom2, custom3, custom4, custom5 As String
            Dim mkdn As Decimal = 0
            Dim grandQty As Decimal = 0
            Dim grandCost As Decimal = 0
            Dim grandRetail As Decimal = 0
            Dim grandMarkdown As Decimal = 0
            custom = Nothing
            custom2 = Nothing
            custom3 = Nothing
            custom4 = Nothing
            custom5 = Nothing
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Orders")
            sql = "CREATE TABLE #dtl (TRANS_ID varchar(30), SEQ_NO int, STORE varchar(10), ITEM varchar(30),  DIM1 varchar(30) NULL, " &
                "DIM2 varchar(30) NULL, DIM3 varchar(30) NULL, LOCATION varchar(10), STATION varchar(30), DRAWER varchar(10) NULL, " &
                "QTY decimal(18,4) NULL, COST decimal(18,4) NULL, RETAIL decimal(18,4) NULL, DATE date, MARKDOWN decimal(18,4) NULL,  " &
                "MKDN_REASON varchar(30) NULL, COUPON_CODE varchar(30) NULL, DEPT varchar(10) NULL, CLASS varchar(10) NULL, " &
                "BUYER varchar(10) NULL, TKT_NO varchar(30) NULL, VENDOR varchar(30), SALES_REP varchar(10), CUST_NO varchar(15), " &
                "ORD_TYPE varchar(10), TKT_DATE datetime, USER_ID varchar(10)) " &
                "INSERT INTO #dtl (TRANS_ID, SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, STATION, DRAWER, QTY, COST, RETAIL, DATE, " &
                "MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, TKT_NO, VENDOR, SALES_REP, CUST_NO, ORD_TYPE, TKT_DATE, USER_ID) " &
                "SELECT ISNULL(l.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO,  ISNULL(l.STR_ID,'NA') STR_ID, " &
                "ISNULL(l.ITEM_NO,'NA') ITEM, ISNULL(c.DIM_1_UPR,'*') DIM1, ISNULL(c.DIM_2_UPR,'*') DIM2, ISNULL(c.DIM_3_UPR,'*') DIM3,  " &
                "ISNULL(l.STK_LOC_ID,'') LOCATION, ISNULL(h.STA_ID,'NA'), 'NA' DRAWER, " &
                "CASE WHEN c.QTY_SHIPPED IS NULL THEN ISNULL(l.QTY_SHIPPED,0) ELSE ISNULL(c.QTY_SHIPPED,0) END QTY,  ISNULL(i.AVG_COST,0) COST, " &
                "ISNULL(l.PRC,0) - ISNULL(l.LIN_DISC_AMT,0) RETAIL, ISNULL(h.BUS_DAT,'1/1/1900') DATE, " &
                "(ISNULL(l.PRC_1,0) - ISNULL(l.PRC,0) + ISNULL(l.LIN_DISC_AMT,0)) MARKDOWN,  " &
                "ISNULL(PRC_OVRD_REAS,'') MKDN_REASON, CONVERT(varchar(30),'') COUPON_CODE, ISNULL(l.CATEG_COD,'NA') DEPT, " &
                "ISNULL(l.SUBCAT_COD,'NA') CLASS,  ISNULL(pv.LST_ORD_BUYER,'OTHER') BUYER, ISNULL(l.TKT_NO,'') TKT_NO, " &
                "ISNULL(i.ITEM_VEND_NO,'') VENDOR_ID, ISNULL(h.SLS_REP,'NA'), ISNULL(h.CUST_NO,'NA'), " &
                "ISNULL(h.DOC_TYP,'UNKNOWN'), ISNULL(h.TKT_DT,'1-1-1900'), ISNULL(h.USR_ID,'UNKNOWN') FROM VI_PS_ORD_HIST_LIN AS l " &
                "INNER JOIN VI_PS_ORD_HIST AS h ON h.DOC_ID = l.DOC_ID " &
                "LEFT JOIN PO_VEND_ITEM AS pv ON pv.ITEM_NO = l.ITEM_NO AND pv.VEND_NO = l.ITEM_VEND_NO " &
                "LEFT JOIN VI_IM_ITEM_WITH_INV i ON i.ITEM_NO = l.ITEM_NO AND i.LOC_ID = l.STK_LOC_ID " &
                "LEFT JOIN VI_PS_ORD_HIST_LIN_CELL c ON c.DOC_ID = l.DOC_ID AND c.LIN_SEQ_NO = l.LIN_SEQ_NO " &
                "LEFT JOIN VI_AR_CUST_WITH_ADDRESS ca ON ca.CUST_NO = h.CUST_NO " &
                "WHERE h.TKT_DAT >= '" & minORDdate & "' AND NOT EXISTS (SELECT * FROM VI_PS_DOC_AUDIT_LOG a " &
                    "WHERE a.DOC_ID = l.DOC_ID And a.DOC_TYP = 'V')  " &
                "SELECT DOC_ID, LIN_SEQ_NO, PROMPT_ALPHA_1 INTO #us2 FROM VI_PS_DOC_LIN WHERE PROMPT_COD_1 = 'MKTG_CODE' " &
                "UPDATE t1 SET t1.COUPON_CODE = PROMPT_ALPHA_1 FROM #dtl t1 " &
                "JOIN #us2 t2 ON t2.DOC_ID = t1.TRANS_ID AND t2.LIN_SEQ_NO = SEQ_NO " &
                "SELECT TRANS_ID, SEQ_NO, STORE, ITEM, DIM1, DIM2, DIM3, LOCATION, STATION, DRAWER, QTY, COST, RETAIL, DATE, MARKDOWN, " &
                "MKDN_REASON,  COUPON_CODE, DEPT, CLASS, BUYER, TKT_NO, SALES_REP, CUST_NO, ORD_TYPE, TKT_DATE, USER_ID FROM #dtl WHERE QTY > 0"
            cmd = New SqlCommand(sql, cpCon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("ORDER")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("TRANS_ID", oTest)
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = 0 Else oTest = CInt(oTest)
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("STORE")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("STR_ID", oTest)
                oTest = rdr("ITEM")
                If Not IsDBNull(oTest) Then oTest = Replace(oTest, "'", "''") Else oTest = ""
                item = oTest
                itemString = oTest
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        itemString &= "~" & oTest & "~" & rdr("DIM2") & "~" & rdr("DIM3")
                    End If
                End If
                xmlWriter.WriteElementString("SKU", Replace(itemString, "'", "''"))
                xmlWriter.WriteElementString("ITEM", item)
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then oTest = Replace(oTest, "'", "''") Else oTest = ""
                xmlWriter.WriteElementString("DIM_1", oTest)
                oTest = rdr("DIM2")
                If Not IsDBNull(oTest) Then oTest = Replace(oTest, "'", "''") Else oTest = ""
                xmlWriter.WriteElementString("DIM_2", oTest)
                oTest = rdr("DIM3")
                If Not IsDBNull(oTest) Then oTest = Replace(oTest, "'", "''") Else oTest = ""
                xmlWriter.WriteElementString("DIM_3", oTest)
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("LOCATION", oTest)
                oTest = rdr("STATION")
                xmlWriter.WriteElementString("STATION", oTest)
                oTest = rdr("DRAWER")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DRAWER", oTest)
                oTest = rdr("QTY")
                If Not IsDBNull(oTest) Then
                    If Not IsNothing(oTest) Then qty = CDec(oTest)
                Else
                    qty = 0
                End If
                xmlWriter.WriteElementString("QTY", qty)
                oTest = rdr("COST")
                If Not IsDBNull(oTest) Then
                    If IsNumeric(oTest) Then cost = CDec(oTest)
                Else
                    cost = 0
                End If
                xmlWriter.WriteElementString("COST", oTest)
                oTest = rdr("RETAIL")
                If Not IsDBNull(oTest) Then
                    If IsNumeric(oTest) Then retail = CDec(oTest)
                Else
                    retail = 0
                End If
                xmlWriter.WriteElementString("RETAIL", oTest)
                oTest = rdr("DATE")
                xmlWriter.WriteElementString("TRANS_DATE", oTest)
                oTest = rdr("MARKDOWN")
                If IsNumeric(oTest) Then mkdn = CDec(oTest) Else mkdn = 0
                mkdn = mkdn * qty
                xmlWriter.WriteElementString("MARKDOWN", mkdn)
                oTest = rdr("MKDN_REASON")
                If IsNumeric(oTest) Then oTest = CDec(oTest) Else oTest = 0
                xmlWriter.WriteElementString("MKDN_REASON", CDec(oTest))
                oTest = rdr("COUPON_CODE")
                xmlWriter.WriteElementString("COUPON_CODE", oTest)
                oTest = rdr("DEPT")
                xmlWriter.WriteElementString("DEPT", oTest)
                oTest = rdr("CLASS")
                xmlWriter.WriteElementString("CLASS", oTest)
                oTest = rdr("BUYER")
                xmlWriter.WriteElementString("BUYER", oTest)
                oTest = rdr("TKT_NO")
                xmlWriter.WriteElementString("TKT_NO", oTest)
                xmlWriter.WriteElementString("VENDOR_ID", oTest)
                oTest = rdr("SALES_REP")
                xmlWriter.WriteElementString("SALES_REP", CStr(oTest))
                oTest = rdr("CUST_NO")
                xmlWriter.WriteElementString("CUST_NO", CStr(oTest))
                oTest = rdr("ORD_TYPE")
                xmlWriter.WriteElementString("ORD_TYPE", CStr(oTest))
                oTest = rdr("TKT_DATE")
                xmlWriter.WriteElementString("TICKET_DATE", oTest)
                oTest = rdr("USER_ID")
                xmlWriter.WriteElementString("USER_ID", CStr(oTest))
                xmlWriter.WriteElementString("CUSTOM_1", custom)
                xmlWriter.WriteElementString("CUSTOM_2", custom2)
                xmlWriter.WriteElementString("CUSTOM_3", custom3)
                xmlWriter.WriteElementString("CUSTOM_4", custom4)
                xmlWriter.WriteElementString("CUSTOM_5", custom5)
                totlQty += qty
                totlCost += (qty * cost)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()

            xmlWriter.WriteStartElement("ORDER")
            xmlWriter.WriteElementString("TRANS_ID", "")
            xmlWriter.WriteElementString("SEQ_NO", "")
            xmlWriter.WriteElementString("STORE", "")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("ITEM", "")
            xmlWriter.WriteElementString("DIM_1", "")
            xmlWriter.WriteElementString("DIM_2", "")
            xmlWriter.WriteElementString("DIM_3", "")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("DRAWER", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", "")
            xmlWriter.WriteElementString("MARKDOWN", "")
            xmlWriter.WriteElementString("MKDN_REASON", "")
            xmlWriter.WriteElementString("COUPON_CODE", "")
            xmlWriter.WriteElementString("DEPT", "")
            xmlWriter.WriteElementString("CLASS", "")
            xmlWriter.WriteElementString("BUYER", "")
            xmlWriter.WriteElementString("TKT_NO", "")
            xmlWriter.WriteElementString("VENDOR_ID", "")
            xmlWriter.WriteElementString("SALES_REP", "")
            xmlWriter.WriteElementString("CUST_NO", "")
            xmlWriter.WriteElementString("TKT_DATE", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            Console.WriteLine(ex.Message & " " & ex.StackTrace)
            Console.ReadLine()

            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Sale Records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try

        Try
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
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            totlMarkdown = 0
            Dim sku As String
            sql = "create table #t1(TYPE varchar(20), EVENT_NO varchar(30), SEQ_NO integer, ITEM_NO varchar(30), " &
                "DIM_1_UPR varchar(30), DIM_2_UPR varchar(30), DIM_3_UPR varchar(30), LOC_ID varchar(30), " &
                 "TRX_DAT date, QTY decimal(18,4), COST decimal(18,4), RETAIL decimal(18,4), LAST_UPDATE datetime, " &
                " LAST_USER varchar(40)) " &
                 "insert into #t1 (TYPE, EVENT_NO, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOC_ID, " &
                 "TRX_DAT, QTY, COST, RETAIL, LAST_UPDATE, LAST_USER) " &
                "SELECT 'Quick OUT', H.EVENT_NO, H.SEQ_NO, H.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, H.FROM_LOC_ID LOC_ID, " &
                "  H.TRX_DAT, coalesce(-C.QTY, -H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.TO_EXT_COST / H.QTY, H.TO_EXT_COST) as decimal(20,8)) end as COST, " &
                "  H.TO_UNIT_RETL_VAL * (H.QTY_DENOM / H.QTY_NUMER) as RETAIL, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT), " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from IM_QXFER_HIST H " &
                "left join IM_QXFER_HIST_CELL C on H.EVENT_NO = C.EVENT_NO And H.BAT_ID = C.BAT_ID " &
                "  And H.ITEM_NO = C.ITEM_NO And H.FROM_LOC_ID = C.FROM_LOC_ID And H.TRX_DAT = C.TRX_DAT " &
                "  And H.SEQ_NO = C.SEQ_NO " &
                "WHERE H.LST_MAINT_DT >= '" & minXFERdate & "' " &
                "UNION ALL " &
                "SELECT 'Quick IN', H.EVENT_NO, H.SEQ_NO, H.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, H.TO_LOC_ID LOC_ID, " &
                "  H.TRX_DAT, coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.TO_EXT_COST / H.QTY, H.TO_EXT_COST) as decimal(20,8)) end as COST, " &
                "  H.TO_UNIT_RETL_VAL * (H.QTY_DENOM / H.QTY_NUMER) as RETAIL, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT), " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from IM_QXFER_HIST H " &
                "left join IM_QXFER_HIST_CELL C on H.EVENT_NO = C.EVENT_NO And H.BAT_ID = C.BAT_ID " &
                "  And H.ITEM_NO = C.ITEM_NO And H.FROM_LOC_ID = C.FROM_LOC_ID And H.TRX_DAT = C.TRX_DAT " &
                "  And H.SEQ_NO = C.SEQ_NO " &
                "WHERE H.LST_MAINT_DT >= '" & minXFERdate & "' " &
                "UNION ALL " &
                "SELECT 'Transfer OUT', H.EVENT_NO, L.XFER_LIN_SEQ_NO SEQ_NO, L.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, H.FROM_LOC_ID LOC_ID, " &
                "  H.SHIP_DAT, coalesce(-C.XFER_QTY, -L.XFER_QTY) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "  case when L.XFER_QTY = 0 then 0 else cast(coalesce(L.FROM_EXT_COST " &
                    "/ L.XFER_QTY, L.FROM_EXT_COST) as decimal(20,8)) end as COST, " &
                "  L.FROM_UNIT_RETL_VAL * (L.XFER_QTY_DENOM / L.XFER_QTY_NUMER) as RETAIL, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT), " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from IM_XFER_HDR H " &
                "join IM_XFER_LIN L on H.XFER_NO = L.XFER_NO " &
                "left join IM_XFER_CELL C on C.XFER_NO = L.XFER_NO And C.XFER_LIN_SEQ_NO = L.XFER_LIN_SEQ_NO " &
                "WHERE H.LST_MAINT_DT >= '" & minXFERdate & "' " &
                "UNION ALL " &
                "SELECT 'Transfer IN', H.EVENT_NO, L.XFER_LIN_SEQ_NO SEQ_NO, L.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, H.TO_LOC_ID LOC_ID, " &
                "  H.RECVD_DAT, coalesce(C.QTY_RECVD, L.QTY_RECVD) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "case when L.QTY_RECVD = 0 then 0 else cast(coalesce(L.TO_TOT_COST_CORR " &
                    "/ L.QTY_RECVD, L.TO_TOT_COST_CORR ) as decimal(20,8)) end as COST, " &
                "  L.TO_UNIT_RETL_VAL * (L.XFER_QTY_DENOM / L.XFER_QTY_NUMER) as RETAIL, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT), " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER from IM_XFER_IN_HIST H " &
                "join IM_XFER_IN_HIST_LIN L on H.XFER_NO = L.XFER_NO And H.EVENT_NO = L.EVENT_NO " &
                "left join IM_XFER_IN_HIST_CELL C on C.EVENT_NO = L.EVENT_NO And C.XFER_NO = L.XFER_NO " &
                    "And C.XFER_IN_LIN_SEQ_NO = L.XFER_IN_LIN_SEQ_NO " &
                "WHERE H.LST_MAINT_DT >= '" & minXFERdate & "' " &
                "UNION ALL " &
                "SELECT 'Recon', H.EVENT_NO, L.XFER_LIN_SEQ_NO SEQ_NO, L.ITEM_NO, coalesce(C.DIM_1_UPR, '*') as DIM_1_UPR, " &
                "  coalesce(C.DIM_2_UPR, '*') as DIM_2_UPR, coalesce(C.DIM_3_UPR, '*') as DIM_3_UPR, " &
                "  CASE WHEN L.RECON_METH = 'S' THEN H.TO_LOC_ID ELSE H.FROM_LOC_ID END LOC_ID, " &
                "  H.RECON_DAT TRX_DAT, coalesce(-C.QTY_VARIANCE, -L.QTY_VARIANCE) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "case " &
                "    when L.RECON_METH = 'S' then cast(coalesce(L.TO_EXT_COST / (case when L.QTY_VARIANCE = 0 then 1 " &
                "		Else L.QTY_VARIANCE End), L.TO_EXT_COST) As Decimal(20, 8)) " &
                "    when L.QTY_VARIANCE = 0 then 0 " &
                "    else cast(coalesce(L.FROM_EXT_COST / L.QTY_VARIANCE, L.FROM_EXT_COST) as decimal(20, 8)) " &
                "end as COST, " &
                "case " &
                "    when L.RECON_METH = 'S' then L.TO_UNIT_RETL_VAL " &
                "    when L.QTY_VARIANCE = 0 then 0 " &
                "    else L.FROM_UNIT_RETL_VAL " &
                "end as RETAIL, coalesce(C.LST_MAINT_DT, H.LST_MAINT_DT) LAST_UPDATE, " &
                "  coalesce(C.LST_MAINT_USR_ID, H.LST_MAINT_USR_ID) LAST_USER " &
                "from IM_XFER_RECON_HIST H " &
                "JOIN IM_XFER_RECON_HIST_LIN L on H.XFER_NO = L.XFER_NO " &
                "LEFT JOIN IM_XFER_RECON_HIST_CELL C on C.XFER_NO = H.XFER_NO And C.XFER_LIN_SEQ_NO = L.XFER_LIN_SEQ_NO " &
                "WHERE L.RECON_METH in ('S', 'R') AND H.LST_MAINT_DT >= '" & minXFERdate & "' " &
                "select TYPE, EVENT_NO, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOC_ID, TRX_DAT, SUM(QTY) QTY, " &
                    "SUM(cost) COST, SUM(retail) RETAIL, LAST_UPDATE, LAST_USER FROM #t1 " &
                    "GROUP BY TYPE, EVENT_NO, SEQ_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, LOC_ID, TRX_DAT, LAST_UPDATE, " &
                    "LAST_USER"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 960
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("QTY")
                If IsDBNull(oTest) Then qty = 0 Else qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                If qty <> 0 Then
                    totlQty += qty
                    xmlWriter.WriteStartElement("TRANSFER")
                    oTest = rdr("EVENT_NO")
                    If IsDBNull(oTest) Then oTest = "NA"
                    xmlWriter.WriteElementString("TRANS_ID", Replace(oTest, "'", "''"))
                    oTest = rdr("SEQ_NO")
                    If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                    xmlWriter.WriteElementString("SEQ_NO", CInt(oTest))
                    sku = rdr("ITEM_NO")
                    If IsDBNull(sku) Then sku = "NA"
                    oTest = rdr("DIM_1_UPR")
                    If Not IsDBNull(oTest) Then
                        If oTest <> "*" Then
                            sku &= "~" & oTest & "~" & rdr("DIM_2_UPR") & "~" & rdr("DIM_3_UPR")
                        End If
                    End If
                    xmlWriter.WriteElementString("SKU", Replace(sku, "'", "''"))
                    oTest = rdr("LOC_ID")
                    If IsDBNull(oTest) Then oTest = "NA"
                    xmlWriter.WriteElementString("LOCATION", Replace(oTest, "'", "''"))
                    xmlWriter.WriteElementString("QTY", qty)
                    oTest = rdr("COST")
                    If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    totlCost += (qty * cost)
                    xmlWriter.WriteElementString("COST", cost)
                    oTest = rdr("RETAIL")
                    If IsDBNull(oTest) Then retail = 0 Else retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    totlRetail += (qty * retail)
                    xmlWriter.WriteElementString("RETAIL", retail)
                    oTest = rdr("TRX_DAT")
                    If IsDBNull(oTest) Then oTest = CDate("1900-01-01")
                    xmlWriter.WriteElementString("TRANS_DATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                    oTest = rdr("LAST_UPDATE")
                    If IsDBNull(oTest) Then oTest = CDate("1900-01-01")
                    xmlWriter.WriteElementString("LAST_UPDATE", Format(oTest, "yyyy-MM-dd HH:mm:ss"))
                    oTest = rdr("LAST_USER")
                    If IsDBNull(oTest) Then oTest = Nothing
                    xmlWriter.WriteElementString("LAST_USER", oTest)
                    xmlWriter.WriteElementString("EXTRACT_DATE", DateTimeString)
                    xmlWriter.WriteEndElement()
                End If
            End While
            cpCon.Close()
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

            Dim m As String = Format(cnt, "###,###,##0") & " Records"
            Call Update_Process_Log("1", "Extract Transfers", m, "")
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
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
            sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " &
               "SELECT '" & thisDateTime & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            Dim el As New XMLExtract_TCM.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Update Message (Cant update the ProcessLog)")
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



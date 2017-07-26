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
    Private tbl, prdTbl As DataTable
    Sub Main()
        Try
            StopWatch = New Stopwatch
            StopWatch.Start()
            Dim yrwk As Integer = 0
            Dim ADJwks As Integer = 2
            Dim RCPTwks As Integer = 2
            Dim RTNwks As Integer = 2
            Dim SALwks As Integer = 2
            Dim XFERwks As Integer = 2
            Dim POwks As Integer = 2
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
            Console.WriteLine("Extracting Calendar")
            conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & "; Integrated Security=True"
            cpCon = New SqlConnection(conString)
            Dim theFirstYear As Boolean = True
            path = xmlPath & "\Calendar.xml"
            'xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            'xmlWriter.WriteStartDocument(True)
            'xmlWriter.Formatting = Formatting.Indented
            'xmlWriter.Indentation = 5
            'xmlWriter.WriteStartElement("Calendar")
            cpCon.Open()
            sql = "SELECT * FROM TCMarketing.dbo.SY_CALNDR ORDER BY CALNDR_ID"
            Dim cmd As New SqlCommand(sql, cpCon)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            Dim row As DataRow
            Dim period As Int16 = 0
            Dim theLasteDate, sDate, eDate As Date
            Dim theYear As Integer
            tbl = New DataTable
            tbl.Columns.Add("Year_Id")
            tbl.Columns.Add("Prd_Id")
            tbl.Columns.Add("Week_Id")
            tbl.Columns.Add("sDate")
            tbl.Columns.Add("eDate")
            tbl.Columns.Add("YrPrd")
            tbl.Columns.Add("YrWks")
            tbl.Columns.Add("PrdWk")
            prdTbl = New DataTable
            prdTbl.Columns.Add("Prd", GetType(System.Int32))
            prdTbl.Columns.Add("sDate", GetType(System.DateTime))
            prdTbl.Columns.Add("eDate", GetType(System.DateTime))
            Dim prow As DataRow
            While rdr.Read
                ' Add records for the year range
                period = 0
                prdTbl.Clear()
                theLasteDate = rdr.Item(1)
                theYear = rdr("CALNDR_ID")
                row = tbl.NewRow
                row.Item("Year_Id") = theYear                                                       ' Year
                row.Item("Prd_Id") = 0                                                              ' Period_Id
                row.Item("Week_Id") = 0                                                             ' Week_Id
                sDate = rdr(1)
                eDate = rdr(2)
                row.Item("sDate") = sDate                                                           ' sDate
                row.Item("eDate") = eDate                                                           ' edate
                row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrPrd
                row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
                row.Item("PrdWk") = 0
                tbl.Rows.Add(row)
                ' Add records for period ranges (called Months in CounterPoint)
                For i = 2 To 28 Step 2
                    oTest = rdr.Item(i + 16)
                    If Not IsDBNull(oTest) Then
                        period += 1
                        eDate = CDate(oTest)
                        If i = 2 Then
                            sDate = rdr(1)
                        Else
                            sDate = theLasteDate
                        End If

                        prow = prdTbl.NewRow
                        row = tbl.NewRow
                        row.Item("Year_Id") = rdr("CALNDR_ID")                                      ' Year
                        row.Item("Prd_Id") = period                                                 ' Period_Id
                        row.Item("Week_Id") = 0                                                     ' Week_Id
                        row.Item("sDate") = sDate                                                   ' sDate
                        row.Item("eDate") = eDate                                                   ' edate
                        row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & Format(period, "00")  ' YrPrd
                        row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
                        row.Item("PrdWk") = 0
                        theLasteDate = DateAdd(DateInterval.Day, 1, eDate)
                        tbl.Rows.Add(row)
                        prow.Item("Prd") = period
                        prow.Item("sDate") = sDate
                        prow.Item("eDate") = eDate
                        prdTbl.Rows.Add(prow)
                    End If
                Next
                Dim prevPrd As Integer = 1
                Dim prdWeek As Integer = 1
                Dim week As Integer = 1
                period = 0
                yrwk = 0
                ' Add records for week ranges
                For i = 2 To 108 Step 2
                    oTest = rdr.Item(i + 44)
                    If Not IsDBNull(oTest) Then
                        yrwk += 1
                        eDate = CDate(oTest)
                        If i = 2 Then
                            sDate = rdr(1)
                        Else
                            sDate = DateAdd(DateInterval.Day, 1, theLasteDate)
                        End If                                                                      ' edate
                        period = findPeriod(eDate)
                        If period <> prevPrd Then
                            prevPrd = period
                            prdWeek = 1
                        End If
                        row = tbl.NewRow
                        row.Item("Year_Id") = rdr("CALNDR_ID")                                      ' Year
                        row.Item("Prd_Id") = period                                                 ' Period_Id
                        row.Item("Week_Id") = week                                                  ' Week_Id
                        row.Item("sDate") = sDate                                                   ' sDate
                        row.Item("eDate") = eDate
                        row.Item("YrPrd") = Microsoft.VisualBasic.Right(theYear, 2) & Format(period, "00")                  ' YrPrd
                        row.Item("YrWks") = Microsoft.VisualBasic.Right(theYear, 2) & Format(yrwk, "00")  ' YrWks
                        row.Item("PrdWk") = prdWeek
                        theLasteDate = DateAdd(DateInterval.Day, 1, eDate)
                        tbl.Rows.Add(row)
                        week += 1
                        prdWeek += 1
                        theLasteDate = eDate
                    End If
                Next
                theFirstYear = False
            End While
            cpCon.Close()
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Calendar")
            Dim year, prd, wk, yp, yw, pw As Integer
            Dim sd, ed As Date
            For Each row In tbl.Rows
                oTest = row("Year_Id")
                If Not IsDBNull(oTest) Then

                    xmlWriter.WriteStartElement("Date")
                    xmlWriter.WriteElementString("Year_Id", oTest)
                    oTest = row("Prd_Id")
                    If Not IsDBNull(oTest) Then year = CInt(oTest)
                    xmlWriter.WriteElementString("Prd_Id", oTest)
                    oTest = row("Week_Id")
                    If Not IsDBNull(oTest) Then wk = CInt(oTest)
                    xmlWriter.WriteElementString("Week_Id", oTest)
                    oTest = row("sDate")
                    If Not IsDBNull(oTest) Then sd = CDate(oTest)
                    xmlWriter.WriteElementString("sDate", oTest)
                    oTest = row("eDate")
                    If Not IsDBNull(oTest) Then ed = CDate(oTest)
                    xmlWriter.WriteElementString("eDate", oTest)
                    oTest = row("YrPrd")
                    If Not IsDBNull(oTest) Then yp = CInt(oTest)
                    xmlWriter.WriteElementString("YrPrd", oTest)
                    oTest = row("YrWks")
                    If Not IsDBNull(oTest) Then yw = CInt(oTest)
                    xmlWriter.WriteElementString("YrWks", oTest)
                    oTest = row("PrdWk")
                    If Not IsDBNull(oTest) Then pw = CInt(oTest)
                    xmlWriter.WriteElementString("PrdWk", oTest)
                    xmlWriter.WriteEndElement()
                End If
            Next
            xmlWriter.WriteEndElement()
            XmlWriter.WriteEndDocument()
            xmlWriter.Close()

            


        Catch ex As Exception
            ''Dim el As New CounterpointXMLExtract.ErrorLogger
            ''el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Extracting Calendar")
            ''If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try
    End Sub

    Private Function findPeriod(eDate As Date)
        For Each r In prdtbl.Rows
            Dim val As Integer = 0
            'Dim e As Date = r("edate")
            'Dim w As Integer = r("Week_Id")
            'Dim p As Integer = r("Prd_Id")
            If eDate >= r("sDate") And eDate <= r("eDate") Then
                val = r("Prd")
                Return val
            End If

        Next
    End Function

    Private Function FirstDayOfMonth(ByVal sourceDate As DateTime) As DateTime    
        Return New DateTime(sourceDate.Year, sourceDate.Month, 1)
    End Function
End Module

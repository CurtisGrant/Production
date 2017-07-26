Imports System.Data.SqlClient
Imports System.Xml
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module Module1
    Private con, con2, con3, rcCon, rcCon2 As SqlConnection
    Private cmd As SqlCommand
    Private rdr, rdr2 As SqlDataReader
    Private oTest As Object
    Private server, dbase, xmlPath, sql, conString, conString2 As String
    Public client, errorPath As String
    Private thisMonday, thisSaturday As Date
    Private cnt, wks As Integer
    Sub Main()
        Try
            thisMonday = Date.Today.AddDays(1 - Date.Today.DayOfWeek)
            thisSaturday = Date.Today.AddDays(6 - Date.Today.DayOfWeek)
            Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
            Dim fld As String = ""
            Dim valu As String = ""
            Dim rcServer, rcConString, rcExePath, rcPassword, location, store, typ As String
            Dim qty, cost, retail, mkdn As Decimal
            cnt = 0
            Dim dte As Date
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
            rcConString = "Server=" & rcServer & ";Initial Catalog=RCClient;Integrated Security=True" & ""
            rcCon = New SqlConnection(rcConString)
            rcCon2 = New SqlConnection(rcConString)
            rcCon.Open()
            sql = "SELECT Client_ID, [Server], [Database], XMLs, ErrorLog FROM Client_Master WHERE Status = 'Active'"
            cmd = New SqlCommand(sql, rccon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt = 0
                client = rdr("Client_ID")
                server = rdr("Server")
                dbase = rdr("Database")
                xmlPath = rdr("XMLs")
                errorPath = rdr("ErrorLog")
                conString = "Server=" & server & ";Initial Catalog=" & dbase & ";Integrated Security=True"
                conString2 = conString
                con = New SqlConnection(conString)
                con2 = New SqlConnection(conString)
                con3 = New SqlConnection(conString)
                con.Open()
                Dim ds As DataSet = New DataSet
                Dim dt As DataTable = New DataTable
                Dim xmlFile As XmlReader
                Dim row As DataRow
                Dim thePath As String = xmlPath & "\TransSummary.xml"
                If System.IO.File.Exists(thePath) Then
                    xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
                    ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
                    If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
                        dt = ds.Tables(0)
                        For Each row In dt.Rows
                            cnt += 1
                            If cnt Mod 100 = 0 Then Console.WriteLine(client & " " & cnt)
                            oTest = row("LOCATION")
                            If IsDBNull(oTest) Then location = "" Else location = CStr(oTest)
                            oTest = row("STORE")
                            If IsDBNull(oTest) Then store = "" Else store = CStr(oTest)
                            oTest = row("DATE")
                            If IsDBNull(oTest) Then dte = "1/1/1900" Else dte = CDate(oTest)
                            oTest = row("TYPE")
                            If IsDBNull(oTest) Then typ = "NA" Else typ = CStr(oTest)
                            oTest = row("QTY")
                            If IsDBNull(oTest) Then qty = 0 Else qty = CDec(oTest)
                            oTest = row("COST")
                            If IsDBNull(oTest) Then cost = 0 Else cost = Math.Round(CDec(oTest), 2, MidpointRounding.AwayFromZero)
                            oTest = row("RETAIL")
                            If IsDBNull(oTest) Then retail = 0 Else retail = Math.Round(CDec(oTest), 2, MidpointRounding.AwayFromZero)
                            oTest = row("MARKDOWN")
                            If IsDBNull(oTest) Then mkdn = 0 Else mkdn = CDec(oTest)
                            sql = "IF NOT EXISTS (SELECT * FROM Weekly_Extract_Log WHERE Location = '" & location & "' " & _
                                "AND [Type] = '" & typ & "' AND eDate = '" & dte & "' "
                            If store <> "NA" Then sql &= " AND Store = '" & store & "') "
                            sql &= "INSERT INTO Weekly_Extract_Log(Type, Location, Store, eDate, Qty, Cost, Retail, Markdown, Last_Update) " & _
                                "SELECT '" & typ & "','" & location & "','" & store & "','" & dte & "'," & qty & "," & cost & "," & _
                                retail & "," & mkdn & ",'" & Date.Now & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()





                            ''     Console.WriteLine(sql)




                            wks = DateDiff(DateInterval.WeekOfYear, dte, Date.Today)
                            ''If dte = thisSaturday Then
                            If wks <= 1 Then
                                Call Update_Table(qty, cost, retail, mkdn, location, typ, dte, store)
                            Else
                                If Left(typ, 4) <> "XFER" And Left(typ, 4) <> "XFRE" Then
                                    Call Check_Numbers(location, store, dte, typ, qty, cost, retail, mkdn, cnt)
                                End If

                            End If
                        Next
                    End If
                Else
                    Call Log_Error(client, "", "", thisSaturday, "", "TransSummary.xml WAS NOT FOUND!", "", "")
                End If
                con.Close()
            End While
            rcCon.Close()

            ''Console.WriteLine("Processed " & cnt & " records")
            ''Console.ReadLine()

        Catch ex As Exception
            Dim el As New TransSummaryUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "TransSummaryUpdate Read & Process XML File")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Check_Numbers(ByVal location, ByVal store, ByVal dte, ByVal typ, ByVal qty, ByVal cost, ByVal retail, ByVal mkdn, ByVal cnt)
        Try
            Dim currQty, currCost, currRetail, currMkdn As Decimal
            Dim maxDate As Date
            con2.Open()
            sql = "SELECT Qty, Cost, Retail, Markdown FROM Weekly_Extract_Log WHERE Type = '" & typ & "' " & _
                "AND Location = '" & location & "' AND Store = '" & store & "' AND eDate = '" & dte & "'"
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                oTest = rdr2("Qty")
                If IsDBNull(oTest) Then currQty = 0 Else currQty = CDec(oTest)
                If currQty <> qty Then
                    Call Log_Error(client, location, store, dte, typ, "TransSummary.xml QTY <> Weekly_Extract_Log QTY", CDec(currQty), qty)
                End If
                oTest = rdr2("Cost")
                If IsDBNull(oTest) Then cost = 0 Else currCost = Math.Round(CDec(oTest), 2, MidpointRounding.AwayFromZero)
                If currCost <> cost Then
                    Call Log_Error(client, location, store, dte, typ, "TransSummary.xml COST <> Weekly_Extract_Log COST", CDec(currCost), cost)
                End If
                oTest = rdr2("Retail")
                If IsDBNull(oTest) Then currRetail = 0 Else currRetail = Math.Round(CDec(oTest), 2, MidpointRounding.AwayFromZero)
                If currRetail <> retail Then
                    Call Log_Error(client, location, store, dte, typ, "TransSummary.xml RETAIL <> Weekly_Extract_Log RETAIL", CDec(currRetail), retail)
                End If
                oTest = rdr2("Markdown")
                If IsDBNull(oTest) Then currMkdn = 0 Else currMkdn = CDec(oTest)
                If currMkdn <> mkdn Then
                    Call Log_Error(client, location, store, dte, typ, "TransSummary.xml MKDN <> Weekly_Extract_Log MKDN", CDec(currMkdn), mkdn)
                End If
            End While
            con2.Close()

            con2.Open()
            Select Case typ
                Case "ADJ"
                    sql = "SELECT SUM(ISNULL(QTY,0)) QTY, CONVERT(Decimal(10,2),SUM(ISNULL(QTY * COST,0))) COST, " & _
                        "CONVERT(Decimal(10,2),SUM(ISNULL(QTY * RETAIL,0))) RETAIL, 0 AS MKDN, " & _
                        "DATEADD(Day,7-DATEPART(Weekday, MAX(POST_DATE)),MAX(POST_DATE)) MAXDATE FROM Daily_Transaction_Log l " & _
                        "LEFT JOIN Item_Master m ON m.SKU = l.SKU AND m.[TYPE] = 'I' " & _
                        "WHERE l.[TYPE] = 'ADJ' AND LEFT(TRANS_ID,1) <> 'P' AND LOCATION = '" & location & "' " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN DATEADD(Day,1-DATEPART(Weekday, '" & dte & "'),'" & dte & "') " & _
                        "AND DATEADD(Day,7-DATEPART(Weekday, '" & dte & "'),'" & dte & "')"
                Case "Physical"
                    sql = "SELECT SUM(ISNULL(QTY,0)) QTY, CONVERT(Decimal(10,2),SUM(ISNULL(QTY * COST,0))) COST, " & _
                        "CONVERT(Decimal(10,2),SUM(ISNULL(QTY * RETAIL,0))) RETAIL, 0 AS MKDN, " & _
                        "DATEADD(Day,7-DATEPART(Weekday, MAX(POST_DATE)),MAX(POST_DATE)) MAXDATE FROM Daily_Transaction_Log l " & _
                        "LEFT JOIN Item_Master m ON m.SKU = l.SKU AND m.[TYPE] = 'I' " & _
                        "WHERE l.[TYPE] = 'ADJ' AND LEFT(TRANS_ID,1) = 'P' AND LOCATION = '" & location & "' " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN DATEADD(Day,1-DATEPART(Weekday, '" & dte & "'),'" & dte & "') " & _
                        "AND DATEADD(Day,7-DATEPART(Weekday, '" & dte & "'),'" & dte & "')"
                Case "Receipt"
                    sql = "SELECT SUM(ISNULL(QTY,0)) QTY, CONVERT(Decimal(10,2),SUM(ISNULL(QTY * COST,0))) COST, " & _
                        "CONVERT(Decimal(10,2),SUM(ISNULL(QTY * RETAIL,0))) RETAIL, 0 AS MKDN, " & _
                        "DATEADD(Day,7-DATEPART(Weekday, MAX(POST_DATE)),MAX(POST_DATE)) MAXDATE FROM Daily_Transaction_Log l " & _
                        "LEFT JOIN Item_Master m ON m.SKU = l.SKU AND m.[TYPE] = 'I' " & _
                        "WHERE l.[TYPE] = 'RECVD' AND LOCATION = '" & location & "' " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN DATEADD(Day,1-DATEPART(Weekday, '" & dte & "'),'" & dte & "') " & _
                        "AND DATEADD(Day,7-DATEPART(Weekday, '" & dte & "'),'" & dte & "')"
                Case "RTV"
                    sql = "SELECT SUM(ISNULL(QTY,0)) QTY, SUM(ISNULL(QTY * COST,0)) COST, " & _
                        "SUM(ISNULL(QTY * RETAIL,0)) RETAIL, 0 AS MKDN, " & _
                        "DATEADD(Day,7-DATEPART(Weekday, MAX(POST_DATE)),MAX(POST_DATE)) MAXDATE FROM Daily_Transaction_Log l " & _
                        "LEFT JOIN Item_Master m ON m.SKU = l.SKU AND m.[TYPE] = 'I' " & _
                        "WHERE l.[TYPE] = 'RTV' AND LOCATION = '" & location & "' " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN DATEADD(Day,1-DATEPART(Weekday, '" & dte & "'),'" & dte & "') " & _
                        "AND DATEADD(Day,7-DATEPART(Weekday, '" & dte & "'),'" & dte & "')"
                Case "Sales"
                    sql = "SELECT SUM(ISNULL(QTY,0)) QTY, CONVERT(Decimal(10,2),SUM(ISNULL(QTY * COST,0))) COST, " & _
                        "CONVERT(Decimal(10,2),SUM(ISNULL(QTY * RETAIL,0))) RETAIL, SUM(ISNULL(MKDN,0)) MKDN, " & _
                        "DATEADD(Day,7-DATEPART(Weekday, MAX(POST_DATE)),MAX(POST_DATE)) MAXDATE FROM Daily_Transaction_Log l " & _
                        "WHERE l.[TYPE] = 'Sold' AND LOCATION = '" & location & "' AND STORE = '" & store & "' " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN DATEADD(Day,1-DATEPART(Weekday, '" & dte & "'),'" & dte & "') " & _
                        "AND DATEADD(Day,7-DATEPART(Weekday, '" & dte & "'),'" & dte & "')"
            End Select
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                oTest = rdr2("MAXDATE")
                If IsDBNull(oTest) Then
                    Call Log_Error(client, location, store, dte, typ, "POST DATE '" & dte & "' IS NULL", 0, 0)
                    GoTo 25
                End If
                maxDate = CDate(oTest)
                oTest = rdr2("QTY")
                If IsDBNull(oTest) Then currQty = 0 Else currQty = CDec(oTest)
                If currQty <> qty Then
                    Call Log_Error(client, location, store, dte, typ, "XMLExtract.xml QTY <> Daily_Transaction_Log QTY", currQty, qty)
                    Call Update_Table(qty, cost, retail, mkdn, location, typ, dte, store)
                End If
                oTest = rdr2("COST")
                If IsDBNull(oTest) Then currCost = 0 Else currCost = oTest
                If currCost <> cost And maxDate < thisSaturday Then
                    Call Log_Error(client, location, store, dte, typ, "XMLExtract.xml COST <> Daily_Transaction_Log COST", currCost, cost)
                    Call Update_Table(qty, cost, retail, mkdn, location, typ, dte, store)
                End If
                oTest = rdr2("RETAIL")
                If IsDBNull(oTest) Then currRetail = 0 Else currRetail = oTest
                If currRetail <> retail And maxDate < thisSaturday Then
                    Call Log_Error(client, location, store, dte, typ, "XMLExtract.xml RETAIL <> Daily_Transaction_Log COST", currRetail, retail)
                    Call Update_Table(qty, cost, retail, mkdn, location, typ, dte, store)
                End If
                oTest = rdr2("MKDN")
                If IsDBNull(oTest) Then currMkdn = 0 Else currMkdn = CDec(oTest)
                If currMkdn <> mkdn And maxDate < thisSaturday Then
                    Call Log_Error(client, location, store, dte, typ, "XMLExtract.xml MKDN <> Daily_Transaction_Log MKDN", currMkdn, mkdn)
                    Call Update_Table(qty, cost, retail, mkdn, location, typ, dte, store)
                End If
25:         End While
            con2.Close()
        Catch ex As Exception
            Dim el As New TransSummaryUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "TransSummaryUpdate Check Numbers " & typ & " " & sql)
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Update_Table(ByVal qty, ByVal cost, ByVal retail, ByVal mkdn, ByVal location, ByVal typ, ByVal dte, ByVal store)
        con3.Open()
        sql = "UPDATE Weekly_Extract_Log SET Qty = " & qty & ", Cost = " & cost & ", Retail = " & retail & ", " & _
              "Markdown = " & mkdn & ", Last_Update = '" & Date.Now & "' " & _
              "WHERE Location = '" & location & "' AND [Type] = '" & typ & "' AND eDate = '" & dte & "' "
        If store <> "NA" Then sql &= " AND Store = '" & store & "'"
        cmd = New SqlCommand(sql, con3)
        cmd.ExecuteNonQuery()
        con3.Close()
    End Sub

    Private Sub Log_Error(ByVal client, ByVal location, ByVal store, ByVal dte, ByVal typ, ByVal msg, ByVal oldval, ByVal newval)
        Try
            rcCon2.Open()
            sql = "IF EXISTS (SELECT ID FROM Extract_Error_Log WHERE Client_Id = '" & client & "' AND Store = '" & store & "' " & _
                "AND eDate = '" & dte & "' AND TransType = '" & typ & "' AND Error = '" & msg & "' AND OrigValue = '" & oldval & "' " & _
                "AND NewValue = '" & newval & "') " & _
                "UPDATE Extract_Error_Log SET LastErrorDate = '" & Date.Now & "' WHERE Client_Id = '" & client & "' AND Store = '" & store & "' " & _
                "AND eDate = '" & dte & "' AND TransType = '" & typ & "' AND Error = '" & msg & "' " & _
                "ELSE " & _
                "INSERT INTO Extract_Error_Log(Client_Id, Location, Store, eDate, TransType, Error, OrigValue, NewValue, OrigErrorDate, LastErrorDate) " & _
                "SELECT '" & client & "','" & location & "','" & store & "','" & dte & "','" & typ & "','" & msg & "','" &
                oldval & "','" & newval & "','" & Date.Now & "','" & Date.Now & "' "
            cmd = New SqlCommand(sql, rcCon2)
            cmd.ExecuteNonQuery()
            rcCon2.Close()
        Catch ex As Exception
            Dim el As New TransSummaryUpdate.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "TransSummaryUpdate Log Error " & typ & " " & sql)
            If rcCon2.State = ConnectionState.Open Then rcCon2.Close()
        End Try
    End Sub
End Module

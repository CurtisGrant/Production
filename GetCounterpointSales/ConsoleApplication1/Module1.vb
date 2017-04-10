Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Private masterCon, con As SqlConnection
    Sub Main()
        Try

            Dim dayofweek As Integer = Date.Now.DayOfWeek
            If dayofweek = 0 Then Exit Sub
            Dim dt1 As DateTime = Date.Now
            Dim timeofday = dt1.TimeOfDay
            Dim oTest As Object
            Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RCCLIENT.xml")
            Dim server, database, userId, conString, mConString, exePath, passWord, client As String
            Dim fld As String = ""
            Dim valu As String = ""
            Dim thisClient As String
            Dim sql, dept As String
            Dim xmlPath As String
            Dim cmd As SqlCommand
            Dim rdr As SqlDataReader
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
            mConString = "Server=" & server & ";Initial Catalog=RCClient;User Id=sa;Password=" & passWord & ""
            masterCon = New SqlConnection(mConString)
            masterCon.Open()
            sql = "SELECT Client_Id, Server, [Database], SQLUserID, SQLPassword, XMLs FROM Client_Master WHERE Client_Id = '" & client & "'"
            cmd = New SqlCommand(sql, masterCon)
            Dim mrdr As SqlDataReader
            mrdr = cmd.ExecuteReader()
            While mrdr.Read
                thisClient = mrdr("Client_Id")
                server = mrdr("Server")
                database = mrdr("Database")
                userId = mrdr("SQLUserID")
                passWord = mrdr("SQLPassword")
                xmlPath = mrdr("XMLs")
            End While
            masterCon.Close()

            conString = "Server=" & server & ";Initial Catalog=" & database & ";User Id=" & userId & ";Password=" & passWord
            con = New SqlConnection(conString)
            Dim tcConString As String = ""

            Dim closeTime As DateTime = Date.Today & " 7:00:00 PM"                        ' default closing time
            Dim openTime As DateTime = Date.Today & " 9:00:00 AM"                         ' default open time
            Dim oTime = openTime.TimeOfDay
            Dim cTime = closeTime.TimeOfDay
            con.Open()
            sql = "SELECT Parameter, Value FROM Controls WHERE ID = 'StoreHours'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Value")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If rdr("Parameter") = "Close" Then closeTime = oTest Else openTime = oTest
                End If
            End While
            con.Close()

            If timeofday < oTime Or timeofday > cTime Then Exit Sub

            con.Open()
            sql = "SELECT Value FROM Controls WHERE ID = 'TimeClock' AND Parameter = 'conString'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Value")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then tcConString = oTest
            End While
            con.Close()

            con = New SqlConnection(tcConString)
            con.Open()
            Dim storeCloses As DateTime = Date.Today & " " & closeTime
            Dim var As String = Now.ToString("HH:mm:ss")
            Dim code As Integer
            Dim minutes As Int32
            Dim dte As Date
            Dim rightNow As DateTime = Date.Now
            Dim path As String = xmlPath & "\DayTimeWorked.xml"
            Dim xmlWriter As XmlTextWriter
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("Minutes")
            sql = "DECLARE @now AS DateTime " & _
               "DECLARE @closeToday AS DateTime " & _
               "SET @closeToday = '" & storeCloses & "' " & _
               "SET @now = (SELECT CASE WHEN @now >= @closeToday THEN @closeToday ELSE GETDATE() END) " & _
               "SELECT l.EmployeeId, Department, CASE WHEN TimeOut IS NULL THEN DATEDIFF(Minute,TimeIn, @now) " & _
               "ELSE DATEDIFF(Minute,TimeIn, TimeOut) END AS Worked INTO #t1 FROM tcp_EmployeeWork.Worksegment h " & _
               "JOIN tcp_Employee.Employee l ON l.RecordId = h.EmployeeRecordId " & _
               "WHERE  CAST(CONVERT(VARCHAR(16),TimeIn,111) AS DATE) = CAST(GETDATE() AS Date) " & _
               "SELECT Department, ISNULL(SUM(Worked),0) AS Minutes FROM #t1 GROUP BY Department"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                'dte = rdr("Date")
                dte = Date.Today
                dept = rdr("Department")
                If dept <> "" Then
                    minutes = rdr("Minutes")
                    xmlWriter.WriteStartElement("Department")
                    xmlWriter.WriteElementString("Date", dte)
                    xmlWriter.WriteElementString("Dept", dept)
                    xmlWriter.WriteElementString("Minutes", minutes)
                    xmlWriter.WriteElementString("LastUpdate", rightNow)
                    xmlWriter.WriteEndElement()
                End If
            End While
            con.Close()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            con = New SqlConnection(conString)
            con.Open()
            path = xmlPath & "\CounterpointSales.xml"
            Dim DateTimeString As String = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
            Dim qty, cost, retail, totlQty, totlCost, totlRetail, totlMarkdown As Decimal
            Dim item, itemString, dim1, dim2, dim3, location, store, drawer As String
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 2
            xmlWriter.WriteStartElement("Sales")
            sql = "SELECT ISNULL(l.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, ISNULL(l.ITEM_NO,'NA') AS ITEM, " & _
                "DIM_1_UPR AS DIM1, DIM_2_UPR AS DIM2, DIM_3_UPR AS DIM3, l.STK_LOC_ID AS LOCATION, DRW_ID AS DRAWER, " & _
               "ISNULL(l.STR_ID,'NA') AS STORE, ISNULL(l.QTY_SOLD,0) AS QTY,  ISNULL(l.UNIT_COST,0) AS COST, " & _
               "ISNULL(l.PRC,0) AS RETAIL, ISNULL(h.TKT_DT,'1/1/1900') AS DATE, ISNULL(((l.QTY_SOLD * l.PRC_1) - l.EXT_PRC),0) AS MARKDOWN, " & _
               "PRC_OVRD_REAS AS MKDN_REASON, CONVERT(varchar(30),'') AS COUPON_CODE, " & _
               "ISNULL(l.CATEG_COD,'NA') AS DEPT, ISNULL(l.SUBCAT_COD,'NA') AS CLASS, ISNULL(pv.LST_ORD_BUYER,'NA') AS BUYER, CUST_NO, l.TKT_NO " & _
               "INTO #t1 FROM CPSQL.dbo.PS_DOC_LIN AS l " & _
               "INNER JOIN CPSQL.dbo.PS_DOC_HDR AS h ON h.DOC_ID = l.DOC_ID " & _
               "LEFT JOIN CPSQL.dbo.PO_VEND_ITEM AS pv ON pv.ITEM_NO = l.ITEM_NO AND pv.VEND_NO =  l.ITEM_VEND_NO " & _
               "JOIN CPSQL.dbo.IM_ITEM i ON i.ITEM_NO = l.ITEM_NO " & _
               "WHERE h.DOC_TYP = 'T' AND l.LIN_TYP IN ('S','A','R') AND DIM_1_UPR IS NOT NULL " & _
               "AND NOT EXISTS (SELECT * FROM CPSQL.dbo.PS_DOC_AUDIT_LOG a WHERE a.DOC_ID = l.DOC_ID AND a.DOC_TYP = 'V') " & _
               "INSERT INTO #t1 (TRANS_ID,SEQ_NO, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, " & _
               "STORE, QTY,COST, RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO) " & _
               "SELECT ISNULL(l.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, ISNULL(l.ITEM_NO,'NA') AS ITEM, " & _
                "DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, l.STK_LOC_ID AS LOCATION, DRW_ID AS DRAWER, " & _
               "ISNULL(l.STR_ID,'NA') AS STORE, ISNULL(l.QTY_SOLD,0) AS QTY, ISNULL(l.UNIT_COST,0) AS COST, ISNULL(l.PRC,0) AS RETAIL, " & _
               "ISNULL(h.TKT_DT,'1/1/1900') AS DATE, ISNULL(((l.QTY_SOLD * l.PRC_1) - l.EXT_PRC),0) AS MARKDOWN, l.PRC_OVRD_REAS, NULL, " & _
               "ISNULL(l.CATEG_COD,'NA') AS DEPT, ISNULL(l.SUBCAT_COD,'NA') AS CLASS, ISNULL(pv.LST_ORD_BUYER,'NA') AS BUYER, CUST_NO, h.TKT_NO " & _
              "FROM CPSQL.dbo.PS_TKT_HIST_LIN AS l " & _
              "JOIN CPSQL.dbo.PS_TKT_HIST h ON h.DOC_ID = l.DOC_ID " & _
              "INNER JOIN CPSQL.dbo.PS_TKT_HIST_EXT AS e ON e.DOC_ID_EXT = l.DOC_ID " & _
              "LEFT JOIN CPSQL.dbo.PO_VEND_ITEM AS pv ON pv.ITEM_NO = l.ITEM_NO AND pv.VEND_NO =  l.ITEM_VEND_NO " & _
              "JOIN CPSQL.dbo.IM_ITEM i ON i.ITEM_NO = l.ITEM_NO " & _
              "WHERE l.DIM_1_UPR IS NOT NULL AND DIM_1_UPR <> '*'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            sql = "SELECT TRANS_ID,SEQ_NO, ITEM, DIM1, DIM2, DIM3, LOCATION, DRAWER, STORE, QTY, COST, " & _
                "RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO FROM #t1"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                item = ""
                itemString = ""
                dim1 = ""
                dim2 = "*"
                dim3 = "*"
                xmlWriter.WriteStartElement("SALE")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TRANS_ID", oTest)
                oTest = rdr("SEQ_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("SEQ_NO", oTest)
                oTest = rdr("ITEM")
                If IsDBNull(oTest) Then oTest = "NA"
                item = oTest
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then
                        dim1 = oTest
                    End If
                End If
                oTest = rdr("DIM2")
                If Not IsDBNull(oTest) Then
                    If dim2 <> "*" Then dim2 = oTest
                End If
                oTest = rdr("DIM3")
                If Not IsDBNull(oTest) Then
                    If oTest <> "*" Then dim3 = oTest
                End If
                If dim1 <> "" Then
                    itemString = item & "~" & dim1 & "~" & dim2 & "~" & dim3
                Else : itemString = item
                End If
                xmlWriter.WriteElementString("ITEM_NO", itemString)
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then location = "1" Else location = oTest
                xmlWriter.WriteElementString("LOCATION", location)
                oTest = rdr("DRAWER")
                If IsDBNull(oTest) Then drawer = "1" Else drawer = oTest
                xmlWriter.WriteElementString("DRAWER", drawer)
                oTest = rdr("STORE")
                If IsDBNull(oTest) Then store = "1" Else store = oTest
                xmlWriter.WriteElementString("STORE", store)
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
                oTest = rdr("MARKDOWN")
                If IsDBNull(oTest) Then oTest = 0
                totlMarkdown += CDec(oTest)
                xmlWriter.WriteElementString("MARKDOWN", oTest)
                oTest = rdr("DEPT")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("DEPT", oTest)
                oTest = rdr("CLASS")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("CLASS", oTest)
                oTest = rdr("BUYER")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("BUYER", oTest)
                oTest = rdr("CUST_NO")
                If IsDBNull(oTest) Then oTest = "U"
                xmlWriter.WriteElementString("CUST_NO", oTest)
                oTest = rdr("TKT_NO")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("TKT_NO", oTest)
                oTest = rdr("MKDN_REASON")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("MKDN_REASON", oTest)
                oTest = rdr("COUPON_CODE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("COUPON_CODE", oTest)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Now)
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteStartElement("SALE")
            xmlWriter.WriteElementString("TRANS_ID", "SALES")
            xmlWriter.WriteElementString("SEQ_NO", 0)
            xmlWriter.WriteElementString("ITEM_NO", "TOTALS")
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
            con.Close()

            Dim p As New ProcessStartInfo
            p.FileName = exePath & "\UpdateSalesTime.exe"
            p.UseShellExecute = True
            p.WindowStyle = ProcessWindowStyle.Normal
            Dim proc As Process = Process.Start(p)

        Catch ex As Exception
            If masterCon.State = ConnectionState.Open Then masterCon.Close()
            If con.State = ConnectionState.Open Then con.Close()
            Dim el As New GetCounterpointSales.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "GetSalesTime,Main")
        End Try
    End Sub

End Module

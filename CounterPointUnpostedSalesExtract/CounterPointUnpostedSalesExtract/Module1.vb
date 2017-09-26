Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Public errorPath As String
    Private path, xmlPath, conString, theProcess As String
    Private cpCon As SqlConnection
    Private xmlReader As XmlTextReader
    Private xmlWriter As XmlTextWriter
    Private minSalDate As Date
    Private SALwks As Integer
    Private thisSdate, thisEdate As Date

    Sub Main()
        Try
            Dim fld, valu, server, database, client, user, password As String
            Dim ADJwks As Integer = 2
            Dim RCPTwks As Integer = 2
            Dim RTNwks As Integer = 2
            Dim SALwks As Integer = 2
            Dim XFERwks As Integer = 2
            Dim POwks As Integer = 2
            Dim cpServer, cpDatabase, cpUserID, cpPassword As String
            Dim logPath As String = ""
            xmlReader = New XmlTextReader("c:\RetailClarity\RCExtract.xml")
            Dim rcExePath As String
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
            '
            '              change error path to same folder as xmls & delete it first thing
            '
            errorPath = xmlPath
            path = xmlPath & "\errlog.txt"
            If System.IO.File.Exists(path) Then System.IO.File.Delete(path)

            conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & ";Integrated Security=SSPI;Connection Timeout=60;"
        Catch ex As Exception
            Dim el As New CounterPointUnpostedSalesExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Inital set up")
        End Try

        Try
            Dim oTest As Object
            cpCon = New SqlConnection(conString)
            Dim cmd As SqlCommand
            Dim rdr As SqlDataReader
            Dim cnt As Integer
            Dim cost, qty, retail, totlCost, totlQty, totlRetail, mkdn As Decimal
            Dim item, itemString, sql As String
            cnt = 0
            totlQty = 0
            totlCost = 0
            totlRetail = 0
            thisSdate = Date.Today.AddDays(0 - Date.Today.DayOfWeek)
            minSALdate = DateAdd(DateInterval.Day, SALwks * -14, thisSdate)
            Console.WriteLine("Extracting Unposted Sales")
            theProcess = "Set up XML writer for Header Data"
            path = xmlPath & "\UnpostedSalesHDR.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("UnpostedSale")
            theProcess = "Open sql connection for header data"
            cpCon.Open()
            theProcess = "Select header data from database"
            sql = "CREATE TABLE #hdr (TRANS_ID varchar(30), STORE varchar(10), LOCATION varchar(10), STATION varchar(10) NULL, " &
                "DRAWER varchar(10) NULL, DATE datetime, CUST_NO varchar(30) NULL, CUST_TYPE varchar(10) NULL, DOC_TYPE varchar(10) NULL, " &
                "TKT_NO varchar(30) NULL, SALES_REP varchar(30), SHIP_COUNTRY varchar(30), AMOUNT_DUE Decimal(18, 4), TERMS varchar(30), " &
                "AMOUNT_APPLIED decimal(18,4), STATUS varchar(1), ORDER_TOTAL decimal(18,4)) " &
                "INSERT INTO #hdr (TRANS_ID, STORE, LOCATION, STATION, DRAWER, Date, CUST_NO, CUST_TYPE, DOC_TYPE, TKT_NO, SALES_REP, " &
                "SHIP_COUNTRY, AMOUNT_DUE, TERMS, AMOUNT_APPLIED, STATUS, ORDER_TOTAL) " &
                "SELECT ISNULL(h.DOC_ID,'NA') AS TRANS_ID, ISNULL(h.STR_ID,'NA') STR_ID, ISNULL(h.STK_LOC_ID,'') LOCATION, " &
                "ISNULL(STA_ID,'') STATION, ISNULL(DRW_ID,'') DRAWER, ISNULL(h.TKT_DT,'1/1/1900') DATE, ISNULL(h.CUST_NO,'') CUST_NO, " &
                "ISNULL(ca.CUST_TYP,'U') CUST_TYP, ISNULL(h.DOC_TYP,'U') DOC_TYP, ISNULL(h.TKT_NO,'') TKT_NO, ISNULL(h.SLS_REP,'') " &
                "SALES_REP, ISNULL(h.SHIP_CNTRY,'') SHIP_COUNTRY, ISNULL(h.ORD_AMT_DUE,0) AMOUNT_DUE, ISNULL(h.TERMS_COD, '') TERMS, " &
                "ISNULL(h.DEP_AMT_APPLIED, 0) AMT_APPLIED, ISNULL(DOC_STAT,'U') DOC_STAT, ISNULL(h.ORD_TOT,0) ORDER_TOTAL FROM VI_PS_DOC_HDR AS h " &
                "LEFT JOIN VI_AR_CUST_WITH_ADDRESS ca ON ca.CUST_NO = h.CUST_NO " &
                "WHERE NOT EXISTS (SELECT * FROM VI_PS_DOC_AUDIT_LOG a WHERE a.DOC_ID = h.DOC_ID " &
                "AND a.DOC_TYP = 'V') " &
                "SELECT TRANS_ID,STORE, LOCATION, STATION, DRAWER, Date, CUST_NO, CUST_TYPE, DOC_TYPE, TKT_NO, SALES_REP, SHIP_COUNTRY, " &
                "AMOUNT_DUE, TERMS, AMOUNT_APPLIED, STATUS, ORDER_TOTAL FROM #hdr"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 960
            theProcess = "Read sql data for headers"
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("SALE")
                oTest = rdr("TRANS_ID")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("TRANS_ID", oTest)
                oTest = rdr("STORE")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("STR_ID", oTest)
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("LOCATION", oTest)
                oTest = rdr("STATION")
                xmlWriter.WriteElementString("STATION", oTest)
                oTest = rdr("DRAWER")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("DRAWER", oTest)
                oTest = rdr("DATE")
                xmlWriter.WriteElementString("TRANS_DATE", oTest)
                oTest = rdr("CUST_NO")
                xmlWriter.WriteElementString("CUST_NO", oTest)
                oTest = rdr("CUST_TYPE")
                xmlWriter.WriteElementString("CUST_TYPE", oTest)
                oTest = rdr("DOC_TYPE")
                xmlWriter.WriteElementString("TRANS_TYPE", oTest)
                oTest = rdr("TKT_NO")
                xmlWriter.WriteElementString("TKT_NO", oTest)
                oTest = rdr("SALES_REP")
                xmlWriter.WriteElementString("SALES_REP", oTest)
                oTest = rdr("SHIP_COUNTRY")
                xmlWriter.WriteElementString("SHIP_COUNTRY", oTest)
                oTest = rdr("AMOUNT_DUE")
                If IsNumeric(oTest) Then oTest = CDec(oTest) Else oTest = 0
                xmlWriter.WriteElementString("AMOUNT_DUE", CDec(oTest))
                oTest = rdr("TERMS")
                xmlWriter.WriteElementString("TERMS", oTest)
                oTest = rdr("AMOUNT_APPLIED")
                If IsNumeric(oTest) Then oTest = CDec(oTest) Else oTest = 0
                xmlWriter.WriteElementString("AMT_APPLIED", CDec(oTest))
                oTest = rdr("STATUS")
                xmlWriter.WriteElementString("STATUS", oTest)
                oTest = rdr("ORDER_TOTAL")
                xmlWriter.WriteElementString("ORDER_TOTAL", CDec(oTest))
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
                xmlWriter.WriteEndElement()
            End While
            theProcess = "Close sql connection and XML writer for header"
            cpCon.Close()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()

            theProcess = "Open sql connection for detail data"
            cpCon.Open()
            theProcess = "Set up XML writer for detail data"
            path = xmlPath & "\UnpostedSalesDTL.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("UnpostedSale")
            theProcess = "Select detail data from database"
            sql = "CREATE TABLE #dtl (TRANS_ID varchar(30), SEQ_NO int, STORE varchar(10), LOCATION varchar(10), QTY decimal(18,4) NULL, " &
                "ITEM varchar(30), DIM1 varchar(30), DIM2 varchar(30), dim3 varchar(30), COST decimal(18,4) NULL, RETAIL decimal(18,4) NULL, " & _
                "DATE datetime, MARKDOWN decimal(18,4) NULL, MKDN_REASON varchar(30) NULL, COUPON_CODE varchar(30) NULL, " &
                "TKT_NO varchar(30) NULL, TRANS_TYPE varchar(1), LINE_TYPE varchar(1)) " &
                "INSERT INTO #dtl (TRANS_ID, SEQ_NO, STORE, LOCATION, ITEM, DIM1, DIM2, DIM3, QTY, COST, RETAIL, DATE, MARKDOWN, " &
                "MKDN_REASON, COUPON_CODE, TKT_NO, TRANS_TYPE, LINE_TYPE) " &
                "SELECT ISNULL(l.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO,  ISNULL(l.STR_ID,'NA') STR_ID, " &
                "ISNULL(l.STK_LOC_ID,'') LOCATION, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, " &
                "CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) ELSE ISNULL(c.QTY_SOLD,0) END QTY,  ISNULL(i.AVG_COST,0) COST, " &
                "ISNULL(l.PRC,0) - ISNULL(l.LIN_DISC_AMT,0) RETAIL, ISNULL(h.TKT_DT,'1/1/1900') DATE, " &
                "(ISNULL(l.PRC_1,0) - ISNULL(l.PRC,0) + ISNULL(l.LIN_DISC_AMT,0)) MARKDOWN,  " &
                "ISNULL(PRC_OVRD_REAS,'') MKDN_REASON, CONVERT(varchar(30),'') COUPON_CODE, ISNULL(l.TKT_NO,'') TKT_NO, " &
                "h.DOC_TYP, l.LIN_TYP FROM VI_PS_DOC_LIN AS l " &
                "INNER JOIN VI_PS_DOC_HDR AS h ON h.DOC_ID = l.DOC_ID " &
                "LEFT JOIN VI_IM_ITEM_WITH_INV i ON i.ITEM_NO = l.ITEM_NO AND i.LOC_ID = l.STK_LOC_ID " &
                "LEFT JOIN VI_PS_DOC_LIN_CELL c ON c.DOC_ID = l.DOC_ID AND c.LIN_SEQ_NO = l.LIN_SEQ_NO " &
                "WHERE NOT EXISTS (SELECT * FROM VI_PS_DOC_AUDIT_LOG a WHERE a.DOC_ID = l.DOC_ID And a.DOC_TYP = 'V')  " &
                "SELECT DOC_ID, LIN_SEQ_NO, PROMPT_ALPHA_1 INTO #us2 FROM VI_PS_DOC_LIN WHERE PROMPT_COD_1 = 'MKTG_CODE' " &
                "UPDATE t1 SET t1.COUPON_CODE = PROMPT_ALPHA_1 FROM #dtl t1 " &
                "JOIN #us2 t2 ON t2.DOC_ID = t1.TRANS_ID AND t2.LIN_SEQ_NO = SEQ_NO " &
                "SELECT TRANS_ID, SEQ_NO, STORE, LOCATION, ITEM, DIM1, DIM2, DIM3, QTY, COST, RETAIL, DATE, " &
                "CONVERT(Decimal(18,4),(QTY * MARKDOWN)) MARKDOWN, MKDN_REASON,  COUPON_CODE, TKT_NO, TRANS_TYPE, LINE_TYPE FROM #dtl"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 960
            theProcess = "Read sql data for detail records"
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("SALE")
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
                oTest = rdr("LOCATION")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("LOCATION", oTest)
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
                xmlWriter.WriteElementString("MKDN_REASON", CStr(oTest))
                oTest = rdr("COUPON_CODE")
                xmlWriter.WriteElementString("COUPON_CODE", oTest)
                oTest = rdr("TKT_NO")
                xmlWriter.WriteElementString("TKT_NO", oTest)
                xmlWriter.WriteElementString("VENDOR_ID", oTest)
                oTest = rdr("TRANS_TYPE")
                xmlWriter.WriteElementString("TRANS_TYPE", oTest)
                oTest = rdr("LINE_TYPE")
                xmlWriter.WriteElementString("LINE_TYPE", oTest)
                totlQty += qty
                totlCost += (qty * cost)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
                xmlWriter.WriteEndElement()
            End While
            theProcess = "Close sql connection for detail data"
            cpCon.Close()
            theProcess = "Write total row to XML file"
            xmlWriter.WriteStartElement("SALE")
            xmlWriter.WriteElementString("TRANS_ID", "")
            xmlWriter.WriteElementString("SEQ_NO", "")
            xmlWriter.WriteElementString("STORE", "")
            xmlWriter.WriteElementString("SKU", "TOTALS")
            xmlWriter.WriteElementString("LOCATION", "")
            xmlWriter.WriteElementString("QTY", totlQty)
            xmlWriter.WriteElementString("COST", totlCost)
            xmlWriter.WriteElementString("RETAIL", totlRetail)
            xmlWriter.WriteElementString("TRANS_DATE", "")
            xmlWriter.WriteElementString("MARKDOWN", "")
            xmlWriter.WriteElementString("MKDN_REASON", "")
            xmlWriter.WriteElementString("COUPON_CODE", "")
            xmlWriter.WriteElementString("TKT_NO", "")
            xmlWriter.WriteElementString("VENDOR_ID", "")
            xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New CounterPointUnpostedSalesExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "CounterpointUnpostedSalesExtract " & theProcess)
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try
    End Sub

End Module

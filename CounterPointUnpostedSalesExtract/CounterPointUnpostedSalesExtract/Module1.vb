Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Public errorPath As String
    Private path, xmlPath, conString As String
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
            '
            '
            '

            conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & ";Integrated Security=True;Connection Timeout=30;"
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
            cpCon.Open()
            path = xmlPath & "\DailySales.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("UnpostedSale")
            sql = "CREATE TABLE #us (TRANS_ID varchar(30), SEQ_NO int, STORE varchar(10), SKU varchar(90), DIM1 varchar(30) NULL, " & _
                "DIM2 varchar(30) NULL, DIM3 varchar(30) NULL, LOCATION varchar(10), DRAWER varchar(10) NULL,  " & _
                "QTY decimal(18,4) NULL, COST decimal(18,2) NULL, RETAIL decimal(18,2) NULL, DATE datetime, MARKDOWN decimal(18,2) NULL, " & _
                "MKDN_REASON varchar(30) NULL,COUPON_CODE varchar(30) NULL, DEPT varchar(10) NULL, CLASS varchar(10) NULL, " & _
                "BUYER varchar(10) NULL, CUST_NO varchar(30) NULL, CUST_TYPE varchar(10) NULL, TRANS_TYPE varchar(10) NULL, " & _
                "TKT_NO varchar(30) NULL) " & _
                "INSERT INTO #us (TRANS_ID,SEQ_NO, STORE, SKU, DIM1, DIM2, DIM3, LOCATION, DRAWER, " & _
                "QTY,COST, RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, TKT_NO) " & _
                "SELECT ISNULL(l.DOC_ID,'NA') AS TRANS_ID, ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, ISNULL(l.STR_ID,'NA'), " & _
                "ISNULL(l.ITEM_NO,'NA') AS ITEM, c.DIM_1_UPR AS DIM1, c.DIM_2_UPR AS DIM2, c.DIM_3_UPR AS DIM3, l.STK_LOC_ID AS LOCATION, " & _
                "DRW_ID AS DRAWER, CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) " & _
                "ELSE c.QTY_SOLD END AS QTY, ISNULL(l.UNIT_COST,0) AS COST, ISNULL(l.PRC,0) AS RETAIL, " & _
                "ISNULL(h.TKT_DT,'1/1/1900') AS DATE, ISNULL(l.PRC_1 - l.PRC,0) AS MARKDOWN, " & _
                "PRC_OVRD_REAS AS MKDN_REASON, CONVERT(varchar(30),'') AS COUPON_CODE, ISNULL(l.CATEG_COD,'NA') AS DEPT, " & _
                "ISNULL(l.SUBCAT_COD,'NA') AS CLASS, ISNULL(pv.LST_ORD_BUYER,'OTHER') AS BUYER, CUST_NO, l.TKT_NO FROM PS_DOC_LIN AS l " & _
                "INNER JOIN PS_DOC_HDR AS h ON h.DOC_ID = l.DOC_ID " & _
                "LEFT JOIN PO_VEND_ITEM AS pv ON pv.ITEM_NO = l.ITEM_NO AND pv.VEND_NO = l.ITEM_VEND_NO " & _
                "LEFT JOIN IM_ITEM i ON i.ITEM_NO = l.ITEM_NO " & _
                "LEFT JOIN PS_DOC_LIN_CELL c ON c.DOC_ID = l.DOC_ID AND c.LIN_SEQ_NO = l.LIN_SEQ_NO " & _
                "WHERE h.DOC_TYP IN ('O','T') " & _
                "AND NOT EXISTS (SELECT * FROM PS_DOC_AUDIT_LOG a WHERE a.DOC_ID = l.DOC_ID AND a.DOC_TYP = 'V') " & _
                "SELECT DOC_ID, LIN_SEQ_NO, PROMPT_ALPHA INTO #us2 FROM PS_TKT_HIST_LIN_PROMPT " & _
                "WHERE PROMPT_COD = 'MKTG_CODE' " & _
                "UPDATE t1 SET t1.COUPON_CODE = PROMPT_ALPHA FROM #us t1 JOIN #us2 t2 ON t2.DOC_ID = t1.TRANS_ID AND t2.LIN_SEQ_NO = SEQ_NO " & _
                "SELECT TRANS_ID, SEQ_NO, STORE, SKU, DIM1, DIM2, DIM3, LOCATION, DRAWER, QTY, COST, " & _
                "RETAIL, DATE, MARKDOWN, MKDN_REASON, COUPON_CODE, DEPT, CLASS, BUYER, CUST_NO, CUST_TYPE, TRANS_TYPE, TKT_NO FROM #us"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
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
                oTest = rdr("SKU")
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
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("TRANS_DATE", oTest)
                oTest = rdr("MARKDOWN")
                If IsDBNull(oTest) Then mkdn = 0 Else mkdn = CDec(oTest)
                mkdn = mkdn * qty
                xmlWriter.WriteElementString("MARKDOWN", mkdn)
                oTest = rdr("MKDN_REASON")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("MKDN_REASON", oTest)
                oTest = rdr("COUPON_CODE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("COUPON_CODE", oTest)
                oTest = rdr("DEPT")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("DEPT", oTest)
                oTest = rdr("CLASS")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("CLASS", oTest)
                oTest = rdr("BUYER")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("BUYER", oTest)
                oTest = rdr("CUST_NO")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CUST_NO", oTest)
                oTest = rdr("CUST_TYPE")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("CUST_TYPE", oTest)
                oTest = rdr("TRANS_TYPE")
                If Not IsDBNull(oTest) Then oTest = Replace(oTest, "'", "''") Else oTest = ""
                xmlWriter.WriteElementString("TRANS_TYPE", oTest)
                oTest = rdr("TKT_NO")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("TKT_NO", oTest)
                totlQty += qty
                totlCost += (qty * cost)
                totlRetail += (qty * retail)
                xmlWriter.WriteElementString("EXTRACT_DATE", Date.Today)
                xmlWriter.WriteEndElement()
            End While
            cpCon.Close()

            xmlWriter.WriteStartElement("SALE")
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
            xmlWriter.WriteElementString("CUST_NO", "")
            xmlWriter.WriteElementString("CUST_TYPE", "")
            xmlWriter.WriteElementString("TRANS_TYPE", "")
            xmlWriter.WriteElementString("TKT_NO", "")
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
        Catch ex As Exception
            Dim el As New CounterPointUnpostedSalesExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Writing XML records")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try
    End Sub

End Module

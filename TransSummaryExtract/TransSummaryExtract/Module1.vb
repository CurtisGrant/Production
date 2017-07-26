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
        End While
        conString = "Server=" & cpServer & ";Initial Catalog=" & cpDatabase & "; Integrated Security=True"
        cpCon = New SqlConnection(conString)
        errorPath = xmlPath                          ' send errors to the xml folder so we can see them

        Try
            Dim thisDay As Integer = Date.Now.DayOfWeek
            path = xmlPath & "\errlog.txt"
            If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
        Catch ex As Exception
            Dim el As New TransSummaryExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Daily_XML_Extract Delete errlog")
        End Try

        Try
            Dim cnt As Integer = 0
            If (Not System.IO.Directory.Exists(xmlPath)) Then
                System.IO.Directory.CreateDirectory(xmlPath)
            End If
            Console.WriteLine("Extracting Data")
            path = xmlPath & "\TransSummary.xml"
            xmlWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            xmlWriter.WriteStartDocument(True)
            xmlWriter.Formatting = Formatting.Indented
            xmlWriter.Indentation = 5
            xmlWriter.WriteStartElement("TRANSACTION")
            cpCon.Open()
            sql = "DECLARE @thiseDate date = (Select CONVERT(Date,DATEADD(DAY , 7-DATEPART(WEEKDAY,GETDATE()),GETDATE())) ) " & _
                "DECLARE @beginDate date = DATEADD(Day,-90,@thiseDate) " & _
                "DECLARE @endDate date = DATEADD(Day,-7,@thiseDate) " & _
                "IF OBJECT_ID('tempDB.dbo.#t1','U') IS NOT NULL DROP TABLE #t1; " & _
                "IF OBJECT_ID('tempDB.dbo.#t2','U') IS NOT NULL DROP TABLE #t2; " & _
                "IF OBJECT_ID('tempDB.dbo.#t3','U') IS NOT NULL DROP TABLE #t3; " & _
                "IF OBJECT_ID('tempDB.dbo.#t4','U') IS NOT NULL DROP TABLE #t4; " & _
                "IF OBJECT_ID('tempDB.dbo.#t5','U') IS NOT NULL DROP TABLE #t5; " & _
                "IF OBJECT_ID('tempDB.dbo.#t6','U') IS NOT NULL DROP TABLE #t6; " & _
                "IF OBJECT_ID('tempDB.dbo.#t7','U') IS NOT NULL DROP TABLE #t7; " & _
                "IF OBJECT_ID('tempDB.dbo.#temp','U') IS NOT NULL DROP TABLE #temp; " & _
                "CREATE TABLE #temp(Location varchar(30), Store varchar(30), Date Date, Type varchar(30), Qty Decimal(18,4),  " & _
                "Cost Decimal(18,4), Retail Decimal(18,4), Markdown Decimal(18,4)) " & _
                "SELECT h.LOC_ID AS LOCATION, CASE WHEN c.QTY IS NOT NULL THEN c.QTY ELSE h.QTY * QTY_NUMER END AS QTY, COST,  " & _
                "UNIT_RETL_VAL AS RETAIL, DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT) AS DATE INTO #t1 FROM IM_ADJ_HIST h  " & _
                "LEFT JOIN IM_ADJ_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.BAT_ID = h.BAT_ID AND c.ITEM_NO = h.ITEM_NO " & _
                "AND c.LOC_ID = h.LOC_ID  AND c.SEQ_NO = h.SEQ_NO WHERE h.TRX_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT LOCATION, DATE, 'ADJ', SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t1 " & _
                "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " & _
                "SELECT h.LOC_ID AS LOCATION, c.CNT_QTY_1, AVG_COST AS COST, UNIT_RETL_VAL AS RETAIL, " & _
                "DATEADD(Day,7-DATEPART(WEEKDAY,POST_DAT),POST_DAT) AS DATE, c.FRZ_QTY_ON_HND,  " & _
                "CASE WHEN c.QTY_CNTD IS NOT NULL THEN c.QTY_CNTD - c.FRZ_QTY_ON_HND ELSE QTY_ADJ END AS QTY INTO #t2 FROM IM_CNT_HIST h  " & _
                "LEFT JOIN IM_CNT_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO AND c.LOC_ID = h.LOC_ID  " & _
                "WHERE ITEM_TYP = 'I' AND QTY_ADJ <> 0 AND h.FRZ_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT LOCATION, DATE, 'Physical',  " & _
                "SUM(ISNULL(QTY,0)), SUM(ISNULL(QTY * COST,0)), SUM(ISNULL(QTY * RETAIL,0)) FROM #t2 " & _
                "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION  " & _
                "SELECT l.RECVR_LOC_ID AS LOCATION, DATEADD(Day,7-DATEPART(Weekday,l.RECVR_DAT),l.RECVR_DAT) AS DATE, " & _
                "l.RECVD_COST / l.QTY_RECVD_NUMER AS COST, l.UNIT_RETL_VAL / l.QTY_RECVD_NUMER AS RETAIL,    " & _
                "CASE WHEN c.QTY_RECVD IS NOT NULL THEN c.QTY_RECVD ELSE l.QTY_RECVD * l.QTY_RECVD_NUMER END AS QTY " & _
                "INTO #t3 FROM PO_RECVR_HIST_LIN AS l  " & _
                "LEFT JOIN PO_RECVR_HIST_CELL c ON c.RECVR_NO = l.RECVR_NO AND c.SEQ_NO = l.SEQ_NO  " & _
                "INNER JOIN PO_RECVR_HIST AS h ON h.RECVR_NO = l.RECVR_NO  " & _
                "LEFT JOIN PO_ORD_HDR p ON p.PO_NO = l.PO_NO  " & _
                "WHERE l.QTY_RECVD > 0 AND ITEM_TYP = 'I' AND l.RECVR_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT LOCATION, DATE, 'Receipt',  " & _
                "SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t3 " & _
                "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " & _
                "SELECT l.RTV_LOC_ID AS LOCATION, DATEADD(Day,7-DATEPART(WEEKDAY,l.RTV_DAT),l.RTV_DAT) AS DATE, " & _
                "COST / ISNULL(l.QTY_RETD_NUMER,1) AS COST,  " & _
                "(UNIT_RETL_VAL / ISNULL(QTY_RETD_NUMER,1)) AS RETAIL,  " & _
                "CASE WHEN c.QTY_RETD IS NOT NULL THEN c.QTY_RETD ELSE l.QTY_RETD END AS QTY INTO #t4 FROM PO_RTV_HIST_LIN as l  " & _
                "LEFT JOIN PO_RTV_HIST_CELL c ON c.RTV_NO = l.RTV_NO AND c.SEQ_NO = l.SEQ_NO  " & _
                "INNER JOIN PO_RTV_HIST AS h ON h.RTV_NO = l.RTV_NO WHERE ITEM_TYP = 'I' AND l.RTV_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT LOCATION, DATE, 'RTV',  SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t4 " & _
                "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " & _
                "SELECT ISNULL(l.STR_ID,'NA') STORE, l.STK_LOC_ID AS LOCATION,  " & _
                "ISNULL(l.ITEM_NO,'NA') AS ITEM, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, DRW_ID AS DRAWER,  " & _
                "CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) ELSE c.QTY_SOLD END AS QTY, " & _
                "ISNULL(CONVERT(DECIMAL(10,2),l.UNIT_COST),0) AS COST, ISNULL(CONVERT(DECIMAL(10,2),l.PRC),0) AS RETAIL,  " & _
                "DATEADD(Day,7-DATEPART(Weekday, h.BUS_DAT),h.BUS_DAT) AS DATE, " & _
                "ISNULL(CONVERT(DECIMAL(10,2),l.PRC_1 - l.PRC),0) AS MARKDOWN INTO #t5 FROM PS_TKT_HIST AS h  " & _
                "JOIN PS_TKT_HIST_LIN l ON h.DOC_ID = l.DOC_ID AND h.BUS_DAT = l.BUS_DAT  " & _
                "LEFT JOIN PS_TKT_HIST_LIN_CELL c ON c.DOC_ID = h.DOC_ID AND c.BUS_DAT = h.BUS_DAT AND c.LIN_SEQ_NO = l.LIN_SEQ_NO  " & _
                "LEFT JOIN PO_VEND_ITEM pv ON pv.VEND_NO = l.ITEM_VEND_NO AND pv.ITEM_NO = l.ITEM_NO  " & _
                "WHERE h.TKT_TYP = 'T' AND l.LIN_TYP IN ('S','R') AND l.BUS_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Store, Date, Type, Qty, Cost, Retail, Markdown) " & _
                "SELECT LOCATION, STORE, DATE, 'Sales',  SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL), SUM(QTY * MARKDOWN) FROM #t5 " & _
                "GROUP BY STORE, LOCATION, DATE ORDER BY DATE, STORE, LOCATION " & _
                "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, " & _
                "c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " & _
                "CASE WHEN c.XFER_QTY IS NULL THEN (l.XFER_QTY_NUMER * l.XFER_QTY * -1)  " & _
                "ELSE (c.XFER_QTY * -1) END AS QTY_OUT, l.xfer_qty_numer,  " & _
                "CASE WHEN l.XFER_QTY = 0 THEN CAST(l.FROM_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " & _
                "ELSE CAST(l.FROM_UNIT_RETL_VAL / l.XFER_QTY_NUMER AS DECIMAL) END AS RETAIL,  " & _
                "CASE WHEN l.XFER_QTY = 0 THEN CAST(ABS(l.FROM_UNIT_COST / (l.XFER_QTY_NUMER)) AS DECIMAL(8,2))  " & _
                "ELSE CAST(ABS(l.FROM_EXT_COST / (l.XFER_QTY * l.XFER_QTY_NUMER)) AS DECIMAL(8,2)) END AS COST,   " & _
                "DATEADD(Day,7-DATEPART(Weekday,h.SHIP_DAT),h.SHIP_DAT) AS Date_Out INTO #t6 FROM IM_XFER_LIN AS l " & _
                "INNER JOIN IM_XFER_HDR as h ON h.XFER_NO = l.XFER_NO  " & _
                "LEFT JOIN IM_XFER_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_LIN_SEQ_NO = l.XFER_LIN_SEQ_NO  " & _
                "WHERE h.SHIP_DAT >= @beginDate " & _
                "INSERT INTO #t6(TRANS_ID, SEQ_NO, XFER_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " & _
                "QTY_OUT, XFER_QTY_NUMER, RETAIL, COST, Date_Out) " & _
                "SELECT h.EVENT_NO, 0, '', h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.FROM_LOC_ID, h.TO_LOC_ID, " & _
                "CASE WHEN c.QTY IS NULL THEN (h.QTY * -1) ELSE (c.QTY * -1) END QTY, h.QTY_NUMER, FROM_UNIT_RETL_VAL RETAIL, " & _
                "FROM_COST COST, h.TRX_DAT FROM IM_QXFER_HIST h " & _
                "LEFT JOIN IM_QXFER_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO  " & _
                "WHERE h.ITEM_TYP = 'I' AND h.TRX_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT FROM_LOC_ID, DATE_OUT, 'XFREOUT',  " & _
                "SUM(QTY_OUT), SUM(QTY_OUT * COST), SUM(QTY_OUT * RETAIL) FROM #t6  " & _
                "GROUP BY FROM_LOC_ID, DATE_OUT ORDER BY DATE_OUT, FROM_LOC_ID " & _
                "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_IN_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, " & _
                "c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID,  " & _
                "CASE WHEN c.QTY_RECVD IS NULL THEN l.XFER_QTY_NUMER * l.QTY_RECVD ELSE c.QTY_RECVD END AS QTY, l.xfer_qty_numer,  " & _
                "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " & _
                "ELSE CAST(TO_UNIT_RETL_VAL / l.XFER_QTY_NUMER AS DECIMAL(8,2)) END AS RETAIL,  " & _
                "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_COST / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " & _
                "ELSE CAST(TO_EXT_COST / (l.QTY_RECVD * l.XFER_QTY_NUMER) AS DECIMAL(8,2)) END AS COST,  " & _
                "DATEADD(Day,7-DATEPART(WEEKDAY,h.RECVD_DAT),h.RECVD_DAT) AS Date_In INTO #t7 FROM IM_XFER_IN_HIST_LIN AS l " & _
                "INNER JOIN IM_XFER_IN_HIST as h ON h.XFER_NO = l.XFER_NO AND h.EVENT_NO = l.EVENT_NO  " & _
                "LEFT JOIN IM_XFER_IN_HIST_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_IN_LIN_SEQ_NO = l.XFER_IN_LIN_SEQ_NO  " & _
                "WHERE h.RECVD_DAT >= @beginDate " & _
                "INSERT INTO #t7(TRANS_ID, SEQ_NO, XFER_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " & _
                "QTY, XFER_QTY_NUMER, RETAIL, COST, Date_In) " & _
                "SELECT h.EVENT_NO, 0, '', h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.FROM_LOC_ID, h.TO_LOC_ID, " & _
                "CASE WHEN c.QTY IS NULL THEN h.QTY ELSE c.QTY END QTY, h.QTY_NUMER, FROM_UNIT_RETL_VAL RETAIL, " & _
                "FROM_COST COST, h.TRX_DAT FROM IM_QXFER_HIST h " & _
                "LEFT JOIN IM_QXFER_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO  " & _
                "WHERE h.ITEM_TYP = 'I' AND h.TRX_DAT >= @beginDate " & _
                "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " & _
                "SELECT TO_LOC_ID, DATE_IN, 'XFERIN',  " & _
                "SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t7 " & _
                "GROUP BY TO_LOC_ID, DATE_IN ORDER BY DATE_IN, TO_LOC_ID " & _
                "SELECT Location, Store, Date, Type, Qty, Cost, Retail, Markdown FROM #temp ORDER BY Type, Date, Location "
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("TRANSACTION")
                oTest = rdr("Location")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("LOCATION", CStr(oTest))
                oTest = rdr("Store")
                If IsDBNull(oTest) Then oTest = ""
                xmlWriter.WriteElementString("STORE", CStr(oTest))
                oTest = rdr("Date")
                If IsDBNull(oTest) Then oTest = "1/1/1900"
                xmlWriter.WriteElementString("DATE", CDate(oTest))
                oTest = rdr("Type")
                If IsDBNull(oTest) Then oTest = "UNKNOWN"
                xmlWriter.WriteElementString("TYPE", CStr(oTest))
                oTest = rdr("Qty")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("QTY", CDec(oTest))
                oTest = rdr("Cost")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("COST", CDec(oTest))
                oTest = rdr("Retail")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("RETAIL", CDec(oTest))
                oTest = rdr("Markdown")
                If IsDBNull(oTest) Then oTest = 0
                xmlWriter.WriteElementString("MARKDOWN", CDec(oTest))
                xmlWriter.WriteEndElement()
            End While
            xmlWriter.WriteEndElement()
            xmlWriter.WriteEndDocument()
            xmlWriter.Close()
            cpCon.Close()

            ''Console.WriteLine("wrote " & cnt & " records")
            ''Console.ReadLine()

        Catch ex As Exception
            Dim el As New TransSummaryExtract.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "TransSummaryExtract Create XML file")
            If cpCon.State = ConnectionState.Open Then cpCon.Close()
        End Try
    End Sub

End Module

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
            ''sql = "DECLARE @thiseDate date = (Select CONVERT(Date,DATEADD(DAY , 7-DATEPART(WEEKDAY,GETDATE()),GETDATE())) ) " &
            ''    "DECLARE @beginDate date = DATEADD(month,-45,@thiseDate) " &
            ''    "DECLARE @endDate date = DATEADD(Day,-7,@thiseDate) " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t1','U') IS NOT NULL DROP TABLE #t1; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t2','U') IS NOT NULL DROP TABLE #t2; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t3','U') IS NOT NULL DROP TABLE #t3; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t4','U') IS NOT NULL DROP TABLE #t4; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t5','U') IS NOT NULL DROP TABLE #t5; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t6','U') IS NOT NULL DROP TABLE #t6; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t7','U') IS NOT NULL DROP TABLE #t7; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#t55','U') IS NOT NULL DROP TABLE #t55; " &
            ''    "IF OBJECT_ID('tempDB.dbo.#temp','U') IS NOT NULL DROP TABLE #temp; " &
            ''    "CREATE TABLE #temp(Location varchar(30), Store varchar(30), Date Date, Type varchar(30), Qty Decimal(18,4),  " &
            ''    "Cost Decimal(18,4), Retail Decimal(18,4), Markdown Decimal(18,4)) " &
            ''    "SELECT h.LOC_ID AS LOCATION, CASE WHEN c.QTY IS NOT NULL THEN c.QTY ELSE h.QTY * QTY_NUMER END AS QTY, COST,  " &
            ''    "UNIT_RETL_VAL AS RETAIL, DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT) AS DATE INTO #t1 FROM IM_ADJ_HIST h  " &
            ''    "LEFT JOIN IM_ADJ_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.BAT_ID = h.BAT_ID AND c.ITEM_NO = h.ITEM_NO " &
            ''    "AND c.LOC_ID = h.LOC_ID  AND c.SEQ_NO = h.SEQ_NO WHERE h.TRX_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT LOCATION, DATE, 'ADJ', SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t1 " &
            ''    "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " &
            ''    "SELECT h.LOC_ID AS LOCATION, c.CNT_QTY_1, AVG_COST AS COST, UNIT_RETL_VAL AS RETAIL, " &
            ''    "DATEADD(Day,7-DATEPART(WEEKDAY,POST_DAT),POST_DAT) AS DATE, c.FRZ_QTY_ON_HND,  " &
            ''    "CASE WHEN c.QTY_CNTD IS NOT NULL THEN c.QTY_CNTD - c.FRZ_QTY_ON_HND ELSE QTY_ADJ END AS QTY INTO #t2 FROM IM_CNT_HIST h  " &
            ''    "LEFT JOIN IM_CNT_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO AND c.LOC_ID = h.LOC_ID  " &
            ''    "WHERE ITEM_TYP = 'I' AND QTY_ADJ <> 0 AND h.FRZ_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT LOCATION, DATE, 'Physical',  " &
            ''    "SUM(ISNULL(QTY,0)), SUM(ISNULL(QTY * COST,0)), SUM(ISNULL(QTY * RETAIL,0)) FROM #t2 " &
            ''    "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION  " &
            ''    "SELECT l.RECVR_LOC_ID AS LOCATION, DATEADD(Day,7-DATEPART(Weekday,l.RECVR_DAT),l.RECVR_DAT) AS DATE, " &
            ''    "l.RECVD_COST / l.QTY_RECVD_NUMER AS COST, l.UNIT_RETL_VAL / l.QTY_RECVD_NUMER AS RETAIL,    " &
            ''    "CASE WHEN c.QTY_RECVD IS NOT NULL THEN c.QTY_RECVD ELSE l.QTY_RECVD * l.QTY_RECVD_NUMER END AS QTY " &
            ''    "INTO #t3 FROM PO_RECVR_HIST_LIN AS l  " &
            ''    "LEFT JOIN PO_RECVR_HIST_CELL c ON c.RECVR_NO = l.RECVR_NO AND c.SEQ_NO = l.SEQ_NO  " &
            ''    "INNER JOIN PO_RECVR_HIST AS h ON h.RECVR_NO = l.RECVR_NO  " &
            ''    "LEFT JOIN PO_ORD_HDR p ON p.PO_NO = l.PO_NO  " &
            ''    "WHERE l.QTY_RECVD > 0 AND ITEM_TYP = 'I' AND l.RECVR_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT LOCATION, DATE, 'Receipt',  " &
            ''    "SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t3 " &
            ''    "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " &
            ''    "SELECT l.RTV_LOC_ID AS LOCATION, DATEADD(Day,7-DATEPART(WEEKDAY,l.RTV_DAT),l.RTV_DAT) AS DATE, " &
            ''    "COST / ISNULL(l.QTY_RETD_NUMER,1) AS COST,  " &
            ''    "(UNIT_RETL_VAL / ISNULL(QTY_RETD_NUMER,1)) AS RETAIL,  " &
            ''    "CASE WHEN c.QTY_RETD IS NOT NULL THEN c.QTY_RETD ELSE l.QTY_RETD END AS QTY INTO #t4 FROM PO_RTV_HIST_LIN as l  " &
            ''    "LEFT JOIN PO_RTV_HIST_CELL c ON c.RTV_NO = l.RTV_NO AND c.SEQ_NO = l.SEQ_NO  " &
            ''    "INNER JOIN PO_RTV_HIST AS h ON h.RTV_NO = l.RTV_NO WHERE ITEM_TYP = 'I' AND l.RTV_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT LOCATION, DATE, 'RTV',  SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t4 " &
            ''    "GROUP BY LOCATION, DATE ORDER BY DATE, LOCATION " &
            ''    "SELECT DATEADD(Day,7-DATEPART(WEEKDAY,h.BUS_DAT),h.BUS_DAT) DATE, ISNULL(h.DOC_ID,'NA') AS TRANS_ID, " &
            ''    "ISNULL(l.LIN_SEQ_NO,0) AS SEQ_NO, ISNULL(l.STR_ID,'NA') STORE, ISNULL(l.ITEM_NO,'NA') AS ITEM, c.DIM_1_UPR, " &
            ''    "c.DIM_2_UPR, c.DIM_3_UPR, l.STK_LOC_ID AS LOCATION, DRW_ID AS DRAWER,  " &
            ''    "CASE WHEN c.QTY_SOLD IS NULL THEN ISNULL(l.QTY_SOLD,0) ELSE c.QTY_SOLD END AS QTY, ISNULL(l.COST,0) AS COST, " &
            ''    "ISNULL(l.PRC,0) - ISNULL(l.LIN_DISC_AMT,0) AS RETAIL,  h.TKT_DT, " &
            ''    "(ISNULL(l.PRC_1,0) - ISNULL(l.PRC,0) + ISNULL(l.LIN_DISC_AMT,0)) AS MARKDOWN, " &
            ''    "l.PRC_OVRD_REAS, ISNULL(l.CATEG_COD,'NA') AS DEPT, ISNULL(l.SUBCAT_COD,'NA') AS CLASS, " &
            ''    "ISNULL(pv.LST_ORD_BUYER,'OTHER') AS BUYER, CUST_NO, h.TKT_NO, USR_ID, l.SLS_REP, l.STA_ID INTO #t5 FROM VI_PS_TKT_HIST_LIN AS l " &
            ''    "JOIN VI_PS_TKT_HIST h ON h.DOC_ID = l.DOC_ID And h.BUS_DAT = l.BUS_DAT " &
            ''    "LEFT JOIN VI_PS_TKT_HIST_LIN_CELL c ON c.DOC_ID = h.DOC_ID AND c.BUS_DAT = h.BUS_DAT AND c.LIN_SEQ_NO = l.LIN_SEQ_NO " &
            ''    "LEFT JOIN PO_VEND_ITEM pv ON pv.VEND_NO = l.ITEM_VEND_NO And pv.ITEM_NO = l.ITEM_NO " &
            ''    "WHERE l.LIN_TYP IN ('S','R') AND h.BUS_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Store, Date, Type, Qty, Cost, Retail, Markdown) " &
            ''    "Select LOCATION, STORE, Date, 'Sales',  SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL), SUM(QTY * MARKDOWN) FROM #t5 " &
            ''    "GROUP BY STORE, LOCATION, DATE ORDER BY DATE, STORE, LOCATION " &
            ''    "SELECT  ISNULL(l.STR_ID,'NA') STR_ID,  ISNULL(l.STK_LOC_ID,'NA') LOCATION,  " &
            ''    "CASE WHEN c.QTY_SHIPPED IS NULL THEN ISNULL(l.QTY_SHIPPED,0) ELSE ISNULL(c.QTY_SHIPPED,0) END QTY,  " &
            ''    "ISNULL(i.AVG_COST,0) COST, ISNULL(l.PRC,0) - ISNULL(l.LIN_DISC_AMT,0) RETAIL, " &
            ''    "DATEADD(Day,7-DATEPART(Weekday, h.BUS_DAT),h.BUS_DAT) Date, " &
            ''    "(ISNULL(l.PRC_1,0) - ISNULL(l.PRC,0) + ISNULL(l.LIN_DISC_AMT,0)) MARKDOWN INTO #t55 FROM VI_PS_ORD_HIST_LIN AS l " &
            ''    "INNER JOIN VI_PS_ORD_HIST AS h ON h.DOC_ID = l.DOC_ID " &
            ''    "LEFT JOIN VI_IM_ITEM_WITH_INV i ON i.ITEM_NO = l.ITEM_NO AND i.LOC_ID = l.STK_LOC_ID " &
            ''    "LEFT JOIN VI_PS_ORD_HIST_LIN_CELL c ON c.DOC_ID = l.DOC_ID And c.LIN_SEQ_NO = l.LIN_SEQ_NO " &
            ''    "WHERE h.TKT_DAT >= @beginDate AND l.QTY_SHIPPED <> 0 " &
            ''    "AND NOT EXISTS (SELECT * FROM VI_PS_DOC_AUDIT_LOG a WHERE a.DOC_ID = l.DOC_ID And a.DOC_TYP = 'V')  " &
            ''    "INSERT INTO #temp(Location, Store, Date, Type, Qty, Cost, Retail, Markdown) " &
            ''    "SELECT LOCATION, STR_ID, DATE, 'Orders',  SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL), SUM(QTY * MARKDOWN) FROM #t55 " &
            ''    "WHERE QTY <> 0 " &
            ''    "GROUP BY LOCATION, STR_ID, DATE ORDER BY DATE, LOCATION " &
            ''    "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, " &
            ''    "c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " &
            ''    "CASE WHEN c.XFER_QTY IS NULL THEN (l.XFER_QTY_NUMER * l.XFER_QTY * -1)  " &
            ''    "ELSE (c.XFER_QTY * -1) END AS QTY_OUT, l.xfer_qty_numer,  " &
            ''    "CASE WHEN l.XFER_QTY = 0 THEN CAST(l.FROM_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " &
            ''    "ELSE CAST(l.FROM_UNIT_RETL_VAL / l.XFER_QTY_NUMER AS DECIMAL) END AS RETAIL,  " &
            ''    "CASE WHEN l.XFER_QTY = 0 THEN CAST(ABS(l.FROM_UNIT_COST / (l.XFER_QTY_NUMER)) AS DECIMAL(8,2))  " &
            ''    "ELSE CAST(ABS(l.FROM_EXT_COST / (l.XFER_QTY * l.XFER_QTY_NUMER)) AS DECIMAL(8,2)) END AS COST,   " &
            ''    "DATEADD(Day,7-DATEPART(Weekday,h.SHIP_DAT),h.SHIP_DAT) AS Date_Out INTO #t6 FROM IM_XFER_LIN AS l " &
            ''    "INNER JOIN IM_XFER_HDR as h ON h.XFER_NO = l.XFER_NO  " &
            ''    "LEFT JOIN IM_XFER_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_LIN_SEQ_NO = l.XFER_LIN_SEQ_NO  " &
            ''    "WHERE h.SHIP_DAT >= @beginDate " &
            ''    "INSERT INTO #t6(TRANS_ID, SEQ_NO, XFER_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " &
            ''    "QTY_OUT, XFER_QTY_NUMER, RETAIL, COST, Date_Out) " &
            ''    "SELECT h.EVENT_NO, 0, '', h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.FROM_LOC_ID, h.TO_LOC_ID, " &
            ''    "CASE WHEN c.QTY IS NULL THEN (h.QTY * -1) ELSE (c.QTY * -1) END QTY, h.QTY_NUMER, FROM_UNIT_RETL_VAL RETAIL, " &
            ''    "FROM_COST COST, h.TRX_DAT FROM IM_QXFER_HIST h " &
            ''    "LEFT JOIN IM_QXFER_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO  " &
            ''    "WHERE h.ITEM_TYP = 'I' AND h.TRX_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT FROM_LOC_ID, DATE_OUT, 'XFREOUT',  " &
            ''    "SUM(QTY_OUT), SUM(QTY_OUT * COST), SUM(QTY_OUT * RETAIL) FROM #t6  " &
            ''    "GROUP BY FROM_LOC_ID, DATE_OUT ORDER BY DATE_OUT, FROM_LOC_ID " &
            ''    "SELECT h.EVENT_NO AS TRANS_ID, l.XFER_IN_LIN_SEQ_NO AS SEQ_NO, h.XFER_NO, l.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, " &
            ''    "c.DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID,  " &
            ''    "CASE WHEN c.QTY_RECVD IS NULL THEN l.XFER_QTY_NUMER * l.QTY_RECVD ELSE c.QTY_RECVD END AS QTY, l.xfer_qty_numer,  " &
            ''    "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_RETL_VAL / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " &
            ''    "ELSE CAST(TO_UNIT_RETL_VAL / l.XFER_QTY_NUMER AS DECIMAL(8,2)) END AS RETAIL,  " &
            ''    "CASE WHEN l.QTY_RECVD = 0 THEN CAST(TO_UNIT_COST / (l.XFER_QTY_NUMER) AS DECIMAL(8,2))  " &
            ''    "ELSE CAST(TO_EXT_COST / (l.QTY_RECVD * l.XFER_QTY_NUMER) AS DECIMAL(8,2)) END AS COST,  " &
            ''    "DATEADD(Day,7-DATEPART(WEEKDAY,h.RECVD_DAT),h.RECVD_DAT) AS Date_In INTO #t7 FROM IM_XFER_IN_HIST_LIN AS l " &
            ''    "INNER JOIN IM_XFER_IN_HIST as h ON h.XFER_NO = l.XFER_NO AND h.EVENT_NO = l.EVENT_NO  " &
            ''    "LEFT JOIN IM_XFER_IN_HIST_CELL c ON c.XFER_NO = l.XFER_NO AND c.XFER_IN_LIN_SEQ_NO = l.XFER_IN_LIN_SEQ_NO  " &
            ''    "WHERE h.RECVD_DAT >= @beginDate " &
            ''    "INSERT INTO #t7(TRANS_ID, SEQ_NO, XFER_NO, ITEM_NO, DIM_1_UPR, DIM_2_UPR, DIM_3_UPR, FROM_LOC_ID, TO_LOC_ID, " &
            ''    "QTY, XFER_QTY_NUMER, RETAIL, COST, Date_In) " &
            ''    "SELECT h.EVENT_NO, 0, '', h.ITEM_NO, c.DIM_1_UPR, c.DIM_2_UPR, c.DIM_3_UPR, h.FROM_LOC_ID, h.TO_LOC_ID, " &
            ''    "CASE WHEN c.QTY IS NULL THEN h.QTY ELSE c.QTY END QTY, h.QTY_NUMER, FROM_UNIT_RETL_VAL RETAIL, " &
            ''    "FROM_COST COST, h.TRX_DAT FROM IM_QXFER_HIST h " &
            ''    "LEFT JOIN IM_QXFER_HIST_CELL c ON c.EVENT_NO = h.EVENT_NO AND c.ITEM_NO = h.ITEM_NO  " &
            ''    "WHERE h.ITEM_TYP = 'I' AND h.TRX_DAT >= @beginDate " &
            ''    "INSERT INTO #temp(Location, Date, Type, Qty, Cost, Retail) " &
            ''    "SELECT TO_LOC_ID, DATE_IN, 'XFERIN',  " &
            ''    "SUM(QTY), SUM(QTY * COST), SUM(QTY * RETAIL) FROM #t7 " &
            ''    "GROUP BY TO_LOC_ID, DATE_IN ORDER BY DATE_IN, TO_LOC_ID " &
            ''    "SELECT Location, Store, Date, Type, Qty, Cost, Retail, Markdown FROM #temp ORDER BY Type, Date, Location "
            sql = "DECLARE @thisEdate Date = (SELECT CONVERT(Date,DATEADD(Day,7-DATEPART(WeekDay,GETDATE()),GETDATE()))) " &
                "DECLARE @beginDate Date = DATEADD(Day,-90,@thisEdate) " &
                "IF OBJECT_ID('tempDB.dbo.#temp','U') IS NOT NULL DROP TABLE #temp; " &
                "CREATE TABLE #temp(TYPE varchar(10), LOC_ID varchar(30), Store varchar(30), TRX_DAT Date, QTY Decimal(18,4), " &
                "Cost Decimal(18,4), Retail Decimal(18,4), Markdown Decimal(18,4)) " &
                "INSERT INTO #temp(TYPE, LOC_ID, Store, TRX_DAT, QTY, COST, RETAIL, MARKDOWN) " &
                "select 'ADJ', H.LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT),  " &
                "  coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.EXT_COST * C.QTY / H.QTY, H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY, H.QTY) * H.UNIT_RETL_VAL as RETAIL, 0 " &
                "from IM_ADJ_HIST H " &
                "left join IM_ADJ_HIST_CELL C " &
                "  on H.EVENT_NO = C.EVENT_NO " &
                " and H.BAT_ID = C.BAT_ID " &
                " and H.ITEM_NO = C.ITEM_NO " &
                " and H.LOC_ID = C.LOC_ID " &
                " and H.SEQ_NO = C.SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'SALE', H.STK_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.POST_DAT),h.POST_DAT),  " &
                "  coalesce(-C.QTY_SOLD, -H.QTY_SOLD) * (H.QTY_NUMER / H.QTY_DENOM) as QTY,   " &
                "  case when H.QTY_SOLD = 0 then 0 else cast(coalesce(-H.EXT_COST * C.QTY_SOLD / H.QTY_SOLD, -H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(-C.QTY_SOLD, -H.QTY_SOLD) * H.UNIT_RETL_AT_POST as EXT_RET, 0 " &
                "from VI_PS_TKT_HIST_LIN H " &
                "left join VI_PS_TKT_HIST_LIN_CELL C " &
                "  on H.SEQ_NO = C.LIN_SEQ_NO " &
                " and H.DOC_ID = C.DOC_ID " &
                " and H.TKT_NO = C.TKT_NO " &
                "where LIN_TYP in ('S', 'R') and H.POST_DAT >= @beginDate " &
                "union all " &
                "select 'PHYS', H.LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.POST_DAT),h.POST_DAT),  " &
                "  coalesce(C.QTY_CNTD - C.QTY_ON_HND_BEFORE, H.QTY_ADJ) as QTY, " &
                "  case when H.QTY_CNTD = 0 then 0 else cast(coalesce(H.EXT_COST * C.QTY_CNTD / H.QTY_CNTD, H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY_CNTD - C.QTY_ON_HND_BEFORE, H.QTY_ADJ) * H.UNIT_RETL_AT_POST as EXT_RET, 0 " &
                "from IM_CNT_HIST H " &
                "left join IM_CNT_HIST_CELL C " &
                "  on H.EVENT_NO = C.EVENT_NO " &
                " and H.ITEM_NO = C.ITEM_NO " &
                " and H.LOC_ID = C.LOC_ID " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'RECVD', H.LOC_ID, '',DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT),  " &
                "  coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.EXT_COST * C.QTY / H.QTY, H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY, H.QTY) * H.UNIT_RETL_VAL as EXT_RET, 0 " &
                "from PO_QRECV_HIST H " &
                "left join PO_QRECV_HIST_CELL C " &
                "  on H.EVENT_NO = C.EVENT_NO " &
                " and H.BAT_ID = C.BAT_ID " &
                " and H.ITEM_NO = C.ITEM_NO " &
                " and H.LOC_ID = C.LOC_ID " &
                " and H.TRX_DAT = C.TRX_DAT " &
                " and H.SEQ_NO = C.SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'RECVD', H.RECVR_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.RECVR_DAT),h.RECVR_DAT),  " &
                "  coalesce(C.QTY_RECVD, H.QTY_RECVD) * (H.QTY_RECVD_NUMER / H.QTY_RECVD_DENOM) as QTY, " &
                "  case when H.QTY_RECVD = 0 then 0 else cast(coalesce(H.EXT_COST * C.QTY_RECVD / H.QTY_RECVD, H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY_RECVD, H.QTY_RECVD) * UNIT_RETL_VAL as EXT_RET, 0 " &
                "from PO_RECVR_HIST_LIN H " &
                "join PO_RECVR_HIST " &
                "  on PO_RECVR_HIST.RECVR_NO = H.RECVR_NO " &
                "left join PO_RECVR_HIST_CELL C " &
                "  on H.RECVR_NO = C.RECVR_NO " &
                " and H.SEQ_NO = C.SEQ_NO " &
                "WHERE PO_RECVR_HIST.IS_DROPSHIP_RECVR = 'N'  " &
                "and H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'RTV', H.RTV_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.RTV_DAT),h.RTV_DAT),  " &
                "  coalesce(-C.QTY_RETD, -H.QTY_RETD) * (H.QTY_RETD_NUMER / H.QTY_RETD_DENOM) as QTY, " &
                "  case when H.QTY_RETD = 0 then 0 else cast(coalesce(H.EXT_COST * C.QTY_RETD / H.QTY_RETD, H.EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(-C.QTY_RETD, -H.QTY_RETD) * UNIT_RETL_VAL as EXT_RET, 0 " &
                "from PO_RTV_HIST_LIN H " &
                "left join PO_RTV_HIST_CELL C " &
                "  on H.SEQ_NO = C.SEQ_NO " &
                " and H.RTV_NO = C.RTV_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'RECVD', H.FROM_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT),  " &
                "  coalesce(-C.QTY, -H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.FROM_EXT_COST * C.QTY / H.QTY, H.FROM_EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(-C.QTY, -H.QTY) * H.FROM_UNIT_RETL_VAL as EXT_RET, 0 " &
                "from IM_QXFER_HIST H " &
                "left join IM_QXFER_HIST_CELL C " &
                "  on H.EVENT_NO = C.EVENT_NO " &
                " and H.BAT_ID = C.BAT_ID " &
                " and H.ITEM_NO = C.ITEM_NO " &
                " and H.FROM_LOC_ID = C.FROM_LOC_ID " &
                " and H.TRX_DAT = C.TRX_DAT " &
                " and H.SEQ_NO = C.SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'XFER', H.TO_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.TRX_DAT),h.TRX_DAT),  " &
                "  coalesce(C.QTY, H.QTY) * (H.QTY_NUMER / H.QTY_DENOM) as QTY, " &
                "  case when H.QTY = 0 then 0 else cast(coalesce(H.TO_EXT_COST * C.QTY / H.QTY, H.TO_EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY, H.QTY) * H.TO_UNIT_RETL_VAL as EXT_RET, 0 " &
                "from IM_QXFER_HIST H " &
                "left join IM_QXFER_HIST_CELL C " &
                "  on H.EVENT_NO = C.EVENT_NO " &
                " and H.BAT_ID = C.BAT_ID " &
                " and H.ITEM_NO = C.ITEM_NO " &
                " and H.FROM_LOC_ID = C.FROM_LOC_ID " &
                " and H.TRX_DAT = C.TRX_DAT " &
                " and H.SEQ_NO = C.SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'XFER', H.FROM_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.SHIP_DAT),h.SHIP_DAT),  " &
                "  coalesce(-C.XFER_QTY, -L.XFER_QTY) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "  case when L.XFER_QTY = 0 then 0 else cast(coalesce(L.FROM_EXT_COST * C.XFER_QTY / L.XFER_QTY, L.FROM_EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(-C.XFER_QTY, -L.XFER_QTY) * L.FROM_UNIT_RETL_VAL as EXT_RET, 0 " &
                "from IM_XFER_HDR H " &
                "join IM_XFER_LIN L on H.XFER_NO = L.XFER_NO " &
                "left join IM_XFER_CELL C " &
                "  on C.XFER_NO = L.XFER_NO " &
                " and C.XFER_LIN_SEQ_NO = L.XFER_LIN_SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'XFER', H.TO_LOC_ID, '', DATEADD(Day,7-DATEPART(Weekday, h.RECVD_DAT),h.RECVD_DAT),  " &
                "  coalesce(C.QTY_RECVD, L.QTY_RECVD) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "  case when L.QTY_RECVD = 0 then 0 else cast(coalesce(L.TO_EXT_COST * C.QTY_RECVD / L.QTY_RECVD, L.TO_EXT_COST) as decimal(20,8)) end as EXT_COST, " &
                "  coalesce(C.QTY_RECVD, L.QTY_RECVD) * L.TO_UNIT_RETL_VAL as EXT_RET, 0 " &
                "from IM_XFER_IN_HIST H " &
                "join IM_XFER_IN_HIST_LIN L " &
                "  on H.XFER_NO = L.XFER_NO " &
                "  and H.EVENT_NO = L.EVENT_NO " &
                "left join IM_XFER_IN_HIST_CELL C " &
                "  on C.EVENT_NO = L.EVENT_NO " &
                " and C.XFER_NO = L.XFER_NO " &
                " and C.XFER_IN_LIN_SEQ_NO = L.XFER_IN_LIN_SEQ_NO " &
                "where H.LST_MAINT_DT >= @beginDate " &
                "union all " &
                "select 'XFER', case when L.RECON_METH = 'S' then H.TO_LOC_ID else H.FROM_LOC_ID end as LOC_ID, '', " &
                "  DATEADD(Day,7-DATEPART(Weekday, h.RECON_DAT),h.RECON_DAT), " &
                "  coalesce(-C.QTY_VARIANCE, -L.QTY_VARIANCE) * (L.XFER_QTY_NUMER / L.XFER_QTY_DENOM) as QTY, " &
                "  case " &
                "    when L.RECON_METH = 'S' then cast(coalesce(L.TO_EXT_COST * C.QTY_VARIANCE / (case when L.QTY_VARIANCE = 0 then 1 else L.QTY_VARIANCE end), L.TO_EXT_COST) as decimal(20, 8)) " &
                "    when L.QTY_VARIANCE = 0 then 0 " &
                "    else cast(coalesce(L.FROM_EXT_COST * C.QTY_VARIANCE / L.QTY_VARIANCE, L.FROM_EXT_COST) as decimal(20, 8)) " &
                "  end as EXT_COST, " &
                "  case " &
                "    when L.RECON_METH = 'S' then coalesce(C.QTY_VARIANCE, L.QTY_VARIANCE) * L.TO_UNIT_RETL_VAL " &
                "    when L.QTY_VARIANCE = 0 then 0 " &
                "    else coalesce(C.QTY_VARIANCE, L.QTY_VARIANCE) * L.FROM_UNIT_RETL_VAL " &
                "  end as EXT_RET, 0 " &
                "from IM_XFER_RECON_HIST H " &
                "join IM_XFER_RECON_HIST_LIN L on H.XFER_NO = L.XFER_NO " &
                "left join IM_XFER_RECON_HIST_CELL C " &
                "  on C.XFER_NO = H.XFER_NO " &
                " and C.XFER_LIN_SEQ_NO = L.XFER_LIN_SEQ_NO " &
                "where L.RECON_METH in ('S', 'R') " &
                "and H.LST_MAINT_DT >= @beginDate " &
                "select type, loc_id location, trx_dat date, sum(qty) qty, sum(cost) cost, sum(retail) retail from #temp " &
                "group by type, loc_id, trx_dat"
            cmd = New SqlCommand(sql, cpCon)
            cmd.CommandTimeout = 960
            rdr = cmd.ExecuteReader
            While rdr.Read
                cnt += 1
                If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                xmlWriter.WriteStartElement("TRANSACTION")
                oTest = rdr("Location")
                If IsDBNull(oTest) Then oTest = "NA"
                xmlWriter.WriteElementString("LOCATION", CStr(oTest))
                ''oTest = rdr("Store")
                ''If IsDBNull(oTest) Then oTest = ""
                ''xmlWriter.WriteElementString("STORE", CStr(oTest))
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
                ''oTest = rdr("Markdown")
                ''If IsDBNull(oTest) Then oTest = 0
                ''xmlWriter.WriteElementString("MARKDOWN", CDec(oTest))
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

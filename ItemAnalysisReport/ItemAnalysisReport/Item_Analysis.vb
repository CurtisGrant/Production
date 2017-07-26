Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Microsoft
Imports excel = Microsoft.Office.Interop.Excel           ' add .net (Microsoft Excel 14.0 Object Library) reference under Projects
Imports System.ComponentModel
Imports System.Xml
Public Class Item_Analysis
    Public Shared conString As String
    Public Shared con As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared itemStat, sql, vendorId, sortBy, displayStores, selectedStores, templatesFolder, reportsFolder As String
    Public Shared oTest As Object
    Public Shared firstPass As Boolean = False
    Public Shared Export_To_Excel As Boolean = True
    Public Shared Combine_Stores As Boolean = True
    Public Shared allGood As Boolean = False
    Public Shared coverageLoYrWk, coverageHiYrWk As Integer
    Public Shared coverageSdate, coverageEdate, salesSdate, salesEdate, thisEdate, onOrdloDate, lyStartDate, lastweeksedate As Date
    Public Shared LYsalesSdate, LYsalesEdate, lowestDate As Date
    Public Shared minQtySold, MinWksOH, minGMROI As Integer
    Public Shared tbl, theDataTable, displayTable As DataTable
    Public Shared stores() As String
    Public Shared thisDateString, thisBuyer, vendorName As String
    Public Shared user As String = Environment.UserName

    Private Sub Item_Analysis_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RCCLIENT.xml")
        Dim server, exePath, passWord, dBase, userid As String
        Dim fld As String = ""
        Dim valu As String = ""
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
        End While
        conString = "Server=" & server & ";Initial Catalog=RCClient;User Id=sa;Password=" & passWord & ""
        Dim rCon As New SqlConnection(conString)
        rCon.Open()
        sql = "SELECT Server, [Database], SQLUserID, SQLPassword, Templates, Reports, Last_XML_Update FROM Client_Master " & _
            "WHERE Client_Id = 'PARGIF'"
        ''sql = "SELECT Server, [Database], SQLUserID, SQLPassword, Templates, Reports, Last_XML_Update FROM Client_Master " & _
        ''   "WHERE Client_Id = 'DEMO3'"
        cmd = New SqlCommand(sql, rCon)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Last_XML_Update")
            If IsDBNull(oTest) Then
                Dim msg As String = "On Hand and On Order are currently being updated. Do you wish to continue?"
                Dim caption As String = "WARNING!"
                Dim ans As Integer = MessageBox.Show(msg, caption, MessageBoxButtons.YesNo)
                If ans = DialogResult.No Then
                    rCon.Close()
                    Application.Exit()
                End If
            End If
            server = rdr("Server")
            dBase = rdr("Database")
            userid = rdr("SQLUserID")
            passWord = rdr("SQLPassword")
            templatesFolder = rdr("Templates")
            reportsFolder = rdr("Reports")
        End While
        rCon.Close()

        conString = "server=" & server & ";Initial Catalog=" & dBase & ";User Id=" & userid & ";Password=" & passWord & ""
        con = New SqlConnection(conString)
        con.Open()
        If chkVendor.Checked Then
            sql = "SELECT DISTINCT m.Vendor_Id, m.Vendor FROM Item_Master m JOIN Vendors v ON v.ID = m.Vendor_Id " & _
                "WHERE v.Status = 'Active' AND m.Buyer = '" & thisBuyer & "' " & _
                "ORDER BY m.Vendor"
        Else : sql = "SELECT DISTINCT ID AS Vendor_Id, Description AS Vendor FROM Vendors ORDER BY Description"
        End If
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim id, vendor As String
        cboVendor.Items.Clear()
        cboVendor.Items.Add("Enter Vendor Code Below or Select From Drop Down at Right")
        While rdr.Read
            id = rdr("Vendor_Id")
            vendor = rdr("Vendor")
            cboVendor.Items.Add(id & " - " & vendor)
        End While
        con.Close()

        con.Open()
        cboVendor.SelectedIndex = 0
        sql = "SELECT sDate, eDate FROM Calendar WHERE Week_Id > 0 ORDER BY sDate DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboSfrom.Items.Add(rdr("sDate"))
            cboCfrom.Items.Add(rdr("sDate"))
            cboSthru.Items.Add(rdr("eDate"))
            cboCthru.Items.Add(rdr("eDate"))
            cboS2from.Items.Add(rdr("sDate"))
            cboS2thru.Items.Add(rdr("eDate"))
        End While
        con.Close()

        con.Open()
        Dim sDate, eDate As Date
        Dim wk, yr As Integer
        sql = "SELECT MAX(eDate) AS eDate FROM Calendar WHERE eDate < CONVERT(Date,GETDATE()) AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            eDate = rdr("eDate")
        End While
        con.Close()

        lastweeksedate = eDate
        thisEdate = DateAdd(DateInterval.Day, 7, eDate)
        salesEdate = thisEdate
        LYsalesEdate = DateAdd(DateInterval.Day, -365, salesEdate)
        sDate = DateAdd(DateInterval.WeekOfYear, -16, thisEdate)
        salesSdate = DateAdd(DateInterval.Day, -6, sDate)
        coverageSdate = DateAdd(DateInterval.Day, 8, eDate)
        eDate = DateAdd(DateInterval.WeekOfYear, 7, coverageSdate)
        coverageEdate = DateAdd(DateInterval.Day, 6, eDate)
        Dim wks As Integer
        con.Open()
        sql = "SELECT Value FROM Controls WHERE ID = 'GMROI' AND Parameter = 'Weeks'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            wks = rdr("Value")
        End While
        con.Close()
        lyStartDate = DateAdd(DateInterval.WeekOfYear, wks * -1, salesEdate)

        con.Open()
        sql = "SELECT Year_Id, Week_Id FROM Calendar WHERE eDate = '" & thisEdate & "' AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            yr = rdr("Year_ID")
            wk = rdr("Week_Id")
        End While
        con.Close()

        con.Open()
        '' sql = "SELECT eDate FROM Calendar WHERE '" & LYsalesEdate & "' BETWEEN sDate AND eDate AND Week_Id > 0"
        sql = "SELECT sDate FROM Calendar WHERE Year_Id = " & yr - 1 & " AND Week_Id = " & wk & " "
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            LYsalesSdate = rdr("sDate")
        End While
        con.Close()

        ''LYsalesSdate = DateAdd(DateInterval.WeekOfYear, -16, LYsalesEdate)
        ''LYsalesSdate = DateAdd(DateInterval.Day, -6, LYsalesSdate)
        LYsalesEdate = DateAdd(DateInterval.Day, 118, LYsalesSdate)                    ' changed 2/24/2017
        cboSfrom.SelectedIndex = cboSfrom.FindString(salesSdate)
        cboSthru.SelectedIndex = cboSthru.FindString(salesEdate)
        cboCfrom.SelectedIndex = cboCfrom.FindString(coverageSdate)
        cboCthru.SelectedIndex = cboCthru.FindString(coverageEdate)
        '' cboS2from.SelectedIndex = cboS2from.FindString(LYsalesSdate)
        '' cboS2thru.SelectedIndex = cboS2thru.FindString(LYsalesEdate)
        cboS2from.SelectedIndex = cboS2from.FindString(LYsalesSdate)                   ' changed 2/24/2017
        cboS2thru.SelectedIndex = cboS2thru.FindString(LYsalesEdate)                   ' changed 2/24/2017

        Dim today As Date = Date.Today
        today = CDate(today)
        Dim day As Integer = DatePart(DateInterval.Day, Date.Today)
        Dim month As Integer = DatePart(DateInterval.Month, Date.Today)
        Dim year As Integer = DatePart(DateInterval.Year, Date.Today)
        thisDateString = year & month & day

        txtMinOH.Text = -99
        txtMaxOH.Text = 999999
        txtGMROI.Text = -9999
        minGMROI = -9999
        txtOnHand.Text = 0
        txtUnitsSold.Text = 0
        minQtySold = 0
        lowestDate = DateAdd(DateInterval.Month, -18, Date.Today)
        If LYsalesSdate < lowestDate Then lowestDate = LYsalesSdate
        txtVendors.Select()
    End Sub

    Private Sub cboVendor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVendor.SelectedIndexChanged
        If cboVendor.SelectedIndex = 0 Then Exit Sub
        Dim vend As String = Me.cboVendor.SelectedItem
        Dim searchString As String = " -"
        Dim len As Integer = vend.IndexOf(searchString)
        If Me.txtVendors.Text = "" Then
            Me.txtVendors.Text = vend.Substring(0, len)
        Else
            Me.txtVendors.Text = Me.txtVendors.Text & ", " & vend.Substring(0, len)
        End If
        firstPass = False
        Me.Refresh()
    End Sub

    Private Sub Get_Coverage_Dates()
        oTest = txtCweeks.Text
        Dim wks As Integer
        If Not IsNothing(oTest) Then
            If IsNumeric(oTest) Then
                wks = CInt(oTest)
                con.Open()
                sql = "SELECT MIN(sDate) AS sDate FROM Calendar WHERE sDate >= CONVERT(Date,GETDATE()) AND Week_Id > 0"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    coverageSdate = rdr("sDate")
                    '' thisEdate = DateAdd(DateInterval.Day, 6, coverageSdate)
                    onOrdloDate = DateAdd(DateInterval.Month, -6, coverageSdate)
                    oTest = DateAdd(DateInterval.WeekOfYear, wks, coverageSdate)
                    coverageEdate = DateAdd(DateInterval.Day, 6, oTest)
                End While
                con.Close()

                ''txtCfrom.Text = coverageSdate
                ''txtCthru.Text = coverageEdate
            End If
        End If
    End Sub

    Private Sub Get_Sales_Dates()
        oTest = txtSweeks.Text
        Dim wks As Integer
        If Not IsNothing(oTest) Then
            If IsNumeric(oTest) Then
                wks = CInt(oTest) * -1
                con.Open()
                sql = "SELECT MAX(eDate) AS eDate, MAX(YrWk) AS YrWk FROM Calendar WHERE eDate < CONVERT(Date,GETDATE()) AND Week_Id > 0"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    salesEdate = rdr("eDate")
                    oTest = DateAdd(DateInterval.WeekOfYear, wks, salesEdate)
                    salesSdate = DateAdd(DateInterval.Day, -6, oTest)
                End While
                con.Close()
            End If
        End If
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        oTest = txtVendors.Text
        If IsNothing(oTest) Or oTest = "" Then
            MessageBox.Show("Select a vendor from the list or enter a vendor id and try again!", "ERROR - VENDOR HAS NOT BEEN SELECTED")
            Exit Sub
        End If
        lblProcessing.Visible = True
        Me.Refresh()
        MinWksOH = Int(Val(txtOnHand.Text))
        minQtySold = Int(Val(txtUnitsSold.Text))
        If chkItem.Checked = True Then
            Call Get_Vendor()
            chkGMROI.Checked = False
            chkWksOnHand.Checked = False
            chkMargin.Checked = False
            chkUnitsSold.Checked = False
            chkQtyOnHand.Checked = False
            sortBy = "Item_No"
        End If
        If chkGMROI.Checked = True Then
            chkItem.Checked = False
            chkWksOnHand.Checked = False
            chkMargin.Checked = False
            chkUnitsSold.Checked = False
            chkQtyOnHand.Checked = False
            sortBy = "GMROI"
        End If
        If chkWksOnHand.Checked = True Then
            chkItem.Checked = False
            chkGMROI.Checked = False
            chkMargin.Checked = False
            chkUnitsSold.Checked = False
            chkQtyOnHand.Checked = False
            sortBy = "WksOH"
        End If
        If chkMargin.Checked = True Then
            chkItem.Checked = False
            chkGMROI.Checked = False
            chkWksOnHand.Checked = False
            chkUnitsSold.Checked = False
            chkQtyOnHand.Checked = False
            sortBy = "GM"
        End If
        If chkUnitsSold.Checked = True Then
            chkItem.Checked = False
            chkGMROI.Checked = False
            chkWksOnHand.Checked = False
            chkMargin.Checked = False
            chkQtyOnHand.Checked = False
            sortBy = "Sold"
        End If
        If chkQtyOnHand.Checked = True Then
            chkItem.Checked = False
            chkGMROI.Checked = False
            chkWksOnHand.Checked = False
            chkMargin.Checked = False
            chkUnitsSold.Checked = False
            sortBy = "OH"
        End If

        If chkExcel.Checked = True Then Export_To_Excel = True Else Export_To_Excel = False
        If Me.chkCombined.Checked = True Then Combine_Stores = True Else Combine_Stores = False

        If txtProfCod2.Text <> Nothing And txtProfCod2.Text <> "" Then
            allGood = True
            Call Check_Lines(allGood)
            If allGood = False Then
                MsgBox("Invalid Product Line!")
                Exit Sub
            End If
        End If

        If txtProfCod1.Text <> Nothing And txtProfCod1.Text <> "" Then
            allGood = True
            Call check_Seasons(allGood)
            If allGood = False Then
                MsgBox("Invalid Season!")
                Exit Sub
            End If
        End If

        If txtCateg.Text <> Nothing And txtCateg.Text <> "" Then
            allGood = True
            Call Check_Departments(allGood)
            If allGood = False Then
                MsgBox("Invalid Department!")
                Exit Sub
            End If
        End If

        If txtSubcat.Text <> Nothing And txtSubcat.Text <> "" Then
            allGood = True
            Call Check_Classes(allGood)
            If allGood = False Then
                MsgBox("Invalid Class!")
                Exit Sub
            End If
        End If

        allGood = True
        Call Check_Vendors(allGood)
        If allGood = False Then
            MsgBox("Vendor not found!")
            Exit Sub
        End If
        Dim str(2) As String
        Dim idx As Integer = -1
        selectedStores = ""
        If chkLP.Checked = True Then
            selectedStores = "1"
            idx += 1
            str(idx) = "1"
        End If

        If chkNP.Checked = True Then
            If chkLP.Checked = True Then
                selectedStores &= ",2"
                idx += 1
                str(idx) = "2"
            Else
                selectedStores = "2"
                idx += 1
                str(idx) = "2"
            End If
        End If
        If chkW1.Checked = True Then
            If selectedStores = "" Then
                idx += 1
                str(idx) = "W1"
                selectedStores = "W1"
            Else
                selectedStores &= ",W1"
                idx += 1
                str(idx) = "W1"
            End If
        End If
        displayStores = selectedStores
        stores = str

        itemStat = "'Active','Inactive'"
        If chkActive.Checked = True Then
            itemStat = "'Active'"
        End If
        If chkInactive.Checked = True Then
            itemStat = "'Inactive'"
        End If

        Dim values As String = txtVendors.Text
        Dim vendors As String() = Nothing
        vendors = values.Split(", ")
        Dim v As String

        For Each v In vendors
            vendorId = UCase(Trim(v.ToString))

            con.Open()
            sql = "SELECT Description AS Name FROM Vendors WHERE ID = '" & vendorId & "' "
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                vendorName = rdr("Name")
            End While
            con.Close()

            txtProgress.Text = "Retreiving and processing data for " & vendorId
            Me.Refresh()
            Cursor.Current = Cursors.WaitCursor

            Call GetTheData()
            firstPass = True

            Cursor.Current = Cursors.Default
        Next
    End Sub

    Private Sub GetTheData()
        oTest = txtGMROI.Text
        If Not IsNothing(oTest) And IsNumeric(oTest) Then minGMROI = CInt(oTest)
        oTest = txtUnitsSold.Text
        If Not IsNothing(oTest) And IsNumeric(oTest) Then minQtySold = CInt(oTest)
        oTest = txtOnHand.Text
        If Not IsNothing(oTest) And IsNumeric(oTest) Then MinWksOH = CInt(oTest)
        Dim dept As String = Me.txtCateg.Text.Replace(", ", "','")
        Dim clss As String = Me.txtSubcat.Text.Replace(", ", "','")
        Dim season As String = Me.txtProfCod1.Text.Replace(", ", "','")
        Dim pline As String = Me.txtProfCod2.Text.Replace(", ", "','")

        con.Open()





        '' GoTo 20




        Try

            sql = "SELECT DISTINCT Str_Id, d.Item_No, Description, Dept, Buyer, Class, Season, PLine, Mktg_Code, " & _
                    "Vend_Item_No, UOM AS Pur_Unit, Curr_Cost AS Cost, Curr_Retail AS Retail, CONVERT(DECIMAL(18,4),0) AS Recvd, " & _
                    "CONVERT(DECIMAL(18,4),0) AS Sold, CONVERT(DECIMAL(18,4),0) AS d2Sold, CONVERT(DECIMAL(18,4),0) AS GrossMargin, " & _
                    "CONVERT(Integer,0) AS WksOH, dbo.fnGMROI(Str_Id, d.Item_No, '" & salesEdate & "') AS gmroi, " & _
                    "CONVERT(DECIMAL(18,4),0) AS InvCost, CONVERT(DECIMAL(18,4),0) AS SalesCost, " & _
                    "CONVERT(DECIMAL(18,4),0) AS SalesRetail, CONVERT(DECIMAL(18,4),0) AS OH, " & _
                    "CONVERT(DECIMAL(18,4),0) AS GProfit, CONVERT(DECIMAL(18,4),0) AS OnOrd, " & _
                    "CONVERT(VARCHAR(10),'') AS Score, CONVERT(DECIMAL(18,1),0) AS ForecastSales, " & _
                    "CONVERT(DECIMAL(18,4),0) AS MarkUp, CONVERT(Integer,0) AS MaxWksOH, Stock_Code INTO #t1 " & _
                    "FROM Item_Detail d JOIN Item_Master m ON m.Item_No = d.Item_No " & _
                    "WHERE Vendor_Id = '" & vendorId & "' AND eDate >= Dateadd(month,-18,GetDate()) "
            If itemStat = "'Active'" Or itemStat = "'Inactive'" Then
                sql &= "AND m.Status = " & itemStat & " "
            End If
            sql &= "GROUP BY Str_Id, d.Item_No, Description, Dept, Buyer, Class, Season, PLine, Mktg_Code, Vend_Item_No, " & _
                    "UOM, Curr_Cost, Curr_Retail, Stock_Code " & _
                    "SELECT d.Str_Id, d.Item_No, Description, Dept, m.Buyer, Class, Season, PLine, Mktg_Code, Vend_Item_No, UOM, Curr_Cost, " & _
                        "Curr_Retail, ISNULL(SUM(Qty_Due),0) AS Qty, 0 AS OH INTO #t1a FROM PO_Detail d " & _
                        "JOIN PO_Header h ON h.PO_NO = d.PO_NO " & _
                        "JOIN Item_Master m ON m.Item_No = d.Item_No " & _
                        "WHERE Qty_Due > 0 AND h.Vendor_Id = '" & vendorId & "' AND h.Due_Date < '" & coverageEdate & "' "
            If itemStat = "'Active'" Or itemStat = "'Inactive'" Then
                sql &= "AND m.Status = " & itemStat & " "
            End If
            sql &= "GROUP BY d.Str_Id, d.Item_No, Description, Dept, m.Buyer, Class, Season, PLine, Mktg_Code, Vend_Item_No, UOM, " & _
                        "Curr_Cost, Curr_Retail " & _
                    "MERGE #t1 AS t USING #t1a AS s ON (t.Str_Id = s.Str_Id AND t.Item_No = s.Item_No)  " & _
                        "WHEN MATCHED THEN Update SET t.OnOrd = s.Qty " & _
                        "WHEN NOT MATCHED BY TARGET THEN INSERT(Str_Id, Item_No, Description, Dept, Buyer, Class, Season, " & _
                            "PLine, Mktg_Code, Vend_Item_No, Pur_Unit, Cost, Retail, OnOrd, OH) " & _
                            "VALUES (Str_Id, Item_No, Description, Dept, Buyer, Class, Season, PLine, Mktg_Code, Vend_Item_No, UOM, " & _
                                "Curr_Cost, Curr_Retail, Qty, 0); " & _
                    "UPDATE t SET OH = (SELECT ISNULL(SUM(i.End_OH),0) FROM Item_Inv i " & _
                        "WHERE i.Item_No = t.Item_No And i.Str_Id = t.Str_Id AND '" & thisEdate & "' BETWEEN i.sDate AND i.eDate) FROM #t1 t " & _
                    "UPDATE t SET Score = (SELECT TOP 1 Score FROM Item_Detail d " & _
                        "WHERE d.Item_NO = t.Item_No AND d.Str_Id = t.Str_Id " & _
                        "AND eDate <= '" & thisEdate & "' AND Score IS NOT NULL ORDER BY eDate DESC) FROM #t1 t " & _
                    "UPDATE t SET ForecastSales = (SELECT ISNULL(SUM(Calculated_Demand),0) FROM Item_Forecast f " & _
                        "WHERE t.Str_Id = f.Str_Id AND t.Item_No = f.Item_No AND f.YearWk BETWEEN " & coverageLoYrWk & " AND " & coverageHiYrWk & " " & _
                        "AND ISNULL(Calculated_Demand,0) > 0) " & _
                            "FROM #t1 t " & _
                    "SELECT d.Str_Id, d.Item_No, " & _
                        "ISNULL(SUM(d.Sold * m.Curr_Cost),0) AS SalesCost, ISNULL(SUM(d.Sales_Retail),0) AS SalesRetail, " & _
                        "ISNULL(SUM(d.Sold),0) AS Sold, CONVERT(DECIMAL(18,4),0) AS d2Sold " & _
                        "INTO #t2 FROM Item_Detail d " & _
                        "JOIN #t1 t ON t.Str_Id = d.Str_Id AND t.Item_No = d.Item_No " & _
                        "JOIN Item_Master m ON m.Item_No = d.Item_No " & _
                        "WHERE d.eDate BETWEEN '" & salesSdate & "' AND '" & salesEdate & "' "
            If itemStat = "'Active'" Or itemStat = "'Inactive'" Then
                sql &= "AND m.Status = " & itemStat & " "
            End If

            ''sql &= "GROUP BY d.Str_Id, d.Item_No " & _
            ''        "UPDATE t SET t.RECVD = (SELECT ISNULL(SUM(d.RECVD),0) + ISNULL(SUM(d.XFER),0) - ISNULL(SUM(d.RTV),0) " & _
            ''            "FROM Item_Detail d JOIN Item_Master m ON m.Item_No = d.Item_No " & _
            ''            "WHERE eDate BETWEEN '" & salesSdate & "' AND '" & salesEdate & "' "
            sql &= "GROUP BY d.Str_Id, d.Item_No " & _
                    "UPDATE t SET t.RECVD = (SELECT ISNULL(SUM(Qty_Recvd),0) FROM PO_Detail p " & _
                    "WHERE Last_Recvd_Date BETWEEN '" & salesSdate & "' AND '" & salesEdate & "' " & _
                    "AND p.Str_Id = t.Str_Id AND p.Item_No = t.Item_No) + (SELECT ISNULL(SUM(d.XFER),0) - ISNULL(SUM(d.RTV),0) " & _
                        "FROM Item_Detail d JOIN Item_Master m ON m.Item_No = d.Item_No " & _
                        "WHERE eDate BETWEEN '" & salesSdate & "' AND '" & salesEdate & "' "
            If itemStat = "'Active'" Or itemStat = "'Inactive'" Then
                sql &= "AND m.Status = " & itemStat & " "
            End If
            sql &= " AND t.Str_Id = d.Str_Id AND t.Item_No = d.Item_No) FROM #t1 t " & _
                    "UPDATE t SET t.d2Sold = (SELECT ISNULL(SUM(Sold),0) FROM Item_Detail d WHERE t.Str_Id = d.Str_Id " & _
                        "AND d.Item_No = t.Item_No AND d.eDate BETWEEN '" & LYsalesSdate & "' AND '" & LYsalesEdate & "') FROM #t1 t " & _
                    "UPDATE t SET t.WksOH = (SELECT COUNT(*) FROM Item_Inv i WHERE Max_OH > 0 AND i.Str_Id = t.Str_Id " & _
                        "AND i.Item_No = t.Item_No AND i.eDate BETWEEN '" & lyStartDate & "' AND '" & lastweeksedate & "') FROM #t1 t " & _
                    "UPDATE t SET t.MaxWksOH = (SELECT COUNT(*) FROM Item_Inv i WHERE Max_OH > 0 AND i.Str_Id = t.Str_Id " & _
                        "AND i.Item_No = t.Item_No) FROM #t1 t " & _
                    "UPDATE t SET InvCost = (SELECT ISNULL(SUM(Max_OH * Curr_Cost),0) FROM Item_Inv i " & _
                        "JOIN Item_Master m ON m.Item_No = i.Item_No " & _
                        "WHERE Max_OH > 0 AND i.Str_Id = t.Str_Id " & _
                        "AND i.Item_No = t.Item_No AND i.eDate BETWEEN '" & lyStartDate & "' and '" & thisEdate & "') FROM #t1 t " & _
                    "UPDATE t SET t.MarkUp = (SELECT ISNULL(SUM(Sales_Retail - Sales_Cost),0) FROM Item_Detail d " & _
                        "WHERE d.Str_Id = t.Str_Id AND d.Item_No = t.Item_No " & _
                        "And d.eDate BETWEEN '" & lyStartDate & "' AND '" & thisEdate & "') FROM #t1 t " & _
                    "UPDATE t1 SET t1.Sold = t2.Sold, " & _
                        "t1.SalesCost = t2.SalesCost, t1.SalesRetail = t2.SalesRetail FROM #t1 t1 " & _
                        "JOIN #t2 t2 on t2.Str_Id = t1.Str_Id AND t1.Item_No = t2.Item_No " & _
                    "UPDATE #t1 SET GrossMargin = CASE Retail WHEN 0 THEN 0 ELSE (Retail - Cost) / Retail END " & _
                    "UPDATE #t1 SET gmroi = CASE WksOh WHEN 0 THEN 0 " & _
                        "ELSE (((MarkUp) / (InvCost / WksOH)) * 100) * 52 / WksOH END WHERE InvCost <> 0"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 240
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            oTest = ex.Message
            MsgBox(ex.Message)
            con.Close()
            Exit Sub
        End Try

        ' Add Warehouse On Hand, Received and On Order to Lakeland store
        If chkW1.Checked And chkLP.Checked And chkCombined.Checked = True Then
            sql = "SELECT DISTINCT '1' AS Str_Id, Item_No, Pur_Unit, Cost, Retail, Description, InvCost, OH, OnOrd, Recvd, " & _
                    "Dept, Class, Season, PLine, Mktg_Code " & _
                    "INTO #t4 FROM #t1 WHERE Str_Id = 'W1' " & _
                    "MERGE #t1 AS TARGET " & _
                        "USING (SELECT Str_Id, Item_No, Pur_Unit, Cost, Retail, Description, InvCost, OH, OnOrd, Recvd, Dept, " & _
                        "Class, Season, PLine, Mktg_Code FROM #t4) AS SOURCE " & _
                            "ON (TARGET.Str_Id = SOURCE.Str_Id AND TARGET.Item_No = SOURCE.Item_No) " & _
                    "WHEN MATCHED THEN " & _
                    "UPDATE SET TARGET.OH = TARGET.OH + SOURCE.OH, TARGET.OnOrd = TARGET.OnOrd + SOURCE.OnOrd, " & _
                        "TARGET.Recvd = TARGET.Recvd + SOURCE.Recvd, TARGET.InvCost = TARGET.InvCost + Source.InvCost " & _
                    "WHEN NOT MATCHED BY TARGET THEN " & _
                    "INSERT (Str_Id, Item_No, Pur_Unit, Cost, Retail, Description, OH, OnOrd, Recvd, InvCost, " & _
                        "GMROI, Dept, Class, Season, PLine, Mktg_Code) " & _
                        "VALUES(SOURCE.Str_Id, SOURCE.Item_No, SOURCE.Pur_Unit, SOURCE.Cost, SOURCE.Retail, " & _
                        "SOURCE.Description, SOURCE.OH, SOURCE.OnOrd, SOURCE.Recvd, SOURCE.InvCost, 0, SOURCE.Dept, SOURCE.Class, " & _
                        "SOURCE.Season, SOURCE.PLine, Mktg_Code); "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 240
            cmd.ExecuteNonQuery()
        End If
20:
        Dim qualifyFirst As Boolean = False
        If minGMROI > -9999 Then qualifyFirst = True
        If minQtySold > 0 Then qualifyFirst = True
        If MinWksOH > 0 Then qualifyFirst = True
        oTest = txtMinOH
        Dim minoh As Integer
        If IsNumeric(txtMinOH.Text) Then minoh = CInt(txtMinOH.Text)
        Dim maxoh As Integer
        If IsNumeric(txtMaxOH.Text) Then maxoh = CInt(txtMaxOH.Text)
        If minoh > -99 Then qualifyFirst = True
        If maxoh < 999999 Then qualifyFirst = True

        tbl = New DataTable
        tbl.Columns.Add("KeyId", GetType(System.String))
        tbl.Columns.Add("SortBy", GetType(System.String))
        tbl.Columns.Add("Item_No", GetType(System.String))
        theDataTable = New DataTable
        Dim keys(1) As DataColumn
        Dim column = New DataColumn()
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "KeyId"
        theDataTable.Columns.Add(column)
        keys(0) = column
        theDataTable.PrimaryKey = keys
        theDataTable.Columns.Add("Str_Id", GetType(System.String))
        theDataTable.Columns.Add("Item_No", GetType(System.String))
        theDataTable.Columns.Add("Score", GetType(System.String))
        theDataTable.Columns.Add("OrdQty", GetType(System.String))
        theDataTable.Columns.Add("Pur_Unit", GetType(System.String))
        theDataTable.Columns.Add("Cost", GetType(System.Decimal))
        theDataTable.Columns.Add("Retail", GetType(System.Decimal))
        theDataTable.Columns.Add("Descr", GetType(System.String))
        theDataTable.Columns.Add("OH", GetType(System.Decimal))
        theDataTable.Columns.Add("OnPO", GetType(System.Decimal))
        theDataTable.Columns.Add("4Cast", GetType(System.Decimal))                 ' added this
        theDataTable.Columns.Add("IF", GetType(System.Decimal))                  ' added this
        theDataTable.Columns.Add("Recv", GetType(System.Decimal))
        theDataTable.Columns.Add("Sold", GetType(System.Decimal))
        theDataTable.Columns.Add("d2Sold", GetType(System.Decimal))
        theDataTable.Columns.Add("SC", GetType(System.String))
        theDataTable.Columns.Add("GrossMargin", GetType(System.Decimal))
        theDataTable.Columns.Add("GM_Pct", GetType(System.Decimal))
        theDataTable.Columns.Add("GMROI", GetType(System.Decimal))
        theDataTable.Columns.Add("WksOH", GetType(System.Int16))
        theDataTable.Columns.Add("GProfit", GetType(System.Int16))
        theDataTable.Columns.Add("Dept", GetType(System.String))
        theDataTable.Columns.Add("Buyer", GetType(System.String))
        theDataTable.Columns.Add("Class", GetType(System.String))
        theDataTable.Columns.Add("Season", GetType(System.String))
        theDataTable.Columns.Add("PLine", GetType(System.String))
        theDataTable.Columns.Add("Mktg_Code", GetType(System.String))
        theDataTable.Columns.Add("Vend_Item_No", GetType(System.String))
        theDataTable.Columns.Add("Pur_Numer", GetType(System.String))
        theDataTable.Columns.Add("MaxWksOH", GetType(System.Int16))
        theDataTable.Columns.Add("MarkUp", GetType(System.Int32))
        theDataTable.Columns.Add("InvCost", GetType(System.Int32))
        displayTable = New DataTable
        Dim keys2(1) As DataColumn
        Dim column2 = New DataColumn()
        column2.DataType = System.Type.GetType("System.String")
        column2.ColumnName = "KeyId"
        displayTable.Columns.Add(column2)
        keys2(0) = column2
        displayTable.PrimaryKey = keys2
        displayTable.Columns.Add("Str_Id", GetType(System.String))
        displayTable.Columns.Add("Item_No", GetType(System.String))
        displayTable.Columns.Add("Descr", GetType(System.String))
        displayTable.Columns.Add("Score", GetType(System.String))
        displayTable.Columns.Add("GMROI", GetType(System.Decimal))
        displayTable.Columns.Add("OH", GetType(System.Decimal))
        displayTable.Columns.Add("OnPO", GetType(System.Decimal))
        displayTable.Columns.Add("4Cast", GetType(System.Decimal))                 ' added this
        displayTable.Columns.Add("IF", GetType(System.Decimal))             ' added this
        displayTable.Columns.Add("Recv", GetType(System.Decimal))
        displayTable.Columns.Add("Sold", GetType(System.Decimal))
        displayTable.Columns.Add("d2Sold", GetType(System.Decimal))
        displayTable.Columns.Add("OrdQty", GetType(System.String))
        displayTable.Columns.Add("SC", GetType(System.String))
        displayTable.Columns.Add("Pur_Unit", GetType(System.String))
        displayTable.Columns.Add("Cost", GetType(System.Decimal))
        displayTable.Columns.Add("Retail", GetType(System.Decimal))
        displayTable.Columns.Add("GrossMargin", GetType(System.Decimal))
        displayTable.Columns.Add("Dept", GetType(System.String))
        displayTable.Columns.Add("Buyer", GetType(System.String))
        displayTable.Columns.Add("Class", GetType(System.String))
        displayTable.Columns.Add("Season", GetType(System.String))
        displayTable.Columns.Add("PLine", GetType(System.String))
        displayTable.Columns.Add("Mktg_Code", GetType(System.String))
        displayTable.Columns.Add("Vend_Item_No", GetType(System.String))
        displayTable.Columns.Add("WksOH", GetType(System.Int16))
        displayTable.Columns.Add("MaxWksOH", GetType(System.Int16))
        displayTable.Columns.Add("MarkUp", GetType(System.Int32))
        displayTable.Columns.Add("InvCost", GetType(System.Int32))
        Dim item As String
        sql = "SELECT * FROM #t1 d WHERE Item_No IS NOT NULL "
        If dept <> "" Then sql = sql & "AND d.Dept IN ('" & dept & "') "
        If clss <> "" Then sql = sql & "AND d.Class IN ('" & clss & "') "
        If season <> "" Then sql = sql & "AND d.Season IN ('" & season & "') "
        If pline <> "" Then sql = sql & "AND d.PLine IN ('" & pline & "') "
        If chkByItem.Checked = True Then sql = sql & "AND d.Item_No='" & txtItem.Text & "' "
        sql &= "ORDER BY d.Item_No, d.Str_Id "
        cmd = New SqlCommand(sql, con)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim row, rw As DataRow
        Dim oh, forecast, val, onord, sold, d2sold, mgn As Decimal
        Dim maxwksoh As Integer
        While rdr.Read
            item = Nothing
            row = theDataTable.NewRow
            rw = tbl.NewRow
            row("KeyID") = rdr("Str_Id") & "*" & rdr("Item_No")
            row("Str_Id") = rdr("Str_Id")
            oTest = rdr("Item_No")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then item = oTest
            row("Item_No") = rdr("Item_No")
            row("Score") = rdr("Score")
            oTest = rdr("onord")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then onord = CDec(oTest)
            row("OnPO") = rdr("OnOrd")
            If rdr("pur_unit") = "EA" Then row("Pur_Unit") = "EACH" Else row("Pur_Unit") = rdr("pur_unit")
            row("Cost") = rdr("Cost")
            row("Retail") = rdr("retail")
            row("Descr") = rdr("Description")
            oTest = rdr("Buyer")
            row("Buyer") = rdr("Buyer")
            oTest = rdr("oh")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oh = CDec(oTest)
            Else : oh = 0
            End If
            row("OH") = oh
            row("OnPo") = rdr("onord")
            row("Recv") = rdr("recvd")
            oTest = rdr("InvCost")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                row("InvCost") = CDec(oTest)
            End If
            oTest = rdr("sold")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then
                    row("Sold") = CDec(oTest)
                    sold = CDec(oTest)
                Else : row("Sold") = 0
                End If
            End If
            oTest = rdr("d2sold")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then
                    row("d2Sold") = CDec(oTest)
                    d2sold = CDec(oTest)
                Else : row("d2Sold") = 0
                End If
            End If
            oTest = rdr("GrossMargin")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                row("GrossMargin") = CDec(oTest) * 100
            End If
            oTest = rdr("MarkUp")
            If Not IsDBNull(oTest) And IsNumeric(oTest) Then
                row("MarkUp") = CInt(oTest)
            Else : row("MarkUp") = 0
            End If

            row("GMROI") = rdr("GMROI")

            oTest = rdr("wksoh")
            row("WksOH") = rdr("wksoh")
            oTest = rdr("ForecastSales")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                forecast = CDec(oTest)
            Else : forecast = 0
            End If
            If forecast >= 0.5 Then
                val = (onord + oh + 0.01) / forecast
            Else : val = 9
            End If
            row("4Cast") = forecast                             ' added this
            row("IF") = val                                     ' added this
            oTest = rdr("MaxWksOH")                             ' set GMROI to Nothing when we've had an item < 10 weeks and haven't sold any

            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then
                    row("MaxWksOH") = CInt(oTest)
                    maxwksoh = CInt(oTest)
                    'If maxwksoh < 10 And sold = 0 Then row("GMROI") = -999999
                    If maxwksoh < 10 And rdr("MarkUp") = 0 Then row("GMROI") = -999999
                    If maxwksoh < 3 Then row("4Cast") = -999999
                End If
            End If

            row("Dept") = rdr("dept")
            oTest = rdr("GProfit")
            row("GProfit") = rdr("GProfit")
            row("Class") = rdr("class")
            row("Season") = rdr("season")
            row("PLine") = rdr("pline")
            row("Mktg_Code") = rdr("Mktg_Code")
            oTest = rdr("Vend_Item_No")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                row("Vend_Item_No") = oTest
            Else
                row("Vend_Item_No") = item
            End If
            oTest = rdr("MarkUp")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                row("MarkUp") = oTest
            End If
            rw("KeyId") = rdr("Str_Id") & "*" & rdr("Item_No")
            oTest = rdr(sortBy)
            If sortBy = "Item_No" Then
                rw("sortBy") = rdr("Item_No")
            Else : rw("sortBy") = String.Format("{0:0000000}", rdr(sortBy))
            End If
            rw("Item_No") = rdr("Item_No")

            If qualifyFirst = True Then
                If rdr("gmroi") >= minGMROI And rdr("wksoh") >= MinWksOH And rdr("oh") >= minoh And rdr("oh") <= maxoh And
                    rdr("sold") >= minQtySold And InStr(selectedStores, rdr("Str_Id")) > 0 Then
                    tbl.Rows.Add(rw)
                End If
            Else
                oTest = rdr("Item_No")
                If InStr(selectedStores, rdr("Str_Id")) > 0 And rdr("oh") >= minoh And rdr("oh") <= maxoh Then
                    tbl.Rows.Add(rw)
                End If
            End If
            row("SC") = rdr("Stock_Code")
            theDataTable.Rows.Add(row)
        End While
        con.Close()


        Dim cx As Integer = tbl.Rows.Count
        If cx = 0 Then
            MessageBox.Show("No records found for Vendor " & vendorId, "ERROR")
            lblProcessing.Visible = False
            txtVendors.Text = Nothing
            txtVendors.Select()
            Exit Sub
        End If

        Dim view As New DataView(tbl)                                      ' sort tbl in "sortBy" order
        If sortBy <> "Item_No" Then
            view.Sort = "sortBy DESC"
        Else
            view.Sort = "sortBy ASC"
        End If

        Dim testRow As DataRow
        Dim totD2Sold As Int32 = 0
        Dim totOH As Int32 = 0
        Dim totOnOrd As Int32 = 0
        Dim totRecvd As Int32 = 0
        Dim totSold As Int32 = 0
        Dim tot4cast As Decimal = 0
        Dim totInvCost As Decimal = 0
        Dim totMgn As Decimal = 0
        Dim totWksOH As Integer = 0
        For Each r As DataRowView In view                                    ' use sorted tbl to retreive data from theDataTable
            Dim key As String = r.Item(0).ToString
            Dim txt As String = key.ToString
            Dim txtArray() As String = txt.Split("*")
            Dim loc As String = txtArray(0).ToString
            Dim itm As String = txtArray(1).ToString
            Dim newRow As DataRow
            If InStr(selectedStores, ",") = 0 Then                                                 '  a single store was selected
                If qualifyFirst Then
                    ''key = "1*" & itm
                    key = selectedStores & "*" & itm
                    newRow = theDataTable.Rows.Find(key)
                    If newRow IsNot Nothing Then
                        Dim xRow As DataRow = displayTable.NewRow
                        Call CopyDX2keyTable(newRow, xRow)
                        testRow = displayTable.Rows.Find(key)                                      ' skip adding the new row to keyTable if its already there
                        If testRow Is Nothing Then
                            displayTable.Rows.Add(xRow)
                        End If
                    End If
                Else
                    If loc = selectedStores Then
                        key = selectedStores & "*" & itm
                        newRow = theDataTable.Rows.Find(key)
                        If newRow IsNot Nothing Then
                            Dim xrow As DataRow = displayTable.NewRow
                            Call CopyDX2keyTable(newRow, xrow)
                            displayTable.Rows.Add(xrow)
                        End If
                    End If
                End If
            Else
                For Each value As String In stores
                    If value <> "W1" Then
                        '  key = "1*" & itm                                                                 '  2 or more stores were selected
                        key = value & "*" & itm
                        newRow = theDataTable.Rows.Find(key)
                        If newRow IsNot Nothing Then
                            Dim xRow As DataRow = displayTable.NewRow
                            Call CopyDX2keyTable(newRow, xRow)
                            testRow = displayTable.Rows.Find(key)                                        ' skip adding the new row to keyTable if its already there
                            If testRow Is Nothing Then
                                displayTable.Rows.Add(xRow)
                            End If
                        End If
                    End If
                Next
            End If
10:     Next
        For Each rw In displayTable.Rows
            '  create a total row and add it to displayTable
            oTest = rw("GrossMargin")
            oTest = rw.Item("OH")
            If Not IsDBNull(oTest) Then
                totOH += rw.Item("OH")
            End If
            oTest = rw.Item("OnPO")
            If Not IsDBNull(oTest) Then
                totOnOrd += rw.Item("OnPO")
            End If
            oTest = rw.Item("Recv")
            If Not IsDBNull(oTest) Then
                totRecvd += rw.Item("Recv")
            End If
            oTest = rw.Item("InvCost")
            If Not IsDBNull(oTest) Then
                totInvCost += CDec(oTest)
            End If
            oTest = rw.Item("Sold")
            If Not IsDBNull(oTest) Then
                totSold += rw.Item("Sold")
            End If
            oTest = rw.Item("d2Sold")
            If Not IsDBNull(oTest) Then
                totD2Sold += rw.Item("d2Sold")
            End If
            oTest = rw("4Cast")
            If Not IsDBNull(oTest) Then
                If CDec(oTest) <> -999999 Then tot4cast += CDec(oTest)
            End If
            oTest = rw("MarkUp")
            If Not IsDBNull(oTest) Then
                totMgn += CDec(rw("MarkUp"))
            End If
            oTest = rw("MaxWksOH")
            If Not IsDBNull(oTest) Then
                If CInt(oTest) > totWksOH Then totWksOH = CInt(oTest)
            End If
        Next
        Dim gmroi As Integer = 0
        If totWksOH > 0 And totInvCost > 0 Then
            gmroi = ((totMgn / (totInvCost / totWksOH)) * 100) * (52 / totWksOH)
        End If
        Preview.totOnPO = totOnOrd
        Preview.totOH = totOH
        Preview.tot4Cast = tot4cast
        Preview.totRecv = totRecvd
        Preview.totSold = totSold
        Preview.totd2Sold = totD2Sold
        Preview.totGMROI = gmroi
        Dim paramsTxt As String
        paramsTxt = vendorId & " Loc(s): " & displayStores & " Coverage: " & coverageSdate & " - " & coverageEdate & "  "
        paramsTxt &= "Sales Dates: " & salesSdate & " - " & salesEdate & " "
        paramsTxt &= "Min GMROI: " & minGMROI & " Min Sold: " & minQtySold & " Min Weeks OnHand: " & MinWksOH
        Dim reviewResults As New Preview
        reviewResults.ShowResults(displayTable, theDataTable, paramsTxt)
        txtProgress.Visible = False
        txtVendors.Text = Nothing
        Me.Refresh()
        reviewResults.Show()
    End Sub

    Public Sub CopyDX2keyTable(ByRef dxRow, ByVal keyRow)
        Try
            keyRow("KeyId") = dxRow("KeyId")
            keyRow("Str_Id") = dxRow("Str_Id")
            keyRow("Item_No") = dxRow("Item_No")
            keyRow("Descr") = dxRow("Descr")
            keyRow("Score") = dxRow("Score")
            keyRow("GMROI") = dxRow("GMROI")
            keyRow("OH") = dxRow("OH")
            keyRow("OnPO") = dxRow("OnPO")
            keyRow("4Cast") = dxRow("4Cast")
            keyRow("Recv") = dxRow("Recv")
            keyRow("Sold") = dxRow("Sold")
            keyRow("d2Sold") = dxRow("d2Sold")
            '' keyRow("QrdQty") = dxRow("OrdQty")
            keyRow("Pur_Unit") = dxRow("Pur_Unit")
            keyRow("Cost") = dxRow("Cost")
            keyRow("Retail") = dxRow("Retail")
            keyRow("GrossMargin") = dxRow("GrossMargin")
            keyRow("Dept") = dxRow("Dept")
            keyRow("Buyer") = dxRow("Buyer")
            keyRow("Class") = dxRow("Class")
            keyRow("Season") = dxRow("Season")
            keyRow("PLine") = dxRow("Pline")
            keyRow("Mktg_Code") = dxRow("Mktg_Code")
            keyRow("Vend_Item_No") = dxRow("Vend_Item_No")
            keyRow("WksOH") = dxRow("Wksoh")
            keyRow("IF") = dxRow("IF")
            keyRow("MaxWksOH") = dxRow("MaxWksOH")
            keyRow("MarkUp") = dxRow("MarkUp")
            keyRow("InvCost") = dxRow("InvCost")
            keyRow("SC") = dxRow("SC")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Get_Vendor()
        Dim sql As String = "SELECT vendor_Id, Vendor FROM Item_Master WHERE Item_No = '" & txtItem.Text & "'"
        con.Open()
        Dim cmd As New SqlCommand(sql, con)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            txtVendors.Text = rdr("vendor_Id")
            vendorId = rdr("vendor_Id")
            vendorName = rdr("Vendor")
        End While
        rdr.Close()
        con.Close()
        Me.Refresh()
    End Sub
    Private Sub Check_Lines(allGood)
        Dim values As String = txtProfCod2.Text
        If values = "" Then Exit Sub
        If values.Contains(",") Then
            If values.Contains(", ") Then
            Else
                Me.txtProfCod2.Text = values.Replace(",", ", ")
                values = txtProfCod2.Text
            End If
        End If
        Dim lines As String() = Nothing
        lines = values.Split(", ")
        Dim pline As String
        Dim v As String
        Dim sql As String
        Dim txt As String
        Dim recs As Integer = 0
        For Each v In lines
            pline = v.ToString
            sql = "SELECT * FROM ProductLines WHERE ID = '" & pline & "'"
            Dim sqlcmd As New SqlCommand(sql, con)
            con.Open()
            Dim sqlReader As SqlDataReader = sqlcmd.ExecuteReader
            While sqlReader.Read
                txt = sqlReader(0).ToString
                recs += 1
                If UCase(txt.ToString) <> UCase(pline) Then
                    MsgBox(vendorId & " is not a valid Product Line")
                    allGood = False
                End If
            End While
            con.Close()
            If recs = 0 Then allGood = False
        Next
    End Sub
    Private Sub check_Seasons(allGood)
        Dim values As String = txtProfCod1.Text
        If values = "" Then Exit Sub
        If values.Contains(",") Then
            If values.Contains(", ") Then
            Else
                Me.txtProfCod1.Text = values.Replace(",", ", ")
                values = txtProfCod1.Text
            End If
        End If
        Dim seasons As String() = Nothing
        seasons = values.Split(", ")
        Dim season As String
        Dim v As String
        Dim sql As String
        Dim txt As String
        Dim recs As Integer = 0
        For Each v In seasons
            season = v.ToString
            sql = "SELECT * FROM Seasons WHERE ID = '" & season & "'"
            Dim sqlcmd As New SqlCommand(sql, con)
            con.Open()
            Dim sqlReader As SqlDataReader = sqlcmd.ExecuteReader
            While sqlReader.Read
                txt = sqlReader(0).ToString
                recs += 1
                If UCase(txt) <> UCase(season) Then
                    allGood = False
                End If
            End While
            con.Close()
            If recs = 0 Then allGood = False
        Next
    End Sub
    Private Sub Check_Departments(allGood)
        Dim values As String = txtCateg.Text
        If values = "" Then Exit Sub
        If values.Contains(",") Then
            If values.Contains(", ") Then
            Else
                Me.txtCateg.Text = values.Replace(",", ", ")
                values = txtCateg.Text
            End If
        End If
        Dim departments As String() = Nothing
        departments = values.Split(", ")
        Dim category As String
        Dim v As String
        Dim sql As String
        Dim txt As String
        Dim recs As Integer = 0
        For Each v In departments
            category = v.ToString
            sql = "SELECT ID FROM Departments WHERE ID='" & category & "'"
            Dim sqlcmd As New SqlCommand(sql, con)
            con.Open()
            Dim sqlReader As SqlDataReader = sqlcmd.ExecuteReader
            While sqlReader.Read
                txt = sqlReader(0).ToString
                recs += 1
                If UCase(txt.ToString) <> UCase(category) Then
                    allGood = False
                End If
            End While
            con.Close()
            If recs = 0 Then allGood = False
        Next
    End Sub
    Private Sub Check_Classes(allGood)
        Dim values As String = txtSubcat.Text
        If values = "" Then Exit Sub
        If values.Contains(",") Then
            If values.Contains(", ") Then
            Else
                Me.txtSubcat.Text = values.Replace(",", ", ")
                values = txtSubcat.Text
            End If
        End If
        Dim classes As String() = Nothing
        classes = values.Split(", ")
        Dim subcat As String
        Dim v As String
        Dim sql As String
        Dim txt As String
        Dim recs As Integer = 0
        For Each v In classes
            subcat = v.ToString
            sql = "SELECT ID FROM Classes WHERE ID ='" & subcat & "'"
            Dim sqlcmd As New SqlCommand(sql, con)
            con.Open()
            Dim sqlReader As SqlDataReader = sqlcmd.ExecuteReader
            While sqlReader.Read
                txt = sqlReader(0).ToString
                recs += 1
                If UCase(txt.ToString) <> UCase(subcat) Then
                    allGood = False
                End If
            End While
            con.Close()
            If recs = 0 Then allGood = False
        Next
    End Sub
    Private Sub Check_Vendors(allGood)
        Dim values As String = txtVendors.Text
        Dim vendors As String() = Nothing
        vendors = values.Split(", ")
        Dim vendorID As String
        Dim v As String
        Dim sql As String
        Dim txt As String
        If txtVendors.Text = "" Then
            MsgBox("No vendor has been selected")
            allGood = False
            Exit Sub
        End If
        For Each v In vendors
            vendorID = UCase(v.ToString)
            sql = "SELECT ID AS Vend_No, Description AS Name FROM Vendors WHERE ID = '" & vendorID & "'"
            Dim sqlcmd As New SqlCommand(sql, con)
            con.Open()
            sqlcmd.ExecuteNonQuery()
            Dim sqlReader As SqlDataReader = sqlcmd.ExecuteReader
            Dim recs As Integer = 0
            While sqlReader.Read
                txt = sqlReader(0).ToString
                recs += 1
                If UCase(txt.ToString) <> UCase(vendorID) Then
                    MsgBox(vendorID & " is not a valid vendor code")
                    allGood = False
                End If
            End While
            If recs = 0 Then allGood = False
            sqlReader.Close()
            sqlReader.Dispose()
            con.Close()
            ''con.Dispose()
        Next
    End Sub

    Private Sub cboCfrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCfrom.SelectedIndexChanged
        coverageSdate = cboCfrom.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, coverageSdate, coverageEdate) + 1
        txtCweeks.Text = diff
        con.Open()
        sql = "SELECT YrWk FROM Calendar WHERE sDate = '" & coverageSdate & "' AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            coverageLoYrWk = rdr("YrWk")
        End While
        con.Close()
    End Sub

    Private Sub cboCthru_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCthru.SelectedIndexChanged
        coverageEdate = cboCthru.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, coverageSdate, coverageEdate) + 1
        txtCweeks.Text = diff
        con.Open()
        sql = "SELECT YrWk FROM Calendar WHERE eDate = '" & coverageEdate & "' AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            coverageHiYrWk = rdr("YrWk")
        End While
        con.Close()
    End Sub

    Private Sub cboSfrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSfrom.SelectedIndexChanged
        salesSdate = cboSfrom.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, salesSdate, salesEdate) + 1
        txtSweeks.Text = diff
    End Sub

    Private Sub cboSthru_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSthru.SelectedIndexChanged
        salesEdate = cboSthru.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, salesSdate, salesEdate) + 1
        txtSweeks.Text = diff
    End Sub

    Private Sub cboS2from_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboS2from.SelectedIndexChanged
        LYsalesSdate = cboS2from.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, LYsalesSdate, LYsalesEdate) + 1
        txtS2Weeks.Text = diff
    End Sub

    Private Sub cboS2thru_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboS2thru.SelectedIndexChanged
        LYsalesEdate = cboS2thru.SelectedItem
        Dim diff = DateDiff(DateInterval.WeekOfYear, LYsalesSdate, LYsalesEdate) + 1
        txtS2Weeks.Text = diff
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub chkVendor_CheckedChanged(sender As Object, e As EventArgs) Handles chkVendor.CheckedChanged
        Select Case UCase(user)
            Case "CURTIS"
                thisBuyer = "CGW"
            Case "CRIS"
                thisBuyer = "CGW"
            Case "GARY"
                thisBuyer = "CGW"
            Case "CHAD"
                thisBuyer = "CGW"
            Case "KAREN"
                thisBuyer = "MKL"
            Case "RENEE"
                thisBuyer = "RCI"
            Case Else
                thisBuyer = "NA"
        End Select
        'conString = "Server=Development;Initial Catalog=PARGIF;Integrated Security=True"
        'con = New SqlConnection(conString)

        cboVendor.Items.Clear()
        cboVendor.Items.Add("Enter Vendor Code Below or Select From Drop Down at Right")

        con.Open()
        If chkVendor.Checked Then
            sql = "SELECT DISTINCT m.Vendor_Id, m.Vendor FROM Item_Master m JOIN Vendors v ON v.ID = m.Vendor_Id " & _
                "WHERE v.Status = 'Active' AND m.Buyer = '" & thisBuyer & "' " & _
                "ORDER BY m.Vendor"
        Else : sql = "SELECT DISTINCT ID AS Vendor_Id, Description AS Vendor FROM Vendors ORDER BY Description"
        End If
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim id, vendor As String
        While rdr.Read
            id = rdr("Vendor_Id")
            vendor = rdr("Vendor")
            cboVendor.Items.Add(id & " - " & vendor)
        End While
        con.Close()
        cboVendor.SelectedIndex = 0
    End Sub

End Class
Imports System.IO
Imports System.Data.SqlClient
Imports System.Web.UI

Public Class VendorInventoryAnalysisReport
    Inherits System.Web.UI.Page
    Private Shared conString As String
    Private Shared con As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared sql As String
    Private Shared rdr As SqlDataReader
    Private Shared oTest As Object
    Private Shared thisStore, thisDept, thisBuyer, thisClass, thisSeason As String
    Private Shared thisEdate, tyHiDate, tyLoDate, lyHiDate, lyLoDate As Date
    Private Shared tbl As DataTable


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        conString = "Server=LP-CURTIS;Initial Catalog=PARGIF;Integrated Security=SSPI"
        Dim con As New SqlConnection(conString)
        con.Open()
        sql = "SELECT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddStore.Items.Add(rdr("ID"))
        End While
        con.Close()

        ddDept.Items.Add("ALL")
        con.Open()
        sql = "SELECT ID FROM Departments WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddDept.Items.Add(rdr("ID"))
        End While
        con.Close()

        ddClass.Items.Add("ALL")
        con.Open()
        sql = "SELECT ID FROM Classes WHERE Status = 'Active' AND Dept = '" & ddDept.SelectedItem.Value & "' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddClass.Items.Add(rdr("ID"))
        End While
        con.Close()

        ddBuyer.Items.Add("ALL")
        con.Open()
        sql = "SELECT ID FROM Buyers WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddBuyer.Items.Add(rdr("ID"))
        End While
        con.Close()

        ddSeason.Items.Add("ALL")
        con.Open()
        sql = "SELECT ID FROM Seasons WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddSeason.Items.Add(rdr("ID"))
        End While
        con.Close()

        Call Get_Dates()

    End Sub


    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles btnRunReport.Click
        Dim minOH As Integer = 0
        If IsNumeric(txtMinOH.Text) Then
            minOH = CInt(txtMinOH.Text)
        Else : minOH = 0
        End If
        thisStore = ddStore.SelectedItem.Value
        thisBuyer = ddBuyer.SelectedItem.Value
        thisDept = ddDept.SelectedItem.Value
        thisSeason = ddSeason.SelectedItem.Value
        If IsDBNull(thisClass) Or IsNothing(thisClass) Then Call Load_Classes(thisDept)
        thisClass = ddClass.SelectedItem.Value
        Dim wk1, wk4, wk5, wk8, wk9, wk12, wk13, wk16, l6, l17, lyp1, lyp17, lyf1, lyf17 As Date
        Dim oh, due, n4, n8, n12, n16, avail, lysales, ly17sales, tysales, lycost, tycost, tlycost, ttycost As Integer
        Dim toh, tdue, tn4, tn8, tn12, tn16, tavail, tlysales, tly17sales, ttysales As Integer
        Dim gm, trend As Decimal
        wk1 = DateAdd(DateInterval.Day, 7, thisEdate)
        wk4 = DateAdd(DateInterval.Day, 28, thisEdate)
        wk5 = DateAdd(DateInterval.Day, 35, thisEdate)
        wk8 = DateAdd(DateInterval.Day, 56, thisEdate)
        wk9 = DateAdd(DateInterval.Day, 63, thisEdate)
        wk12 = DateAdd(DateInterval.Day, 84, thisEdate)
        wk13 = DateAdd(DateInterval.Day, 91, thisEdate)
        wk16 = DateAdd(DateInterval.Day, 112, thisEdate)
        l6 = DateAdd(DateInterval.Month, -6, thisEdate)
        l17 = DateAdd(DateInterval.Day, -119, thisEdate)
        lyp1 = DateAdd(DateInterval.Day, -371, thisEdate)
        lyp17 = DateAdd(DateInterval.Day, -483, thisEdate)
        lyf1 = DateAdd(DateInterval.Day, -364, thisEdate)
        lyf17 = DateAdd(DateInterval.Day, -252, thisEdate)
        tbl = New DataTable
        Dim row As DataRow
        tbl.Columns.Add("Vendor ID", Type.GetType("System.String"))
        tbl.Columns.Add("GMROI", Type.GetType("System.String"))
        tbl.Columns.Add("OnHand", Type.GetType("System.String"))
        tbl.Columns.Add("Due", Type.GetType("System.String"))
        tbl.Columns.Add("Next4", Type.GetType("System.String"))
        tbl.Columns.Add("Next8", Type.GetType("System.String"))
        tbl.Columns.Add("Next12", Type.GetType("System.String"))
        tbl.Columns.Add("Next16", Type.GetType("System.String"))
        tbl.Columns.Add("INV Avail", Type.GetType("System.String"))
        tbl.Columns.Add("LYSales", Type.GetType("System.String"))
        tbl.Columns.Add("GM%", Type.GetType("System.String"))
        tbl.Columns.Add("TYSales", Type.GetType("System.String"))
        tbl.Columns.Add("GM %", Type.GetType("System.String"))
        tbl.Columns.Add("Trend", Type.GetType("System.String"))

        con = New SqlConnection(conString)
        con.Open()
        sql = "SELECT Vendor_Id, ISNULL(SUM(End_OH * Retail),0) AS onHand, " & _
            "CONVERT(Decimal,0) AS Due, CONVERT(Decimal,0) AS next4, CONVERT(Decimal,0) AS next8, " & _
            "CONVERT(Decimal,0) AS next12, CONVERT(Decimal,0) AS next16, CONVERT(Decimal,0) AS totalonorder, " & _
            "CONVERT(Decimal,0) AS lySales, CONVERT(Decimal,0) AS lyCost, CONVERT(Decimal,0) AS L17Sales, " & _
            "CONVERT(Decimal,0) AS L17Cost, CONVERT(Decimal,0) AS LY17Sales, CONVERT(Decimal,0) AS LY17Cost " & _
            "INTO #inv FROM Item_Inv i " & _
            "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
            "JOIN Item_Master m ON m.Sku = i.Sku " & _
            "WHERE eDate = '" & thisEdate & "' "
        If thisStore <> "All" Then sql &= "AND s.ID = '" & thisStore & "' "
        If thisDept <> "ALL" Then sql &= "AND m.Dept = '" & thisDept & "' "
        If thisBuyer <> "ALL" Then sql &= "AND m.Buyer = '" & thisBuyer & "' "
        If thisClass <> "ALL" Then sql &= "AND m.Class = '" & thisClass & "' "
        If thisSeason <> "ALL" Then sql &= "AND m.Custom1 = '" & thisSeason & "' "
        sql &= "GROUP BY Vendor_Id "
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "SELECT m.Vendor_Id, Due_Date, CONVERT(Decimal,0) AS onHand, " & _
            "CASE WHEN Due_Date >= '" & l6 & "' AND Due_Date < '" & wk1 & "' THEN ISNULL(SUM(Qty_Due * Curr_Retail),0) ELSE 0 END AS Due, " & _
            "CASE WHEN Due_Date BETWEEN '" & wk1 & "' AND '" & wk4 & "' THEN ISNULL(SUM(Qty_Due * Curr_Retail),0) ELSE 0 END AS next4, " & _
            "CASE WHEN Due_Date BETWEEN '" & wk5 & "' AND '" & wk8 & "' THEN ISNULL(SUM(Qty_Due * Curr_Retail),0) ELSE 0 END AS next8, " & _
            "CASE WHEN Due_Date BETWEEN '" & wk9 & "' AND '" & wk12 & "' THEN ISNULL(SUM(Qty_Due * Curr_Retail),0) ELSE 0 END AS next12, " & _
            "CASE WHEN Due_Date BETWEEN '" & wk13 & "' AND '" & wk16 & "' THEN ISNULL(SUM(Qty_Due * Curr_Retail),0) ELSE 0 END AS next16, " & _
            "CONVERT(Decimal,0) AS lyCost, CONVERT(Decimal,0) AS lySales, CONVERT(Decimal,0) AS L17Sales, " & _
            "CONVERT(Decimal,0) AS L17Cost, CONVERT(Decimal,0) AS LY17Sales, CONVERT(Decimal,0) AS LY17Cost " & _
            "INTO #t1 FROM PO_Detail d " & _
            "JOIN Stores s ON s.Inv_Loc = d.Loc_Id " & _
            "JOIN PO_Header h ON h.PO_No = d.Po_No AND h.Loc_Id = d.Loc_Id " & _
            "JOIN Item_Master m ON m.Sku = d.Sku "
        If thisStore <> "ALL" Then sql &= "AND s.ID = '" & thisStore & "' "
        If thisDept <> "ALL" Then sql &= "AND m.Dept = '" & thisDept & "' "
        If thisBuyer <> "ALL" Then sql &= "AND m.Buyer = '" & thisBuyer & "' "
        If thisClass <> "ALL" Then sql &= "AND m.Class = '" & thisClass & "' "
        If thisSeason <> "ALL" Then sql &= "AND m.Season = '" & thisSeason & "' "
        sql &= "GROUP BY m.Vendor_Id, Due_Date"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "SELECT Vendor_Id, SUM(Due) AS Due, SUM(next4) AS Next4, SUM(next8) AS Next8, SUM(next12) As Next12, " & _
            "SUM(next16) AS Next16, SUM(Due) + SUM(next4) + SUM(next8) + SUM(next12) + SUM(next16) AS totalonorder, " & _
            "SUM(lySales) AS lySales, SUM(lyCost) AS lyCost, SUM(L17Sales) As L17Sales, SUM(L17Cost) AS L17Cost, " & _
            "SUM(LY17Sales) AS LY17Sales, SUM(LY17Cost) AS LY17Cost  INTO #due FROM #t1 GROUP BY Vendor_Id"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "SELECT Vendor_Id, eDate, CONVERT(Decimal,0) AS onHand, CONVERT(Decimal,0) AS Due, " & _
            "CONVERT(Decimal,0) AS next4, CONVERT(Decimal,0) AS next8, CONVERT(Decimal,0) AS next12, " & _
            "CONVERT(Decimal,0) AS next16, CONVERT(Decimal,0) AS totalonorder, " & _
            "CASE WHEN eDate BETWEEN '" & lyf1 & "' AND '" & lyf17 & "' THEN SUM(Sales_retail) ELSE 0 END AS lySales, " & _
            "CASE WHEN eDate BETWEEN '" & lyf1 & "' AND '" & lyf17 & "' THEN SUM(Sales_Cost) ELSE 0 END AS lyCost, " & _
            "CASE WHEN eDate BETWEEN '" & l17 & "' AND '" & tyHiDate & "' THEN SUM(Sales_Retail) ELSE 0 END AS L17Sales, " & _
            "CASE WHEN eDate BETWEEN '" & l17 & "' AND '" & tyHiDate & "' THEN SUM(Sales_Cost) ELSE 0 END AS L17Cost, " & _
            "CASE WHEN eDate BETWEEN '" & lyp17 & "' AND '" & lyp1 & "' THEN SUM(Sales_Retail) ELSE 0 END AS LY17Sales, " & _
            "CASE WHEN eDate BETWEEN '" & lyp17 & "' AND '" & lyp1 & "' THEN SUM(Sales_Cost) ELSE 0 END AS LY17Cost  " & _
            "INTO #t2 FROM Item_Sales as d " & _
            "JOIN Item_Master AS m ON m.Sku = d.Sku " & _
            "WHERE eDate BETWEEN '" & lyp17 & "' AND '" & tyHiDate & "' "
        If thisStore <> "ALL" Then sql &= "AND d.Loc_Id = '" & thisStore & "' "
        If thisDept <> "ALL" Then sql &= "AND Dept = '" & thisDept & "' "
        If thisBuyer <> "ALL" Then sql &= "AND Buyer = '" & thisBuyer & "' "
        If thisClass <> "ALL" Then sql &= "AND Class = '" & thisClass & "' "
        If thisSeason <> "ALL" Then sql &= "AND Season = '" & thisSeason & "' "
        sql &= "GROUP BY Vendor_Id, eDate " & _
            "SELECT Vendor_Id, SUM(onHand) As onHand, SUM(Due) AS Due, SUM(next4) AS next4, SUM(next8) AS next8, " & _
            "SUM(next12) AS next12, SUM(next16) AS next16, SUM(totalonorder) AS totalonorder, SUM(lySales) AS lySales,  " & _
            "SUM(lyCost) As lyCost, SUM(L17Sales) AS L17Sales, SUM(L17Cost) AS L17Cost, SUM(LY17Sales) AS LY17Sales, " & _
            "SUM(LY17Cost) AS LY17Cost INTO #sales FROM #t2 GROUP BY Vendor_Id"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "MERGE #sales AS Target " & _
             "USING #inv AS Source " & _
                "ON Target.Vendor_Id = Source.Vendor_Id " & _
            "WHEN MATCHED THEN " & _
                "UPDATE SET " & _
                    "Target.onHand = Source.onHand, " & _
                    "Target.Due = Source.Due, " & _
                    "Target.next4 = Source.next4, " & _
                    "Target.next8 = Source.next8, " & _
                    "Target.next12 = Source.next12, " & _
                    "Target.next16 = Source.next16, " & _
                    "Target.totalonorder = Source.totalonorder " & _
            "WHEN NOT MATCHED BY TARGET THEN " & _
                    "INSERT(Vendor_id, due, next4, next8, next12, next16, totalonorder, lysales, " & _
                    "lycost, L17Sales, L17Cost, ly17sales, LY17Cost) " & _
               "VALUES (Source.Vendor_id,Source.Due, Source.next4, Source.next8, Source.next12, Source.next16, Source.totalonorder, " & _
                        "Source.lySales, Source.lyCost, Source.L17Sales, Source.L17Cost, Source.LY17Sales, Source.LY17Cost);"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()

        sql = "MERGE  #sales AS Target " & _
             "USING #due AS Source " & _
                "ON Target.Vendor_Id = Source.Vendor_Id " & _
            "WHEN MATCHED THEN " & _
                "UPDATE SET " & _
                    "Target.Due = Source.Due, " & _
                    "Target.next4 = Source.next4, " & _
                    "Target.next8 = Source.next8, " & _
                    "Target.next12 = Source.next12, " & _
                    "Target.next16 = Source.next16, " & _
                    "Target.totalonorder = Source.totalonorder " & _
            "WHEN NOT MATCHED BY TARGET THEN " & _
                    "INSERT(Vendor_id, due, next4, next8, next12, next16, totalonorder, lysales, " & _
                    "lycost, L17Sales, L17Cost, ly17sales, LY17Cost) " & _
               "VALUES (Source.Vendor_id,Source.Due, Source.next4, Source.next8, Source.next12, Source.next16, Source.totalonorder, " & _
                        "Source.lySales, Source.lyCost, Source.L17Sales, Source.L17Cost, Source.LY17Sales, Source.LY17Cost);"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()


        ''sql = "SELECT s.vendor_id, (select gmroi from vendor_stats v1 where v1.vendor_id = s.vendor_id and dept = '" & thisDept & "' " & _
        ''    "and edate = (select max(edate) from vendor_stats v2 where v2.vendor_id = s.vendor_id and dept = '" & thisDept & "')) as gmroi, " & _
        ''    "onHand, due, next4, next8, next12, next16, lycost, lysales, L17Sales, L17Cost, LY17Sales, LY17Cost FROM #sales s ORDER BY onHand DESC"
        sql = "SELECT s.vendor_id, 0 as gmroi, " & _
            "onHand, due, next4, next8, next12, next16, lycost, lysales, L17Sales, L17Cost, LY17Sales, LY17Cost FROM #sales s ORDER BY onHand DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        While rdr.Read
            row = tbl.NewRow
            oTest = rdr("vendor_id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row("Vendor ID") = oTest
            oTest = rdr("gmroi")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row("GMROI") = oTest
            oTest = rdr("onHand")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oh = CInt(oTest)
                toh += CInt(oTest)
                row("OnHand") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("Due")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                due = CInt(oTest)
                tdue += CInt(oTest)
                row("Due") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("Next4")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                n4 = CInt(oTest)
                tn4 += CInt(oTest)
                row("Next4") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("Next8")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                n8 = CInt(oTest)
                tn8 += CInt(oTest)
                row("Next8") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("Next12")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                n12 = CInt(oTest)
                tn12 += CInt(oTest)
                row("Next12") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("Next16")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                n16 = CInt(oTest)
                tn16 += CInt(oTest)
                row("Next16") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("onHand")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                avail = oh + due + n4 + n8 + n12 + n16
                tavail += oh + due + n4 + n8 + n12 + n16
                row("Inv Avail") = Format(CInt(avail), "###,###,###")
            End If
            oTest = rdr("LYSales")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                lysales = CInt(oTest)
                tlysales += CInt(oTest)
                row("LYSales") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("LYCost")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                lycost = CInt(oTest)
                tlycost += CInt(oTest)
            End If
            If lycost <> 0 And lysales <> 0 Then
                gm = (lysales - lycost) / lysales
            Else : gm = 0
            End If
            row("GM%") = Format(gm, "##0%")
            oTest = rdr("L17Sales")
            If Not IsDBNull(oTest) And Not IsDBNull(oTest) Then
                tysales = CInt(oTest)
                ttysales += CInt(oTest)
                row("TYSales") = Format(CInt(oTest), "###,###,###")
            End If
            oTest = rdr("L17Cost")
            If Not IsDBNull(oTest) And Not IsDBNull(oTest) Then
                tycost = CInt(oTest)
                ttycost += CInt(oTest)
            End If
            If tycost <> 0 And tysales <> 0 Then
                gm = (tysales - tycost) / tysales
            Else : gm = 0
            End If
            row("GM %") = Format(gm, "###%")
            oTest = rdr("LY17Sales")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                ly17sales = CInt(oTest)
                tly17sales += CInt(oTest)
            End If
            If tysales <> 0 And ly17sales <> 0 Then
                trend = (tysales / ly17sales) - 1
                row("Trend") = Format(trend, "###%")
            End If
            If oh >= minOH Then tbl.Rows.Add(row)
        End While
        con.Close()

        row = tbl.NewRow
        row("Vendor ID") = "Totals"
        row("onHand") = Format(toh, "###,###,###")
        row("Due") = Format(tdue, "###,###,###")
        row("Next4") = Format(tn4, "###,###,###")
        row("Next8") = Format(tn8, "###,###,###")
        row("Next12") = Format(tn12, "###,###,###")
        row("Next16") = Format(tn16, "###.###,###")
        row("Inv Avail") = Format(tavail, "###,###,###")
        row("LYSales") = Format(tlysales, "###,###,###")
        If tlycost <> 0 And tlysales <> 0 Then
            gm = (tlysales - tlycost) / tlysales
        Else : gm = 0
        End If
        row("GM%") = Format(gm, "##0%")
        row("TYSales") = Format(ttysales, "###,###,###")
        If ttycost <> 0 And ttysales <> 0 Then
            gm = (ttysales - ttycost) / ttysales
        Else : gm = 0
        End If
        row("GM %") = Format(gm, "##0%")
        If tlysales <> 0 And ttysales <> 0 Then
            trend = (ttysales / tly17sales) - 1
            row("Trend") = Format(trend, "##0%")
        End If
        tbl.Rows.Add(row)


        'Dim rws As Integer = tbl.Rows.Count


        dgv1.DataSource = tbl.DefaultView
        dgv1.DataBind()


        'dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'For i As Integer = 1 To dgv1.Columns.Count - 1
        'dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        'dgv1.Columns(i).Width = 70
        'If i > 1 And i < 7 Then
        '    dgv1.Columns(i).DefaultCellStyle.Font = New Font("Veranda", 8, FontStyle.Underline)
        '    dgv1.Columns(i).DefaultCellStyle.ForeColor = Color.Blue
        'End If
        'Next
        'Dim r As Integer = dgv1.Rows.Count - 1
        'dgv1.Rows(r).DefaultCellStyle.Font = New Font("Veranda", 8, FontStyle.Bold)
        'dgv1.Rows(r).DefaultCellStyle.ForeColor = Color.Black
        'dgv1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        con.Open()
        sql = "DECLARE @smalldate smalldatetime=getdate() " &
            "DECLARE @today date = @smalldate " &
            "SELECT Loc_Id, Dept, Buyer, Class, sDate, eDate, " &
            "Case When @today Between sDate and eDate Then Sum(ISNULL(OTB,0)) Else 0 End As otb, " &
            "Case When Datediff(day, @today, eDate) Between 29 And 35 Then OTB Else 0 End As otb1, " &
            "Case When Datediff(day, @today, eDate) Between 57 And 63 Then OTB Else 0 End As otb2, " &
            "Case When Datediff(day, @today, eDate) Between 85 And 91 Then OTB Else 0 End As otb3, " &
            "Case When Datediff(day, @today, eDate) Between 113 And 119 Then OTB Else 0 End As otb4 " &
            "Into #t3 From VI_OTB Where Loc_Id Is Not NULL "
        If thisStore <> "ALL" Then sql &= "AND Loc_Id = '" & thisStore & "' "
        If thisDept <> "ALL" Then sql &= "AND Dept = '" & thisDept & "' "
        If thisClass <> "ALL" Then sql &= "AND Class = '" & thisClass & "' "
        If thisBuyer <> "ALL" Then sql &= "AND Buyer = '" & thisBuyer & "' "
        sql &= "group by Loc_Id, Dept, Buyer, Class, sDate, eDate, OTB " &
            "Select distinct dept, " &
            "Case When Datediff(day, @today, eDate) Between 29 And 35 Then sDate  End As dte1, " &
            "Case When Datediff(day, @today, eDate) Between 57 And 63 Then sDate  End As dte2, " &
            "Case When Datediff(day, @today, eDate) Between 85 And 91 Then sDate  End As dte3, " &
            "Case When Datediff(day, @today, eDate) Between 113 And 119 Then sDate  End As dte4 " &
            "into #t4 from Sales_Summary "
        If thisDept <> "ALL" Then sql &= " where dept='" & thisDept & "' "
        If thisClass <> "ALL" Then sql &= " AND Class='" & thisClass & "' "
        sql &= "select dept, sum(otb) as now, sum(otb1) as next4, sum(otb2) as next8, sum(otb3) as next12, sum(otb4) as next16 " & _
            "into #t5 from #t3 group by dept " & _
            "select #t5.now, " & _
            "Case When Now <> 0 THEN NEXT4 - Now ELSE Next4 End AS Next4, " & _
            "CASE When Next4 <> 0 Then Next8 - Next4 Else Next8 End As Next8, " & _
            "Case When Next8 <> 0 THEN NEXT12 - Next8 ELSE Next12 End AS Next12, " & _
            "CASE When Next12 <> 0 Then Next16 - Next12 Else Next16 End As Next16, " & _
            "max(dte1) as dte1, max(dte2) as dte2, max(dte3) as dte3, max(dte4) as dte4 " & _
            "into #x from #t5 inner join #t4 on #t4.dept=#t5.dept "
        If thisDept = "ALL" Then
            sql &= "group by now, next4, next8, next12, next16"
        Else
            sql &= "group by #t5.dept, now, next4, next8, next12, next16"
        End If
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        sql = "SELECT dte1, dte2, dte3, dte4, SUM(now) as Today, SUM(Next4) AS Next4, SUM(Next8) AS Next8, " & _
            "SUM(Next12) AS Next12, SUM(Next16) AS Next16 FROM #x GROUP BY dte1, dte2, dte3, dte4"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        ''row = otbTbl.NewRow
        Dim dadate As Date
        Dim x As String
        While rdr.Read
            oTest = rdr("dte1")
            If Not IsDBNull(oTest) Then
                dadate = rdr("dte1")
                x = dadate.ToString("MM/dd/yyyy")
                lblDate4.Text = Microsoft.VisualBasic.Left(x, 5)
            End If
            oTest = rdr("dte2")
            If Not IsDBNull(oTest) Then
                dadate = rdr("dte2")
                x = dadate.ToString("MM/dd/yyyy")
                lblDate8.Text = Microsoft.VisualBasic.Left(x, 5)
            End If
            oTest = rdr("dte3")
            If Not IsDBNull(oTest) Then
                dadate = rdr("dte3")
                x = dadate.ToString("MM/dd/yyyy")
                lblDate12.Text = Microsoft.VisualBasic.Left(x, 5)
            End If
            oTest = rdr("dte4")
            If Not IsDBNull(oTest) Then
                dadate = rdr("dte4")
                x = dadate.ToString("MM/dd/yyyy")
                lblDate16.Text = Microsoft.VisualBasic.Left(x, 5)
            End If
            lblNow.Text = Format(rdr("Today"), "###,###,###")
            lblNex4.Text = Format(rdr("Next4"), "###,###,###")
            lblDate8.Text = Format(rdr("Next8"), "###,###,###")
            lblDate12.Text = Format(rdr("Next12"), "###,###,###")
            lblDate16.Text = Format(rdr("Next16"), "###,###,###")
        End While
        con.Close()

        If thisSeason = "ALL" Then
            lblToday.Visible = True
            lblDate12.Visible = True
            lblDate16.Visible = True
            lblDate4.Visible = True
            lblDate8.Visible = True
            lblDate16.Visible = True
            lblNex4.Visible = True
            lblDate12.Visible = True
            lblDate16.Visible = True
            lblDate8.Visible = True
            lblNow.Visible = True
        End If
    End Sub
    Private Sub Get_Dates()
        Dim yrwk, thisYear, thisWeek As Integer
        lblReportDate.Text = Date.Today
        con = New SqlConnection(conString)
        con.Open()
        sql = "SELECT MAX(eDate) AS eDate, MAX(YrWks) AS YrWk FROM Calendar " & _
            "WHERE eDate <= '" & Date.Today & "' AND Week_Id > 0 "
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            tyHiDate = rdr("eDate")
            yrwk = rdr("YrWk")
        End While
        con.Close()
        tyLoDate = DateAdd(DateInterval.Day, -111, tyHiDate)
        thisEdate = DateAdd(DateInterval.Day, 7, tyHiDate)
        con.Open()
        sql = "SELECT Year_Id, Week_Id FROM Calendar WHERE eDate = '" & tyHiDate & "' AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            thisYear = rdr("Year_Id")
            thisWeek = rdr("Week_Id")
        End While
        con.Close()

        con.Open()

        sql = "SELECT eDate AS eDate FROM Calendar WHERE Year_Id = " & thisYear - 1 & " AND Week_Id = " & thisWeek + 1 & " "
        ''sql = "SELECT eDate FROM Calendar WHERE '" & DateAdd(DateInterval.Day, -365, tyHiDate) & "' BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            lyLoDate = rdr("eDate")
        End While
        con.Close()

        lyHiDate = DateAdd(DateInterval.Day, 111, lyLoDate)
        lblLYSales.Text = lyLoDate & " - " & lyHiDate
        lblTYSales.Text = tyLoDate & " - " & tyHiDate
    End Sub

    Protected Sub ddStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddStore.SelectedIndexChanged
        oTest = ddStore.SelectedItem.Value
        If Not IsDBNull(oTest) Then
            thisStore = CStr(oTest)
        End If
    End Sub

    Protected Sub ddDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddDept.SelectedIndexChanged
        oTest = ddDept.SelectedItem.Value
        If Not IsDBNull(oTest) Then
            thisDept = CStr(oTest)
            Call Load_Classes(thisDept)
        End If
    End Sub

    Protected Sub ddClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddClass.SelectedIndexChanged
        oTest = ddClass.SelectedItem.Value
        If Not IsDBNull(oTest) Then
            thisClass = ddClass.SelectedItem.Value
        End If
    End Sub

    Protected Sub ddBuyer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddBuyer.SelectedIndexChanged
        oTest = ddBuyer.SelectedItem.Value
        If Not IsDBNull(oTest) Then
            thisBuyer = CStr(oTest)
        End If
    End Sub

    Protected Sub ddSeason_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddSeason.SelectedIndexChanged
        oTest = ddSeason.SelectedItem.Value
        If Not IsDBNull(oTest) Then
            thisSeason = CStr(oTest)
        End If
    End Sub

    Private Sub Load_Classes(ByVal thisDept As String)
        con = New SqlConnection(conString)
        con.Open()
        sql = "SELECT ID FROM Classes WHERE Status = 'Active' AND Dept = '" & thisDept & "' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddClass.Items.Add(rdr("ID"))
        End While
        con.Close()
        ddClass.SelectedIndex = 0
        thisClass = ddClass.SelectedItem.Value
    End Sub
End Class
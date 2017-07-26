Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class ScoringSetup
    Public Shared con As SqlConnection
    Private Shared cmd As SqlCommand
    Public Shared sql, thisStore, conString As String
    Private Shared rdr As SqlDataReader
    Private storeTbl, scoreTbl, scoreTbl2, scoreTbl3, dateTbl, atbl, btbl, ctbl, tbl, tbl2, universeTbl As DataTable
    Private Shared startDate, endDate As Date
    Private Shared lastUpdate As DateTime = Nothing
    Private Shared oTest As Object
    Private Shared changesMade As Boolean
    Private Shared weeks As Integer
    Private Shared aPct, Bpct, Cpct, Dpct, Epct, pctArray() As Decimal
    Public Shared client, server, dbase, sqluserid, sqlpassword, exepath As String
    Private Shared row, row2, row3, foundRow As DataRow

    Private Sub ScoringSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conString = MainMenu.conString
        con = MainMenu.con
        client = MainMenu.Client_Id
        server = MainMenu.Server
        dbase = MainMenu.dBase
        sqluserid = MainMenu.UserID
        sqlpassword = MainMenu.Password
        exepath = MainMenu.exePath
        lblServer.Text = MainMenu.serverLabel.Text
        lblServer2.Text = MainMenu.serverLabel.Text
        lblServer3.Text = MainMenu.serverLabel.Text
        lblServer4.Text = MainMenu.serverLabel.Text
        lblServer5.Text = MainMenu.serverLabel.Text
        lblServer6.Text = MainMenu.serverLabel.Text
        lblServer7.Text = MainMenu.serverLabel.Text
        lblServer8.Text = MainMenu.serverLabel.Text
        changesMade = False
        con.Open()
        Dim today As Date = Date.Today
        sql = "SELECT MAX(eDate) AS eDate FROM Item_Inv WHERE eDate <= '" & today & "'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            endDate = rdr("eDate")
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT sDate FROM Sales_Summary WHERE sDate < GETDATE() ORDER BY sDate DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboStartDate.Items.Add(rdr("sDate"))
        End While
        cboStartDate.SelectedIndex = 0
        con.Close()

        con.Open()
        storeTbl = New DataTable
        storeTbl.Columns.Add("Store")
        Dim srow As DataRow
        sql = "SELECT Str_Id FROM Stores WHERE Status = 'Active' ORDER BY Str_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            srow = storeTbl.NewRow
            srow("Store") = rdr("Str_Id")
            storeTbl.Rows.Add(srow)
            cboDeptStore.Items.Add(rdr("Str_Id"))
            cboBuyerStore.Items.Add(rdr("Str_Id"))
            cboClassStore.Items.Add(rdr("Str_Id"))
            cboVendorStore.Items.Add(rdr("Str_Id"))
            cboPLineStore.Items.Add(rdr("Str_Id"))
            cboItemStore.Items.Add(rdr("Str_Id"))
            cboScoreStore.Items.Add(rdr("Str_Id"))
        End While
        con.Close()
        scoreTbl = New DataTable
        Dim column = New DataColumn()
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "Code"
        scoreTbl.Columns.Add(column)
        Dim PrimaryKey(1) As DataColumn
        PrimaryKey(0) = scoreTbl.Columns("Code")
        scoreTbl.PrimaryKey = PrimaryKey
        scoreTbl.Columns.Add("Break$", GetType(System.Decimal))
        scoreTbl.Columns.Add("GM%", GetType(System.Decimal))
        scoreTbl.Columns.Add("%SKU", GetType(System.Decimal))
        scoreTbl.Columns.Add("CumGM%", GetType(System.Decimal))
        scoreTbl.Columns.Add("Skus", GetType(System.Int32))
        Dim row As DataRow
        row = scoreTbl.NewRow
        Dim arr() As String = {"A", "B", "C", "D", "E"}
        pctArray = {0.5, 0.3, 0.15, 0.04, 0.01}
        Dim tpct As Decimal = 0
        For Each element As String In arr
            row = scoreTbl.NewRow
            row("Code") = element
            scoreTbl.Rows.Add(row)
        Next
        '                 Set Defaults
        txtWeeks.Text = 26
        aPct = pctArray(0)
        Bpct = pctArray(1)
        Cpct = pctArray(2)
        Dpct = pctArray(3)
        Epct = pctArray(4)
        con.Open()
        sql = "SELECT Last_Update, Weeks, Apct, Bpct, Cpct, Dpct, Epct FROM Score_Setup_History " & _
            "WHERE Last_Update = (SELECT MAX(Last_Update) FROM Score_Setup_History)"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            lastUpdate = rdr("Last_Update")
            weeks = rdr("Weeks")
            aPct = rdr("Apct")
            pctArray(0) = aPct
            Bpct = rdr("Bpct")
            pctArray(1) = Bpct
            Cpct = rdr("Cpct")
            pctArray(2) = Cpct
            Dpct = rdr("Dpct")
            pctArray(3) = Dpct
            Epct = rdr("Epct")
            pctArray(4) = Epct
        End While
        con.Close()

        If Not IsNothing(weeks) Then
            txtWeeks.Text = weeks
        End If
        txtGP1.Text = Format(aPct, "###%")
        scoreTbl.Rows(0).Item("GM%") = aPct
        txtGP2.Text = Format(Bpct, "###%")
        scoreTbl.Rows(1).Item("GM%") = Bpct
        txtGP3.Text = Format(Cpct, "###%")
        scoreTbl.Rows(2).Item("GM%") = Cpct
        txtGP4.Text = Format(Dpct, "###%")
        scoreTbl.Rows(3).Item("GM%") = Dpct
        txtGP5.Text = Format(Epct, "###%")
        If Not IsNothing(lastUpdate) Then
            lblLastUpdate.Text = lastUpdate
            lblLastUpdate.Visible = True
            lblLastUpdte.Visible = True
            Me.Refresh()
        End If

        ''For i As Integer = 0 To 4
        ''    tpct += pctArray(i)
        ''    scoreTbl.Rows(i).Item("GM%") = pctArray(i)
        ''    scoreTbl.Rows(i).Item("CumGM%") = tpct
        ''Next


    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        lblProcessing.Visible = True
        Me.Refresh()
        Call Save_Changes()
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub

    Private Sub Save_Changes()
        'Try
        Dim totl As Integer = 0
        oTest = Replace(txtGP1.Text, "%", "")
        If IsNumeric(oTest) Then
            aPct = oTest * 0.01
            pctArray(0) = aPct
            txtGP1.Text = Format(aPct, "###%")
            totl += CInt(oTest)
            scoreTbl.Rows(0).Item("GM%") = CDec(oTest * 0.01)
            scoreTbl.Rows(0).Item("CumGM%") = CDec(totl * 0.01)
        End If
        oTest = Replace(txtGP2.Text, "%", "")
        If IsNumeric(oTest) Then
            Bpct = oTest * 0.01
            pctArray(1) = Bpct
            txtGP2.Text = Format(Bpct, "###%")
            totl += CInt(oTest)
            scoreTbl.Rows(1).Item("GM%") = CDec(oTest * 0.01)
            scoreTbl.Rows(1).Item("CumGM%") = CDec(totl * 0.01)
        End If
        oTest = Replace(txtGP3.Text, "%", "")
        If IsNumeric(oTest) Then
            Cpct = oTest * 0.01
            pctArray(2) = Cpct
            txtGP3.Text = Format(Cpct, "###%")
            totl += CInt(oTest)
            scoreTbl.Rows(2).Item("GM%") = CDec(oTest * 0.01)
            scoreTbl.Rows(2).Item("CumGM%") = CDec(totl * 0.01)
        End If
        oTest = Replace(txtGP4.Text, "%", "")
        If IsNumeric(oTest) Then
            Dpct = oTest * 0.01
            pctArray(3) = Dpct
            txtGP4.Text = Format(Dpct, "###%")
            totl += CInt(oTest)
            scoreTbl.Rows(3).Item("GM%") = CDec(oTest * 0.01)
            scoreTbl.Rows(3).Item("CumGM%") = CDec(totl * 0.01)
        End If
        oTest = Replace(txtGP5.Text, "%", "")
        If IsNumeric(oTest) Then
            Epct = oTest * 0.01
            pctArray(4) = Epct
            txtGP5.Text = Format(Epct, "###%")
            totl += CInt(oTest)
            scoreTbl.Rows(4).Item("GM%") = CDec(oTest * 0.01)
            scoreTbl.Rows(4).Item("CumGM%") = CDec(totl * 0.01)
        End If

        If totl <> 100 Then
            MessageBox.Show("OOPS! GM% distribution does not equal 100%", "Gross Profit Break Point")
            Exit Sub
        End If
        weeks = CInt(txtWeeks.Text)
        oTest = con.State
        con.Open()
        sql = "INSERT INTO Score_Setup_History (Last_Update, Weeks, Apct, Bpct, Cpct, Dpct, Epct) " & _
            "SELECT '" & Date.Now & "'," & weeks & "," & aPct & "," & Bpct & "," & Cpct & "," & Dpct & "," & Epct & " "
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        txtProgress.Visible = True
        Dim daycnt As Integer = CInt(txtWeeks.Text) * -7
        startDate = DateAdd(DateInterval.Day, daycnt + 1, endDate)
        Dim scoreTypeArray() As String = {"Sku", "Dept", "Buyer", "Class", "Vendor_Id", "PLine"}


        ''Call Get_GrossMargin_Params("FS", "Sku")
        ''Call Get_Turns_Params("BD", "Vendor_Id")
        ''Call Get_UnitCost_Params("TP", "Dept")


        For Each row As DataRow In storeTbl.Rows
            thisStore = row("Store")
            For Each typ As String In scoreTypeArray
                txtProgress.Text = "Processing " & typ & " for Store " & thisStore
                Me.Refresh()
                Call Get_GrossMargin_Params(thisStore, typ)
                Call Get_Turns_Params(thisStore, typ)
                Call Get_UnitCost_Params(thisStore, typ)
            Next
        Next
        changesMade = True
        txtProgress.Text = "Scoring setup complete."
        Me.Refresh()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Save")
        'End Try
    End Sub

    Private Sub Get_GrossMargin_Params(thisStore, typ)                                     ' Compute paramaters for Gross Margin
        Try
            ''For Each drow As DataRow In dateTbl.Rows
            Dim thisEDate, thisSDate As Date
            Dim weeksBack As Integer
            thisEDate = endDate
            weeksBack = weeks
            txtProgress.Text = "Computing Gross Margin for " & typ
            Me.Refresh()
            Dim tgm As Integer = 0
            Dim summgn As Integer = 0
            Dim tcnt As Integer = 0
            Dim sumcnt As Integer = 0
            Dim val As Decimal
            Dim records As Integer
            '
            '                                       Get the date range
            '
            con.Open()
            sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                "WHERE Week_Id > 0 AND eDate = '" & thisEDate & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSDate = rdr("sDate")
                thisEDate = rdr("eDate")
            End While
            con.Close()

            tbl = New DataTable
            tbl.Columns.Add(typ, GetType(System.String))
            tbl.Columns.Add("gm", GetType(System.Int16))
            tbl.Columns.Add("pctmgn", GetType(System.Decimal))
            tbl.Columns.Add("pctsku", GetType(System.Decimal))


            aTbl = New DataTable
            aTbl.Columns.Add(typ, GetType(System.String))
            aTbl.Columns.Add("GrossMargin", GetType(System.Decimal))
            aTbl.Columns.Add("Pct", GetType(System.Decimal))
            aTbl.Columns.Add("iPct", GetType(System.Decimal))
            '
            '                        Zero Score Table Gross Margin
            '
            For i As Integer = 0 To scoreTbl.Rows.Count - 1
                scoreTbl.Rows(i).Item(1) = 0
                scoreTbl.Rows(i).Item(5) = 0
            Next
            '
            '                        Get the gross margin for every sku within the date range and stick it in a table in descending GM sequence
            ' 
            con.Open()
            sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Retail - Sales_Cost),0) AS grossmargin " & _
                    "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                    "WHERE Str_id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' " & _
                    "AND eDate <= '" & thisEDate & "' AND Sales_Retail > 0 " & _
                    "GROUP BY m." & typ & " ORDER BY grossmargin DESC"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = aTbl.NewRow
                oTest = rdr("typ")
                row(typ) = rdr("typ")
                val = rdr("grossmargin")
                If val > 0 Then
                    tcnt += 1
                    summgn += val
                End If
                row("GrossMargin") = val
                aTbl.Rows.Add(row)
            End While
            con.Close()

            records = atbl.Rows.Count
            If records = 0 Then Exit Sub

            Dim cnt As Integer = 0
            Dim pct As Decimal
            Dim view As New DataView(aTbl)
            view.Sort = "GrossMargin DESC"                     ' sort the table in descending gross margin sequence
            For Each rw As DataRowView In view
                cnt += 1
                oTest = rw(0)
                val = rw("grossmargin")
                tgm += val                                         ' add gross margin to total gross margin
                If summgn > 0 Then pct = tgm / summgn Else pct = 0 ' compute this items percentage of the total gross margin
                rw("Pct") = pct                                    ' save the result in a table
                If tcnt > 0 Then rw("iPct") = cnt / tcnt ' save this items percentage of the total record count
                For j As Integer = 0 To 4                       ' compare this items GM percentage to the percentage that correspond with A-E
                    Dim it As Decimal = scoreTbl(j).Item(4)
                    oTest = scoreTbl.Rows(j).Item("Break$")
                    If pct >= scoreTbl.Rows(j).Item("CumGM%") And scoreTbl.Rows(j).Item("Break$") = 0 Then
                        scoreTbl.Rows(j).Item("Break$") = val
                    End If
                Next
            Next
            Dim Acnt, bCnt, cCnt, Dcnt, Ecnt, gm As Decimal
            Acnt = 0
            bCnt = 0
            cCnt = 0
            Dcnt = 0
            Ecnt = 0
            '
            '                            Loop thru aTbl and determine how many A's, B's, C's, D's and E's we have based on each rows GM
            '

            For Each row As DataRow In aTbl.Rows
                'If typ = "Dept" Then
                '    oTest = scoreTbl.Rows(0).Item(1)
                'End If
                gm = row("grossmargin")
                If gm >= scoreTbl.Rows(0).Item("Break$") Then Acnt += 1
                If gm >= scoreTbl.Rows(1).Item("Break$") And gm < scoreTbl.Rows(0).Item("Break$") Then bCnt += 1
                If gm >= scoreTbl.Rows(2).Item("Break$") And gm < scoreTbl.Rows(1).Item("Break$") Then cCnt += 1
                If gm >= scoreTbl.Rows(3).Item("Break$") And gm < scoreTbl.Rows(2).Item("Break$") Then Dcnt += 1
                If gm >= scoreTbl.Rows(4).Item("Break$") And gm < scoreTbl.Rows(3).Item("Break$") Then Ecnt += 1
            Next
            '
            '                           Save the record count and percentage of total record count in the scoreTbl by Score
            '
            If tcnt > 0 Then
                scoreTbl.Rows(0).Item("%Sku") = CDec(Acnt / tcnt)
                scoreTbl.Rows(1).Item("%Sku") = CDec(bCnt / tcnt)
                scoreTbl.Rows(2).Item("%Sku") = CDec(cCnt / tcnt)
                scoreTbl.Rows(3).Item("%Sku") = CDec(Dcnt / tcnt)
                scoreTbl.Rows(4).Item("%Sku") = CDec(Ecnt / tcnt)
            End If
            scoreTbl.Rows(0).Item("Skus") = Acnt
            scoreTbl.Rows(1).Item("Skus") = bCnt
            scoreTbl.Rows(2).Item("Skus") = cCnt
            scoreTbl.Rows(3).Item("Skus") = Dcnt
            scoreTbl.Rows(4).Item("Skus") = Ecnt
            Dim id, code As String
            Dim amt, ipct As Decimal
            Dim now As DateTime = Date.Now
            For Each row As DataRow In scoreTbl.Rows
                amt += row(2)
            Next
            con.Open()
            For Each row As DataRow In scoreTbl.Rows
                id = "GrossMargin"
                ''type = typ
                code = row(0)
                oTest = row(1)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = row(1) Else amt = 0
                oTest = row(2)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(2) Else pct = 0
                oTest = row(3)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ipct = row(3) Else ipct = 0
                Dim dbTyp As String = ""
                Select Case typ
                    Case "Sku"
                        dbTyp = "Sku"
                    Case "Dept"
                        dbTyp = "Dept"
                    Case "Class"
                        dbTyp = "Class"
                    Case "Buyer"
                        dbTyp = "Buyer"
                    Case "Vendor_Id"
                        dbTyp = "Vendor"
                    Case "PLine"
                        dbTyp = "PLine"
                End Select
                sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                    "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "') " & _
                    "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                    "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','" & dbTyp & "','" & _
                    code & "'," & amt & "," & pct & "," & ipct & ",'" & now & "' " & _
                    "ELSE " & _
                    "UPDATE Score_Params SET Value1 = " & amt & ", Value2 = " & pct & ", Value3 = " & ipct & ", Last_Update = '" & now & "' " & _
                    "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "' " & _
                    "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
            Next
            con.Close()
            ''Next
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Get_Turns_Params(thisStore, typ)
        '                                                    Compute paramaters from Turns
        Try
            Dim records As Integer
            Dim dbTyp As String
            Dim turns As Decimal
            Select Case typ
                Case "Sku"
                    dbTyp = "Sku"
                Case "Dept"
                    dbTyp = "Dept"
                Case "Class"
                    dbTyp = "Class"
                Case "Buyer"
                    dbTyp = "Buyer"
                Case "Vendor_Id"
                    dbTyp = "Vendor_Id"
                Case "PLine"
                    dbTyp = "PLine"
            End Select

            txtProgress.Text = "Computing Turns Parameters " & thisStore & " " & typ
            Me.Refresh()
            ''For Each drow As DataRow In dateTbl.Rows
            Dim thisSDate, thisEDate As Date
            thisEDate = endDate
            Dim weeksBack As Integer = weeks
            con.Open()
            sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                "WHERE eDate = '" & thisEDate & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSDate = rdr("sDate")
                thisEDate = rdr("eDate")
            End While
            con.Close()

            scoreTbl2 = New DataTable
            scoreTbl2.Columns.Add("Dept", GetType(System.String))
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = scoreTbl2.Columns("Dept")
            scoreTbl2.PrimaryKey = PrimaryKey2
            scoreTbl2.Columns.Add("F", GetType(System.Decimal))
            scoreTbl2.Columns.Add("M", GetType(System.Decimal))
            scoreTbl2.Columns.Add("S", GetType(System.Decimal))
            scoreTbl2.Columns.Add("records", GetType(System.Int32))

            Dim now As DateTime = Date.Now
            If typ = "Sku" Then
                con.Open()
                sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    row2 = scoreTbl2.NewRow
                    oTest = rdr("Dept")
                    row2(0) = rdr("Dept")
                    scoreTbl2.Rows.Add(row2)
                End While
                con.Close()

                btbl = New DataTable
                btbl.Columns.Add("cnt", GetType(System.Int32))
                btbl.Columns.Add("dept", GetType(System.String))
                btbl.Columns.Add("sku", GetType(System.String))
                btbl.Columns.Add("invcost", GetType(System.Decimal))
                btbl.Columns.Add("wksoh", GetType(System.Int32))
                btbl.Columns.Add("COGS", GetType(System.Decimal))
                btbl.Columns.Add("turns", GetType(System.Decimal))
                Dim cnt, wks, fCnt, mCnt As Integer
                Dim val, cogs, invcost As Decimal
                Dim dept, item As String
                Dim prevDept As String = ""
                Dim foundRow As DataRow
                con.Open()
                sql = "SELECT m.Dept, d.Sku, CONVERT(Decimal,0) AS invcost, ISNULL(SUM(Sales_Cost),0) As COGS, " & _
                  "CONVERT(Integer,0) AS wksOH INTO #t1 FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                  "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                  "GROUP BY Str_Id, m.Dept, d.Sku " & _
                  "SELECT m.Dept, i.Sku, ISNULL(SUM(Max_OH * Curr_Retail),0) AS invcost, " & _
                  "COUNT(*) As wksOH INTO #t2 FROM Item_Inv i JOIN Item_Master m ON m.Sku = i.Sku " & _
                  "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                  "WHERE s.Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                  "GROUP BY Str_Id, m.Dept, i.Sku " & _
                  "UPDATE #t1 SET #t1.invcost = #t2.invcost, #t1.wksOH = #t2.wksOH FROM #t1 " & _
                  "JOIN #t2 ON #t2.Dept = #t1.Dept AND #t2.Sku = #t1.Sku " & _
                  "SELECT Dept, Sku, invcost, wksoh, COGS FROM #t1 ORDER BY Dept"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    dept = rdr("Dept")
                    If dept <> prevDept Then
                        If prevDept = "" Then
                            prevDept = dept
                        Else
                            foundRow = scoreTbl2.Rows.Find(prevDept)
                            foundRow(4) = cnt
                            cnt = 0
                            prevDept = dept
                        End If
                    End If
                    If rdr("COGS") > 0 Then
                        cnt += 1
                    End If
                    row = btbl.NewRow
                    row("cnt") = cnt
                    row("dept") = rdr("Dept")
                    row("Sku") = rdr("Sku")
                    row("invcost") = rdr("invcost")
                    row("wksoh") = rdr("wksOH")
                    row("COGS") = rdr("COGS")
                    row("turns") = 0
                    btbl.Rows.Add(row)
                    item = rdr("Sku")

                End While
                con.Close()

                foundRow = scoreTbl2.Rows.Find(prevDept)
                If IsNothing(foundRow) Then Exit Sub

                foundRow(4) = cnt
                prevDept = ""
                For Each row As DataRow In btbl.Rows
                    wks = row("wksoh")
                    cogs = row("COGS")
                    invcost = row("invcost")
                    If wks > 0 And invcost > 0 Then
                        val = (row("cogs") / (invcost / row("wksoh"))) * (52 / row("wksoh"))
                        row("turns") = val
                    Else : row("turns") = 0
                    End If
                Next
                cnt = 0
                fCnt = 0
                mCnt = 0
                Dim view As New DataView(btbl)
                view.Sort = "Dept ASC, turns DESC, Sku ASC"
                For Each rw As DataRowView In view
                    cnt += 1
                    dept = rw("dept")
                    val = rw("turns")
                    item = rw("Sku")
                    If cnt = fCnt Then
                        foundRow = scoreTbl2.Rows.Find(prevDept)
                        foundRow(1) = val
                    End If
                    If cnt = mCnt Then
                        foundRow = scoreTbl2.Rows.Find(prevDept)
                        foundRow(2) = val
                        foundRow(3) = 0.01
                    End If
                    If dept <> prevDept Then
                        foundRow = scoreTbl2.Rows.Find(dept)
                        fCnt = foundRow(4) / 3
                        mCnt = fCnt * 2
                        cnt = 1
                        prevDept = dept
                    End If
                Next
                '                                   Save to Sql table
                Dim id, type, code As String
                Dim pct, pct2, pct3 As Decimal


                con.Open()
                For Each row As DataRow In scoreTbl2.Rows
                    id = "Turns"
                    type = "Item"
                    code = row(0)
                    oTest = row(1)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(1) Else pct = 0
                    oTest = row(2)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct2 = row(2) Else pct2 = 0
                    oTest = row(3)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct3 = row(3) Else pct3 = 0
                    sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                        "AND eDate = '" & thisEDate & "' AND Type = 'Item' AND Code = '" & code & "') " & _
                        "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                        "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','Item','" & _
                        code & "'," & pct & "," & pct2 & "," & pct3 & ",'" & now & "' " & _
                        "ELSE " & _
                        "UPDATE Score_Params SET Value1 = " & pct & ", Value2 = " & pct2 & ", Value3 = " & pct3 & ", Last_Update = '" & now & "' " & _
                        "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = '" & code & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                Next
                con.Close()
            Else
                Dim cnt As Integer = 0
                btbl = New DataTable
                btbl.Columns.Add("dept", GetType(System.String))
                btbl.Columns.Add("turns", GetType(System.String))
                con.Open()
                If typ = "Buyer" Or typ = "Class" Or typ = "Dept" Then
                    sql = "SELECT w." & typ & " AS Dept, ISNULL(SUM(Act_Sales - Act_Sales_Cost),0) / (ISNULL(SUM(Act_Inv_Cost),0) / 52) AS Turns " & _
                        "FROM Sales_Summary w " & _
                        "JOIN Stores s ON s.Str_Id = w.Str_Id " & _
                        "JOIN Inv_Summary i ON i.Loc_Id = s.Inv_Loc AND i.Dept = w.Dept AND i.Buyer = w.Buyer " & _
                        "AND i.Class = w.Class AND i.eDate = w.eDate " & _
                        "WHERE w.Str_Id = '" & thisStore & "' AND w.sDate >= '" & thisSDate & "' AND w.eDate <= '" & thisEDate & "' " & _
                        "AND Act_Inv_Cost > 0 GROUP BY w." & typ & " ORDER BY Turns DESC"
                    '' MsgBox("you need this code")
                Else
                    sql = "IF OBJECT_ID('tempdb.dbo.#t1', 'U') IS NOT NULL DROP TABLE #t1 " & _
                        "IF OBJECT_ID('tempdb.dbo.#t2', 'U') IS NOT NULL DROP TABLE #t2 " & _
                        "SELECT m." & typ & " AS Dept,i.Sku, ISNULL(SUM(Max_OH * Curr_Cost),0) AS invcost, CONVERT(Decimal,0) AS COGS, " & _
                        "COUNT(*) AS wksoh, CONVERT(Decimal,0) AS turns INTO #t1 FROM Item_Inv i " & _
                        "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                        "JOIN Item_Master m ON m.Sku = i.Sku " & _
                        "WHERE s.Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                        "GROUP BY Str_Id, i.Sku, m." & typ & " " & _
                        "DELETE FROM #t1 WHERE invcost = 0 " & _
                        "SELECT m." & typ & " AS Dept, d.Sku, CONVERT(Decimal,0) AS invcost, ISNULL(SUM(Sales_Cost),0) AS COGS, " & _
                        "CONVERT(Decimal,0) AS wksoh INTO #t2 FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                        "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                        "GROUP BY Str_Id, d.Sku, m." & typ & " " & _
                        "UPDATE #t1 SET #t1.COGS = #t2.COGS FROM #t1 " & _
                        "JOIN #t2 ON #t2.Dept = #t1.Dept AND #t2.Sku = #t1.Sku " & _
                        "UPDATE #t1 SET turns = CASE WHEN invcost > 0 AND wksoh > 0 THEN (COGS/(invcost/wksoh))*52/wksoh ELSE 0 END " & _
                        "SELECT Dept, AVG(turns) AS turns FROM #t1 GROUP BY Dept ORDER BY turns DESC"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                End If
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    turns = rdr("Turns")
                    If turns > 0 Then
                        row = btbl.NewRow
                        row("dept") = rdr("Dept")
                        row("turns") = rdr("Turns")
                        btbl.Rows.Add(row)
                        cnt += 1
                    End If
                End While
                con.Close()

                records = btbl.Rows.Count                       ' Get out if there aren't any records
                If records = 0 Then Exit Sub

                Dim fcnt As Integer = cnt / 3
                Dim mcnt As Integer = fcnt * 2
                Dim fast, medium As Decimal
                cnt = 0
                For Each row As DataRow In btbl.Rows
                    cnt += 1
                    If cnt = fcnt Then fast = row("turns")
                    If cnt = mcnt Then medium = row("turns")
                Next
                con.Open()
                sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' " & _
                    "AND Type = '" & dbTyp & "' AND Code = '" & dbTyp & "' AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "') " & _
                    "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
                    "SELECT 'Turns','" & thisStore & "','" & dbTyp & "','" & dbTyp & "','" & thisSDate & "','" & thisEDate & "'," & _
                    fast & "," & medium & "," & 0.01 & ",'" & now & "' " & _
                    "ELSE " & _
                    "UPDATE Score_Params SET Value1 = " & fast & ", Value2 = " & medium & ", Last_Update = '" & now & "' " & _
                    "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = '" & dbTyp & "' AND Code = '" & dbTyp & "' " & _
                    "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
            End If
            ''Next
            Dim m As String = "Get Turns " & typ
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Get_UnitCost_Params(thisStore, typ)
        Try
            Dim records As Integer
            Dim dbTyp As String
            Select Case typ
                Case "Sku"
                    dbTyp = "Item"
                Case "Dept"
                    dbTyp = "Dept"
                Case "Class"
                    dbTyp = "Class"
                Case "Buyer"
                    dbTyp = "Buyer"
                Case "Vendor_Id"
                    dbTyp = "Vendor_Id"
                Case "PLine"
                    dbTyp = "Pline"
            End Select
            txtProgress.Text = "Computing Unit Cost parameters"
            Me.Refresh()
            ''For Each drow As DataRow In dateTbl.Rows
            Dim thisSDate, thisEDate As Date
            Dim weeksBack As Integer = weeks
            thisEDate = endDate
            scoreTbl3 = New DataTable
            scoreTbl3.Columns.Add("Dept", GetType(System.String))
            Dim PrimaryKey3(1) As DataColumn
            PrimaryKey3(0) = scoreTbl3.Columns("Dept")
            scoreTbl3.PrimaryKey = PrimaryKey3
            scoreTbl3.Columns.Add("H", GetType(System.Decimal))
            scoreTbl3.Columns.Add("M", GetType(System.Decimal))
            scoreTbl3.Columns.Add("L", GetType(System.Decimal))
            scoreTbl3.Columns.Add("records", GetType(System.Int32))

            con.Open()
            sql = "Select eDate, CONVERT(varchar(10),DATEADD(day," & (weeksBack * -7) + 1 & ",eDate),101) AS sDate FROM Calendar " & _
                "WHERE eDate = '" & thisEDate & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSDate = rdr("sDate")
                thisEDate = rdr("eDate")
            End While
            con.Close()

            If typ = "Sku" Then
                con.Open()
                sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    row3 = scoreTbl3.NewRow
                    row3(0) = rdr("Dept")
                    scoreTbl3.Rows.Add(row3)
                End While
                con.Close()

                ctbl = New DataTable
                ctbl.Columns.Add("Dept", GetType(System.String))
                ctbl.Columns.Add("Sku", GetType(System.String))
                ctbl.Columns.Add("avgCost", GetType(System.Decimal))
                ctbl.Columns.Add("score", GetType(System.String))
                Dim dept, prevDept As String
                Dim cnt As Integer
                Dim val, cost, div, sold As Decimal
                Dim foundRow As DataRow
                prevDept = ""
                con.Open()
                sql = "SELECT Dept, d.Sku, ISNULL(SUM(Sold),0) AS sold, ISNULL(SUM(Sales_Cost),0) AS cost " & _
                    "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                    "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                    "GROUP BY Dept, d.Sku ORDER BY Dept"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                rdr = cmd.ExecuteReader
                While rdr.Read
                    dept = rdr("Dept")
                    If dept <> prevDept Then
                        If prevDept = "" Then
                            prevDept = dept
                        Else
                            foundRow = scoreTbl3.Rows.Find(prevDept)
                            foundRow(4) = cnt
                            cnt = 0
                            prevDept = dept
                        End If
                    End If
                    sold = rdr("sold")
                    cost = rdr("cost")
                    val = 0
                    If Not IsDBNull(sold) And Not IsNothing(sold) And Not IsDBNull(cost) And Not IsNothing(cost) Then
                        If IsNumeric(sold) And IsNumeric(cost) Then
                            If sold > 0 Then val = cost / sold
                        End If
                    End If
                    If val > 0 Then cnt += 1
                    row = ctbl.NewRow
                    row("Dept") = dept
                    row("avgCost") = val
                    ctbl.Rows.Add(row)
                End While
                con.Close()

                records = ctbl.Rows.Count                ' Get out if there are no records to process
                If records = 0 Then Exit Sub

                Dim hCnt, mCnt As Integer
                Dim view As New DataView(ctbl)
                view.Sort = "Dept ASC, avgCost DESC"                       ' was Dept ASC, avgCosr DESC

                cnt = 0
                prevDept = ""
                For Each rw As DataRowView In view
                    cnt += 1
                    dept = rw("Dept")
                    val = rw("avgCost")
                    If cnt = hCnt Then
                        foundRow = scoreTbl3.Rows.Find(dept)
                        foundRow(1) = val
                    End If
                    If cnt = mCnt Then
                        foundRow = scoreTbl3.Rows.Find(dept)
                        foundRow(2) = val
                        foundRow(3) = 0.01
                    End If
                    If dept <> prevDept Then
                        foundRow = scoreTbl3.Rows.Find(dept)
                        oTest = foundRow(4)
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then
                                hCnt = foundRow(4) / 3
                                mCnt = hCnt * 2
                                cnt = 1
                            End If
                        End If
                        prevDept = dept
                    End If
                Next

                Dim id, type, code, score As String
                Dim amt, pct, pct2, pct3 As Decimal
                Dim now As DateTime = Date.Now

                con.Open()
                For Each row As DataRow In scoreTbl3.Rows
                    id = "UnitCost"
                    code = row(0)
                    oTest = row(1)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = row(1) Else pct = 0
                    oTest = row(2)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct2 = row(2) Else pct2 = 0
                    oTest = row(3)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct3 = row(3) Else pct3 = 0

                    sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
                        "AND eDate = '" & thisEDate & "' AND Type = 'Item' AND Code = '" & code & "') " & _
                        "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
                        "SELECT '" & id & "','" & thisStore & "','" & thisSDate & "','" & thisEDate & "','Item','" & _
                        code & "'," & pct & "," & pct2 & "," & pct3 & ",'" & now & "' " & _
                        "ELSE " & _
                        "UPDATE Score_Params SET Value1 = " & pct & ", Value2 = " & pct2 & ", Value3 = " & pct3 & ", Last_Update = '" & now & "' " & _
                        "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = '" & code & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                Next
                con.Close()
            Else
                ctbl = New DataTable
                ctbl.Columns.Add("cost", GetType(System.String))
                Dim cnt As Integer = 0
                con.Open()
                sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Cost),0) / ISNULL(SUM(Sold),0) AS cost " & _
                    "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                    "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & thisSDate & "' AND eDate <= '" & thisEDate & "' " & _
                    "AND Sold > 0 GROUP BY m." & typ & " ORDER BY cost DESC"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    row = ctbl.NewRow
                    oTest = rdr("cost")
                    row("cost") = rdr("cost")
                    ctbl.Rows.Add(row)
                    cnt += 1
                End While
                con.Close()

                records = ctbl.Rows.Count                      ' Get out if there aren't any records
                If records = 0 Then Exit Sub

                Dim hcnt As Integer = cnt / 3
                Dim mcnt As Integer = hcnt * 2
                Dim high, medium As Decimal
                cnt = 0
                For Each row As DataRow In ctbl.Rows
                    cnt += 1
                    If cnt = hcnt Then high = row("cost")
                    If cnt = mcnt Then medium = row("cost")
                Next
                con.Open()
                sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' " & _
                    "AND Type = '" & dbTyp & "' AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "') " & _
                    "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
                    "SELECT 'UnitCost','" & thisStore & "','" & dbTyp & "','" & dbTyp & "','" & thisSDate & "','" & thisEDate & "'," & _
                    high & "," & medium & "," & 0.01 & ",'" & Now & "' " & _
                    "ELSE " & _
                    "UPDATE Score_Params SET Value1 = " & high & ", Value2 = " & medium & ", Last_Update = '" & Now & "' " & _
                    "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = '" & dbTyp & "' " & _
                    "AND sDate = '" & thisSDate & "' AND eDate = '" & thisEDate & "'"
                cmd = New SqlCommand(sql, con)
                cmd.CommandTimeout = 120
                cmd.ExecuteNonQuery()
                con.Close()
            End If
            ''Next
            Dim m As String = "Get UnitCost " & typ
        Catch ex As Exception
            Dim theMessage As String = ex.Message
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    ''Private Sub Get_GrossProfit_Params(ByVal thisStore As String, ByVal typ As String)                                     ' Compute paramaters for Gross Margin

    ''    Dim tgm As Integer = 0
    ''    Dim summgn As Integer = 0
    ''    Dim tcnt As Integer = 0
    ''    Dim sumcnt As Integer = 0
    ''    Dim val As Decimal
    ''    tbl = New DataTable
    ''    tbl.Columns.Add(typ, GetType(System.String))
    ''    tbl.Columns.Add("gm", GetType(System.Int16))
    ''    tbl.Columns.Add("pctmgn", GetType(System.Decimal))
    ''    tbl.Columns.Add("pctsku", GetType(System.Decimal))
    ''    scoreTbl = New DataTable
    ''    Dim column = New DataColumn()
    ''    column.DataType = System.Type.GetType("System.String")
    ''    column.ColumnName = "Code"
    ''    scoreTbl.Columns.Add(column)
    ''    Dim PrimaryKey(1) As DataColumn
    ''    PrimaryKey(0) = scoreTbl.Columns("Code")
    ''    scoreTbl.PrimaryKey = PrimaryKey
    ''    scoreTbl.Columns.Add("GM%", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("Break$", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("SKU%", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("CumGM%", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("Skus", GetType(System.Int32))

    ''    Dim arr() As String = {"A", "B", "C", "D", "E"}
    ''    Dim row As DataRow
    ''    For Each element As String In arr
    ''        row = scoreTbl.NewRow
    ''        row(0) = element
    ''        row(1) = 0
    ''        scoreTbl.Rows.Add(row)
    ''    Next
    ''    Dim tpct As Decimal = 0
    ''    For i As Integer = 0 To 4
    ''        tpct += pctArray(i)
    ''        scoreTbl.Rows(i).Item("GM%") = pctArray(i)
    ''        scoreTbl.Rows(i).Item("CumGM%") = tpct
    ''    Next

    ''    aTbl = New DataTable
    ''    aTbl.Columns.Add(typ, GetType(System.String))
    ''    aTbl.Columns.Add("GrossProfit", GetType(System.Decimal))
    ''    aTbl.Columns.Add("Pct", GetType(System.Decimal))
    ''    aTbl.Columns.Add("iPct", GetType(System.Decimal))
    ''    con.Open()
    ''    sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Retail - Sales_Cost),0) AS GrossProfit " & _
    ''            "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
    ''            "WHERE Str_id = '" & thisStore & "' AND sDate >= '" & startDate & "' " & _
    ''            "AND eDate <= '" & endDate & "' AND Sales_Retail > 0 " & _
    ''            "GROUP BY m." & typ & " ORDER BY GrossProfit DESC"
    ''    cmd = New SqlCommand(sql, con)
    ''    cmd.CommandTimeout = 120
    ''    rdr = cmd.ExecuteReader
    ''    While rdr.Read
    ''        row = aTbl.NewRow
    ''        oTest = rdr("typ")
    ''        row(typ) = rdr("typ")
    ''        val = rdr("GrossProfit")
    ''        If val > 0 Then
    ''            tcnt += 1
    ''            summgn += val
    ''        End If
    ''        row("GrossProfit") = val
    ''        aTbl.Rows.Add(row)
    ''    End While
    ''    con.Close()

    ''    If tcnt > 0 Then
    ''        Dim cnt As Integer = 0
    ''        Dim pct As Decimal
    ''        Dim view As New DataView(aTbl)
    ''        view.Sort = "GrossProfit DESC"
    ''        For Each rw As DataRowView In view
    ''            cnt += 1
    ''            ''If cnt Mod 1000 Then Console.WriteLine(cnt & " " & thisSDate & " " & thisEDate)
    ''            oTest = rw(0)
    ''            val = rw("grossprofit")
    ''            tgm += val
    ''            pct = tgm / summgn
    ''            rw("Pct") = pct
    ''            rw("iPct") = cnt / tcnt
    ''            For j As Integer = 0 To 4
    ''                Dim cummGP As Decimal = 0
    ''                Dim breakamt As Decimal = 0
    ''                oTest = scoreTbl(j).Item("CumGM%")
    ''                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then cummGP = CDec(oTest)
    ''                oTest = scoreTbl.Rows(j).Item("Break$")
    ''                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then breakamt = CDec(oTest)
    ''                Dim it As Decimal = scoreTbl(j).Item("CumGM%")
    ''                If pct >= cummGP And breakamt = 0 Then
    ''                    scoreTbl.Rows(j).Item("Break$") = val
    ''                End If
    ''            Next
    ''        Next
    ''        Dim Acnt, bCnt, cCnt, Dcnt, Ecnt, gp As Decimal
    ''        For Each arow As DataRow In aTbl.Rows
    ''            gp = arow("grossprofit")
    ''            If gp >= scoreTbl.Rows(0).Item("Break$") Then Acnt += 1
    ''            If gp >= scoreTbl.Rows(1).Item("Break$") And gp < scoreTbl.Rows(0).Item("Break$") Then bCnt += 1
    ''            If gp >= scoreTbl.Rows(2).Item("Break$") And gp < scoreTbl.Rows(1).Item("Break$") Then cCnt += 1
    ''            If gp >= scoreTbl.Rows(3).Item("Break$") And gp < scoreTbl.Rows(2).Item("Break$") Then Dcnt += 1
    ''            If gp >= scoreTbl.Rows(4).Item("Break$") And gp < scoreTbl.Rows(3).Item("Break$") Then Ecnt += 1
    ''        Next


    ''        scoreTbl.Rows(0).Item("SKU%") = CDec(Acnt / tcnt)
    ''        scoreTbl.Rows(1).Item("SKU%") = CDec(bCnt / tcnt)
    ''        scoreTbl.Rows(2).Item("SKU%") = CDec(cCnt / tcnt)
    ''        scoreTbl.Rows(3).Item("SKU%") = CDec(Dcnt / tcnt)
    ''        scoreTbl.Rows(4).Item("SKU%") = CDec(Ecnt / tcnt)
    ''        scoreTbl.Rows(0).Item("Skus") = Acnt
    ''        scoreTbl.Rows(1).Item("Skus") = bCnt
    ''        scoreTbl.Rows(2).Item("Skus") = cCnt
    ''        scoreTbl.Rows(3).Item("Skus") = Dcnt
    ''        scoreTbl.Rows(4).Item("Skus") = Ecnt
    ''        Dim id, code As String
    ''        Dim amt, ipct As Decimal
    ''        Dim now As DateTime = Date.Now
    ''        For Each srow As DataRow In scoreTbl.Rows
    ''            amt += srow(2)
    ''        Next
    ''        con.Open()
    ''        For Each srow As DataRow In scoreTbl.Rows
    ''            id = "GrossProfit"
    ''            ''type = typ
    ''            code = srow(0)
    ''            oTest = srow(1)
    ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = srow(1) Else amt = 0
    ''            oTest = srow(2)
    ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = srow(2) Else pct = 0
    ''            oTest = srow(3)
    ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ipct = srow(3) Else ipct = 0
    ''            Dim dbTyp As String = ""
    ''            Select Case typ
    ''                Case "Sku"
    ''                    dbTyp = "Item"
    ''                Case "Dept"
    ''                    dbTyp = "Dept"
    ''                Case "Class"
    ''                    dbTyp = "Class"
    ''                Case "Buyer"
    ''                    dbTyp = "Buyer"
    ''                Case "Vendor_Id"
    ''                    dbTyp = "Vendor"
    ''                Case "PLine"
    ''                    dbTyp = "PLine"
    ''            End Select
    ''            sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' " & _
    ''                "AND sDate = '" & startDate & "' AND eDate = '" & endDate & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "') " & _
    ''                "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
    ''                "SELECT '" & id & "','" & thisStore & "','" & startDate & "','" & endDate & "','" & dbTyp & "','" & _
    ''                code & "'," & amt & "," & pct & "," & ipct & ",'" & now & "' " & _
    ''                "ELSE " & _
    ''                "UPDATE Score_Params SET Value1 = " & amt & ", Value2 = " & pct & ", Value3 = " & ipct & ", Last_Update = '" & now & "' " & _
    ''                "WHERE ID = '" & id & "' AND Str_Id = '" & thisStore & "' AND Type = '" & dbTyp & "' AND Code = '" & code & "' " & _
    ''                "AND sDate = '" & startDate & "' AND eDate = '" & endDate & "'"
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.ExecuteNonQuery()
    ''        Next
    ''        con.Close()
    ''    End If
    ''End Sub

    ''Private Sub Get_Turns_Params(ByVal thisStore As String, ByVal typ As String)
    ''    txtProgress.Text = "Computing Turns at Cost"
    ''    Me.Refresh()
    ''    scoreTbl = New DataTable
    ''    scoreTbl.Columns.Add("Dept", GetType(System.String))
    ''    Dim PrimaryKey2(1) As DataColumn
    ''    PrimaryKey2(0) = scoreTbl.Columns("Dept")
    ''    scoreTbl.PrimaryKey = PrimaryKey2
    ''    scoreTbl.Columns.Add("F", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("M", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("S", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("records", GetType(System.Int32))
    ''    Dim row As DataRow
    ''    Dim now As DateTime = Date.Now
    ''    If typ = "Sku" Then
    ''        row = scoreTbl.NewRow
    ''        row(0) = "ALL"
    ''        scoreTbl.Rows.Add(row)

    ''        tbl = New DataTable
    ''        tbl.Columns.Add("cnt", GetType(System.Int32))
    ''        tbl.Columns.Add("Sku", GetType(System.String))
    ''        tbl.Columns.Add("invcost", GetType(System.Decimal))
    ''        tbl.Columns.Add("wksoh", GetType(System.Int32))
    ''        tbl.Columns.Add("COGS", GetType(System.Decimal))
    ''        tbl.Columns.Add("turns", GetType(System.Decimal))
    ''        Dim cnt, fCnt, mCnt As Integer
    ''        Dim val As Decimal
    ''        Dim prevDept As String = ""
    ''        Dim foundRow As DataRow
    ''        con.Open()
    ''        sql = "SELECT d.Sku, ISNULL(SUM(End_OH * Cost),0) AS invcost, ISNULL(SUM(Sales_Cost),0) AS COGS, " & _
    ''            "dbo.fnTurns(d.Str_Id, d.Sku, '" & startDate & "','" & endDate & "') AS Turns " & _
    ''            "FROM Item_Sales d JOIN Item_Inv v ON v.Str_Id = d.Str_Id AND v.Sku = d.Sku " & _
    ''            "WHERE d.Str_Id = '" & thisStore & "' AND d.sDate >= '" & startDate & "' AND d.eDate <= '" & endDate & "' " & _
    ''            "GROUP BY d.Str_Id, d.Sku"
    ''        cmd = New SqlCommand(sql, con)
    ''        cmd.CommandTimeout = 120
    ''        rdr = cmd.ExecuteReader
    ''        While rdr.Read
    ''            val = rdr("Turns")
    ''            If val > 0 Then
    ''                cnt += 1
    ''                row = tbl.NewRow
    ''                row("cnt") = cnt
    ''                row("Sku") = rdr("Sku")
    ''                row("invcost") = rdr("invcost")
    ''                'row("wksoh") = rdr("wksOH")
    ''                row("COGS") = rdr("COGS")
    ''                row("turns") = val
    ''                tbl.Rows.Add(row)
    ''            End If
    ''        End While
    ''        con.Close()

    ''        If cnt > 0 Then
    ''            foundRow = scoreTbl.Rows.Find("ALL")
    ''            foundRow("records") = cnt
    ''            fCnt = cnt / 3
    ''            mCnt = fCnt * 2
    ''            Dim fast, medium As Decimal
    ''            cnt = 0
    ''            Dim view As New DataView(tbl)
    ''            view.Sort = "turns DESC"
    ''            For Each rw As DataRowView In view
    ''                cnt += 1
    ''                If cnt Mod 1000 = 0 Then
    ''                    txtProgress.Text = "Computing Item Turns at Cost " & cnt
    ''                    Me.Refresh()
    ''                End If
    ''                If cnt = fCnt Then fast = rw("turns")
    ''                If cnt = mCnt Then medium = rw("turns")
    ''            Next

    ''            con.Open()
    ''            For Each srow As DataRow In scoreTbl.Rows
    ''                sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = 'Turns' AND Str_Id = '" & thisStore & "' " & _
    ''                    "AND sDate = '" & startDate & "' AND Type = 'Item' AND Code = 'Item') " & _
    ''                    "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
    ''                    "SELECT 'Turns','" & thisStore & "','" & startDate & "','" & endDate & "','Item','Item'," & _
    ''                     fast & "," & medium & ",.01,'" & now & "' " & _
    ''                    "ELSE " & _
    ''                    "UPDATE Score_Params SET Value1 = " & fast & ", Value2 = " & medium & ", Value3 = .01 , Last_Update = '" & now & "' " & _
    ''                    "WHERE ID = 'Turns' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = 'Item' AND sDate = '" & startDate & "'"
    ''                cmd = New SqlCommand(sql, con)
    ''                cmd.ExecuteNonQuery()
    ''            Next
    ''            con.Close()
    ''        End If
    ''    Else
    ''        Dim cnt As Integer = 0
    ''        tbl = New DataTable
    ''        tbl.Columns.Add("dept", GetType(System.String))
    ''        tbl.Columns.Add("turns", GetType(System.String))
    ''        con.Open()
    ''        If typ = "Buyer" Or typ = "Class" Or typ = "Dept" Then
    ''            sql = "SELECT Dept, ISNULL(SUM(Act_Sales - Act_Sales_Cost),0) / (ISNULL(SUM(Act_Inv_Cost),0) / 52) AS Turns " & _
    ''                "FROM Item_Sales WHERE sDate >= '" & startDate & "' AND eDate <= '" & endDate & "' " & _
    ''                "GROUP BY Dept ORDER BY Turns DESC"
    ''        Else
    ''            sql = "CREATE TABLE #t1 (dept varchar(20) NOT NULL, Sku varchar(30), invcost decimal(18,4), " & _
    ''                "COGS decimal(18,4), wksoh decimal(18,4)) " & _
    ''                "INSERT INTO #t1 (dept, Sku, invcost, COGS, wksoh) " & _
    ''                "SELECT m." & typ & ", d.Sku,  ISNULL(SUM(Inv_Cost),0) AS invcost, ISNULL(SUM(Sales_Cost),0) AS COGS, " & _
    ''           "dbo.fnWksOH(d.Str_Id, d.Sku, '" & startDate & "','" & endDate & "') AS wksOH " & _
    ''           "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
    ''           "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & startDate & "' AND eDate <= '" & endDate & "' " & _
    ''           "GROUP BY d.Str_Id, d.Sku, m." & typ & " "
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.CommandTimeout = 120
    ''            cmd.ExecuteNonQuery()
    ''            sql = "CREATE TABLE #t2 (Dept varchar(20), Sku varchar(30), invcost decimal(18,4), COGS decimal(18,4), " & _
    ''                "wksoh decimal(18,4), turns decimal(18,4)) " & _
    ''                "INSERT INTO #t2 (Dept, Sku, invcost, COGS, wksoh, turns) " & _
    ''                "SELECT Dept, Sku, invcost, cogs, wksoh, CASE WHEN invcost=0 or wksoh=0 THEN 0 " & _
    ''                "ELSE (COGS/(invcost/wksoh))*(52/wksoh) END AS turns FROM #t1"
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.CommandTimeout = 120
    ''            cmd.ExecuteNonQuery()
    ''            sql = "SELECT Dept, AVG(turns) AS turns FROM #t2 GROUP BY Dept ORDER BY turns DESC"
    ''        End If
    ''        cmd = New SqlCommand(sql, con)
    ''        rdr = cmd.ExecuteReader
    ''        While rdr.Read
    ''            row = tbl.NewRow
    ''            row("dept") = rdr("Dept")
    ''            row("turns") = rdr("Turns")
    ''            tbl.Rows.Add(row)
    ''            cnt += 1
    ''        End While
    ''        con.Close()
    ''        If cnt > 0 Then
    ''            Dim fcnt As Integer = cnt / 3
    ''            Dim mcnt As Integer = fcnt * 2
    ''            Dim fast, medium As Decimal
    ''            cnt = 0
    ''            For Each trow As DataRow In tbl.Rows
    ''                cnt += 1
    ''                If cnt = fcnt Then fast = trow("turns")
    ''                If cnt = mcnt Then medium = trow("turns")
    ''            Next
    ''            con.Open()
    ''            If typ = "Vendor_Id" Then typ = "Vendor"
    ''            sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' " & _
    ''            "AND Type = '" & typ & "' AND Code = '" & typ & "' AND sDate = '" & startDate & "' AND eDate = '" & endDate & "') " & _
    ''            "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
    ''            "SELECT 'Turns','" & thisStore & "','" & typ & "','" & typ & "','" & startDate & "','" & endDate & "'," & _
    ''            fast & "," & medium & "," & 0.01 & ",'" & now & "' " & _
    ''            "ELSE " & _
    ''            "UPDATE Score_Params SET Value1 = " & fast & ", Value2 = " & medium & ", Last_Update = '" & now & "' " & _
    ''            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = '" & typ & "' AND Code = '" & typ & "' " & _
    ''            "AND sDate = '" & startDate & "' AND eDate = '" & endDate & "'"
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.ExecuteNonQuery()
    ''            con.Close()
    ''        End If
    ''    End If
    ''End Sub

    ''Private Sub Get_UnitCost_Params(ByVal thisStore As String, ByVal typ As String)

    ''    scoreTbl = New DataTable
    ''    scoreTbl.Columns.Add("Dept", GetType(System.String))
    ''    Dim PrimaryKey3(1) As DataColumn
    ''    PrimaryKey3(0) = scoreTbl.Columns("Dept")
    ''    scoreTbl.PrimaryKey = PrimaryKey3
    ''    scoreTbl.Columns.Add("H", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("M", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("L", GetType(System.Decimal))
    ''    scoreTbl.Columns.Add("records", GetType(System.Int32))

    ''    con.Open()
    ''    sql = "Select eDate, CONVERT(varchar(10),DATEADD(day,-181,eDate),101) AS sDate FROM Calendar " & _
    ''        "WHERE eDate = '" & endDate & "'"
    ''    cmd = New SqlCommand(sql, con)
    ''    rdr = cmd.ExecuteReader
    ''    While rdr.Read
    ''        startDate = rdr("sDate")
    ''        endDate = rdr("eDate")
    ''    End While
    ''    con.Close()

    ''    Dim row As DataRow
    ''    If typ = "Sku" Then
    ''        row = scoreTbl.NewRow
    ''        row(0) = "ALL"
    ''        scoreTbl.Rows.Add(row)
    ''        tbl = New DataTable
    ''        tbl.Columns.Add("Sku", GetType(System.String))
    ''        tbl.Columns.Add("avgCost", GetType(System.Decimal))
    ''        tbl.Columns.Add("score", GetType(System.String))
    ''        Dim prevDept As String
    ''        Dim cnt As Integer
    ''        Dim val, cost, sold As Decimal
    ''        prevDept = ""
    ''        con.Open()
    ''        sql = "SELECT Sku, ISNULL(SUM(Sold),0) AS sold, ISNULL(SUM(Sales_Cost),0) AS cost " & _
    ''            "FROM Item_Sales WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & startDate & "' AND eDate <= '" & endDate & "' " & _
    ''            "GROUP BY Sku"
    ''        cmd = New SqlCommand(sql, con)
    ''        cmd.CommandTimeout = 120
    ''        rdr = cmd.ExecuteReader
    ''        While rdr.Read
    ''            sold = rdr("sold")
    ''            cost = rdr("cost")
    ''            val = 0
    ''            If Not IsDBNull(sold) And Not IsNothing(sold) And Not IsDBNull(cost) And Not IsNothing(cost) Then
    ''                If IsNumeric(sold) And IsNumeric(cost) Then
    ''                    If sold > 0 Then val = cost / sold
    ''                End If
    ''            End If
    ''            If val > 0 Then cnt += 1
    ''            row = tbl.NewRow
    ''            'row("Dept") = dept
    ''            row("avgCost") = val
    ''            tbl.Rows.Add(row)
    ''        End While
    ''        con.Close()

    ''        Dim hCnt As Integer = cnt / 3
    ''        Dim mCnt As Integer = hCnt * 2
    ''        Dim high, medium As Decimal
    ''        Dim view As New DataView(tbl)
    ''        view.Sort = "avgCost DESC"

    ''        cnt = 0
    ''        prevDept = ""
    ''        For Each rw As DataRowView In view
    ''            cnt += 1
    ''            If cnt = hCnt Then high = rw("avgCost")
    ''            If cnt = mCnt Then medium = rw("avgCost")
    ''        Next
    ''        If high + medium > 0 Then
    ''            Dim now As DateTime = Date.Now
    ''            con.Open()
    ''            sql = "IF Not EXISTS (SELECT ID FROM Score_Params WHERE ID = 'UnitCost' AND Str_Id = '" & thisStore & "' " & _
    ''                    "AND eDate = '" & endDate & "' AND Type = 'Item' AND Code = 'Item') " & _
    ''                    "INSERT INTO Score_Params (ID, Str_Id, sDate, eDate, Type, Code, Value1, Value2, Value3, Last_Update) " & _
    ''                    "SELECT 'UnitCost','" & thisStore & "','" & startDate & "','" & endDate & "','Item','Item'," & _
    ''                    high & "," & medium & ",.01,'" & now & "' " & _
    ''                    "ELSE " & _
    ''                    "UPDATE Score_Params SET Value1 = " & high & ", Value2 = " & medium & ", Value3 = .01, Last_Update = '" & now & "' " & _
    ''                    "WHERE ID = 'UnitCost' AND Str_Id = '" & thisStore & "' AND Type = 'Item' AND Code = 'Item' " & _
    ''                    "AND sDate = '" & startDate & "'"
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.CommandTimeout = 120
    ''            cmd.ExecuteNonQuery()
    ''            con.Close()
    ''        End If
    ''    Else
    ''        tbl = New DataTable
    ''        tbl.Columns.Add("cost", GetType(System.String))
    ''        Dim cnt As Integer = 0
    ''        con.Open()
    ''        sql = "SELECT m." & typ & " AS typ, ISNULL(SUM(Sales_Cost),0) / ISNULL(SUM(Sold),0) AS cost " & _
    ''            "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
    ''            "WHERE Str_Id = '" & thisStore & "' AND sDate >= '" & startDate & "' AND eDate <= '" & endDate & "' " & _
    ''            "AND Sold > 0 GROUP BY m." & typ & " ORDER BY cost DESC"
    ''        cmd = New SqlCommand(sql, con)
    ''        rdr = cmd.ExecuteReader
    ''        While rdr.Read
    ''            row = tbl.NewRow
    ''            oTest = rdr("cost")
    ''            row("cost") = rdr("cost")
    ''            tbl.Rows.Add(row)
    ''            cnt += 1
    ''        End While
    ''        con.Close()
    ''        Dim hcnt As Integer = cnt / 3
    ''        Dim mcnt As Integer = hcnt * 2
    ''        Dim high, medium As Decimal
    ''        cnt = 0
    ''        For Each trow As DataRow In tbl.Rows
    ''            cnt += 1
    ''            If cnt = hcnt Then high = trow("cost")
    ''            If cnt = mcnt Then medium = trow("cost")
    ''        Next
    ''        If high + medium > 0 Then
    ''            con.Open()
    ''            If typ = "Vendor_Id" Then typ = "Vendor"
    ''            sql = "IF NOT EXISTS (SELECT ID FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' " & _
    ''                "AND Type = '" & typ & "' AND sDate = '" & startDate & "' AND eDate = '" & endDate & "') " & _
    ''                "INSERT INTO Score_Params (ID, Str_Id, Type, Code, sDate, eDate, Value1, Value2, Value3, Last_Update) " & _
    ''                "SELECT 'UnitCost','" & thisStore & "','" & typ & "','" & typ & "','" & startDate & "','" & endDate & "'," & _
    ''                high & "," & medium & "," & 0.01 & ",'" & Now & "' " & _
    ''                "ELSE " & _
    ''                "UPDATE Score_Params SET Value1 = " & high & ", Value2 = " & medium & ", Last_Update = '" & Now & "' " & _
    ''                "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = '" & typ & "' " & _
    ''                "AND sDate = '" & startDate & "' AND eDate = '" & endDate & "'"
    ''            cmd = New SqlCommand(sql, con)
    ''            cmd.CommandTimeout = 120
    ''            cmd.ExecuteNonQuery()
    ''            con.Close()
    ''        End If
    ''    End If

    ''End Sub

    Private Sub cboDeptStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDeptStore.SelectedIndexChanged
        thisStore = cboDeptStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvDept1.View = View.Details
        lvDept1.GridLines = True
        lvDept1.Columns.Clear()
        lvDept1.Items.Clear()

        'lvDept1.Width = 300

        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvDept1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvDept1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvDept1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvDept1.Columns.Add(colx)

        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Dept' AND eDate = " & _
            "(SELECT MAX(eDate) FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Dept')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvDept1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next

        For Each Column As ColumnHeader In lvDept1.Columns
            Column.Width = -2
        Next

        lvDept2.View = View.Details
        lvDept2.GridLines = True
        lvDept2.Columns.Clear()
        lvDept2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvDept2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvDept2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvDept2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvDept2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow

        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type = 'Dept' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = 'Dept')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            lst2 = lvDept2.Items.Add("Turns at Cost")
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvDept2.Columns
            Column.Width = -2
        Next

        lvDept3.View = View.Details
        lvDept3.GridLines = True
        lvDept3.Columns.Clear()
        lvDept3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvDept3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvDept3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvDept3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvDept3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type = 'Dept' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = 'Dept')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            lst3 = lvDept3.Items.Add("Unit Cost")
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvDept3.Columns
            Column.Width = -2
        Next
    End Sub

    Private Sub cboBuyerStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuyerStore.SelectedIndexChanged
        thisStore = cboBuyerStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvBuyer1.View = View.Details
        lvBuyer1.GridLines = True
        lvBuyer1.Columns.Clear()
        lvBuyer1.Items.Clear()
        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvBuyer1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvBuyer1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvBuyer1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvBuyer1.Columns.Add(colx)

        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Buyer' AND eDate = " & _
            "(SELECT MAX(eDate) FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Buyer')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvBuyer1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next

        For Each Column As ColumnHeader In lvBuyer1.Columns
            Column.Width = -2
        Next

        lvBuyer2.View = View.Details
        lvBuyer2.GridLines = True
        lvBuyer2.Columns.Clear()
        lvBuyer2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvBuyer2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvBuyer2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvBuyer2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvBuyer2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow

        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type = 'Buyer' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = 'Buyer')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            lst2 = lvBuyer2.Items.Add("Turns at Cost")
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvBuyer2.Columns
            Column.Width = -2
        Next

        lvBuyer3.View = View.Details
        lvBuyer3.GridLines = True
        lvBuyer3.Columns.Clear()
        lvBuyer3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvBuyer3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvBuyer3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvBuyer3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvBuyer3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type = 'Buyer' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = 'Buyer')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            lst3 = lvBuyer3.Items.Add("Unit Cost")
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvBuyer3.Columns
            Column.Width = -2
        Next
    End Sub

    Private Sub cboClassStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClassStore.SelectedIndexChanged
        thisStore = cboClassStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvClass1.View = View.Details
        lvClass1.GridLines = True
        lvClass1.Columns.Clear()
        lvClass1.Items.Clear()
        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvClass1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvClass1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvClass1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvClass1.Columns.Add(colx)

        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Class' AND eDate = " & _
            "(SELECT MAX(eDate) FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Class')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvClass1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next

        For Each Column As ColumnHeader In lvClass1.Columns
            Column.Width = -2
        Next

        lvClass2.View = View.Details
        lvClass2.GridLines = True
        lvClass2.Columns.Clear()
        lvClass2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvClass2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvClass2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvClass2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvClass2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow

        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type = 'Class' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = 'Class')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            lst2 = lvClass2.Items.Add("Turns at Cost")
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvClass2.Columns
            Column.Width = -2
        Next

        lvClass3.View = View.Details
        lvClass3.GridLines = True
        lvClass3.Columns.Clear()
        lvClass3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvClass3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvClass3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvClass3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvClass3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type = 'Class' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = 'Class')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            lst3 = lvClass3.Items.Add("Unit Cost")
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
            oTest = row3(3)
        End While
        con.Close()

        For Each Column As ColumnHeader In lvClass3.Columns
            Column.Width = -2
        Next
    End Sub

    Private Sub cboVendorStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVendorStore.SelectedIndexChanged
        thisStore = cboVendorStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvVendor1.View = View.Details
        lvVendor1.GridLines = True
        lvVendor1.Columns.Clear()
        lvVendor1.Items.Clear()
        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvVendor1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvVendor1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvVendor1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvVendor1.Columns.Add(colx)

        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type LIKE 'Vendor%' AND eDate = " & _
            "(SELECT MAX(eDate) FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type LIKE 'Vendor%')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvVendor1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next


        For Each Column As ColumnHeader In lvVendor1.Columns
            Column.Width = -2
        Next

        lvVendor2.View = View.Details
        lvVendor2.GridLines = True
        lvVendor2.Columns.Clear()
        lvVendor2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvVendor2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvVendor2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvVendor2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvVendor2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow

        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type LIKE 'Vendor%' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type LIKE 'Vendor%')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            lst2 = lvVendor2.Items.Add("Turns at Cost")
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()


        For Each Column As ColumnHeader In lvVendor2.Columns
            Column.Width = -2
        Next
        lvVendor3.View = View.Details
        lvVendor3.GridLines = True
        lvVendor3.Columns.Clear()
        lvVendor3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvVendor3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvVendor3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvVendor3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvVendor3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type LIKE 'Vendor%' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type LIKE 'Vendor%')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            lst3 = lvVendor3.Items.Add("Unit Cost")
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvVendor3.Columns
            Column.Width = -2
        Next
    End Sub

    Private Sub cboPLineStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPLineStore.SelectedIndexChanged
        thisStore = cboPLineStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvPLine1.View = View.Details
        lvPLine1.GridLines = True
        lvPLine1.Columns.Clear()
        lvPLine1.Items.Clear()
        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvPLine1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvPLine1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvPLine1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvPLine1.Columns.Add(colx)

        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'PLine' AND eDate = " & _
            "(SELECT MAX(eDate) FROM Score_Params WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'PLine')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvPLine1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next

        For Each Column As ColumnHeader In lvPLine1.Columns
            Column.Width = -2
        Next

        lvPLine2.View = View.Details
        lvPLine2.GridLines = True
        lvPLine2.Columns.Clear()
        lvPLine2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvPLine2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvPLine2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvPLine2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvPLine2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow

        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type = 'PLine' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = 'PLine')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            lst2 = lvPLine2.Items.Add("Turns at Cost")
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()


        For Each Column As ColumnHeader In lvPLine2.Columns
            Column.Width = -2
        Next

        lvPLine3.View = View.Details
        lvPLine3.GridLines = True
        lvPLine3.Columns.Clear()
        lvPLine3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = ""
            .TextAlign = HorizontalAlignment.Center
            .Width = 100
        End With
        lvPLine3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvPLine3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvPLine3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvPLine3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type = 'PLine' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = 'PLine')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            lst3 = lvPLine3.Items.Add("Unit Cost")
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
        End While
        con.Close()

        For Each Column As ColumnHeader In lvPLine3.Columns
            Column.Width = -2
        Next
    End Sub

    Private Sub cboItemStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboItemStore.SelectedIndexChanged
        thisStore = cboItemStore.SelectedItem
        tbl = New DataTable
        tbl.Columns.Add("Code")
        tbl.Columns.Add("GM%")
        tbl.Columns.Add("Break$")
        tbl.Columns.Add("SKU%")

        lvItem1.View = View.Details
        lvItem1.GridLines = True
        lvItem1.Columns.Clear()
        lvItem1.Items.Clear()
        Dim colx As New ColumnHeader
        With colx
            .Text = "Code"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvItem1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "GM%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvItem1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Break$"
            .TextAlign = HorizontalAlignment.Right
            .Width = 100
        End With
        lvItem1.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "SKU%"
            .TextAlign = HorizontalAlignment.Right
            .Width = 50
        End With
        lvItem1.Columns.Add(colx)
        Dim row As DataRow

        con.Open()
        sql = "SELECT Code, Value2 AS GPPct, Value1 AS GPAmt, Value3 AS SKUPct FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Sku' " & _
            "AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'GrossMargin' AND Type = 'Sku')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Code") = rdr("Code")
            row("GM%") = Format(rdr("GPPct"), "p0")
            row("Break$") = Format(rdr("GPAmt"), "$###,###,###")
            row("SKU%") = Format(rdr("SKUPct"), "p0")
            tbl.Rows.Add(row)
        End While
        con.Close()

        For Each trow As DataRow In tbl.Rows
            Dim lst As ListViewItem
            lst = lvItem1.Items.Add(trow("Code"))
            For i As Integer = 1 To tbl.Columns.Count - 1
                lst.SubItems.Add(trow(i))
            Next
        Next

        For Each Column As ColumnHeader In lvItem1.Columns
            Column.Width = -2
        Next

        lvItem2.View = View.Details
        lvItem2.GridLines = True
        lvItem2.Columns.Clear()
        lvItem2.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = "Dept"
            .TextAlign = HorizontalAlignment.Center
            .Width = 90
        End With
        lvItem2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Fast"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvItem2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvItem2.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Slow"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvItem2.Columns.Add(colx)

        Dim tbl2 As New DataTable
        tbl2.Columns.Add("Dept")
        tbl2.Columns.Add("Fast")
        tbl2.Columns.Add("Medium")
        tbl2.Columns.Add("Slow")
        Dim row2 As DataRow


        con.Open()
        sql = "SELECT Code, Value1, Value2, Value3 FROM Score_Params WHERE ID = 'Turns' AND Type = 'Item' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'Turns' AND Type = 'Item')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader

        Dim lst2 As ListViewItem
        While rdr.Read
            row2 = tbl2.NewRow
            row2(0) = rdr("Code")
            row2(1) = Format(rdr("Value1"), "###.00")
            row2(2) = Format(rdr("Value2"), "###.00")
            row2(3) = Format(rdr("Value3"), "###.00")

            ''lst2 = lvItem2.Items.Add("Turns at Cost")
            lst2 = lvItem2.Items.Add(row2(0))
            lst2.SubItems.Add(row2(1))
            lst2.SubItems.Add(row2(2))
            lst2.SubItems.Add(row2(3))
        End While
        con.Close()

        'For Each Column As ColumnHeader In lvItem2.Columns
        '    Column.Width = -2
        'Next

        lvItem3.View = View.Details
        lvItem3.GridLines = True
        lvItem3.Columns.Clear()
        lvItem3.Items.Clear()
        colx = New ColumnHeader
        With colx
            .Text = "Dept"
            .TextAlign = HorizontalAlignment.Center
            .Width = 80
        End With
        lvItem3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "High"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvItem3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Medium"
            .TextAlign = HorizontalAlignment.Center
            .Width = 60
        End With
        lvItem3.Columns.Add(colx)
        colx = New ColumnHeader
        With colx
            .Text = "Low"
            .TextAlign = HorizontalAlignment.Center
            .Width = 50
        End With
        lvItem3.Columns.Add(colx)

        Dim tbl3 As New DataTable
        tbl3.Columns.Add("Dept")
        tbl3.Columns.Add("High")
        tbl3.Columns.Add("Medium")
        tbl3.Columns.Add("Low")
        Dim row3 As DataRow
        con.Open()
        sql = "SELECT Code, Value1, Value2, Value3 FROM Score_Params WHERE ID = 'UnitCost' AND Type = 'Item' " & _
            "AND Str_Id = '" & thisStore & "' AND eDate = (SELECT MAX(eDate) FROM Score_Params " & _
            "WHERE Str_Id = '" & thisStore & "' AND ID = 'UnitCost' AND Type = 'Item')"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim lst3 As ListViewItem
        While rdr.Read
            row3 = tbl3.NewRow
            row3(0) = rdr("Code")
            row3(1) = Format(rdr("Value1"), "$###.00")
            row3(2) = Format(rdr("Value2"), "$###.00")
            row3(3) = Format(rdr("Value3"), "$###.00")
            ''lst3 = lvItem3.Items.Add("Unit Cost")
            lst3 = lvItem3.Items.Add(row3(0))
            lst3.SubItems.Add(row3(1))
            lst3.SubItems.Add(row3(2))
            lst3.SubItems.Add(row3(3))
        End While
        con.Close()

        'For Each Column As ColumnHeader In lvItem3.Columns
        '    Column.Width = -2
        'Next
    End Sub


    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                If changesMade Then
                    Select Case MessageBox.Show("Save changes before exiting?", "CHANGE(S) DETECTED!",
                                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        Case DialogResult.Yes
                            Call Save_Changes()
                            e.Cancel = False
                        Case DialogResult.No
                            e.Cancel = False
                        Case Windows.Forms.DialogResult.Cancel
                            e.Cancel = True
                            Exit Sub
                    End Select
                End If
            Case Else
                e.Cancel = False
        End Select
    End Sub

    Private Sub txtWeeks_TextChanged(sender As Object, e As EventArgs) Handles txtWeeks.TextChanged
        oTest = txtWeeks.Text
        If Not IsNumeric(oTest) Then
            MessageBox.Show("Look Back Weeks Must be Numeric and Between 1 and 52.", "ERROR ENTERING LOOK BACK WEEKS")
            txtWeeks.Text = Nothing
            Exit Sub
        End If
        If oTest < 1 Or oTest > 52 Then
            MessageBox.Show("Look Back Weeks Must be Numeric and Between 1 and 52.", "ERROR ENTERING LOOK BACK WEEKS")
            txtWeeks.Text = Nothing
            Exit Sub
        End If
        weeks = CInt(oTest)
    End Sub

    Private Sub txtWeeks_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWeeks.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtGP1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGP1.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtGP2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGP2.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtgp3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGP3.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtgp4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGP4.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtgp5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGP5.KeyPress
        If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub cboScoreStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboScoreStore.SelectedIndexChanged
        thisStore = cboScoreStore.SelectedItem
    End Sub

    Private Sub cboStartDate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStartDate.SelectedIndexChanged
        startDate = cboStartDate.SelectedItem
    End Sub

    Private Sub btnBackScore_Click(sender As Object, e As EventArgs) Handles btnBackScore.Click
        Dim wk As Integer = Nothing
        con.Open()
        sql = "SELECT Week FROM Score_Setup_History WHERE Last_Update = (SELECT MAX(Last_Update FROM Score_Setup_History)"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            wk = rdr("Week")
        End While
        con.Close()
        If IsNothing(wk) Then
            MessageBox.Show("Run scoring setup and try again.", "BACK SCORE")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exepath & "\Score.exe"
        p.Arguments = "" & client & "," & server & "," & dbase & "," & sqluserid & "," & sqlpassword & "," & thisStore & "," & startDate
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        Process.Start(exepath & "\Score.exe")                        ' Scoring is set up to run on Sunday's only
    End Sub

End Class
Imports System.Data.SqlClient
Module Module1

    Sub Main()
        Try
            Dim con, con2 As SqlConnection
            Dim cmd As SqlCommand
            Dim rdr As SqlDataReader
            Dim conString, client, server, dbase, userid, password, sql As String
            Dim startDate, endDate As Date
            Dim cnt As Integer = 0
            Dim added As Integer = 0
            Dim oTest As Object
            Dim stopwatch As New Stopwatch
            Dim firstRecord As Boolean = False
            Dim location, sku, item, prevItem, prevDim1, prevDim2, prevDim3 As String
            Dim dim1, dim2, dim3 As Object
            Dim cost, retail As Decimal
            Dim edate, preveDate, nextEdate, nextSdate, sdate As Date
            Dim prevLocation As String = ""
            Dim prevSku As String = ""
            Dim prevOH, prevMax As Decimal
            Dim prevCost, prevRetail As Decimal
            Dim oh, boh, max, adj, xfer, rtv, recvd, sold, test, thisBeginOH, thisEndOH, prevBegin, prevEnd As Decimal
            Dim prevLeadTime, leadTime, yrwk As Integer
            stopwatch.Start()

            ''Dim args As String() = Environment.GetCommandLineArgs()
            ''Dim txtArray() As String = args(1).Split(";")
            ''client = txtArray(0).ToString
            ''server = txtArray(1).ToString
            ''dbase = txtArray(2).ToString
            ''userid = txtArray(3).ToString
            ''password = txtArray(4).ToString
            ''startDate = CDate(txtArray(5).ToString)
            ''endDate = CDate(txtArray(6).ToString)
            ''conString = "Server=" & server & ";Initial Catalog=" & dbase & ";User Id=" & userid & ";Password=" & password & ""




            ''conString = "Server=RetailClarity1;Initial Catalog=TCM;Integrated Security=True"
            conString = "Server=CURTIS-MOBILE;Initial Catalog=PARGIF;Integrated Security=True"
            startDate = "1/1/2014"
            endDate = "4/1/2017"

            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)

            MsgBox("WARNING - startDate and endDate MUST BE Calendar eDates!!!!!")

            GoTo here



            Console.WriteLine("Merging Item_Inv with transactions from Daily_Transaction_Log")
            con.Open()
            sql = "IF OBJECT_ID('tempDB.dbo.#xfer','U') IS NOT NULL DROP TABLE #xfer; " & _
"IF OBJECT_ID('tempDB.dbo.#adj','U') IS NOT NULL DROP TABLE #adj; " & _
"IF OBJECT_ID('tempDB.dbo.#rtv','U') IS NOT NULL DROP TABLE #rtv; " & _
"IF OBJECT_ID('tempDB.dbo.#rcvd','U') IS NOT NULL DROP TABLE #rcvd; " & _
"IF OBJECT_ID('tempDB.dbo.#sold','U') IS NOT NULL DROP TABLE #sold; " & _
"IF OBJECT_ID('tempDB.dbo.#inv','U') IS NOT NULL DROP TABLE #inv; " & _
"ALTER TABLE Item_Inv ADD Sold decimal(18,4) NULL " & _
"DECLARE @fromDate date = '1/1/2014' " & _
"DECLARE @thruDate date = '4/1/2017' " & _
"select * into #inv from item_inv where edate = @thrudate " & _
"delete from item_inv " & _
"insert into Item_Inv(loc_id, sku, sdate, edate, cost, retail, yrwk,onhand, item) " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, convert(decimal(18,4),0) qty, item from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between sdate and edate and week_id > 0 " & _
"where convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, yrwk, item " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) XFER, Item into #xfer from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between c.sdate and c.edate and c.week_id > 0 " & _
"where [type]='xfer' and convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, Yrwk, Item " & _
"update i set i.xfer=t.xfer from item_inv i " & _
"join #xfer t on t.location=i.loc_id and t.sku=i.sku and t.edate=i.edate " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) ADJ, Item into #adj from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between c.sdate and c.edate and c.week_id > 0 " & _
"where [type]='adj' and convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, Yrwk, Item " & _
"update i set i.adj=t.adj from item_inv i " & _
"join #adj t on t.location=i.loc_id and t.sku=i.sku and t.edate=i.edate " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) RTV, Item into #rtv from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between c.sdate and c.edate and c.week_id > 0 " & _
"where [type]='rtv' and convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, Yrwk,Item " & _
"update i set i.rtv=t.rtv from item_inv i " & _
"join #rtv t on t.location=i.loc_id and t.sku=i.sku and t.edate=i.edate " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) RECVD, Item into #rcvd from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between c.sdate and c.edate and c.week_id > 0 " & _
"where [type]='recvd' and convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, Yrwk, Item " & _
"update i set i.recvd=t.recvd from item_inv i " & _
"join #rcvd t on t.location=i.loc_id and t.sku=i.sku and t.edate=i.edate " & _
"select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) SOLD, Item into #sold from Daily_Transaction_Log l " & _
"join calendar c on convert(date,trans_date) between c.sdate and c.edate and c.week_id > 0 " & _
"where [type]='Sold' and convert(date,trans_date) between @fromDate and @thruDate " & _
"group by location, sku, sdate, edate, Yrwk, Item " & _
"update i set i.sold=t.sold from item_inv i " & _
"join #sold t on t.location=i.loc_id and t.sku=i.sku and t.edate=i.edate " & _
"merge item_inv t using #inv s " & _
"on t.loc_id=s.loc_id and t.sku=s.sku and t.edate=s.edate and t.item=s.item " & _
"when not matched by target " & _
"then insert(loc_id, sku, sdate, edate, cost, retail, yrwk, end_oh, item) " & _
"values(s.loc_id, s.sku, s.sdate, s.edate, s.cost, s.retail, s.yrwk, s.end_oh, s.item) " & _
"when matched then update set t.end_oh=s.end_oh; "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            cmd.ExecuteNonQuery()
            con.Close()

here:       Console.WriteLine("Selecting Skus")
            con.Open()
            con2.Open()
            sql = "SELECT i.Loc_Id, i.Sku, i.sDate, i.eDate, ISNULL(Cost,0) AS Cost, ISNULL(Retail,0) AS Retail, c.YrWk, " & _
                "i.Sold, ISNULL(ADJ,0) AS ADJ, ISNULL(XFER,0) AS XFER, ISNULL(RTV,0) AS RTV, ISNULL(RECVD,0) AS RECVD, " & _
                "ISNULL(Begin_OH,0) AS Begin_OH, ISNULL(End_OH,0) AS End_OH, ISNULL(Max_OH,0) AS Max_OH, " & _
                "Lead_Time, i.Item, i.DIM1, i.DIM2, i.DIM3 FROM Item_Inv i " & _
                "JOIN Item_Master m ON m.Sku = i.Sku " & _
                "JOIN Calendar c ON c.eDate = i.eDate AND c.Week_Id > 0 " & _
                "WHERE Type = 'I' AND i.eDate BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
                "ORDER BY i.Loc_Id, i.Sku, eDate DESC"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            rdr = cmd.ExecuteReader

            While rdr.Read
                location = rdr("Loc_Id")
                sku = rdr("Sku")
                sdate = rdr("sDate")
                edate = rdr("eDate")
                oTest = rdr("Begin_OH")
                If Not IsDBNull(oTest) Then boh = CDec(oTest) Else boh = 0
                oTest = rdr("End_OH")
                If Not IsDBNull(oTest) Then oh = CDec(oTest) Else oh = 0
                If boh > oh Then max = boh Else max = oh
                oTest = rdr("Cost")
                If Not IsDBNull(oTest) Then cost = CDec(oTest) Else cost = 0
                oTest = rdr("Retail")
                If Not IsDBNull(oTest) Then retail = CDec(oTest) Else retail = 0
                oTest = rdr("ADJ")
                If Not IsDBNull(oTest) Then adj = CDec(oTest) Else adj = 0
                oTest = rdr("XFER")
                If Not IsDBNull(oTest) Then xfer = CDec(oTest) Else xfer = 0
                oTest = rdr("RTV")
                If Not IsDBNull(oTest) Then rtv = CDec(oTest) Else rtv = 0
                oTest = rdr("RECVD")
                If Not IsDBNull(oTest) Then recvd = CDec(oTest) Else recvd = 0
                oTest = rdr("Sold")
                If Not IsDBNull(oTest) Then sold = CDec(oTest) Else sold = 0
                oTest = rdr("Lead_Time")
                If Not IsDBNull(oTest) Then leadTime = CInt(oTest) Else leadTime = 0
                oTest = rdr("Item")
                If Not IsDBNull(oTest) Then item = CStr(oTest) Else item = sku
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then dim1 = CStr(oTest) Else dim1 = Nothing
                oTest = rdr("DIM2")
                If Not IsDBNull(oTest) Then dim2 = CStr(oTest) Else dim2 = Nothing
                oTest = rdr("DIM3")
                If Not IsDBNull(oTest) Then dim3 = CStr(oTest) Else dim3 = Nothing
                cnt += 1
                If cnt Mod 1000 = 0 Then
                    Console.WriteLine(cnt & " " & location & " " & sku & " " & edate)
                End If
                test = xfer + adj + recvd + rtv

                If location <> prevLocation Then
                    prevLocation = location
                    prevSku = sku
                    preveDate = edate
                    sdate = DateAdd(DateInterval.Day, -6, edate)
                    thisBeginOH = oh - adj - recvd + rtv - xfer + sold
                    prevBegin = thisBeginOH
                    prevEnd = oh
                    prevCost = cost
                    prevRetail = retail
                    prevLeadTime = leadTime
                    prevItem = item
                    prevDim1 = dim1
                    prevDim2 = dim2
                    prevDim3 = dim3
                    max = oh + sold
                    If oh < 0 Then
                        max = sold
                    Else
                        max = sold + oh
                    End If
                    If max < 0 Then max = 0
                    prevMax = max
                    sql = "UPDATE Item_Inv SET Begin_OH = " & thisBeginOH & ",End_OH = " & oh & ",Max_OH = " & max & ", " & _
                        "Cost = " & cost & ", Retail = " & retail & " " & _
                        "WHERE Loc_Id = '" & prevLocation & "' AND Sku = '" & prevSku & "' AND eDate = '" & preveDate & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                    GoTo 100
                End If
                '
                '          Different sku below
                '
                If sku <> prevSku Then
                    If prevBegin <> 0 Then
                        nextEdate = DateAdd(DateInterval.Day, -7, preveDate)
                        nextSdate = DateAdd(DateInterval.Day, -6, nextEdate)
                        Do While startDate <= nextEdate

                            ''If nextSdate > startDate Then
                            prevMax = thisEndOH
                            If prevMax < 0 Then prevMax = 0
                            sql = "IF NOT EXISTS (SELECT * FROM Item_Inv WHERE Loc_Id = '" & location & "' AND Sku = '" & prevSku & "' " & _
                                "AND sDate = '" & nextSdate & "' AND eDate = '" & nextEdate & "' AND Item = '" & prevItem & "') " & _
                                "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, Begin_OH, End_OH, Max_OH, " & _
                                "Lead_Time, Item, DIM1, DIM2, DIM3) " & _
                                "SELECT '" & location & "', '" & prevSku & "', '" & nextSdate & "', '" & nextEdate & "', " & prevCost & ", " & _
                                prevRetail & ", YrWk, " & prevBegin & ", " & prevBegin & ", " & prevMax & ", " & prevLeadTime & ", '" & _
                                prevItem & "','" & prevDim1 & "','" & prevDim2 & "','" & prevDim3 & "' FROM Calendar " & _
                                "WHERE eDate = '" & nextEdate & "' AND Week_Id > 0 "
                            cmd = New SqlCommand(sql, con2)
                            cmd.ExecuteNonQuery()
                            ''End If
                            nextSdate = DateAdd(DateInterval.Day, -7, nextSdate)
                            nextEdate = DateAdd(DateInterval.Day, -7, nextEdate)
                        Loop
                        ''Console.WriteLine("nextSdate=" & nextSdate)
                        ''Console.ReadLine()

                    End If
                    prevLocation = location
                    prevSku = sku
                    prevItem = item
                    prevDim1 = dim1
                    prevDim2 = dim2
                    prevDim3 = dim3
                    preveDate = edate
                    thisBeginOH = oh - adj - recvd + rtv - xfer + sold
                    prevBegin = thisBeginOH
                    prevEnd = oh
                    prevCost = cost
                    prevRetail = retail
                    max = oh + sold
                    If oh < 0 Then
                        max = sold
                    Else
                        max = sold + oh
                    End If
                    If max < 0 Then max = 0
                    prevMax = max
                    sql = "UPDATE Item_Inv SET Begin_OH = " & thisBeginOH & ",End_OH = " & oh & ",Max_OH = " & max & ", " & _
                        "Cost = " & cost & ", Retail = " & retail & " " & _
                        "WHERE Loc_Id = '" & prevLocation & "' AND Sku = '" & prevSku & "' AND eDate = '" & preveDate & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                    GoTo 100
                End If
                '
                ' Same location and same sku                '
                '
                thisEndOH = prevBegin
                thisBeginOH = thisEndOH - adj - recvd + rtv - xfer + sold
                max = thisEndOH + sold
                If thisEndOH < 0 Then
                    max = sold
                Else
                    max = sold + thisEndOH
                End If
                If max < 0 Then max = 0
                sql = "UPDATE Item_Inv SET Cost = " & cost & ", Retail = " & retail & ", Begin_OH = " & thisBeginOH & ", " & _
                        "End_OH = " & thisEndOH & ", Max_OH = " & max & " WHERE Loc_Id = '" & location & "' AND Sku = '" & sku & "' " & _
                        "AND eDate = '" & edate & "'"
                cmd = New SqlCommand(sql, con2)
                cmd.ExecuteNonQuery()
                nextEdate = DateAdd(DateInterval.Day, -7, preveDate)
                nextSdate = DateAdd(DateInterval.Day, -6, nextEdate)
                prevMax = thisEndOH
                If prevMax < 0 Then prevMax = 0
                If prevBegin <> 0 Then
                    Do While edate < nextEdate
                        nextSdate = DateAdd(DateInterval.Day, -6, nextEdate)
                        If nextSdate > startDate Then
                            sql = "IF NOT EXISTS (SELECT * FROM Item_Inv WHERE Loc_Id = '" & location & "' AND Sku = '" & prevSku & "' " & _
                                "AND sDate = '" & nextSdate & "' AND eDate = '" & nextEdate & "') " & _
                                "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, Begin_OH, End_OH, Max_OH, " & _
                                "Item, DIM1, DIM2, DIM3) " & _
                                "SELECT '" & location & "', '" & prevSku & "', '" & nextSdate & "', '" & nextEdate & "', " & cost & ", " & _
                                retail & ", YrWk, " & prevBegin & ", " & thisEndOH & ", " & prevMax & ",'" & _
                                prevItem & "','" & prevDim1 & "','" & prevDim2 & "','" & prevDim3 & "' " & _
                                "From Calendar WHERE eDate = '" & nextEdate & "' AND Week_Id > 0"
                            cmd = New SqlCommand(sql, con2)
                            cmd.ExecuteNonQuery()

                        End If
                        nextEdate = DateAdd(DateInterval.Day, -7, nextEdate)
                    Loop
                End If
                prevEnd = thisEndOH
                prevBegin = thisBeginOH
                preveDate = edate
100:        End While
            con.Close()
            con2.Close()
            stopwatch.Stop()
            Dim ts As TimeSpan = stopwatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim pgm As String = "RCSetup"
            con.Open()
            sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
                "SELECT '" & Date.Now & "','" & pgm & "',1,'BackFillItems','Completed','','" & et & "'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Module

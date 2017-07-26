Imports System.Data
Imports System.Xml
Imports System.Data.SqlClient
Module Module1
    Private con, con2, con3, masterCon, masterCon2 As SqlConnection
    Private cmd As SqlCommand
    Private rdr, rdr2, mrdr As SqlDataReader
    Private sql As String
    Private tbl As DataTable
    Private row As DataRow
    Private oTest As Object
    Private thisClient As String
    Private thisStore As String = "1"
    Private thiseDate, lastWeekseDate As Date
    Private lastForecast As DateTime
    Private thisYear, startWk, endWk As Integer
    Private lastWeeksYrWk As Integer
    Private forecastHistory As Boolean = False
    Private Modifier As Decimal = 1
    Private stopWatch As Stopwatch
    Private maxYear As Integer
    Sub Main()
        stopWatch = New Stopwatch
        stopwatch.Start()
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RCCLIENT.xml")
        Dim server, database, userId, conString, mConString, exePath, passWord, client As String
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
            If fld = "CLIENTID" Then client = valu
        End While
        mConString = "Server=" & server & ";Initial Catalog=RCClient;Integrated Security=True"
        masterCon = New SqlConnection(mConString)
        masterCon2 = New SqlConnection(mConString)
        masterCon.Open()
        sql = "SELECT Client_Id, Server, [Database], SQLUserID, SQLPassword FROM Client_Master WHERE Status = 'Active' " & _
            "AND Item4Cast = 'Y' ORDER BY Client_Id"
        cmd = New SqlCommand(sql, masterCon)
        mrdr = cmd.ExecuteReader
        While mrdr.Read
            thisClient = mrdr("Client_Id")
            server = mrdr("Server")
            database = mrdr("Database")
            userId = mrdr("SQLUserID")
            passWord = mrdr("SQLPassword")
            conString = "Server=" & server & ";Initial Catalog=" & database & ";Integrated Security=True"
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)

            thisYear = DatePart(DateInterval.Year, Date.Today)
            con.Open()
            sql = "SELECT MIN(eDate) as eDate FROM Calendar WHERE Year_Id = " & thisYear & " "
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thiseDate = rdr("eDate")
            End While
            con.Close()

            con.Open()
            sql = "SELECT MAX(eDate) AS eDate, MAX(YrWk) AS YrWk FROM Calendar WHERE GETDATE() > eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                lastWeekseDate = rdr("eDate")
                lastWeeksYrWk = rdr("YrWk")
            End While
            con.Close()

            con.Open()
            maxYear = 0
            sql = "SELECT MAX(Year) AS Year FROM Seasonality_Index"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                maxYear = rdr("Year")
            End While
            con.Close()

            If maxYear = 0 Then
                Console.WriteLine("Could not find any seasonality indexes,", "ERROR!")
                Exit Sub
            End If




            GoTo 100   ' skip the code below to prevent rebuilding Item_Forecast_History every week




            Console.WriteLine("Recreating Item_Forecast_History for " & thisClient)
            con.Open()
            sql = "IF OBJECT_ID('dbo.Item_Forecast_History','U') IS NOT NULL DROP TABLE dbo.Item_Forecast_History " & _
                    "CREATE TABLE [dbo].[Item_Forecast_History](" & _
                    "[Loc_Id] [varchar](10) NOT NULL, " & _
                    "[Sku] [varchar](90) NOT NULL, " & _
                    "[Item] [varchar](30) NOT NULL, " & _
                    "[DIM1] [varchar](30) NULL, " & _
                    "[DIM2] [varchar](30) NULL, " & _
                    "[DIM3] [varchar](30) NULL, " & _
                    "[YearWk] [int] NOT NULL, " & _
                    "[Season_Code] [varchar](10) NOT NULL, " & _
                    "[Season_Index] [decimal](18, 4) NULL, " & _
                    "[Event_Modifier] [decimal](18, 4) NULL, " & _
                    "[Normalized_Demand] [decimal](18, 4) NULL, " & _
                    "[Smoothed_Demand] [decimal](18, 4) NULL, " & _
                    "[Demand_Trend] [decimal](18, 4) NULL, " & _
                    "[Smoothed_Trend] [decimal](18, 4) NULL, " & _
                    "[Smoothed_Forecast] [decimal](18, 4) NULL, " & _
                    "[Calculated_Demand] [decimal](18, 4) NULL, " & _
                    "[Override_Demand] [decimal](18, 4) NULL, " & _
                    "[Event_Code] [varchar](10) NULL, " & _
                    "[Modified_Actual] [decimal](18,4) NULL, " & _
                    "[Class]  AS ([Season_Code]), " & _
                    "[Year] AS Left(YearWk,4), " & _
                    "[Week] AS Right(YearWk,2), " & _
                 "CONSTRAINT [PK_Item_Forecast_History] PRIMARY KEY CLUSTERED (" & _
                    "[Loc_Id] ASC," & _
                    "[Sku] ASC," & _
                    "[Item] ASC, " & _
                    "[YearWk] ASC," & _
                    "[Season_Code] Asc) " & _
                "WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, " & _
                "ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) " & _
                "ON [PRIMARY];"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Threading.Thread.Sleep(3000)

            con.Open()
            sql = "INSERT INTO Item_Forecast_History (Loc_Id, Sku, Item, DIM1, DIM2, DIM3, YearWk, Season_Code) " & _
                "SELECT Loc_Id, i.Sku, i.Item, i.DIM1, i.DIM2, i.DIM3, c.YrWk, Class from Item_Inv i " & _
                "JOIN Item_Master m on m.Sku = i.Sku " & _
                "JOIN Calendar c ON c.eDate = i.eDate AND Week_Id > 0 " & _
                "WHERE i.eDate BETWEEN '6/14/2014' AND '" & lastWeekseDate & "' " & _
                "AND Loc_Id = '" & thisStore & "' "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            Console.WriteLine("Updating Season Code for " & thisClient)
            cmd.ExecuteNonQuery()
            con.Close()
            Call Update_Process_Log("1", "Recreate Forecast_History", "", "")

            con.Open()
            stopWatch.Start()
            sql = "UPDATE h SET Season_Code = Season FROM Item_Forecast_History h " & _
                "JOIN Item_Master m ON m.Sku = h.Sku " & _
                "WHERE Season <> 'N'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
            Call Update_Process_Log("1", "Update Season_Code", "", "")

            stopWatch.Start()
            Console.WriteLine("Updating Season Index for " & thisClient)

            con.Open()
            sql = "select Loc_Id, season_code, week, season_index into #t1 from seasonality_index where year= " & maxYear & " " & _
                    "update f set f.season_index=s.season_index from Item_Forecast_History f " & _
                    "join #t1 s on s.Loc_Id=f.Loc_Id and s.Season_Code=f.Season_Code and s.[Week]=f.[Week] "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            Console.WriteLine("Cleaning Up " & thisClient)
            sql = "DELETE FROM Item_Forecast_History WHERE Season_Index = 0"         ' Delete records with Seasonality Index of 0
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            sql = "update item_forecast_history set event_modifier=1"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
            Call Update_Process_Log("1", "Update Seasonalty_Indexes", "", "")
            forecastHistory = True

            stopWatch.Start()
            Call Process_Data(forecastHistory)
            Call Update_Process_Log("1", "Process Item_Forecast_History", "", "")

100:
            stopWatch.Start()
            Console.WriteLine("Creating Item_Forecast")
            con.Open()
            sql = "DELETE FROM Item_Forecast"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()

            Dim rr, l18Wks As Integer
            Dim Sku, item, dim1, dim2, dim3, season As String
            Dim startDate, endDate, last18Months, next6Months, lastEdate As Date
            Dim smoothedDemand As Decimal
            Dim eDate As Date
            endDate = DateAdd(DateInterval.Day, 182, lastWeekseDate)
            last18Months = DateAdd(DateInterval.Month, -18, lastWeekseDate)

            con.Open()
            sql = "SELECT eDate, Yrwk FROM Calendar WHERE '" & last18Months & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                eDate = rdr("eDate")
                l18Wks = rdr("YrWk")
            End While
            con.Close()

            con.Open()
            sql = "SELECT MAX(eDate) AS eDate FROM Calendar WHERE Convert(Date,GETDATE()) > eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                lastEdate = rdr("eDate")
            End While
            con.Close()

            con.Open()
            sql = "SELECT MAX(YearWk) AS YearWk FROM Item_Forecast_History WHERE Loc_Id = '" & thisStore & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                startWk = rdr("YearWk")
            End While
            con.Close()
            '           Bump startwk up by one week
            con.Open()
            sql = "SELECT MIN(YrWk) AS YrWk, MIN(eDate) AS eDate FROM Calendar WHERE YrWk > " & startWk & " "
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                startWk = rdr("YrWk")
                startDate = rdr("eDate")
            End While
            con.Close()
            '           Get forecast weeks from Controls - set it to 52 weeks first in case there's no record
            Dim forecastWeeks As Integer = 52
            con.Open()
            sql = "SELECT Value FROM Controls WHERE ID = 'ForecastWeeks' AND Parameter = 'Future'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                forecastWeeks = rdr("Value")
            End While
            con.Close()

            Dim forecastWeek As Date = DateAdd(DateInterval.WeekOfYear, forecastWeeks, startDate)
            endWk = 0
            con.Open()
            sql = "SELECT MIN(YrWk) AS YrWk FROM Calendar WHERE eDate > '" & forecastWeek & "' AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("YrWk")
                If Not IsDBNull(oTest) Then endWk = rdr("YrWk")
            End While
            con.Close()

            If endWk = 0 Then
                con.Open()
                sql = "SELECT Max(YrWk) AS YrWk FROM Calendar "
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    oTest = rdr("YrWk")
                    If Not IsDBNull(oTest) Then endWk = rdr("YrWk")
                End While
                con.Close()
            End If

            con.Open()
            con2.Open()
            sql = "INSERT INTO Item_Forecast (Loc_Id, Sku, Item, YearWk, Season_Code, Event_Modifier) " & _
                "SELECT Loc_Id, Sku, Item, MAX(YearWk) AS YrWk, '', 1 " & _
                "FROM Item_Forecast_History WHERE YearWk >= " & l18Wks & " AND Loc_Id = '" & thisStore & "' " & _
                "GROUP BY Loc_Id, Sku, Item ORDER BY Sku "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 360
            cmd.ExecuteNonQuery()
            con.Close()
            Call Update_Process_Log("1", "Recreate Item_Forecast", "", "")

            stopWatch.Start()
            sql = "UPDATE t SET t.Season_Code = i.Season_Code, " & _
                "t.Season_Index = ISNULL(i.Season_Index,0), " & _
                "t.Smoothed_Demand = ISNULL(i.Smoothed_Demand,0), " & _
                "t.Normalized_Demand = ISNULL(i.Normalized_Demand,0), " & _
                "t.Demand_Trend = ISNULL(i.Demand_Trend,0), " & _
                "t.Smoothed_Trend = ISNULL(i.Smoothed_Trend,0), " & _
                "t.Smoothed_Forecast = ISNULL(i.Smoothed_Forecast,0), " & _
                "t.Calculated_Demand = ISNULL(i.Calculated_Demand,0) FROM Item_Forecast t " & _
                "JOIN Item_Forecast_History i ON i.Loc_Id = t.Loc_Id AND i.Sku = t.Sku " & _
                "AND i.YearWk = t.YearWk "
            con.Open()
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 360
            cmd.ExecuteNonQuery()
            sql = "SELECT t.Loc_Id, t.Sku, t.Item, m.DIM1, m.DIM2, m.DIM3, MaxYrWk, t.Season_Code, ISNULL(Smoothed_Demand,0) AS Smoothed_Demand FROM Item_Forecast t " & _
                "JOIN Calendar c ON c.YrWk = t.YearWk AND c.Week_Id > 0 " & _
                "JOIN Item_Master m ON m.Sku = t.Sku " & _
                "INNER JOIN (SELECT Loc_Id, Sku, Item, MAX(YearWk) AS MaxYrWk FROM Item_Forecast a " & _
                "GROUP BY Loc_Id, Sku, Item) a " & _
                "ON a.Loc_Id = t.Loc_Id AND a.Sku = t.Sku AND a.Item = t.Item " & _
                "ORDER BY t.Loc_Id, t.Sku"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 360
            rdr = cmd.ExecuteReader
            While rdr.Read
                eDate = startDate
                Sku = rdr("Sku")
                oTest = rdr("Item")
                If Not IsDBNull(oTest) Then item = CStr(oTest) Else item = Nothing
                oTest = rdr("DIM1")
                If Not IsDBNull(oTest) Then dim1 = CStr(oTest) Else dim1 = Nothing
                oTest = rdr("DIM2")
                If Not IsDBNull(oTest) Then dim2 = CStr(oTest) Else dim2 = Nothing
                oTest = rdr("DIM3")
                If Not IsDBNull(oTest) Then dim3 = CStr(oTest) Else dim3 = Nothing
                season = rdr("Season_Code")
                smoothedDemand = rdr("Smoothed_Demand")
                rr += 1
                If rr Mod 1000 = 0 Then
                    Console.WriteLine("Processed " & rr & " " & item)
                End If
                sql = "INSERT INTO Item_Forecast (Loc_Id, Sku, Item, DIM1, DIM2, DIM3, YearWk, Season_Code, Event_Modifier) " & _
                    "SELECT '" & thisStore & "','" & Sku & "','" & item & "','" & dim1 & "','" & dim2 & "','" & dim3 & "',YrWk,'" & season & "',1 FROM Calendar " & _
                    "WHERE YrWk BETWEEN " & startWk & " AND " & endWk & " AND Week_Id > 0"
                cmd = New SqlCommand(sql, con2)
                cmd.CommandTimeout = 360
                cmd.ExecuteNonQuery()
            End While
            con.Close()
            con2.Close()
            Console.WriteLine("Updating Season Code")
            con.Open()
            sql = "UPDATE f SET Season_Code = m.Class FROM Item_Forecast f JOIN Item_Master m ON m.Sku = f.Sku " & _
                "WHERE f.YearWk >= " & startWk & " "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            sql = "UPDATE f SET Season_Code = Season FROM Item_Forecast f JOIN Item_Master m ON m.Sku = f.Sku " & _
                "WHERE Season <> 'N' AND f.YearWk >= " & startWk & " "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            sql = "UPDATE f SET f.Season_Index = (SELECT Season_Index FROM Seasonality_Index s " & _
                "WHERE s.Year = " & maxYear & " AND s.Week = f.Week AND s.Season_Code = f.Season_Code AND s.Loc_Id = f.Loc_Id) " & _
                "FROM Item_Forecast f WHERE f.Season_Index IS NULL"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
            Call Update_Process_Log("1", "Update Seasonality Code", "", "")


            forecastHistory = False
            Call Process_Data(forecastHistory)
            Call Update_Process_Log("1", "Process Item Forecast", "", "")

            con.Open()
            Dim thisDateTime As DateTime = CDate(Now)
            Dim ts As TimeSpan = stopwatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            Dim title As String = "Item Forecast"
            Dim message As String = "Forecast completed in  " & ts.Minutes & " "
            sql = "INSERT INTO Message (mDate, Title, Message) " & _
                "SELECT '" & thisDateTime & "','" & title & "','" & message & "'"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 120
            cmd.ExecuteNonQuery()
            con.Close()

            masterCon2.Open()
            sql = "UPDATE Client_Master SET Last_Forecast = GETDATE() WHERE Client_Id = '" & thisClient & "'"
            cmd = New SqlCommand(sql, masterCon2)
            cmd.ExecuteNonQuery()
            masterCon2.Close()
        End While

        masterCon.Close()

    End Sub

    Private Sub Process_Data(forecastHistory)
        Dim cnt As Integer = 0
        Dim rr As Integer = 0
        Dim prevNormalizedDemand As Decimal = 0
        Dim prevSmoothedDemand As Decimal = 0
        Dim prevSmoothedTrend As Decimal = 0
        Dim eventModifiedActual As Decimal = 0
        Dim prevYrWk As Integer = 0
        Dim onhand As Decimal = 0
        Dim prevSku As String = ""
        Dim prevStore As String = ""
        Dim store, sku, item, dim1, dim2, dim3, season As String
        Dim evnt As String = Nothing
        Dim progress As String = ""
        Dim hasEventCode As Boolean = False
        Dim yrwk, wk As Integer
        Dim demandCoeff, trendCoeff, normalizedDemand, smoothedDemand, smoothedTrend, demandTrend As Decimal
        Dim smoothedForecast, sold, seasonIndex, eventModifier, calculatedDemand, avgSmoothedTrend, val As Decimal
        Dim prevSmoothedForecast As Decimal

        con.Open()
        sql = "SELECT Value FROM Controls WHERE ID = 'ItemForecast' AND Parameter = 'DemandCoefficient'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            demandCoeff = CDec(rdr("Value"))
        End While
        con.Close()

        con.Open()
        sql = "SELECT Value FROM Controls WHERE ID = 'ItemForecast' AND Parameter = 'TrendCoefficient'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            trendCoeff = CDec(rdr("Value"))
        End While
        con.Close()

        con.Open()
        con2.Open()

        If forecastHistory Then
            Console.WriteLine("Updating Forecast History")
            sql = "SELECT f.Loc_Id, f.Sku, f.Item, f.DIM1, f.DIM2, f.DIM3, YearWk, Season_Code, " & _
            "ISNULL(Normalized_Demand,0) AS Normalized_Demand, " & _
            "ISNULL(Smoothed_Demand,0) AS Smoothed_Demand , " & _
            "ISNULL(Smoothed_Trend,0) AS Smoothed_Trend, " & _
            "ISNULL(Sold,0) AS Sold, " & _
            "ISNULL(Season_Index,0) AS Season_Index, " & _
            "ISNULL(Event_Modifier,1) AS Event_Modifier " & _
            "FROM Item_Forecast_History f " & _
            "LEFT JOIN Item_Sales d ON d.Loc_Id = f.Loc_Id AND d.Sku = f.Sku AND d.YrWk = f.YearWk " & _
            "WHERE f.Loc_Id = '1' " & _
            "ORDER BY Loc_Id, Sku, YearWk"
        Else
            Console.WriteLine("Updating Forecast")
            sql = "SELECT f.Loc_Id, f.Sku, f.Item, f.DIM1, f.DIM2, f.DIM3, YearWk, Season_Code, " & _
            "ISNULL(Normalized_Demand,0) AS Normalized_Demand, " & _
            "ISNULL(Smoothed_Demand,0) AS Smoothed_Demand , " & _
            "ISNULL(Smoothed_Trend,0) AS Smoothed_Trend, " & _
            "ISNULL(Season_Index,0) AS Season_Index, " & _
            "ISNULL(Event_Modifier,1) AS Event_Modifier, " & _
            "ISNULL(Sold,0) AS Sold " & _
            "FROM Item_Forecast f " & _
            "LEFT JOIN Item_Sales d ON d.Loc_Id = f.Loc_Id AND d.Sku = f.Sku AND d.YrWk = f.YearWk " & _
            "WHERE f.Loc_Id = '1' " & _
            "ORDER BY Loc_Id, f.Sku, YearWk"
        End If
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 360
        rdr = cmd.ExecuteReader
        While rdr.Read
            rr += 1
            If rr Mod 1000 = 0 Then
                progress = rr & " " & store & " " & sku & " " & yrwk
                Console.WriteLine(progress)
            End If
            store = rdr("Loc_Id")
            sku = rdr("Sku")
            yrwk = rdr("YearWk")
            season = rdr("Season_Code")
            normalizedDemand = rdr("Normalized_Demand")
            smoothedDemand = rdr("Smoothed_Demand")
            smoothedTrend = rdr("Smoothed_Trend")
            seasonIndex = rdr("Season_Index")
            eventModifier = Modifier
            sold = rdr("Sold")

            ''If seasonIndex = 0 Then GoTo 100 '                         Skip records without a seasonality index

            If store <> prevStore Then
                prevStore = store
                prevSku = sku
                prevNormalizedDemand = normalizedDemand
                prevSmoothedDemand = smoothedDemand
                prevSmoothedTrend = smoothedTrend
                cnt = 1

            End If
            If sku <> prevSku Then
                prevSku = sku
                prevNormalizedDemand = normalizedDemand
                prevSmoothedDemand = smoothedDemand
                prevSmoothedTrend = smoothedTrend
                cnt = 1
            End If
            If seasonIndex = 0 Then
                normalizedDemand = 0
            Else
                normalizedDemand = sold / seasonIndex / eventModifier
            End If
            If forecastHistory Then
                Select Case cnt
                    Case 1
                        sql = "UPDATE Item_Forecast_History SET Normalized_Demand = " & normalizedDemand & ", Smoothed_Demand = 0, " & _
                                "Demand_Trend = 0, Smoothed_Trend = 0 " & _
                                "WHERE Loc_Id = '" & store & "' AND Sku = '" & sku & "' AND YearWk = " & yrwk & " " & _
                                "AND Season_Code = '" & season & "'"
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        prevNormalizedDemand = normalizedDemand
                        cnt = 2
                    Case 2
                        sql = "UPDATE Item_Forecast_History SET Normalized_Demand = " & normalizedDemand & ", " & _
                              "Smoothed_Demand = " & prevNormalizedDemand & ", " & _
                              "Demand_Trend = 0, Smoothed_Trend = 0 " & _
                              "WHERE Loc_Id = '" & store & "' AND Sku = '" & sku & "' AND YearWk = " & yrwk & " " & _
                              "AND Season_Code = '" & season & "'"
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        prevSmoothedDemand = prevNormalizedDemand
                        prevNormalizedDemand = normalizedDemand
                        prevSmoothedTrend = smoothedTrend
                        cnt = 3
                    Case Else
                        smoothedDemand = (prevNormalizedDemand * demandCoeff) + (prevSmoothedDemand * (1 - demandCoeff))
                        demandTrend = smoothedDemand - prevSmoothedDemand
                        If cnt = 3 Then
                            smoothedTrend = demandTrend
                        Else
                            smoothedTrend = (demandTrend * trendCoeff) + (prevSmoothedTrend * (1 - trendCoeff))
                        End If
                        Dim c As Integer = 0
                        Dim t As Decimal = 0
                        If cnt > 3 Then
                            smoothedForecast = smoothedDemand + smoothedTrend
                            calculatedDemand = (smoothedForecast * seasonIndex) * eventModifier
                        Else
                            smoothedForecast = 0
                            calculatedDemand = 0
                        End If
                        sql = "UPDATE Item_Forecast_History SET Normalized_Demand = " & normalizedDemand & ", " & _
                                "Smoothed_Demand = " & smoothedDemand & ", Demand_Trend = " & demandTrend & ", " & _
                                "Smoothed_Trend = " & smoothedTrend & ", " & _
                                "Smoothed_Forecast = " & smoothedForecast & ", " & _
                                "Calculated_Demand = " & calculatedDemand & " " & _
                                "WHERE Loc_Id = '" & store & "' AND Sku = '" & sku & "' AND YearWk = " & yrwk & " " & _
                                "AND Season_Code = '" & season & "'"
                        cmd = New SqlCommand(sql, con2)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        prevNormalizedDemand = normalizedDemand
                        prevSmoothedDemand = smoothedDemand
                        prevSmoothedTrend = smoothedTrend
                        cnt += 1
                End Select
            Else

                Select Case cnt
                    Case 1
                        prevNormalizedDemand = normalizedDemand
                        prevSmoothedDemand = smoothedDemand
                        cnt = 2
                        GoTo 100
                    Case 2
                        Dim t As Decimal = 0
                        con3.Open()
                        sql = "SELECT ISNULL(d.Sold,0) AS Sold, ISNULL(i.Max_OH,0) AS OnHand, Event_Code FROM Item_Sales d " & _
                            "JOIN Item_Inv i ON i.Loc_Id = d.Loc_Id AND i.Sku = d.Sku AND i.eDate = d.eDate " & _
                            "LEFT JOIN Forecast_Events p ON p.Str_Id = d.Str_Id AND p.Sku = d.Sku AND p.YearWk = d.YrWk " & _
                            "WHERE d.Loc_Id = '" & thisStore & "' AND d.Sku = '" & sku & "' AND d.YrWk = " & yrwk & ""
                        cmd = New SqlCommand(sql, con3)
                        rdr2 = cmd.ExecuteReader
                        While rdr2.Read
                            sold = rdr2("Sold")
                            onhand = rdr2("OnHand")
                            oTest = rdr2("Event_Code")

                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                If IsNumeric(oTest) Then evnt = oTest
                            End If

                        End While
                        con3.Close()
                        If seasonIndex = 0 Then
                            normalizedDemand = 0
                        Else
                            normalizedDemand = sold / seasonIndex / eventModifier
                        End If
                        smoothedDemand = (prevNormalizedDemand * demandCoeff) + (prevSmoothedDemand * (1 - demandCoeff))
                        demandTrend = smoothedDemand - prevSmoothedDemand
                        smoothedTrend = (demandTrend * trendCoeff) + (prevSmoothedTrend * (1 - trendCoeff))
                        smoothedForecast = (prevNormalizedDemand * demandCoeff) + prevSmoothedDemand * (1 - demandCoeff)
                        calculatedDemand = (smoothedForecast * seasonIndex) * eventModifier
                        If Not IsNothing(evnt) And calculatedDemand = 0 Then
                            eventModifiedActual = 0
                            hasEventCode = True
                        Else
                            If calculatedDemand <> 0 Then
                                eventModifiedActual = sold / calculatedDemand
                            Else : eventModifiedActual = 0
                            End If
                        End If
                        sql = "UPDATE Item_Forecast SET Smoothed_Forecast = " & smoothedForecast & ", " & _
                             "Calculated_Demand = " & calculatedDemand & ", " & _
                             "Normalized_Demand = " & normalizedDemand & ", " & _
                             "Smoothed_Demand = " & smoothedDemand & ", " & _
                             "Demand_Trend = " & demandTrend & ", " & _
                             "Smoothed_Trend = " & smoothedTrend & ", " & _
                             "Modified_Actual = " & eventModifiedActual & " " & _
                             "WHERE Loc_Id = '" & store & "' AND Sku = '" & sku & "' AND YearWk = " & yrwk & " " & _
                             "AND Season_Code = '" & season & "'"
                            cmd = New SqlCommand(sql, con2)
                            cmd.CommandTimeout = 120
                            cmd.ExecuteNonQuery()
                            prevSmoothedForecast = smoothedForecast
                            prevNormalizedDemand = normalizedDemand
                            prevSmoothedDemand = smoothedDemand
                            If onhand > 0 Then
                            sql = "IF NOT EXISTS (SELECT * FROM Item_Forecast_History " & _
                                "WHERE Loc_Id = '" & thisStore & "' AND Sku = '" & sku & "' " & _
                                "AND YearWk = " & yrwk & " AND Season_Code = '" & season & "' AND Season_Index <> 0) " & _
                                "INSERT INTO Item_Forecast_History (Loc_Id, Sku, Item, DIM1, DIM2, DIM3, YearWk, Season_Code, Season_Index, Event_Modifier, " & _
                                "Smoothed_Forecast, Calculated_Demand, Normalized_Demand, Smoothed_Demand, " & _
                                "Demand_Trend, Smoothed_Trend, Event_Code, Modified_Actual) " & _
                                "SELECT Loc_Id, Sku, Item, DIM1, DIM2, DIM3, YearWk, Season_Code, Season_Index, 1, ISNULL(Smoothed_Forecast,0), " & _
                                "ISNULL(Calculated_Demand,0), ISNULL(Normalized_Demand,0), ISNULL(Smoothed_Demand,0), " & _
                                "ISNULL(Demand_Trend,0), ISNULL(Smoothed_Trend,0), Event_Code, ISNULL(Modified_Actual,0) " & _
                                "FROM Item_Forecast WHERE Loc_Id = '" & thisStore & "' AND Sku = '" & sku & "' " & _
                                "AND YearWk = " & yrwk & " AND Season_Code = '" & season & "' AND Season_Index <> 0 "
                                cmd = New SqlCommand(sql, con2)
                                cmd.ExecuteNonQuery()
                            End If
                            cnt += 1
                    Case Else
                            If hasEventCode And eventModifiedActual > 1 Then
                                smoothedForecast = prevSmoothedForecast
                                hasEventCode = False
                            Else
                                smoothedForecast = prevSmoothedForecast
                            End If
                            calculatedDemand = (smoothedForecast * seasonIndex) * eventModifier
                            If IsDBNull(calculatedDemand) Then calculatedDemand = 0
                        sql = "UPDATE Item_Forecast SET Smoothed_Forecast = " & smoothedForecast & ", " & _
                            "Calculated_Demand = " & calculatedDemand & " WHERE Loc_Id = '" & thisStore & "' " & _
                            "AND Sku = '" & sku & "' AND YearWk = " & yrwk & " AND Season_Code = '" & season & "'"
                            cmd = New SqlCommand(sql, con2)
                            cmd.CommandTimeout = 240
                            cmd.ExecuteNonQuery()
                            prevSmoothedForecast = smoothedForecast
                            cnt += 1
                End Select

            End If
100:
        End While
        con.Close()
        con2.Close()
    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        stopWatch.Stop()
        Dim ts As TimeSpan = stopWatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "Item Forecasting"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & Date.Now & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
End Module

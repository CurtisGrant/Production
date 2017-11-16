Imports System
Imports System.Xml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports System.Globalization

Public Class DBAdmin
    Private Shared oTest As Object
    Private Shared itExists As Boolean = False
    Private Shared servr, theVendor, theBuyer, theDept, theClass, thePLine, tableChecked, dbName, theButton As String
    Private Shared server, clientServer, path, log, theFile, xmlPath, xmlDirectory, dbpath, tableName, sql, constr As String
    Private Shared RCClientConString, RCClientServer As String
    Private Shared thisDate As Date = Now
    Private Shared today As Date = Date.Today
    Private Shared thisYear As Integer = DatePart("yyyy", thisDate)
    Private Shared BuildStartDate, BuildEndDate As Date
    Private Shared invDateConfirmed As Boolean = False
    Public Shared conString, errorLog, thisClient, message, arguments As String
    Public Shared con, con2, con3, con4, con5, rcCon As SqlConnection
    Private Shared sqlUserID, sqlPassword As String
    Private Shared adapter As SqlDataAdapter
    Private Shared ds As DataSet
    Private Shared cmd As SqlCommand
    Private Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Private Shared xmlReader As XmlTextReader
    Private Shared cnt As Int32
    Private Shared tcost, tretail, tmkdn As Decimal
    Private Shared tRecs As Int32 = 0
    Private Shared haveCalendar As Boolean = False
    Private Shared serverTable As DataTable
    Private Shared stopwatch As Stopwatch
    Public Shared exePath As String

    Private Sub DBAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            '
            '   Comment line below for CURTIS-MOBILE only!
            '   
            serverTable = InitialSetup.serverTable
            RCClientConString = InitialSetup.RCClientConString
            RCClientServer = InitialSetup.RCClientServer
            ''exePath = InitialSetup.exePath
            '
            '  Comment next 4 lines for CURTIS-MOBILE only!
            '  Comment line 53 in InitialSetup also
            '
            cboServers.Items.Clear()
            For Each rw In serverTable.Rows
                cboServers.Items.Add(rw("ServerName"))
            Next
            cboServers.Items.Add("CURTIS-MOBILE")
            cboServers.SelectedIndex = 0
            dtp.Format = DateTimePickerFormat.Short
            DTP1.Format = DateTimePickerFormat.Short
        Catch ex As Exception
            MsgBox("ERROR! CAN NOT LOCATE ANY SERVERS ON THIS NETWORK.")
            servr = cboServers.SelectedItem

        Finally
            ' Application.Exit()
        End Try
    End Sub

    Private Sub cboServers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboServers.SelectedIndexChanged
        Try
            Dim db As String = ""
            clientServer = cboServers.SelectedItem
            If clientServer = "ADD NEW" Then Exit Sub
            conString = "Server=" & clientServer & ";Initial Catalog=Master;Integrated Security=True"
            con = New SqlConnection(conString)
            con.Open()
            sql = "USE Master; " & _
                "SELECT name FROM sys.databases ORDER BY name"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            cboDatabase.Items.Clear()
            cboDatabase.Items.Add("ADD NEW")
            While rdr.Read
                db = rdr("name")
                cboDatabase.Items.Add(db)
            End While
            con.Close()

            conString = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR SELECTING SERVER")
        End Try
    End Sub

    Private Sub cboDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDatabase.SelectedIndexChanged
        dbName = cboDatabase.SelectedItem









        Dim x As String = "Server=" & clientServer & ";Initial Catalog=RCClient;Integrated Security=True"
        rcCon = New SqlConnection(x)
        '' Dim conx As New SqlConnection(RCClientConString)







        rcCon.Open()
        sql = "SELECT Client_Id FROM Client_Master WHERE [Database] = '" & dbName & "'"
        cmd = New SqlCommand(sql, rcCon)
        rdr = cmd.ExecuteReader
        cboClient.Items.Clear()
        cboClient.Items.Add("ADD NEW")
        While rdr.Read
            cboClient.Items.Add(rdr("Client_Id"))
        End While
        rcCon.Close()
        If dbName = "ADD NEW" Then btnInstallDatabase.Visible = True
        If Not IsNothing(cboClient.Items) Then cboClient.SelectedIndex = 0

        lblClientAdd.Visible = True
        cboClient.Visible = True
        lblClient.Visible = True
        txtClient.Visible = True
        btnEditClient.Visible = True

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub


    Private Sub btnEditClient_Click(sender As Object, e As EventArgs) Handles btnEditClient.Click
        Try
            thisClient = txtClient.Text
            If thisClient = "ADD NEW" Then
                MessageBox.Show("Enter an ID for the Client and try again.", "ERROR!")
                Exit Sub
            End If
            Dim Client_Setup As New Client(clientServer, dbName, thisClient, RCClientConString, serverTable)
            Client_Setup.Show()
            GroupBox1.Visible = False
            GroupBox5.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnInstallDatabase_Click(sender As Object, e As EventArgs) Handles btnInstallDatabase.Click
        Try
            server = cboServers.SelectedItem

            If cboDatabase.SelectedItem = "ADD NEW" Then
                Dim prompt, title As String
                prompt = "Enter ID for new database."
                title = "Retail Clarity Client Setup"
                dbName = InputBox(prompt, title, "", 100, 100)
                cboDatabase.Items.Add(dbName)
                Dim idx As Integer = cboDatabase.FindString(dbName)
                cboDatabase.SelectedIndex = idx
            End If
            If dbName Is Nothing Or dbName = "" Then
                MsgBox("Enter a name for the Database and try again!")
                Exit Sub
            End If

            Dim Client_Setup As New Client(server, dbName, thisClient, RCClientConString, serverTable)
            'Client_Setup.Show()
            'GroupBox5.Visible = True

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBuyers_Click(sender As Object, e As EventArgs) Handles btnBuyers.Click
        If haveCalendar = False Then
            MsgBox("You need a calendar for this year before you can continue!")
            Exit Sub
        End If
        ''arguments = thisClient & ";" & server & ";" & dbName & ";" & xmlPath & ";" & sqlUserID & ";" &
        ''        sqlPassword & ";" & Date.Today & ";" & Date.Today & ";" & errorLog & ";" & Client.DoMarketing
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim cnt As Integer = 0
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";BUYERS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Buyers"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtBuyers.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnClasses.Enabled = True
        GroupBox9.Visible = True
        Me.Refresh()
    End Sub

    Private Sub btnClasses_Click(sender As Object, e As EventArgs) Handles btnClasses.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";CLASSES"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Classes"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtClasses.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnDept.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnDept_Click(sender As Object, e As EventArgs) Handles btnDept.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";DEPARTMENTS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Departments"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtDepartments.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnLocations.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnLocations_Click(sender As Object, e As EventArgs) Handles btnLocations.Click
        If haveCalendar = False Then
            MsgBox("You need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";LOCATIONS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Locations"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtLocations.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        ''                      Add GMROI Weeks to Controls Table
        con.Open()
        sql = "IF NOT EXISTS (SELECT * FROM Controls WHERE ID = 'GMROI' AND Parameter = 'Weeks') " & _
            "INSERT INTO Controls (ID, Parameter, Value) SELECT 'GMROI','Weeks',26"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        sql = "IF NOT EXISTS (SELECT * FROM Controls WHERE ID = 'OptionalSoftware' AND Parameter = 'SalesChart') " & _
            "INSERT INTO Controls(ID, Parameter, Value) SELECT 'OptionalSoftware','SalesChart','YES'"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        btnPLine.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnPLine_Click(sender As Object, e As EventArgs) Handles btnPLine.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";PLINES"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM ProductLines"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtProductlines.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnSeasons.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnSeasons_Click(sender As Object, e As EventArgs) Handles btnSeasons.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";SEASONS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Seasons"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtSeasons.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnStores.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnStores_Click(sender As Object, e As EventArgs) Handles btnStores.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";STORES"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        txtStores.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Stores"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtStores.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnVendors.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnVendors_Click(sender As Object, e As EventArgs) Handles btnVendors.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";VENDORS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Vendors"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtVendors.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnItems.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnItems_Click(sender As Object, e As EventArgs) Handles btnItems.Click
        If haveCalendar = False Then
            MsgBox("Your need a calendar for this year before you can continue!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";ITEMS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        Dim thisDate As Date = CDate(Date.Today)
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Item_Master"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtItems.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        btnBarcodes.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnBarcodes_Click(sender As Object, e As EventArgs) Handles btnBarcodes.Click
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";BARCODES"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        cnt = 0
        Dim thisDate As Date = CDate(Date.Today)
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Item_Barcodes"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtBarcodes.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        GroupBox8.Visible = True
        btnAdj.Enabled = True
        Me.Refresh()
    End Sub
    Private Sub btnInv_Click(sender As Object, e As EventArgs) Handles btnInv.Click
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min and Max Dates and try again", "ERROR - CAN NOT CONTINUE!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()

        txtprogress.Text = "Inserting item records from Daily_Transaction_Log"
        Me.Refresh()
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
        End If            '                                  update inventory from the Inventory.xml extract

        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";INVENTORY"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        '                 Change sDate and eDate to match BuildEndDate
        '                 Clean out Begin_OH, Max_OH and all the tranaction fields
        '                 The backfill code will add all the deleted records back in
        con.Open()
        sql = "UPDATE i SET i.sDate = c.sDate, i.eDate = c.eDate, Begin_OH = NULL, Max_OH = NULL, " &
            "ADJ = NULL, XFER = NULL, RTV = NULL, RECVD = NULL, SOLD = NULL FROM Item_Inv i" &
            "JOIN Calendar c On '" & BuildEndDate & "' BETWEEN c.sDate AND c.eDate AND c.Week_Id > 0 "
        cmd = New SqlCommand(sql, con)
        con.Close()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, SUM(ISNULL(End_OH,0) * ISNULL(Cost,0)) cost, SUM(ISNULL(End_OH,0) * ISNULL(Retail,0)) retail FROM Item_Inv"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()

        txtInvRecords.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtInvCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtInvRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        Me.Refresh()

        stopwatch = New Stopwatch
        stopwatch.Start()


12:


        Call Update_Process_Log("1", "Update Item_Sales", "Updated Item_Sales", "")
        txtprogress.Text = ""
        GroupBox6.Visible = True
        Me.Refresh()

    End Sub

    Private Sub btnAdj_Click(sender As Object, e As EventArgs) Handles btnAdj.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min and Max Dates and try again", "ERROR - CAN NOT CONTINUE!")
            Exit Sub
        End If
        If IsDBNull(arguments) Then Console.WriteLine("its null")
        If IsNothing(arguments) Then Console.WriteLine("its nothing")
        If arguments = "" Then Console.WriteLine("its double quotes")
        Console.WriteLine(arguments)
        Console.ReadLine()

        arguments = thisClient & ";" & server & ";" & dbName & ";" & exePath & ";" & xmlPath & ";" & sqlUserID & ";" &
                sqlPassword & ";" & BuildStartDate & ";" & BuildEndDate & ";" & errorLog & ";" & Client.DoMarketing
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";ADJUSTMENTS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(ISNULL(QTY,0) * ISNULL(COST,0)),0) cost, " &
            "ISNULL(SUM(ISNULL(QTY,0) * ISNULL(RETAIL,0)),0) retail " &
            "FROM Inv_Transaction_Log WHERE [TYPE] = 'ADJ'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtAdjustments.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtAdjCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtAdjRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)

        ''theFile = "\Adjustment.xml"
        ''Call Process_Data(theFile, "ADJ")                                                  ' Add new records to Daily Transaction Log
        btnPhys.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnPhys_Click(sender As Object, e As EventArgs) Handles btnPhys.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";PHYSICAL"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Inv_Transaction_Log WHERE [TYPE] = 'ADJ' AND TRANS_ID LIKE 'P%'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtPhysical.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtPhysCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtPhysRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        stopwatch = New Stopwatch
        stopwatch.Start()
        'theFile = "\Physical.xml"
        'Call Process_Data(theFile, "ADJ")                                                         ' Add new records to Daily Transaction Log
        btnRecv.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnRecv_Click(sender As Object, e As EventArgs) Handles btnRecv.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";RECEIPTS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Inv_Transaction_Log WHERE [TYPE] = 'RECVD'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtReceipts.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtRecvCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtRecvRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        stopwatch = New Stopwatch
        stopwatch.Start()
        'theFile = "\Receipt.xml"
        'Call Process_Data(theFile, "RECVD")                                                      ' Add new records to Daily Transaction Log
        'btnRtn.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnRtn_Click(sender As Object, e As EventArgs) Handles btnRtn.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";RETURNS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Inv_Transaction_Log WHERE [TYPE] = 'RTV'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtReturns.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtRtnCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtRtnRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        stopwatch = New Stopwatch
        stopwatch.Start()
        'theFile = "\Return.xml"
        'Call Process_Data(theFile, "RTV")                                                       ' Add new records to Daily Transaction Log
        btnOrders.Enabled = True
        Me.Refresh()
    End Sub


    Private Sub btnOrders_Click(sender As Object, e As EventArgs) Handles btnOrders.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";ORDERS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Daily_Transaction_Log WHERE [TYPE] = 'Sold' AND SALE_TYPE = 'O'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtOrders.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtOrdersCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtOrdersRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        stopwatch = New Stopwatch
        stopwatch.Start()
        btnSales.Enabled = True
        Me.Refresh()
    End Sub
    Private Sub btnSales_Click(sender As Object, e As EventArgs) Handles btnSales.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";SALES"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, " &
            "ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Daily_Transaction_Log WHERE [TYPE] = 'Sold' AND SALE_TYPE = 'T'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtSales.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtSalesCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtSalesRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        btnXfer.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnXfer_Click(sender As Object, e As EventArgs) Handles btnXfer.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";TRANSFERS"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(QTY * COST),0) cost, ISNULL(SUM(QTY * RETAIL),0) retail " &
            "FROM Inv_Transaction_Log WHERE [TYPE] = 'XFER'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtTransfers.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtXferCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtXferRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        stopwatch = New Stopwatch
        stopwatch.Start()
        'theFile = "\Transfer.xml"
        'Call Process_Data(theFile, "XFER")
        btnLoadPOs.Enabled = True
        Me.Refresh()
    End Sub


    Private Sub btnLoadPOs_Click(sender As Object, e As EventArgs) Handles btnLoadPOs.Click
        txtprogress.Text = "Loading POs"
        Me.Refresh()
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";POHEADER"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        Dim p2 As New ProcessStartInfo
        p2.FileName = exePath & "\XMLUpdate.exe"
        p2.Arguments = arguments & ";PODETAIL"
        p2.UseShellExecute = True
        p2.WindowStyle = ProcessWindowStyle.Normal
        Dim proc2 As Process = Process.Start(p2)
        proc2.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, SUM(ISNULL(COST,0)) cost, SUM(ISNULL(RETAIL,0)) retail FROM PO_Detail"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            retail = rdr("retail")
        End While
        con.Close()
        txtPO.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtPOCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtPORetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        txtprogress.Text = ""
        btnLoadPReq.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnLoadPReq_Click(sender As Object, e As EventArgs) Handles btnLoadPReq.Click
        txtprogress.Text = "Load Purchase Requests"
        Me.Refresh()
        If Not invDateConfirmed Then
            MessageBox.Show("Set Min And Max Dates And try again", "ERROR - CAN Not CONTINUE!")
            Exit Sub
        End If
        stopwatch = New Stopwatch
        stopwatch.Start()
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\XMLUpdate.exe"
        p.Arguments = arguments & ";PREQHEADER"
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
        proc.WaitForExit()

        Dim p2 As New ProcessStartInfo
        p2.FileName = exePath & "\XMLUpdate.exe"
        p2.Arguments = arguments & ";PREQDETAIL"
        p2.UseShellExecute = True
        p2.WindowStyle = ProcessWindowStyle.Normal
        Dim proc2 As Process = Process.Start(p2)
        proc.WaitForExit()

        Dim cost As Decimal = 0
        Dim retail As Decimal = 0
        cnt = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt, ISNULL(SUM(COST),0) cost FROM Purchase_Request_Detail"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
            cost = rdr("cost")
            ''retail = rdr("retail")
        End While
        con.Close()
        txtPReq.Text = cnt.ToString("0,0", CultureInfo.InvariantCulture)
        txtPReqCost.Text = cost.ToString("0,0", CultureInfo.InvariantCulture)
        txtPReqRetail.Text = retail.ToString("0,0", CultureInfo.InvariantCulture)
        txtprogress.Text = ""
        Me.Refresh()
        btnInv.Enabled = True
        btnCreateWeekly.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub Load_Data_Table(thefile, tblName, altId, altDescription)
        Dim thePath As String = ""
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        thePath = xmlPath & thefile
        If System.IO.File.Exists(thePath) Then
        Else
            MsgBox(thePath & " Not found. Copy it there And try again.")
            Exit Sub
        End If
        cnt = 0
        Dim records As Integer = 0
        Dim aok As Boolean = True
        tableName = tblName
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim ID, desc, loc, stat As String
        Dim thisDept As String = ""
        Dim xmlFile As XmlReader
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)

        txtprogress.Text = "Inserting records into " & tableName
        Me.Refresh()

        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            con.Open()
            For Each row In dt.Rows
                ID = CStr(row("ID"))
                oTest = row("DESCRIPTION")
                If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = "NA"
                If ID <> "TOTALS" Then
                    If cnt Mod 1000 = 0 Then
                        txtprogress.Text = "Processed " & cnt & " records."
                        Me.Refresh()
                    End If
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        Select Case tableName
                            Case "Classes"
                                oTest = row("DEPT")
                                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then thisDept = oTest
                                sql = "IF Not EXISTS (SELECT * FROM Classes WHERE ID = '" & ID & "' AND Dept = '" & thisDept & "') " &
                                  "INSERT INTO " & tableName & " (ID, Dept, Description, Status, Orig_Date) " &
                                  "SELECT '" & ID & "','" & thisDept & "', '" & desc & "', 'Active', '" & today & "'"
                            Case "Stores"
                                oTest = row("LOCATION")
                                If Not IsDBNull(oTest) Then loc = CStr(oTest) Else loc = ID
                                sql = "IF NOT EXISTS (SELECT * FROM Stores WHERE ID = '" & ID & "') " & _
                                      "INSERT INTO " & tableName & " (ID, Description, Inv_Loc, Status, Orig_Date) " & _
                                      "SELECT '" & ID & "','" & desc & "', '" & loc & "', 'Active', '" & today & "'"
                            Case Else
                                sql = "IF NOT EXISTS (SELECT * From " & tableName & " WHERE ID = '" & ID & "') " & _
                                    "INSERT INTO " & tableName & " (ID, Description, Status, Orig_Date) " & _
                                    "SELECT '" & ID & "','" & desc & "','Active','" & today & "'"
                        End Select

                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        cnt += 1
                    End If
                Else
                    records = CInt(row("DESCRIPTION"))
                End If
            Next
            con.Close()
            xmlFile.Close()
        End If
        Call Update_Process_Log("1", "Update Tables", "Updated " & thefile, "")
        txtprogress.Text = ""
        Me.Refresh()
        If cnt <> records Then
            aok = False
            message = "Record count " & cnt & " does not match XML Total Record " & records & "."
            Call Log_Error(message)
        End If
        records = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM " & tblName & ""
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            records = rdr("cnt")
        End While
        con.Close()
        If records <> cnt Then
            aok = False
            message = "Record count in SQL " & cnt & " does not match records counted in " & thefile & " table."
            Call Log_Error(message)
        End If
        If aok Then
            Select Case tblName
                Case "Buyers"
                    chkBuyers.Checked = True
                Case "Classes"
                    chkClasses.Checked = True
                Case "Departments"
                    chkDepartments.Checked = True
                Case "Locations"
                    chkLoactions.Checked = True
                Case "ProductLines"
                    chkProdLines.Checked = True
                Case "Seasons"
                    chkSeasons.Checked = True
                Case "Stores"
                    chkStores.Checked = True
                Case "Vendors"
                    chkVendors.Checked = True
            End Select
        End If
        If thefile = "Buyers" Then
            con.Open()
            sql = "IF NOT EXISTS (SELECT * FROM Buyers WHERE ID = 'OTHER' " & _
                "INSERT INTO Buyers (ID, Description, Status, Orig_Date) " & _
                "SELECT 'OTHER','MISCELLANEOUS BUYER','Active','" & Date.Today & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
        End If
    End Sub

    Private Sub Process_Items(ByVal theFile)
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        Dim stopWatch As New Stopwatch
        stopWatch.Start()
        cnt = 0
        Dim isNumber As Boolean
        Dim aok As Boolean = True
        Dim records As Integer = 0
        Dim item, descr, vid, vendor, vitem, note, buyer, dept, clss, custom1, custom2, custom3, custom4, custom5, uom,
            sku, dim1, dim2, dim3, dim1_descr, dim2_descr, dim3_descr, status, sql, type As String
        Dim buyunit, sellunit As Int16
        Dim cost, retail As Decimal
        Dim thePath As String = ""
        txtprogress.Text = "Cleaning XML import file"
        Me.Refresh()
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim dt2 As DataTable = New DataTable
        dt2.Columns.Add("Sku")
        dt2.Columns.Add("Description")
        dt2.Columns.Add("Vendor_Id")
        dt2.Columns.Add("Vendor")
        dt2.Columns.Add("Vend_Item_No")
        dt2.Columns.Add("Dept")
        dt2.Columns.Add("Buyer")
        dt2.Columns.Add("Class")
        dt2.Columns.Add("Curr_Cost")
        dt2.Columns.Add("Curr_Retail")
        dt2.Columns.Add("Custom1")
        dt2.Columns.Add("Custom2")
        dt2.Columns.Add("Custom3")
        dt2.Columns.Add("Custom4")
        dt2.Columns.Add("Custom5")
        dt2.Columns.Add("UOM")
        dt2.Columns.Add("BuyUnit")
        dt2.Columns.Add("SellUnit")
        dt2.Columns.Add("Note")
        dt2.Columns.Add("Status")
        dt2.Columns.Add("Type")
        dt2.Columns.Add("Initial_Date", GetType(System.DateTime))
        dt2.Columns.Add("Last_Change_Date", GetType(System.DateTime))
        dt2.Columns.Add("Item")
        dt2.Columns.Add("DIM1")
        dt2.Columns.Add("DIM2")
        dt2.Columns.Add("DIM3")
        dt2.Columns.Add("DIM1_DESCR")
        dt2.Columns.Add("DIM2_DESCR")
        dt2.Columns.Add("DIM3_DESCR")
        Dim xmlFile As XmlReader
        thePath = path & theFile
        Dim rw As DataRow
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            For Each row In dt.Rows
                sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                If sku <> "TOTALS" Then
                    cnt += 1
                    If cnt Mod 1000 = 0 Then
                        txtprogress.Text = "Processed " & cnt & " records."
                        Me.Refresh()
                    End If

                    If InStr(sku, "~") > 0 Then
                        Dim parts() As String = sku.Split("~"c)
                        item = parts(0)
                        dim1 = parts(1)
                        dim2 = parts(2)
                        dim3 = parts(3)
                    Else
                        item = sku
                        dim1 = Nothing
                        dim2 = Nothing
                        dim3 = Nothing
                    End If
                    dim1_descr = Trim(Microsoft.VisualBasic.Left(row("DIM1_DESCR"), 30))
                    dim2_descr = Trim(Microsoft.VisualBasic.Left(row("DIM2_DESCR"), 30))
                    dim3_descr = Trim(Microsoft.VisualBasic.Left(row("DIM3_DESCR"), 30))
                    descr = Trim(Microsoft.VisualBasic.Left(row("DESCRIPTION"), 40))
                    vid = Trim(Microsoft.VisualBasic.Left(row("VENDOR_ID"), 20))
                    vendor = Trim(Microsoft.VisualBasic.Left(row("VENDOR"), 40))
                    vitem = Trim(Microsoft.VisualBasic.Left(row("VEND_ITEM_NO"), 20))
                    buyer = Trim(Microsoft.VisualBasic.Left(row("BUYER"), 10))
                    If buyer = "" Then buyer = "NA"
                    dept = Trim(Microsoft.VisualBasic.Left(row("DEPT"), 10))
                    clss = Trim(Microsoft.VisualBasic.Left(row("CLASS"), 10))
                    custom1 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM1"), 30))
                    custom2 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM2"), 30))
                    custom3 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM3"), 30))
                    custom4 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM4"), 30))
                    custom5 = Trim(Microsoft.VisualBasic.Left(row("CUSTOM5"), 30))
                    note = Trim(Microsoft.VisualBasic.Left(row("NOTE"), 20))
                    uom = Trim(Microsoft.VisualBasic.Left(row("UOM"), 10))
                    status = Trim(Microsoft.VisualBasic.Left(row("STATUS"), 10))
                    type = Trim(Microsoft.VisualBasic.Left(row("TYPE"), 1))
                    oTest = row("CURR_COST")
                    isNumber = IsNumeric(oTest)
                    If isNumber Then cost = CDec(oTest) Else cost = 0
                    oTest = row("CURR_RTL")
                    isNumber = IsNumeric(oTest)
                    If isNumber Then retail = CDec(oTest) Else retail = 0
                    uom = row("UOM")
                    If uom <> "" Then
                        oTest = row("BUYUNIT")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then buyunit = oTest
                        oTest = row("SELLUNIT")
                        isNumber = IsNumeric(oTest)
                        If isNumber Then sellunit = oTest
                    Else
                        uom = "EA"
                        buyunit = 1
                        sellunit = 1
                    End If
                    rw = dt2.NewRow
                    rw("Sku") = sku
                    rw("Description") = descr
                    rw("Vendor_Id") = vid
                    rw("Vendor") = vendor
                    rw("Vend_Item_No") = vitem
                    rw("Dept") = dept
                    rw("Buyer") = buyer
                    rw("Class") = clss
                    rw("Curr_Cost") = cost
                    rw("Curr_Retail") = retail
                    rw("Custom1") = custom1
                    rw("Custom2") = custom2
                    rw("Custom3") = custom3
                    rw("Custom4") = custom4
                    rw("Custom5") = custom5
                    rw("UOM") = uom
                    rw("BuyUnit") = buyunit
                    rw("SellUnit") = sellunit
                    rw("Note") = note
                    rw("Status") = status
                    rw("Type") = type
                    rw("Initial_Date") = Date.Today
                    rw("Last_Change_Date") = Date.Today
                    rw("Item") = item
                    rw("DIM1") = dim1
                    rw("DIM2") = dim2
                    rw("DIM3") = dim3
                    rw("DIM1_DESCR") = dim1_descr
                    rw("DIM2_DESCR") = dim2_descr
                    rw("DIM3_DESCR") = dim3_descr
                    dt2.Rows.Add(rw)
                Else
                    records = CInt(row("DESCRIPTION"))
                End If
            Next
        End If

        txtprogress.Text = "Inserting records into Item_Master"
        con.Open()
        sql = "DELETE FROM Item_Master"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        Dim connection As SqlConnection = New SqlConnection(constr)
        Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
        connection.Open()
        bulkCopy.DestinationTableName = "dbo.Item_Master"
        bulkCopy.BulkCopyTimeout = 960
        bulkCopy.WriteToServer(dt2)
        connection.Close()

        If cnt <> records Then
            aok = False
            message = "Record count " & cnt & " does not match XML Total Record " & records & "."
            Call Log_Error(message)
        End If
        records = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Item_Master"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            records = rdr("cnt")
        End While
        con.Close()
        If records <> cnt Then
            aok = False
            message = "Record count " & cnt & " does not match records counted in Item_Master table."
            Call Log_Error(message)
        End If
        If aok Then
            chkItems.Checked = True
        End If
        Dim m As String = "Update Item_Master with " & cnt & " Items"
        Call Update_Process_Log("1", "Update Item_Master", m, "")
        xmlFile.Close()

        txtItems.Text = Format(cnt, "0,0")
        txtprogress.Text = ""
        Me.Refresh()

    End Sub

    Private Sub Process_Barcodes(ByVal thefile)
        txtprogress.Text = "Processing Barcodes"
        stopwatch.Start()
        Dim cnt As Integer = 0
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Sku")
        dt.Columns.Add("Item")
        dt.Columns.Add("DIM1")
        dt.Columns.Add("DIM2")
        dt.Columns.Add("DIM3")
        dt.Columns.Add("Type")
        dt.Columns.Add("Barcode")
        dt.Columns.Add("Date")
        Dim thePath As String = ""
        thePath = path & thefile
        Dim xmlFile As XmlReader
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            cnt = dt.Rows.Count
            Dim connection As SqlConnection = New SqlConnection(conString)
            Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
            connection.Open()
            bulkCopy.DestinationTableName = "dbo.Item_Barcodes"
            bulkCopy.BulkCopyTimeout = 960
            bulkCopy.WriteToServer(dt)
            connection.Close()
        End If
        txtBarcodes.Text = cnt
        txtprogress.Text = ""
    End Sub

    Private Sub Process_Inventory(ByVal thefile)
        Dim stopWatch As New Stopwatch
        stopWatch.Start()
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        Dim thePath As String = ""
        thePath = path & thefile
        Dim records As Integer = 0
        Dim aok As Boolean = True
        cnt = 0
        Dim avail, onhand, committed As Decimal
        Dim tonhand As Decimal = 0
        Dim invOH As Decimal = 0
        Dim invCost As Decimal = 0
        Dim invRetail As Decimal = 0
        Dim sku, item, dim1, dim2, dim3, loc As String
        Dim cost, retail As Decimal
        Dim dte As Date = Date.Today
        Dim dttime As DateTime
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim inventorysDate As Date = DateAdd(DateInterval.Day, -6, BuildEndDate)
        tcost = 0
        tretail = 0
        tonhand = 0
        txtprogress.Text = "Inserting records into Item_Inv"
        Me.Refresh()



        ''GoTo 55



        con.Open()
        Dim xmlFile As XmlReader
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            For Each row In dt.Rows
                sku = Trim(row("SKU"))
                If sku <> "TOTALS" Then
                    cnt += 1
                    If cnt Mod 1000 = 0 Then
                        txtprogress.Text = "Processed " & cnt & " records."
                        Me.Refresh()
                    End If
                    If InStr(sku, "~") > 0 Then
                        Dim parts() As String = sku.Split("~"c)
                        item = parts(0)
                        dim1 = parts(1)
                        dim2 = parts(2)
                        dim3 = parts(3)
                    Else
                        item = sku
                        dim1 = Nothing
                        dim2 = Nothing
                        dim3 = Nothing
                    End If

                    loc = Trim(row("LOCATION"))
                    oTest = row("ONHAND")
                    If IsNumeric(oTest) Then onhand = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero) Else onhand = 0
                    oTest = row("COMMITTED")
                    If IsNumeric(oTest) Then committed = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero) Else committed = 0
                    oTest = row("AVAIL")
                    If IsNumeric(oTest) Then avail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero) Else avail = 0
                    oTest = row("COST")
                    If IsNumeric(oTest) Then cost = Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero) Else cost = 0
                    oTest = row("RETAIL")
                    If IsNumeric(oTest) Then retail = Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero) Else retail = 0
                    tonhand += onhand
                    tcost += (onhand * cost)
                    tretail += (onhand * retail)
                    oTest = row("EXTRACT_DATE")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            dttime = oTest
                        Else : dttime = DateAndTime.Now
                        End If
                    Else : dttime = DateAndTime.Now
                    End If
                    sql = "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Avail, End_OH, Committed, Cost, " & _
                        "Retail, YrWk, Item, DIM1, DIM2, DIM3) " & _
                        "SELECT '" & loc & "','" & sku & "','" & inventorysDate & "','" & BuildEndDate &
                           "'," & avail & "," & onhand & "," & committed & "," & cost & "," & retail & ", YrWk, '" & _
                        item & "','" & dim1 & "','" & dim2 & "','" & dim3 & "' FROM Calendar c " & _
                        "JOIN Item_Master m ON m.Sku = '" & sku & "' AND m.[Type] = 'I' " & _
                        "WHERE eDate = '" & BuildEndDate & "' AND Week_Id > 0"
                    cmd = New SqlCommand(sql, con)
                    cmd.CommandTimeout = 120
                    cmd.ExecuteNonQuery()
                Else
                    invOH = Decimal.Round(row("ONHAND"), 4, MidpointRounding.AwayFromZero)
                    invCost = Decimal.Round(row("COST"), 2, MidpointRounding.AwayFromZero)
                    invRetail = Decimal.Round(row("RETAIL"), 2, MidpointRounding.AwayFromZero)
                End If
            Next
        End If
        con.Close()

        If records <> cnt Then
            aok = False
            message = "Record count " & cnt & " does not match XML Total Record " & records & "."
            Call Log_Error(message)
        End If
        records = 0
        con.Open()
        sql = "SELECT COUNT(*) cnt FROM Item_Inv"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            records = rdr("cnt")
        End While
        con.Close()
        If records <> cnt Then
            aok = False
            message = "Record count " & Format(cnt, "###,###") & " does not match " &
                            Format(records, "###,###,###") & " records counted in Item_Inv table."
            Call Log_Error(message)
        End If
        If aok Then chkItems.Checked = True
        txtprogress.Text = "Select records from Transaction_Log"
        Me.Refresh()
55:
        ''con.Open()
        ''        sql = "UPDATE i SET i.Cost = Curr_Cost FROM Item_Inv i " & _
        ''            "JOIN Item_Master m ON m.Sku = i.Sku " & _
        ''            "WHERE ISNULL(Cost,0) = 0"
        ''        cmd = New SqlCommand(sql, con)
        ''  cmd.ExecuteNonQuery()

        ''sql = "UPDATE i SET i.Retail = Curr_Retail FROM Item_Inv i " & _
        ''    "JOIN Item_Master m ON m.Sku = i.Sku " & _
        ''    "WHERE ISNULL(Retail,0) = 0"
        ''cmd = New SqlCommand(sql, con)
        ''cmd.ExecuteNonQuery()
        ''con.Close()

        txtInventory.Text = Format(cnt, "0,0")
        txtInvCost.Text = Format(tcost, "0,0.00")
        txtInvRetail.Text = Format(tretail, "0,0.00")

        If tonhand <> invOH Then
            aok = False
            message = "Summed On Hand Quantity (" & Format(tonhand, "###,###") & ") is not equal XML Total Quantity (" &
                            Format(invOH, "###,###") & ")."
        End If
        If tcost <> invCost Then
            aok = False
            message = "Summed Cost (" & Format(tcost, "###,###,###") & ") is not equal XML Cost (" &
                            Format(invCost, "###,###,###") & ")."
            Call Log_Error(message)
        End If
        If tretail <> invRetail Then
            aok = False
            message = "Summed Retail (" & Format(tretail, "###,###,###") & ") is not equal XML Retail (" &
                            Format(invRetail, "###,###,###") & ")."
            Call Log_Error(message)
        End If
        con.Open()
        sql = "SELECT ISNULL(SUM(End_OH),0) OH, ISNULL(SUM(End_OH * Cost),0) Cost, ISNULL(SUM(End_OH * Retail),0) Retail " & _
            "FROM Item_Inv"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            invOH = CDec(rdr("OH"))
            invCost = CDec(rdr("Cost"))
            invRetail = CDec(rdr("Retail"))
        End While
        con.Close()

        If invOH <> tonhand Then
            aok = False
            message = "Summed XML OnHand (" & Format(tonhand, "###,###") & ") is not equal to Item_Inv.End_OH (" &
                            Format(invOH, "###,###,###") & ")."
            Call Log_Error(message)
        End If
        If Decimal.Round(invCost, 2, MidpointRounding.AwayFromZero) <> Decimal.Round(tcost, 2, MidpointRounding.AwayFromZero) Then
            aok = False
            message = "Summed XML Cost (" & Format(tcost, "###,###,###") & ") is not equal Item_Inv.Cost (" &
                            Format(invCost, "###,###,###") & ")."
            Call Log_Error(message)
        End If
        If Decimal.Round(invRetail, 2, MidpointRounding.AwayFromZero) <> Decimal.Round(tretail, 2, MidpointRounding.AwayFromZero) Then
            aok = False
            message = "Summed XML Retail (" & Format(tretail, "###,###,###") & ") is not equal Item_Inv.Retail (" &
                            Format(invRetail, "###,###,###") & ")."
            Call Log_Error(message)
        End If
        message = "Updated " & cnt & " records - Total Cost = " & tcost & " Total Retail = " & tretail & " "
        Call Update_Process_Log("1", "Update Inventory", message, "")
        txtprogress.Text = ""
        Me.Refresh()
    End Sub

    Private Sub Process_POs()
        Dim sql As String
        Dim cmd As SqlCommand
        Dim cnt As Integer = 0
        Dim store, po, vendor, buyer, stat, terms As String
        Dim ordDate, dueDate, canDate As Date
        Dim amt, recvd_cost, ord_qty, open_amt As Decimal
        Dim lines, open_lines As Integer
        Dim datetime As DateTime
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim xmlFile As XmlReader
        Dim row As DataRow
        Dim thePath As String = xmlPath & "\POHeader.xml"
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            con.Open()
            dt = ds.Tables(0)
            For Each row In dt.Rows
                cnt += 1
                If cnt Mod 1000 = 0 Then
                    txtprogress.Text = "Processed " & cnt & " records."
                End If
                oTest = row("LOCATION")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = CStr(oTest) Else store = "1"
                oTest = row("PO")
                If oTest <> "TOTALS" Then
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then po = CStr(oTest) Else po = "NA"
                    oTest = row("OrderDate")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then ordDate = CDate(oTest) Else ordDate = "1900-01-01"
                    oTest = row("DueDate")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then dueDate = CDate(oTest) Else dueDate = "1900-01-01"
                    oTest = row("CancelDate")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then canDate = CDate(oTest) Else canDate = "1900-01-01"
                    oTest = row("Vendor")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then vendor = CStr(oTest) Else vendor = "NA"
                    oTest = row("BUYER")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = CStr(oTest) Else buyer = "NA"
                    oTest = row("STATUS")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then stat = CStr(oTest) Else stat = "X"
                    oTest = row("AMT")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = CDec(oTest) Else amt = 0
                    oTest = row("RECVD_COST")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then recvd_cost = CDec(oTest) Else recvd_cost = 0
                    oTest = row("LINES")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lines = CInt(oTest) Else lines = 0
                    oTest = row("ORD_QTY")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then ord_qty = CDec(oTest) Else ord_qty = 0
                    oTest = row("OPEN_LINES")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then open_lines = CInt(oTest) Else open_lines = 0
                    oTest = row("OPEN_AMT")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then open_amt = CDec(oTest) Else open_amt = 0
                    oTest = row("TERMS")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then terms = CStr(oTest) Else terms = "NA"
                    oTest = row("EXTRACT_DATE")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then datetime = CDate(oTest) Else datetime = Nothing
                    sql = "IF NOT EXISTS (SELECT PO_NO FROM PO_Header WHERE Loc_Id = '" & store & "' AND PO_NO = '" & po & "') " & _
                        "INSERT INTO PO_Header (Loc_Id, PO_NO, Order_Date, Due_Date, Cancel_Date, Vendor_Id, Buyer, Amount, " & _
                        "Recvd_Cost, Lines, Ord_Qty, Open_Lines, Open_Amt, Terms, Status) " & _
                        "SELECT '" & store & "','" & po & "','" & ordDate & "','" & dueDate & "','" & canDate & "','" &
                        vendor & "','" & buyer & "'," & amt & "," & recvd_cost & "," & lines & "," & ord_qty & "," &
                        open_lines & "," & open_amt & ",'" & terms & "','" & stat & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                End If
            Next
            con.Close()
        End If
        txtPO.Text = Format(cnt, "###,###")
        txtprogress.Text = "Processing Purchase Order Detail records"
        Me.Refresh()
        cnt = 0
        Dim item, loc, lastDate As String
        Dim seq As Int32
        Dim cost, retail, ordQty, recvdQty, expQty As Decimal
        Dim tqty As Decimal = 0
        Dim tcost As Decimal = 0
        Dim tretail As Decimal = 0
        Dim poCnt As Integer = 0
        Dim poQty As Decimal = 0
        Dim poCost As Decimal = 0
        Dim poRetail As Decimal = 0
        Dim aok As Boolean = True
        Dim extractDate As DateTime
        ds = New DataSet
        dt = New DataTable
        thePath = xmlPath & "\PODetail.xml"
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            con.Open()
            dt = ds.Tables(0)
            For Each row In dt.Rows
                oTest = row("SKU")
                If Not IsNothing(oTest) Then item = CStr(oTest) Else item = "UNKNOWN"
                If item <> "TOTALS" Then
                    cnt += 1
                    If cnt Mod 1000 = 0 Then
                        txtprogress.Text = "processed " & cnt
                        Me.Refresh()
                    End If
                    oTest = row("LOCATION")
                    If Not IsNothing(oTest) Then loc = CStr(oTest) Else loc = "NA"
                    oTest = row("PO_NO")
                    If Not IsNothing(oTest) Then po = oTest Else po = "UNKNOWN"
                    oTest = row("SEQ_NO").ToString
                    If IsNumeric(oTest) Then seq = CInt(oTest) Else seq = 0
                    oTest = row("ORD_QTY")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        ordQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    Else : ordQty = 0
                    End If
                    tqty += ordQty
                    oTest = row("QTY_RECVD")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        recvdQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    Else : recvdQty = 0
                    End If
                    oTest = row("EXP_QTY")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        expQty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    Else : expQty = 0
                    End If
                    oTest = row("LAST_RECVD_DATE")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then lastDate = CDate(oTest) Else lastDate = ""
                    oTest = row("COST")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        cost = Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero)
                    Else : cost = 0
                    End If
                    tcost += (cost * ordQty)
                    oTest = row("RETAIL")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        retail = Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero)
                    Else : retail = 0
                    End If
                    tretail += (retail * ordQty)
                    extractDate = row("EXTRACT_DATE")
                    ''sql = "IF NOT EXISTS (SELECT PO_NO FROM PO_Detail WHERE Loc_Id = '" & loc & "' AND PO_NO = '" & po & "' " & _
                    ''    "AND Seq_No = " & seq & " AND Sku = '" & item & "') " & _
                    ''    "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, Cost, Retail, Last_Recvd_Date, " & _
                    ''    "Item, Dim1, Dim2, Dim3) " & _
                    ''    "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & ordQty & "," & recvdQty & "," & expQty & "," &
                    ''    cost & "," & retail & ",'" & lastDate & "', Item, Dim1, Dim2, Dim3 FROM Item_Master WHERE Sku = '" & item & "'"
                    sql = "INSERT INTO PO_Detail (Loc_Id, PO_NO, Seq_No, Sku, Qty_Ordered, Qty_Recvd, Qty_Due, Cost, Retail, Last_Recvd_Date, " & _
                       "Item, Dim1, Dim2, Dim3) " & _
                       "SELECT '" & loc & "','" & po & "'," & seq & ",'" & item & "'," & ordQty & "," & recvdQty & "," & expQty & "," &
                       cost & "," & retail & ",'" & lastDate & "', Item, Dim1, Dim2, Dim3 FROM Item_Master WHERE Sku = '" & item & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                Else
                    poQty = CDec(row("ORD_QTY"))
                    poCost = CDec(row("COST"))
                    poRetail = CDec(row("RETAIL"))

                End If
                ''End If
            Next
            con.Close()
        End If
        txtPOCost.Text = Format(tcost, "###,###,###.00")
        txtPORetail.Text = Format(tretail, "###,###,###.00")
        Me.Refresh()
        message = "Updated " & txtPO.Text & "POs  - Total Cost = " & txtPOCost.Text & " Total Retail = " & txtPORetail.Text & " "
        Call Update_Process_Log("1", "Update Purchase Orders", message, "")
        If tqty <> poQty Then
            aok = False
            message = "Summed Quantity Due (" & Format(tqty, "###,###") & ") is not equal XML Total Quantity (" &
                            Format(poQty, "###,###") & ")."
        End If
        If tcost <> poCost Then
            aok = False
            message = "Summed Cost (" & Format(tcost, "###,###,###") & ") is not equal XML Cost (" &
                            Format(poCost, "###,###,###") & ")."
        End If
        If tretail <> poRetail Then
            aok = False
            message = "Summed Retail (" & Format(tretail, "###,###,###") & ") is not equal XML Retail (" &
                            Format(poRetail, "###,###,###") & ")."
        End If
        con.Open()
        sql = "SELECT COUNT(*) Cnt, ISNULL(SUM(Qty_Ordered),0) Qty, ISNULL(SUM(Cost * Qty_Ordered),0) Cost, " & _
            "ISNULL(SUM(Retail * Qty_Ordered),0) Retail FROM PO_Detail"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            poCnt = CInt(rdr("Cnt"))
            poQty = CDec(rdr("Qty"))
            poCost = CDec(rdr("Cost"))
            poRetail = CDec(rdr("Retail"))
        End While
        con.Close()

        If cnt <> poCnt Then
            aok = False
            message = "Record count in XML (" & Format(cnt, "###,###") & ") is not equal Record Count in PO_Detail (" &
                            Format(poCnt, "###,###") & ")."
        End If
        If poQty <> tqty Then
            aok = False
            message = "Xml Quantity Due (" & Format(tqty, "###,###") & ") is not equal to PO_Detail Qty Due (" &
                            Format(poQty, "###,###") & ")."
        End If
        If poCost <> tcost Then
            aok = False
            message = "XML Total Cost (" & Format(tcost, "###,###,###") & ") is not equal PO_Detail Total Cost (" &
                            Format(poCost, "###,###,###") & ")."


        End If
        If tretail <> poRetail Then
            aok = False
            message = "XML Total Retail (" & Format(tretail, "###,###,###") & ") is not equal PO_Detail Total Retail (" &
                            Format(poRetail, "###,###,###") & "),"
        End If
    End Sub

    Private Sub Process_Preqs()
        Dim cnt As Integer = 0
        Dim preq, batch, vend_no, vend, loc, buyer, alloc, mrg As String
        Dim ord, del, can, extract As Date
        Dim amt, totlCost As Decimal
        totlCost = 0

        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim thePath As String = xmlPath & "\Purchase_Request_Header.xml"
        Dim xmlFile As XmlReader
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            con.Open()
            con2.Open()
            sql = "DELETE FROM Purchase_Request_Header"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            dt = ds.Tables(0)
            For Each row In dt.Rows
                oTest = row("VEND_NO")
                If oTest <> "TOTALS" Then
                    cnt += 1
                    If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                    oTest = row("PREQ_NO")
                    If Not IsDBNull(oTest) Then preq = CStr(oTest) Else preq = ""
                    oTest = row("BATCH")
                    If Not IsDBNull(oTest) Then batch = CStr(oTest)
                    oTest = row("VEND_NO")
                    If Not IsDBNull(oTest) Then vend_no = CStr(oTest)
                    oTest = row("VENDOR")
                    If Not IsDBNull(oTest) Then vend = CStr(oTest)
                    oTest = row("LOC_ID")
                    If IsDBNull(oTest) Then oTest = "UNKNOWN"
                    loc = CStr(oTest)
                    oTest = row("BUYER")
                    If IsDBNull(oTest) Then buyer = "UNKNOWN"
                    buyer = CStr(oTest)
                    oTest = row("ORDER_DATE")
                    If IsDBNull(oTest) Then oTest = 0
                    ord = CDate(oTest)
                    oTest = row("DELIVER_DATE")
                    If IsDBNull(oTest) Then oTest = 0
                    del = CDate(oTest)
                    oTest = row("CANCEL_DATE")
                    If IsDBNull(oTest) Then oTest = 0
                    can = CDate(oTest)
                    oTest = row("ORDER_TOTAL")
                    If IsDBNull(oTest) Then amt = 0 Else amt = CDec(oTest)
                    totlCost += amt
                    oTest = row("ISALLOCATED")
                    If IsDBNull(oTest) Then oTest = "N"
                    alloc = CStr(oTest)
                    oTest = row("MERGED")
                    If IsDBNull(oTest) Then oTest = "S"
                    mrg = CStr(oTest)
                    oTest = row("EXTRACT_DATE")
                    If IsDBNull(oTest) Then oTest = 0
                    extract = CDate(oTest)
                    sql = "INSERT INTO Purchase_Request_Header(PREQ_NO, Batch_Id, Loc_Id, " & _
                        "Vendor_Id, Buyer, Order_Date, Deliver_Date, Cancel_Date, Order_Total, " & _
                        "Allocated, Merged, Last_Update) " & _
                        "SELECT '" & preq & "','" & batch & "','" & loc & "','" & vend_no & "','" & _
                        buyer & "','" & ord & "','" & del & "','" & can & "'," & totlCost & ",'" & _
                        alloc & "','" & mrg & "','" & Date.Today & "'"
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                Else
                    oTest = row("ORDER_TOTAL")
                    If IsDBNull(oTest) Then amt = 0 Else amt = CDec(oTest)
                    txtPReqCost.Text = Format(amt, "###,###,###.00")
                    Me.Refresh()
                End If
            Next
            con.Close()
            con2.Close()
        End If
        '
        '   Purchase Request Detail
        '
        Dim item, desc, uom, dim1, dim2, dim3, sku As String
        Dim seq, numer, denom As Integer
        Dim cost, qty, totlQty As Decimal
        Dim aok As Boolean = True
        totlCost = 0
        totlQty = 0
        ds = New DataSet
        dt = New DataTable
        thePath = xmlPath & "\Purchase_Request_Detail.xml"
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            con.Open()
            con2.Open()
            sql = "DELETE FROM Purchase_Request_Detail"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            dt = ds.Tables(0)
            For Each row As DataRow In dt.Rows
                oTest = row("SKU")
                If oTest <> "TOTALS" Then
                    cnt += 1
                    If cnt Mod 1000 = 0 Then Console.WriteLine(cnt)
                    oTest = row("PREQ_NO")
                    If Not IsDBNull(oTest) Then preq = CStr(oTest) Else preq = ""
                    oTest = row("SEQ_NO")
                    If IsDBNull(oTest) Or IsNothing(oTest) Then seq = 0
                    If IsNumeric(oTest) Then seq = CInt(oTest)
                    sku = Trim(Microsoft.VisualBasic.Left(row("SKU"), 90))
                    If InStr(sku, "~") > 0 Then
                        Dim parts() As String = sku.Split("~"c)
                        item = parts(0)
                        dim1 = parts(1)
                        dim2 = parts(2)
                        dim3 = parts(3)
                    Else
                        item = sku
                        dim1 = Nothing
                        dim2 = Nothing
                        dim3 = Nothing
                    End If

                    oTest = row("DESC")
                    If Not IsDBNull(oTest) Then desc = CStr(oTest) Else desc = ""
                    oTest = row("UOM")
                    If Not IsDBNull(oTest) Then uom = CStr(oTest) Else uom = "EACH"
                    oTest = row("NUMER")
                    If IsDBNull(oTest) Or Not IsNumeric(oTest) Then numer = 1 Else numer = CInt(oTest)
                    oTest = row("DENOM")
                    If IsDBNull(oTest) Or Not IsNothing(oTest) Then denom = 1 Else denom = CInt(oTest)
                    oTest = row("COST")
                    If IsDBNull(oTest) Then cost = 0 Else cost = Decimal.Round(oTest, 2, MidpointRounding.AwayFromZero)
                    oTest = row("QTY")
                    If IsDBNull(oTest) Or IsNothing(oTest) Then oTest = 0
                    qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                    totlQty += qty
                    totlCost += (qty * cost)
                    sql = "INSERT INTO Purchase_Request_Detail(PREQ_NO, Seq_No, Sku, Description, " & _
                        "Stock_Units, Numer, Denom, Cost, Item, DIM1, DIM2, DIM3, Qty) " & _
                        "SELECT '" & preq & "'," & seq & ",'" & sku & "','" & desc & "','" & _
                        uom & "'," & numer & "," & denom & "," & cost & ",'" & item & "','" & _
                        dim1 & "','" & dim2 & "','" & dim3 & "'," & qty & " "
                    cmd = New SqlCommand(sql, con2)
                    cmd.ExecuteNonQuery()
                Else
                    qty = row("QTY")
                    cost = row("COST")
                End If
            Next
            con.Close()
            con2.Close()
            message = "Updated " & txtPReq.Text & "POs  - Total Cost = " & txtPReqCost.Text & " "
            Call Update_Process_Log("1", "Update Purchase Requests", message, "")
            If totlCost <> cost Then
                aok = False
                message = "Summed Cost (" & Format(totlCost, "###,###") & ") is not equal XML Total Cost (" &
                                Format(cost, "###,###") & ")."
                Call Log_Error(message)
            End If
            If totlQty <> qty Then
                aok = False
                message = "Summed Quantity (" & Format(totlQty, "###,###") & ") is not equal XML Total Quantity (" &
                                Format(qty, "###,###") & ")."
                Call Log_Error(message)
            End If
            con.Open()
            sql = "SELECT ISNULL(SUM(Qty),0) Qty, ISNULL(SUM(Cost * Qty),0) Cost FROM Purchase_Request_Detail"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                qty = CDec(rdr("Qty"))
                cost = CDec(rdr("Cost"))
            End While
            con.Close()
            If qty <> totlQty Then
                aok = False
                message = "Summed Quantity (" & Format(totlQty, "###,###") & ") is not equal Purchase_Request Detail.Qty (" &
                    Format(qty, "###,###") & ")."
                Call Log_Error(message)
            End If
            If cost <> totlCost Then
                aok = False
                message = "Summed Extended Cost (" & Format(totlCost, "###,###") & ") is not equal Purchase_Request Detail Extended Cost (" &
                    Format(cost, "###,###") & ")."
                Call Log_Error(message)
            End If
        End If
    End Sub

    Private Sub Process_Data(ByVal thefile, ByVal theField)                              ' Add new records to Daily Transaction Log
        txtprogress.Text = "Processing " & thefile
        Me.Refresh()
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        Dim stopWatch As New Stopwatch
        stopWatch.Start()
        cnt = 0
        tcost = 0
        tretail = 0
        Dim id, sku, item, dim1, dim2, dim3, loc, store, dept, buyer, clss, sql, cust, tkt, reason, coupon,
            ordtype, whsl As String
        Dim tktDate As Nullable(Of DateTime) = Nothing
        Dim seq As Integer
        Dim qty As Decimal = 0
        Dim tqty As Decimal = 0
        Dim xmlQty As Decimal = 0
        Dim xmlCost As Decimal = 0
        Dim xmlRetail As Decimal = 0
        Dim grandQty As Decimal = 0
        Dim grandCost As Decimal = 0
        Dim grandRetail As Decimal = 0
        Dim grandMarkdown As Decimal = 0
        Dim aok As Boolean = True
        Dim cost, retail, mkdn As Decimal
        Dim dte As DateTime
        Dim ordDate As Nullable(Of Date) = Nothing
        Dim expDate As Nullable(Of Date) = Nothing
        Dim extractDate, datetimenow As DateTime
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim xmlFile As XmlReader
        Dim thePath As String = ""
        thePath = xmlPath & thefile
        If theField = "Sold" Then thePath = thefile ' use the file path name from the OpenFileDialog for sales
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            For Each row In dt.Rows
                If row("TRANS_ID") <> "SALES" And row("SKU") <> "TOTALS" Then
                    qty = 0
                    cnt += 1
                    If cnt Mod 1000 = 0 Then
                        txtprogress.Text = "Processed " & cnt & " records."
                        Me.Refresh()
                    End If
                    oTest = row("TRANS_DATE")
                    If oTest = "0" Then oTest = "1900-01-01"
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then
                        dte = CDate(oTest)
                        oTest = row("TRANS_ID")
                        id = oTest
                        If theField = "ADJ" Then
                            If InStr(thePath, "Physical") > 0 Then id = "P" & id Else id = "A" & id
                        End If
                        oTest = row("SEQ_NO").ToString
                        seq = CInt(oTest)
                        oTest = row("SKU")
                        If Not IsDBNull(oTest) Then sku = Trim(Microsoft.VisualBasic.Left(oTest, 90)) Else sku = ""
                        If InStr(sku, "~") > 0 Then
                            Dim parts() As String = sku.Split("~"c)
                            item = parts(0)
                            dim1 = parts(1)
                            dim2 = parts(2)
                            dim3 = parts(3)
                        Else
                            item = sku
                            dim1 = Nothing
                            dim2 = Nothing
                            dim3 = Nothing
                        End If
                        oTest = row("LOCATION")
                        If Not IsDBNull(oTest) Then loc = Trim(Microsoft.VisualBasic.Left(oTest, 30))
                        oTest = row("QTY")
                        qty = 0
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then qty = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                        End If
                        oTest = row("COST")
                        cost = 0
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then cost = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                        End If
                        oTest = row("RETAIL")
                        retail = 0
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then retail = Decimal.Round(oTest, 4, MidpointRounding.AwayFromZero)
                        End If
                        If theField = "Sold" Then
                            oTest = Trim(Microsoft.VisualBasic.Left(row("STR_ID"), 30))
                            If IsDBNull(oTest) Then store = "UNKNOWN" Else store = CStr(oTest)
                            oTest = Trim(Microsoft.VisualBasic.Left(row("DEPT"), 30))
                            If IsDBNull(oTest) Then dept = "UNKNOWN" Else dept = CStr(oTest)
                            oTest = Trim(Microsoft.VisualBasic.Left(row("BUYER"), 30))
                            If IsDBNull(oTest) Then buyer = "UNKNOWN" Else buyer = CStr(oTest)
                            oTest = Trim(Microsoft.VisualBasic.Left(row("CLASS"), 30))
                            If IsDBNull(oTest) Then clss = "UNKNOWN" Else clss = CStr(oTest)
                            oTest = RTrim(Microsoft.VisualBasic.Left(row("CUST_NO"), 30))
                            If IsDBNull(oTest) Then cust = "UNKNOWN" Else cust = CStr(oTest)
                            oTest = row("TKT_NO")
                            If Not IsDBNull(oTest) Then tkt = oTest
                            oTest = row("TICKET_DATE")
                            If oTest <> "" Then tktDate = CDate(oTest) Else tktDate = "1900-01-01"
                            oTest = row("MARKDOWN")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then mkdn = CDec(oTest) Else mkdn = 0
                            oTest = Trim(Microsoft.VisualBasic.Left(row("MKDN_REASON"), 30))
                            If Not IsDBNull(oTest) Then reason = CStr(oTest)
                            oTest = Trim(Microsoft.VisualBasic.Left(row("COUPON_CODE"), 30))
                            If Not IsDBNull(oTest) Then coupon = CStr(oTest)
                            oTest = row("ORD_TYPE")
                            If Not IsDBNull(oTest) Then ordtype = CStr(oTest)
                            oTest = row("WHOLESALE")
                            If Not IsDBNull(oTest) Then whsl = CStr(oTest)
                        Else
                            mkdn = 0
                            store = ""
                            dept = ""
                            buyer = ""
                            clss = ""
                            reason = ""
                            coupon = ""
                            cust = ""
                            tkt = ""
                            tktDate = Nothing
                            ordtype = ""
                            whsl = ""
                        End If

                        oTest = row("EXTRACT_DATE")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then extractDate = oTest
                        tqty += qty
                        tcost += qty * cost
                        tretail += qty * retail
                        tmkdn += mkdn
                        datetimenow = DateAndTime.Now
                        con.Open()
                        sql = "IF NOT EXISTS (SELECT TRANS_ID FROM Daily_Transaction_Log WHERE LOCATION = '" & loc & "' " & _
                            "AND TRANS_ID = '" & id & "' AND SEQUENCE_NO = '" & seq & "' AND SKU = '" & sku & "' " & _
                            "AND TRANS_DATE = '" & dte & "') " & _
                            "INSERT INTO Daily_Transaction_Log (TRANS_ID, SEQUENCE_NO, STORE, SKU, LOCATION, QTY, COST, RETAIL, " & _
                                "MKDN, TRANS_DATE, [TYPE], POST_DATE, DEPT,BUYER, CLASS, EXTRACT_DATE, CUST_NO, TKT_NO, TKT_DATE, " & _
                                "MKDN_REASON, COUPON_CODE, ITEM, DIM1, DIM2, DIM3, ORD_TYPE, WHOLESALE) " & _
                            "SELECT '" & id & "', " & seq & ", '" & store & "', '" & sku & "','" & loc & "'," &
                                qty & "," & cost & "," & retail & ", " & mkdn & ",'" & dte & "','" & theField & "','" & datetimenow & "','" & _
                                dept & "','" & buyer & "','" & clss & "', '" & extractDate & "','" & cust & "','" & tkt & "','" & _
                                tktDate & "','" & reason & "','" & coupon & "','" & item & "','" & dim1 & "','" & dim2 & "','" &
                                dim3 & "','" & ordtype & "','" & whsl & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.CommandTimeout = 120
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Else
                    xmlQty = CDec(row("QTY"))
                    xmlCost = CDec(row("COST"))
                    xmlRetail = CDec(row("RETAIL"))
                    If theField = "Sold" Then
                        oTest = row("TKT_NO")
                        If IsNumeric(oTest) Then grandQty = CDec(oTest) Else grandQty = 0
                        oTest = row("MKDN_REASON")
                        If IsNumeric(oTest) Then grandCost = CDec(oTest) Else grandCost = 0
                        oTest = row("COUPON_CODE")
                        If IsNumeric(oTest) Then grandRetail = CDec(oTest) Else grandRetail = 0
                        oTest = row("ORD_TYPE")
                        If IsNumeric(oTest) Then grandMarkdown = CDec(oTest) Else grandMarkdown = 0
                    End If
                End If
            Next
        End If
        Dim isPhysical As Boolean = False
        Select Case theField
            Case "ADJ"
                If InStr(thePath, "Physical") > 0 Then
                    isPhysical = True
                    txtPhysical.Text = Format(cnt, "0,0")
                    txtPhysCost.Text = Format(tcost, "0,0.00")
                    txtPhysRetail.Text = Format(tretail, "0,0.00")
                Else
                    txtAdjustments.Text = Format(cnt, "0,0")
                    txtAdjCost.Text = Format(tcost, "0,0.00")
                    txtAdjRetail.Text = Format(tretail, "0,0.00")
                End If
            Case "RECVD"
                txtReceipts.Text = Format(cnt, "0,0")
                txtRecvCost.Text = Format(tcost, "0,0.00")
                txtRecvRetail.Text = Format(tretail, "0,0.00")
            Case "RTV"
                txtReturns.Text = Format(cnt, "0,0")
                txtRtnCost.Text = Format(tcost, "0,0.00")
                txtRtnRetail.Text = Format(tretail, "0,0.00")
            Case "Sold"
                txtSales.Text = Format(cnt, "0,0")
                txtSalesCost.Text = Format(tcost, "0,0.00")
                txtSalesRetail.Text = Format(tretail, "0,0.00")
            Case "XFER"
                txtTransfers.Text = Format(cnt, "0,0")
                txtXferCost.Text = Format(tcost, "0,0.00")
                txtXferRetail.Text = Format(tretail, "0,0.00")
        End Select
        Me.Refresh()

        If isPhysical Then theField = "PHYS"
        Dim message As String = "Updated " & cnt & " records - Qty = " & tqty &
            " Total Cost = " & tcost & " Total Retail = " & tretail & " Total Markdowns = " & tmkdn
        Call Update_Process_Log("1", "Update " & theField & "", message, "")
        If tqty <> xmlQty Then
            aok = False
            message = "Quantity summed (" & tqty & ") does not equal Total Qty in XML file (" & xmlQty & ")."
            Call Log_Error(message)
        End If
        If tcost <> xmlCost Then
            aok = False
            message = "Cost summed (" & tcost & ") does not equal Total Cost inc XML file (" & xmlCost & ")."
            Call Log_Error(message)
        End If
        If tretail <> xmlRetail Then
            aok = False
            message = "Retail summed (" & tretail & ") does not equal Total Retail in XML file (" & xmlRetail & ")."
            Call Log_Error(message)
        End If
        If aok Then
            Select Case theField
                Case "ADJ"
                    chkAdj.Checked = True
                Case "PHYS"
                    chkPhys.Checked = True
                Case "RECVD"
                    chkRecv.Checked = True
                Case "RTV"
                    chkRTN.Checked = True
                Case "Sold"
                    chkSales.Checked = True
                Case "XFER"
                    chkXFER.Checked = True
            End Select
        End If
        txtprogress.Text = ""
        Me.Refresh()
    End Sub

    Private Function Check_Log(ByVal con2, ByVal id, ByVal seq, ByVal location, ByVal item, ByVal dte)
        con2.open()
        Dim sql As String = "SELECT TRANS_ID FROM Daily_Transaction_Log WHERE TRANS_ID = '" & id & "' AND SEQ_NO = " & seq & " AND LOCATION = '" &
            location & "' AND ITEM_NO = '" & item & "' AND DATE = '" & dte & "'"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If Not IsNothing(rdr(0)) And Not IsDBNull(rdr(0)) Then
                con2.close()
                Return 1
            End If
        End While
        con2.close()
        Return 0
    End Function

    Private Function CleanData(ByVal strng)
        Dim cleanedData As String
        cleanedData = Regex.Replace(strng, "[^A-Za-z0-9\-/ -&.$]", "")
        Return cleanedData
    End Function

    Private Sub btnAddWeeklyRecords_Click(sender As Object, e As EventArgs) Handles btnAddWeeklyRecords.Click
        Try
            Call Connect_Database()
            txtprogress.Text = "Creating Inventory records for weeks with inventory ON HAND. "
            Me.Refresh()

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
            Dim prevMax As Decimal
            Dim prevCost, prevRetail As Decimal
            Dim oh, boh, max, adj, xfer, rtv, recvd, sold, test, thisBeginOH, thisEndOH, prevBegin, prevEnd As Decimal
            Dim prevLeadTime, leadTime, yrwk As Integer
            stopwatch.Start()
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' CLEAR OUT EVERYTHING EXCEPT THE LATEST INVENTORY AND REBUILD Item_Inv FROM THE 2 TRANSACTION LOGS
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            con.Open()
            sql = "update Item_Inv set adj=null, rtv=null, recvd=null, xfer=null, sold=null, begin_oh=null, max_oh=null " &
                "IF OBJECT_ID('tempDB.dbo.#xfer','U') IS NOT NULL DROP TABLE #xfer; " &
                "IF OBJECT_ID('tempDB.dbo.#adj','U') IS NOT NULL DROP TABLE #adj; " &
                "IF OBJECT_ID('tempDB.dbo.#rtv','U') IS NOT NULL DROP TABLE #rtv; " &
                "IF OBJECT_ID('tempDB.dbo.#rcvd','U') IS NOT NULL DROP TABLE #rcvd; " &
                "IF OBJECT_ID('tempDB.dbo.#sold','U') IS NOT NULL DROP TABLE #sold; " &
                "IF OBJECT_ID('tempDB.dbo.#inv','U') IS NOT NULL DROP TABLE #inv; " &
                "DECLARE @fromDate date = '" & BuildStartDate & "' " &
                "DECLARE @thruDate date = '" & BuildEndDate & "' " &
                "select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) XFER, " &
                    "item, dim1, dim2, dim3 into #xfer from Inv_Transaction_Log l " &
                "Join Calendar c On convert(Date, trans_date) between c.sdate And c.edate And c.week_id > 0 " &
                "where [Type] ='xfer' and convert(date,trans_date) between @fromDate and @thruDate " &
                "Group by location, sku, sdate, edate, yrwk, item, dim1, dim2, dim3 " &
                "merge item_inv t using #xfer s on s.location=t.loc_id And s.sku=t.sku And s.edate=t.edate " &
                "when Not matched by target then insert(loc_id, sku, sdate, edate, cost, retail, xfer, item, dim1, dim2, dim3, yrwk) " &
                "values(s.location, s.sku, s.sdate, s.edate, s.cost, s.retail, s.xfer, s.item, s.dim1, s.dim2, s.dim3, s.yrwk) " &
                "when matched then update set t.xfer=s.xfer; " &
                "Select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) ADJ, " &
                    "item, dim1, dim2, dim3 into #adj from Inv_Transaction_Log l " &
                "Join Calendar c On convert(Date, trans_date) between c.sdate And c.edate And c.week_id > 0 " &
                "where [Type] ='adj' and convert(date,trans_date) between @fromDate and @thruDate " &
                "Group by location, sku, sdate, edate, yrwk, item, dim1, dim2, dim3 " &
                "merge item_inv t using #adj s on s.location=t.loc_id And s.sku=t.sku And s.edate=t.edate " &
                "when Not matched by target then insert(loc_id, sku, sdate, edate, cost, retail, adj, item, dim1, dim2, dim3, yrwk) " &
                "values(s.location, s.sku, s.sdate, s.edate, s.cost, s.retail, s.adj, s.item, s.dim1, s.dim2, s.dim3, s.yrwk) " &
                "when matched then update set t.adj=s.adj; " &
                "Select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) RTV, " &
                    "item, dim1, dim2, dim3 into #rtv from Inv_Transaction_Log l " &
                "Join Calendar c On convert(Date, trans_date) between c.sdate And c.edate And c.week_id > 0 " &
                "where [Type] ='rtv' and convert(date,trans_date) between @fromDate and @thruDate " &
                "Group by location, sku, sdate, edate, yrwk, item, dim1, dim2, dim3 " &
                "merge item_inv t using #rtv s on s.location=t.loc_id And s.sku=t.sku And s.edate=t.edate " &
                "when Not matched by target then insert(loc_id, sku, sdate, edate, cost, retail, rtv, item, dim1, dim2, dim3, yrwk) " &
                "values(s.location, s.sku, s.sdate, s.edate, s.cost, s.retail, s.rtv, s.item, s.dim1, s.dim2, s.dim3, s.yrwk) " &
                "when matched then update set t.rtv=s.rtv; " &
                "Select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) RECVD, " &
                    "item, dim1, dim2, dim3 into #rcvd from Inv_Transaction_Log l " &
                "Join Calendar c On convert(Date, trans_date) between c.sdate And c.edate And c.week_id > 0 " &
                "where [Type] ='recvd' and convert(date,trans_date) between @fromDate and @thruDate " &
                "Group by location, sku, sdate, edate, yrwk, item, dim1, dim2, dim3 " &
                "merge item_inv t using #rcvd s on s.location=t.loc_id And s.sku=t.sku And s.edate=t.edate " &
                "when Not matched by target then insert(loc_id, sku, sdate, edate, cost, retail, recvd, item, dim1, dim2, dim3, yrwk) " &
                "values(s.location, s.sku, s.sdate, s.edate, s.cost, s.retail, s.recvd, s.item, s.dim1, s.dim2, s.dim3, s.yrwk) " &
                "when matched then update set t.recvd=s.recvd; " &
                "Select location, sku, sdate, edate, avg(cost) cost, avg(retail) retail, YrWk, sum(qty) SOLD, " &
                    "item, dim1, dim2, dim3 into #sold from Daily_Transaction_Log l " &
                "Join Calendar c On convert(Date, trans_date) between c.sdate And c.edate And c.week_id > 0 " &
                "where [Type] ='Sold' and sale_type='t' and convert(date,trans_date) between @fromDate and @thruDate " &
                "Group by location, sku, sdate, edate, yrwk, item, dim1, dim2, dim3 " &
                "merge item_inv t using #sold s on s.location=t.loc_id And s.sku=t.sku And s.edate=t.edate " &
                "when Not matched by target then insert(loc_id, sku, sdate, edate, cost, retail, sold, item, dim1, dim2, dim3, yrwk) " &
                "values(s.location, s.sku, s.sdate, s.edate, s.cost, s.retail, s.sold, s.item, s.dim1, s.dim2, s.dim3, s.yrwk) " &
                "when matched then update set t.sold=s.sold; "
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            cmd.ExecuteNonQuery()
            con.Close()

60:

            ''              Get rid of records with no Inventory and no Transactions
            con.Open()
            sql = "DELETE FROM Item_Inv WHERE ISNULL(XFER,0)=0 AND ISNULL(ADJ,0)=0 AND ISNULL(RTV,0)=0 AND ISNULL(RECVD,0)=0 " &
                "AND ISNULL(Sold,0)=0 AND ISNULL(End_OH,0)=0 " &
                "UPDATE Item_Inv SET Begin_OH = NULL, End_OH = NULL, Max_OH = NULL WHERE eDate <> '" & BuildEndDate & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            con2.Open()
            sql = "Select i.Loc_Id, i.Sku, i.sDate, i.eDate, ISNULL(Cost,0) As Cost, ISNULL(Retail,0) As Retail, c.YrWk, " &
                "i.Sold, ISNULL(ADJ,0) As ADJ, ISNULL(XFER,0) As XFER, ISNULL(RTV,0) As RTV, ISNULL(RECVD,0) As RECVD, " &
                "ISNULL(Begin_OH,0) As Begin_OH, ISNULL(End_OH,0) As End_OH, ISNULL(Max_OH,0) As Max_OH, " &
                "Lead_Time, i.Item, i.DIM1, i.DIM2, i.DIM3 FROM Item_Inv i " &
                "JOIN Item_Master m On m.Sku = i.Sku " &
                "JOIN Calendar c On c.eDate = i.eDate And c.Week_Id > 0 " &
                "WHERE Type = 'I' AND i.eDate BETWEEN '" & BuildStartDate & "' AND '" & BuildEndDate & "' " &
                "ORDER BY i.Loc_Id, i.Sku, eDate DESC"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            rdr = cmd.ExecuteReader
            While rdr.Read
                location = rdr("Loc_Id")
                sku = rdr("Sku")
                cnt += 1
                If cnt Mod 1000 = 0 Then
                    txtprogress.Text = cnt & " " & location & " " & sku
                    Me.Refresh()
                End If
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
                    GoTo BF100
                End If
                '
                '          Different sku below
                '
                If sku <> prevSku Then
                    ''If prevBegin <> 0 Then
                    If thisBeginOH > 0 Then
                        nextEdate = DateAdd(DateInterval.Day, -7, preveDate)
                        nextSdate = DateAdd(DateInterval.Day, -6, nextEdate)
                        Do While BuildStartDate <= nextEdate
                            prevMax = prevBegin
                            If prevMax < 0 Then prevMax = 0
                            sql = "IF NOT EXISTS (SELECT * FROM Item_Inv WHERE Loc_Id = '" & location & "' AND Sku = '" & prevSku & "' " &
                                "AND sDate = '" & nextSdate & "' AND eDate = '" & nextEdate & "' AND Item = '" & prevItem & "') " &
                                "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, Begin_OH, End_OH, Max_OH, " &
                                "Lead_Time, Item, DIM1, DIM2, DIM3) " &
                                "SELECT '" & location & "', '" & prevSku & "', '" & nextSdate & "', '" & nextEdate & "', " & prevCost & ", " &
                                prevRetail & ", YrWk, " & prevBegin & ", " & prevBegin & ", " & prevMax & ", " & prevLeadTime & ", '" &
                                prevItem & "','" & prevDim1 & "','" & prevDim2 & "','" & prevDim3 & "' FROM Calendar " &
                                "WHERE eDate = '" & nextEdate & "' AND Week_Id > 0 "
                            cmd = New SqlCommand(sql, con2)
                            cmd.ExecuteNonQuery()
                            nextSdate = DateAdd(DateInterval.Day, -7, nextSdate)
                            nextEdate = DateAdd(DateInterval.Day, -7, nextEdate)
                        Loop
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
                    GoTo BF100
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
                        If nextSdate > BuildStartDate Then
                            sql = "IF NOT EXISTS (SELECT * FROM Item_Inv WHERE Loc_Id = '" & location & "' AND Sku = '" & prevSku & "' " &
                                "AND sDate = '" & nextSdate & "' AND eDate = '" & nextEdate & "') " &
                                "INSERT INTO Item_Inv (Loc_Id, Sku, sDate, eDate, Cost, Retail, YrWk, Begin_OH, End_OH, Max_OH, " &
                                "Item, DIM1, DIM2, DIM3) " &
                                "SELECT '" & location & "', '" & prevSku & "', '" & nextSdate & "', '" & nextEdate & "', " & cost & ", " &
                                retail & ", YrWk, " & prevBegin & ", " & thisEndOH & ", " & prevMax & ",'" &
                                prevItem & "','" & prevDim1 & "','" & prevDim2 & "','" & prevDim3 & "' " &
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
BF100:      End While
            con.Close()
            con2.Close()

            txtprogress.Text = "Updating Inv_Summary"
            Me.Refresh()

            con.Open()
            sql = "SELECT eDate, Loc_Id, Dept, Buyer, Class, ISNULL(SUM(End_OH * Retail),0) AS OnHand INTO #t1 FROM Item_Inv i " & _
                "JOIN Item_Master m ON m.Sku = i.Sku " & _
                "GROUP BY eDate, Loc_Id, Dept, Buyer, Class " & _
                "INSERT Inv_Summary(Loc_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, eDate, YrPrd, Week_Num, Act_Inv_Retail) " & _
                "SELECT Loc_Id, c.Year_Id, c.Prd_Id, c.PrdWk, t.Dept, t.Buyer, t.Class, c.sDate, t.eDate, c.YrPrd, " & _
                "c.Week_Num, t.OnHand FROM #t1 t " & _
                "JOIN Calendar c ON c.eDate = t.eDate AND c.Week_Id > 0 " & _
                "WHERE NOT EXISTS (SELECT Loc_ID FROM Inv_Summary i WHERE i.Loc_Id = t.Loc_Id AND i.eDate = t.eDate " & _
                "AND i.Dept = t.Dept AND i.Buyer = t.Buyer AND i.Class = t.Class) " & _
                "SELECT eDate, Loc_Id, Dept, Buyer, Class, ISNULL(SUM(End_OH * Cost),0) AS OnHand INTO #t2 FROM Item_Inv i " & _
                "JOIN Item_Master m ON m.Sku = i.Sku " & _
                "GROUP BY eDate, Loc_Id, Dept, Buyer, Class " & _
                "UPDATE w SET w.Act_Inv_Cost = t.OnHand FROM Inv_Summary w " & _
                "JOIN #t2 t ON t.eDate = w.edate AND t.Loc_Id = w.Loc_Id AND t.Dept = w.Dept " & _
                "AND t.Buyer = w.Buyer AND t.Class = w.Class"
            cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 960
            cmd.ExecuteNonQuery()
            con.Close()
righthere:
            con.Open()
            sql = "SELECT COUNT(*) AS cnt FROM Item_Inv"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader()
            While rdr.Read
                added = rdr("cnt")
            End While
            con.Close()

            txtInvRecords.Text = Format(added, "###,###,###")
            Me.Refresh()

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

            txtprogress.Text = Nothing
            GroupBox7.Visible = True
            If Client.DoForecasting Then btnForecast.Enabled = True
            If Client.DoScoring Then btnScore.Enabled = True
            If Client.DoMarketing Then btnCustomers.Enabled = True
            If Client.DoSalesPlan Then btnSalesPlan.Enabled = True
            Me.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Connect_Database()
        Try
            thisClient = txtClient.Text
            If IsNothing(thisClient) Or thisClient = "" Then
                MsgBox("Select or enter the Client_ID and try again.")
                Exit Sub
            End If
            If thisClient = "ADD NEW" Then
                MsgBox("Select Client and try again.")
                Exit Sub
            End If
            Dim RTRACcon As New SqlConnection(RCClientConString)
            RTRACcon.Open()
            Dim sql As String = "SELECT Server, [Database] AS dBase, EXEs, XMLs, SQLUserID, SQLPassword, errorLog " &
                "FROM CLIENT_MASTER WHERE Client_Id = '" & thisClient & "'"
            cmd = New SqlCommand(sql, RTRACcon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("dBase")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = rdr("Server")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then server = CStr(oTest)
                    oTest = rdr("dBase")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dbName = CStr(oTest)
                    oTest = rdr("EXEs")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then exePath = CStr(oTest)
                    oTest = rdr("XMLs")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then path = CStr(oTest)
                    oTest = rdr("XMLs")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then xmlPath = CStr(oTest)
                    oTest = rdr("SQLUserID")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then sqlUserID = CStr(oTest)
                    oTest = rdr("SQLPassword")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then sqlPassword = CStr(oTest)
                    oTest = rdr("errorLog")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then errorLog = CStr(oTest)
                End If
            End While
            RTRACcon.Close()
            ''conString = "Server=" & server & ";Initial Catalog=" & dbName & ";User Id=" & sqluserid & ";Password=" & sqlPassword & " "
            conString = "Server=" & server & ";Initial Catalog=" & dbName & ";Integrated Security=True"
            constr = conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con4 = New SqlConnection(conString)
            con5 = New SqlConnection(conString)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnUpdate_Calendar_Click(sender As Object, e As EventArgs) Handles btnUpdate_Calendar.Click
        Try
            If IsNothing(thisClient) Or thisClient = "Add New" Then
                MsgBox("Select Client and try again.")
                Exit Sub
            End If
            Call Connect_Database()
            con = New SqlConnection(conString)
            itExists = False
            Dim theYear As Int16 = CInt(txtYear.Text)
            con.Open()
            Dim sql As String = "SELECT * FROM SYS.DATABASES WHERE NAME = '" & dbName & "'"
            Dim cmd As New SqlCommand(sql, con)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            Dim tbl As DataTable = New DataTable
            itExists = False
            While rdr.Read
                oTest = rdr(0)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then itExists = True
            End While
            con.Close()
            If itExists = False Then
                MsgBox("Database was not found on " & server & ". Create it then try again.")
                Exit Sub
            End If
            ''    con = New SqlConnection(conString)
            ''    con.Open()
            ''    sql = "IF NOT EXISTS (SELECT * FROM SYSOBJECTS WHERE NAME = 'Calendar' AND TYPE = 'U') " & _
            ''    "CREATE TABLE [dbo].[Calendar](" & _
            ''    "[Year_id] [smallint] NOT NULL," & _
            ''    "[Prd_id] [smallint] NOT NULL," & _
            ''    "[Week_id] [smallint] NOT NULL," & _
            ''    "[Sdate] [date] NOT NULL," & _
            ''    "[Edate] [date] NOT NULL," & _
            ''    "[YrPrd] [int] NULL," & _
            ''    "[YrWks] [int] NULL," & _
            ''    "[PrdWk] [int] NULL," & _
            '' "CONSTRAINT [PK_Calendar] PRIMARY KEY CLUSTERED " & _
            ''    "([Year_id] ASC," & _
            ''    "[Prd_id] ASC, " & _
            ''    "[Week_id] ASC) " & _
            ''"WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, " & _
            ''    "ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]" & _
            ''") ON [PRIMARY]"
            ''    cmd = New SqlCommand(sql, con)
            ''    cmd.ExecuteNonQuery()
            ''    con.Close()

            Dim pathName As String = ""
            Dim ofd As New OpenFileDialog
            If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                pathName = ofd.FileName
            End If
            If Not IsNothing(pathName) And pathName <> "" Then
                'Dim dt As DataTable = New DataTable
                Dim ws As DataTable
                ds = New DataSet
                Dim pos, i As Integer
                Dim testPath As String
                For i = 1 To Len(pathName)
                    If Mid(pathName, i, 1) = "\" Then pos = i
                Next
                testPath = Mid(pathName, pos, 999)

                Dim flArray() As String = testPath.Split("\")
                Dim tArray() As String = testPath.Split(".")
                Dim extn As String = UCase(tArray(1).ToString)
                If extn <> "XLSX" Then
                    MsgBox("Wrong file type was selected. Please select an XLSX!")
                    Exit Sub
                End If
                Dim sheetname As String
                Dim oleCon As OleDb.OleDbConnection
                Dim oleAdapter As System.Data.OleDb.OleDbDataAdapter
                oleCon = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" &
                                                               pathName & ";Excel 12.0 xml;HDR=Yes;")
                oleCon.Open()
                ws = oleCon.GetSchema("Tables")
                sheetname = ws.Rows(0).Item("Table_Name")
                Dim connect As OleDb.OleDbConnection
                connect = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" & pathName & ";Excel 12.0 xml;HDR=Yes;")
                oleAdapter = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sheetname & "]", connect)
                oleAdapter.Fill(tbl)
            End If

            ' ''Dim CPSQLCon As SqlConnection
            ' ''CPSQLCon = New SqlConnection("Data Source=CURTIS-MOBILE;Initial Catalog=CPSQL;User ID=sa ; Password=CounterPoint8")
            ' ''CPSQLCon.Open()
            ' ''sql = "SELECT * FROM CPSQL.dbo.SY_CALNDR WHERE CALNDR_ID = '" & txtYear.Text & "'"
            ' ''Dim cmd2 As New SqlCommand(sql, CPSQLCon)
            ' ''Dim rdr2 As SqlDataReader = cmd2.ExecuteReader
            ' ''Dim row As DataRow
            ' ''Dim prdcnt As Int16 = 0
            ' ''While rdr2.Read
            ' ''    ' Add records for the year range
            ' ''    theLasteDate = rdr2.Item(1)
            ' ''    row = tbl.NewRow
            ' ''    row.Item(0) = theYear                                                         ' Year
            ' ''    row.Item(1) = 0                                                               ' Period_Id
            ' ''    row.Item(2) = 0                                                               ' Week_Id
            ' ''    row.Item(3) = theLasteDate                                                    ' sDate
            ' ''    row.Item(4) = rdr2(2)                                                         ' edate
            ' ''    row.Item(5) = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrPrd
            ' ''    row.Item(6) = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
            ' ''    tbl.Rows.Add(row)
            ' ''    ' Add records for period ranges (called Months in CounterPoint)
            ' ''    For i = 2 To 28 Step 2
            ' ''        oTest = rdr2.Item(i + 16)
            ' ''        If Not IsDBNull(oTest) Then
            ' ''            prdcnt += 1
            ' ''            row = tbl.NewRow
            ' ''            row.Item(0) = theYear                                                         ' Year
            ' ''            row.Item(1) = prdcnt                                                          ' Period_Id
            ' ''            row.Item(2) = 0                                                               ' Week_Id
            ' ''            row.Item(3) = theLasteDate                                                    ' sDate
            ' ''            row.Item(4) = rdr2(i + 16)                                                    ' edate
            ' ''            row.Item(5) = Microsoft.VisualBasic.Right(theYear, 2) & Format(prdcnt, "00")  ' YrPrd
            ' ''            row.Item(6) = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrWks
            ' ''            theLasteDate = rdr2.Item(i + 16)
            ' ''            theLasteDate = DateAdd(DateInterval.Day, 1, theLasteDate)
            ' ''            tbl.Rows.Add(row)
            ' ''        End If
            ' ''    Next
            ' ''    prdcnt = 0
            ' ''    theLasteDate = rdr2.Item(1)
            ' ''    ' Add records for week ranges
            ' ''    For i = 2 To 108 Step 2
            ' ''        oTest = rdr2.Item(i + 44)
            ' ''        If Not IsDBNull(oTest) Then
            ' ''            prdcnt += 1
            ' ''            row = tbl.NewRow
            ' ''            row.Item(0) = theYear                                                         ' Year
            ' ''            row.Item(1) = 0                                                               ' Period_Id
            ' ''            row.Item(2) = prdcnt                                                          ' Week_Id
            ' ''            row.Item(3) = theLasteDate                                                    ' sDate
            ' ''            row.Item(4) = rdr2(i + 44)                                                    ' edate
            ' ''            row.Item(5) = Microsoft.VisualBasic.Right(theYear, 2) & "00"                  ' YrPrd
            ' ''            row.Item(6) = Microsoft.VisualBasic.Right(theYear, 2) & Format(prdcnt, "00")  ' YrWks
            ' ''            theLasteDate = rdr2.Item(i + 44)
            ' ''            theLasteDate = DateAdd(DateInterval.Day, 1, theLasteDate)
            ' ''            tbl.Rows.Add(row)
            ' ''        End If
            ' ''    Next
            ' ''End While
            ' ''CPSQLCon.Close()

            ' Delete and replace records in OTB_Calender matching selected year
            con.Open()
            sql = "DELETE FROM Calendar WHERE Year_Id = " & theYear & ""
            Dim cmd3 As New SqlCommand(sql, con)
            cmd3.ExecuteNonQuery()
            con.Close()

            con.Open()
            Dim s As SqlBulkCopy = New SqlBulkCopy(con)
            s.DestinationTableName = "Calendar"
            s.WriteToServer(tbl)
            haveCalendar = True
            s.Close()
            con.Close()

            ''         Fix Period end eDate
            con.Open()
            sql = "SELECT Year_Id, Prd_Id, PrdWk, eDate INTO #t1 FROM Calendar c2 " & _
                "WHERE Week_Id = (SELECT MAX(Week_Id) FROM Calendar c3 WHERE c3.Year_Id = c2.Year_Id AND c3.Prd_Id = c2.Prd_Id) " & _
                "AND Prd_Id > 0 ORDER BY eDate " & _
                "UPDATE c1 SET c1.eDate = c2.eDate FROM Calendar c1 " & _
                "JOIN #t1 c2 on c2.Year_Id = c1.Year_Id AND c2.Prd_Id = c1.Prd_Id AND c1.Week_id = 0"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            '' Set Prd_Id and YrPrd fields for week records
            'con.Open()
            'sql = "UPDATE o1 SET o1.Prd_Id = o2.Prd_Id, o1.YrPrd = o2.YrPrd " & _
            '"FROM OTB_Calendar As o1 " & _
            '"INNER JOIN OTB_Calendar AS o2 ON o1.Year_Id = o2.Year_Id " & _
            '"WHERE o1.Sdate BETWEEN o2.Sdate AND o2.Edate " & _
            '"AND o1.Prd_Id = 0 AND o2.Prd_Id > 0 " & _
            '"AND o1.Week_Id > 0"
            'Dim cmd4 As New SqlCommand(sql, con)
            'cmd4.ExecuteNonQuery()
            'con.Close()

            '' Set PrdWk for week records
            'con.Open()
            'sql = "SELECT * FROM OTB_Calendar WHERE Year_Id = '" & theYear & "' ORDER BY Sdate"
            'Dim adapter As SqlDataAdapter = New SqlDataAdapter(sql, con)
            'Dim dt As DataTable = New DataTable("Calendar")
            'Dim cb As New SqlCommandBuilder(adapter)
            'adapter.Fill(dt)
            'con.Close()
            'Dim wkId, PrdWk As Integer
            'For Each r In dt.Rows
            '    wkId = r.item(2)
            '    If wkId = 0 Then
            '        PrdWk = 0
            '    Else
            '        PrdWk += 1
            '    End If
            '    r.item(7) = PrdWk
            'Next
            'adapter.Update(dt)
            MsgBox(theYear & " Calendar loaded")
            GroupBox2.Visible = True
            btnBuyers.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCreateWeekly.Click
        If IsNothing(conString) Then
            Call Connect_Database()
            '' con = New SqlConnection(conString)
            Exit Sub
        End If
        '
        '       clean up Item_Master so that only Active buyers make it to Weekly_Summary
        '
        con.Open()
        sql = "UPDATE m SET m.Buyer = 'OTHER' FROM Item_Master m " & _
            "LEFT JOIN Buyers b ON b.ID = m.Buyer " & _
            "WHERE b.Status IS NULL " & _
            "UPDATE m SET m.Buyer = 'OTHER' FROM Item_Master m " & _
            "JOIN Buyers b ON b.ID = m.Buyer " & _
            "WHERE b.Status <> 'Active'"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        txtprogress.Text = "Updating Weekly Summary"
        Me.Refresh()
        con.Open()
        ''sql = "SELECT Str_Id, Year_Id, Prd_Id, PrdWk AS Week_Id, Dept, Buyer, Class, d.sDate, d.eDate, c.YrPrd, c.Week_Num, " & _
        ''    "SUM(ISNULL(Sales_Retail,0)) AS Retail, SUM(ISNULL(Sales_Cost,0)) AS Cost, SUM(ISNULL(Markdown,0)) AS Markdown " & _
        ''    "INTO #t1 FROM Item_Sales d " & _
        ''    "JOIN Item_Master m ON m.Item_No = d.Item_No " & _
        ''    "JOIN Calendar c ON c.eDate = d.eDate AND c.Week_Id > 0 " & _
        ''    "GROUP BY Str_Id, Year_Id, Prd_Id, PrdWk, Dept, Buyer, Class, d.sDate, d.eDate, c.YrPrd, c.Week_Num " & _
        ''    "MERGE Weekly_Summary AS t USING #t1 AS s ON t.Str_Id = s.Str_Id AND t.Dept = s.Dept AND t.Buyer = s.Buyer " & _
        ''    "AND t.class = s.Class AND t.eDate = s.eDate " & _
        ''    "WHEN NOT MATCHED BY TARGET THEN INSERT(Str_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, eDate, " & _
        ''    "YrPrd, Week_Num, Act_Sales, Act_Sales_Cost, Act_Mkdn) " & _
        ''    "VALUES(s.Str_id, s.Year_Id, s.Prd_Id, s.Week_Id, s.Dept, s.Buyer, s.Class, s.sDate, s.eDate, s.YrPrd, " & _
        ''    "s.Week_Num, s.Retail, s.Cost, s.Markdown) " & _
        ''    "WHEN MATCHED THEN UPDATE SET t.Act_Sales = s.Retail, t.Act_Sales_Cost = s.Cost, t.Act_Mkdn = s.Markdown; " & _
        ''    "SELECT Loc_Id, Year_Id, Prd_Id, PrdWk AS Week_Id, Dept, Buyer, Class, i.sDate, i.eDate, c.YrPrd, c.Week_Num, " & _
        ''    "SUM(ISNULL(End_OH,0) * ISNULL(Retail,0)) AS InvRetail, SUM(ISNULL(End_OH,0) * ISNULL(Cost,0)) AS InvCost, " & _
        ''    "SUM(ISNULL(Max_OH,0) * ISNULL(Cost,0)) AS MaxCost INTO #t2 FROM Item_Inv i " & _
        ''    "JOIN Item_Master m ON m.Item_No = i.Item_No " & _
        ''    "JOIN Calendar c ON c.eDate = i.eDate AND c.Week_Id > 0 " & _
        ''    "GROUP BY Loc_Id, Year_Id, Prd_Id, PrdWk, Dept, Buyer, Class, i.sDate, i.eDate, c.YrPrd, c.Week_Num " & _
        ''    "MERGE Weekly_Summary AS t USING #t2 AS s ON t.Str_Id = s.Str_Id AND t.Dept = s.Dept AND t.Buyer = s.Buyer " & _
        ''    "AND t.Class = s.Class AND t.eDate = s.eDate " & _
        ''    "WHEN NOT MATCHED BY TARGET THEN INSERT(Str_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, edate, " & _
        ''    "YrPrd, Week_Num, Act_Inv_Retail, Act_Inv_Cost, Max_OH_Cost) " & _
        ''    "VALUES(s.Str_Id, s.Year_Id, s.Prd_Id, s.Week_Id, s.Dept, s.Buyer, s.Class, s.sDate, s.eDate, " & _
        ''    "s.YrPrd, s.Week_Num, s.InvRetail, s.InvCost, s.MaxCost) " & _
        ''    "WHEN MATCHED THEN UPDATE SET t.Act_Inv_Retail = s.InvRetail, t.Act_Inv_Cost = s.InvCost, t.Max_OH_Cost = s.MaxCost;"


        '     Weekly Summary no longer has inventory


        sql = "SELECT Str_Id, Loc_Id, Year_Id, Prd_Id, PrdWk AS Week_Id, Dept, Buyer, Class, d.sDate, d.eDate, c.YrPrd, c.Week_Num, " & _
    "SUM(ISNULL(Sales_Retail,0)) AS Retail, SUM(ISNULL(Sales_Cost,0)) AS Cost, SUM(ISNULL(Markdown,0)) AS Markdown " & _
    "INTO #t1 FROM Item_Sales d " & _
    "JOIN Item_Master m ON m.Sku = d.Sku " & _
    "JOIN Calendar c ON c.eDate = d.eDate AND c.Week_Id > 0 " & _
    "GROUP BY Str_Id, Loc_Id, Year_Id, Prd_Id, PrdWk, Dept, Buyer, Class, d.sDate, d.eDate, c.YrPrd, c.Week_Num " & _
    "MERGE Sales_Summary AS t USING #t1 AS s ON t.Str_Id = s.Str_Id AND t.Loc_Id = s.Loc_Id AND " & _
    "t.Dept = s.Dept AND t.Buyer = s.Buyer AND t.class = s.Class AND t.eDate = s.eDate " & _
    "WHEN NOT MATCHED BY TARGET THEN INSERT(Str_Id, Loc_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, sDate, eDate, " & _
    "YrPrd, Week_Num, Act_Sales, Act_Sales_Cost, Act_Mkdn) " & _
    "VALUES(s.Str_id, s.Loc_Id, s.Year_Id, s.Prd_Id, s.Week_Id, s.Dept, s.Buyer, s.Class, s.sDate, s.eDate, s.YrPrd, " & _
    "s.Week_Num, s.Retail, s.Cost, s.Markdown) " & _
    "WHEN MATCHED THEN UPDATE SET t.Act_Sales = s.Retail, t.Act_Sales_Cost = s.Cost, t.Act_Mkdn = s.Markdown; "
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 480
        cmd.ExecuteNonQuery()
        sql = "SELECT COUNT(*) AS cnt FROM Sales_Summary"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cnt = rdr("cnt")
        End While
        con.Close()
        txtprogress.Text = "Updating Period_Pct Table"
        Me.Refresh()

        con.Open()
        sql = "SELECT CASE WHEN LEN(STORE) = 0 THEN 'UNKNOWN' ELSE STORE END AS Store, " & _
            "CASE WHEN LEN(DEPT) = 0 THEN 'UNKNOWN' ELSE DEPT END AS Dept, Prd_Id, PrdWk, " & _
            "DATEPART(dw,TRANS_DATE) AS DOW, SUM(QTY * RETAIL) AS Sales, " & _
            "CONVERT(Decimal(18,4),0) AS DOWPct, CONVERT(Decimal(18,4),0) AS PrdDayPct INTO #t1 FROM Daily_Transaction_Log l " & _
            "JOIN Calendar c ON TRANS_DATE BETWEEN sDate AND eDate AND Week_Id > 0 " & _
            "WHERE TRANS_DATE BETWEEN DATEADD(day,-365,GETDATE()) AND GETDATE() " & _
            "GROUP BY Store, Dept, Prd_Id, PrdWk, DATEPART(dw,TRANS_DATE) " & _
            "INSERT INTO #t1(Store, Dept, Prd_Id, PrdWk, DOW, Sales) " & _
            "SELECT Store, '**', Prd_Id, PrdWk, DOW, SUM(Sales) FROM #t1 " & _
            "GROUP BY Store, Prd_Id, PrdWk, DOW " & _
            "SELECT Store, Dept, Prd_Id, PrdWk, SUM(Sales) AS Sales INTO #t2 FROM #t1 " & _
            "GROUP BY Store, Dept, Prd_Id, PrdWk " & _
            "UPDATE t1 SET t1.DOWPct = CASE WHEN t2.Sales <> 0 THEN t1.Sales / t2.Sales ELSE 0 END FROM #t1 t1 " & _
            "JOIN #t2 t2 ON t2.Store = t1.Store AND t2.Dept = t1.Dept AND t2.Prd_id = t1.Prd_Id AND t2.PrdWk = t1.PrdWk " & _
            "SELECT Store, Dept, Prd_Id, DOW, SUM(Sales) AS Sales, CONVERT(Decimal(18,4),0) AS Pct INTO #t3 FROM #t1 " & _
            "GROUP BY Store, Dept, Prd_Id, DOW " & _
            "SELECT Store, Dept, Prd_Id, SUM(Sales) As Sales INTO #t4 FROM #t3 " & _
            "GROUP BY Store, Dept, Prd_Id " & _
            "UPDATE t3 SET t3.PCT = CASE WHEN t4.Sales = 0 THEN 0 ELSE t3.Sales / t4.Sales END FROM #t3 t3 " & _
            "JOIN #t4 t4 ON t4.Store = t3.Store AND t4.Dept = t3.Dept " & _
            "UPDATE t1 SET t1.PrdDayPct = t3.PCT FROM #t1 t1 " & _
            "JOIN #t3 t3 ON t3.Store = t1.Store AND t3.Dept = t1.Dept AND t3.Prd_Id = t1.Prd_Id AND t3.DOW = t1.DOW " & _
            "INSERT INTO PeriodWeekPct SELECT * FROM #t1 "
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        txtprogress.Text = "Updating DayOfWeekPct"
        Me.Refresh()
        sql = "SELECT CASE WHEN LEN(STORE) = 0 THEN 'UNKNOWN' ELSE STORE END As Store, Prd_Id, " & _
            "DATEPART(dw,TRANS_DATE) AS DOW, SUM(Qty * Retail) AS Sales, " & _
            "CONVERT(Decimal(18,4),0) AS Pct INTO #t4a FROM Daily_Transaction_Log l " & _
            "JOIN Calendar c ON TRANS_DATE BETWEEN sDate AND eDate AND Week_Id > 0 " & _
            "WHERE TRANS_DATE BETWEEN DATEADD(day,-365,GETDATE()) AND GETDATE() " & _
            "AND l.[TYPE] = 'Sold' " & _
            "GROUP BY Store, Prd_Id, DATEPART(dw,TRANS_DATE) " & _
            "SELECT Store, Prd_Id, SUM(Sales) AS Sales into #t5 FROM #t4a " & _
            "GROUP BY Store, Prd_Id " & _
            "UPDATE t4 SET t4.Pct = t4.Sales / t5.Sales FROM #t4a t4 " & _
            "JOIN #t5 t5 ON t5.Store = t4.Store AND t5.Prd_Id = t4.Prd_Id " & _
            "INSERT INTO DayOfWeekPct(Str_Id, Prd_Id, DOW, Sales, Pct) " & _
            "SELECT Store, Prd_id, DOW, Sales, Pct FROM #t4a"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        txtWeeklySummary.Text = Format(cnt, "###,###,##0")
        GroupBox4.Visible = True
        btnAddWeeklyRecords.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub btnClndr_Click(sender As Object, e As EventArgs) Handles btnClndr.Click
        Calendar.Show()

    End Sub

    Private Sub Split_Sales()
        Dim yr As Integer
        txtprogress.Text = "Splitting Sales.xml file"
        Me.Refresh()
        Dim path As String
        path = xmlPath & "\Sales1.xml"
        If System.IO.File.Exists("" & path & "") Then IO.File.Delete(path)
        path = xmlPath & "\Sales2.xml"
        If System.IO.File.Exists("" & path & "") Then IO.File.Delete(path)
        path = xmlPath & "\Sales3.xml"
        If System.IO.File.Exists("" & path & "") Then IO.File.Delete(path)
        path = xmlPath & "\Sales4.xml"
        If System.IO.File.Exists("" & path & "") Then IO.File.Delete(path)
        path = xmlPath & "\Sales5.xml"
        If System.IO.File.Exists("" & path & "") Then IO.File.Delete(path)

        path = xmlPath & "\Sales.xml"
        Dim xmlReader As XmlTextReader = New XmlTextReader("" & path & "")
        Dim xmlWriter1, xmlWriter2, xmlWriter3, xmlWriter4, xmlWriter5 As XmlTextWriter
        path = xmlPath & "\Sales1.xml"
        xmlWriter1 = New XmlTextWriter("" & path & "", System.Text.Encoding.UTF8)
        path = xmlPath & "\Sales2.xml"
        xmlWriter2 = New XmlTextWriter("" & path & "", System.Text.Encoding.UTF8)
        path = xmlPath & "\Sales3.xml"
        xmlWriter3 = New XmlTextWriter("" & path & "", System.Text.Encoding.UTF8)
        path = xmlPath & "\Sales4.xml"
        xmlWriter4 = New XmlTextWriter("" & path & "", System.Text.Encoding.UTF8)
        path = xmlPath & "\Sales5.xml"
        xmlWriter5 = New XmlTextWriter("" & path & "", System.Text.Encoding.UTF8)

        xmlWriter1.WriteStartDocument(True)
        xmlWriter2.WriteStartDocument(True)
        xmlWriter3.WriteStartDocument(True)
        xmlWriter4.WriteStartDocument(True)
        xmlWriter5.WriteStartDocument(True)
        xmlWriter1.Formatting = Formatting.Indented
        xmlWriter2.Formatting = Formatting.Indented
        xmlWriter3.Formatting = Formatting.Indented
        xmlWriter4.Formatting = Formatting.Indented
        xmlWriter5.Formatting = Formatting.Indented
        xmlWriter1.Indentation = 5
        xmlWriter2.Indentation = 5
        xmlWriter3.Indentation = 5
        xmlWriter4.Indentation = 5
        xmlWriter5.Indentation = 5
        xmlWriter1.WriteStartElement("Sales")
        xmlWriter2.WriteStartElement("Sales")
        xmlWriter3.WriteStartElement("Sales")
        xmlWriter4.WriteStartElement("Sales")
        xmlWriter5.WriteStartElement("Sales")
        Dim cnt, w1, w2, w3, w4, w5 As Int32
        Dim id As String = ""
        Dim seq As String = ""
        Dim item As String = ""
        Dim loc As String = ""
        Dim qty As String = ""
        Dim cost As String = ""
        Dim retail As String = ""
        Dim mkdn As String = ""
        Dim dept As String = ""
        Dim buyer As String = ""
        Dim cls As String = ""
        Dim cust As String = ""
        Dim tkt As String = ""
        Dim fld As String = ""
        Dim val As String = ""
        Dim reason As String = ""
        Dim coupon As String = ""
        Dim dte As Date
        Dim extrct As DateTime
        While xmlReader.Read
            Dim otest As Object = xmlReader.Value

            Select Case xmlReader.NodeType
                Case XmlNodeType.Element
                    fld = xmlReader.Name
                Case XmlNodeType.Text
                    val = xmlReader.Value
                Case XmlNodeType.EndElement
                    If fld = "TRANS_ID" Then
                        cnt += 1
                        If cnt Mod 10000 = 0 Then
                            txtprogress.Text = "read " & cnt & " sales records " & " wrote1 " & w1 & " wrote2 " & w2 &
                                " wrote3 " & w3 & " wrote4 " & w4 & " wrote5 " & w5
                            Me.Refresh()
                        End If
                    End If
                    If fld = "TRANS_ID" Then id = val
                    If fld = "SEQ_NO" Then seq = val
                    If fld = "ITEM_NO" Then item = val
                    If fld = "LOCATION" Then loc = val
                    If fld = "QTY" Then qty = val
                    If fld = "COST" Then cost = val
                    If fld = "RETAIL" Then retail = val
                    If fld = "TRANS_DATE" Then dte = val
                    If fld = "MARKDOWN" Then mkdn = val
                    If fld = "DEPT" Then dept = val
                    If fld = "CLASS" Then cls = val
                    If fld = "BUYER" Then buyer = val
                    If fld = "DATE" Then dte = val
                    If fld = "CUST_NO" Then cust = val
                    If fld = "TKT_NO" Then tkt = val
                    If fld = "EXTRACT_DATE" Then extrct = val
                    If fld = "Mkdn_REASON" Then reason = val
                    If fld = "COUPON_CODE" Then coupon = val
                    If fld = "BUYER" Then
                        yr = DatePart(DateInterval.Year, dte)
                        If DatePart(DateInterval.Year, CDate(dte)) = 2012 Then
                            xmlWriter1.WriteStartElement("SALE")
                            xmlWriter1.WriteElementString("TRANS_ID", id)
                            xmlWriter1.WriteElementString("SEQ_NO", seq)
                            xmlWriter1.WriteElementString("ITEM_NO", item)
                            xmlWriter1.WriteElementString("LOCATION", loc)
                            xmlWriter1.WriteElementString("QTY", qty)
                            xmlWriter1.WriteElementString("COST", cost)
                            xmlWriter1.WriteElementString("RETAIL", retail)
                            xmlWriter1.WriteElementString("TRANS_DATE", dte)
                            xmlWriter1.WriteElementString("MARKDOWN", mkdn)
                            xmlWriter1.WriteElementString("DEPT", dept)
                            xmlWriter1.WriteElementString("CLASS", cls)
                            xmlWriter1.WriteElementString("BUYER", buyer)
                            xmlWriter1.WriteElementString("CUST_NO", cust)
                            xmlWriter1.WriteElementString("TKT_NO", tkt)
                            xmlWriter1.WriteElementString("EXTRACT_DATE", extrct)
                            xmlWriter1.WriteElementString("MKDN_REASON", reason)
                            xmlWriter1.WriteElementString("COUPON_CODE", coupon)
                            xmlWriter1.WriteEndElement()
                            w1 += 1
                        End If
                        If DatePart(DateInterval.Year, CDate(dte)) = 2013 Then
                            xmlWriter2.WriteStartElement("SALE")
                            xmlWriter2.WriteElementString("TRANS_ID", id)
                            xmlWriter2.WriteElementString("SEQ_NO", seq)
                            xmlWriter2.WriteElementString("ITEM_NO", item)
                            xmlWriter2.WriteElementString("LOCATION", loc)
                            xmlWriter2.WriteElementString("QTY", qty)
                            xmlWriter2.WriteElementString("COST", cost)
                            xmlWriter2.WriteElementString("RETAIL", retail)
                            xmlWriter2.WriteElementString("TRANS_DATE", dte)
                            xmlWriter2.WriteElementString("MARKDOWN", mkdn)
                            xmlWriter2.WriteElementString("DEPT", dept)
                            xmlWriter2.WriteElementString("CLASS", cls)
                            xmlWriter2.WriteElementString("BUYER", buyer)
                            xmlWriter2.WriteElementString("CUST_NO", cust)
                            xmlWriter2.WriteElementString("TKT_NO", tkt)
                            xmlWriter2.WriteElementString("EXTRACT_DATE", extrct)
                            xmlWriter2.WriteElementString("MKDN_REASON", reason)
                            xmlWriter2.WriteElementString("COUPON_CODE", coupon)
                            xmlWriter2.WriteEndElement()
                            w2 += 1
                        End If
                        If DatePart(DateInterval.Year, CDate(dte)) = 2014 Then
                            xmlWriter3.WriteStartElement("SALE")
                            xmlWriter3.WriteElementString("TRANS_ID", id)
                            xmlWriter3.WriteElementString("SEQ_NO", seq)
                            xmlWriter3.WriteElementString("ITEM_NO", item)
                            xmlWriter3.WriteElementString("LOCATION", loc)
                            xmlWriter3.WriteElementString("QTY", qty)
                            xmlWriter3.WriteElementString("COST", cost)
                            xmlWriter3.WriteElementString("RETAIL", retail)
                            xmlWriter3.WriteElementString("TRANS_DATE", dte)
                            xmlWriter3.WriteElementString("MARKDOWN", mkdn)
                            xmlWriter3.WriteElementString("DEPT", dept)
                            xmlWriter3.WriteElementString("CLASS", cls)
                            xmlWriter3.WriteElementString("BUYER", buyer)
                            xmlWriter3.WriteElementString("CUST_NO", cust)
                            xmlWriter3.WriteElementString("TKT_NO", tkt)
                            xmlWriter3.WriteElementString("EXTRACT_DATE", extrct)
                            xmlWriter3.WriteElementString("MKDN_REASON", reason)
                            xmlWriter3.WriteElementString("COUPON_CODE", coupon)
                            xmlWriter3.WriteEndElement()
                            w3 += 1
                        End If
                        If DatePart(DateInterval.Year, CDate(dte)) = 2015 Then
                            xmlWriter4.WriteStartElement("SALE")
                            xmlWriter4.WriteElementString("TRANS_ID", id)
                            xmlWriter4.WriteElementString("SEQ_NO", seq)
                            xmlWriter4.WriteElementString("ITEM_NO", item)
                            xmlWriter4.WriteElementString("LOCATION", loc)
                            xmlWriter4.WriteElementString("QTY", qty)
                            xmlWriter4.WriteElementString("COST", cost)
                            xmlWriter4.WriteElementString("RETAIL", retail)
                            xmlWriter4.WriteElementString("TRANS_DATE", dte)
                            xmlWriter4.WriteElementString("MARKDOWN", mkdn)
                            xmlWriter4.WriteElementString("DEPT", dept)
                            xmlWriter4.WriteElementString("CLASS", cls)
                            xmlWriter4.WriteElementString("BUYER", buyer)
                            xmlWriter4.WriteElementString("CUST_NO", cust)
                            xmlWriter4.WriteElementString("TKT_NO", tkt)
                            xmlWriter4.WriteElementString("EXTRACT_DATE", extrct)
                            xmlWriter4.WriteElementString("MKDN_REASON", reason)
                            xmlWriter4.WriteElementString("COUPON_CODE", coupon)
                            xmlWriter4.WriteEndElement()
                            w4 += 1
                        End If
                        If DatePart(DateInterval.Year, CDate(dte)) = 2016 Then
                            xmlWriter5.WriteStartElement("SALE")
                            xmlWriter5.WriteElementString("TRANS_ID", id)
                            xmlWriter5.WriteElementString("SEQ_NO", seq)
                            xmlWriter5.WriteElementString("ITEM_NO", item)
                            xmlWriter5.WriteElementString("LOCATION", loc)
                            xmlWriter5.WriteElementString("QTY", qty)
                            xmlWriter5.WriteElementString("COST", cost)
                            xmlWriter5.WriteElementString("RETAIL", retail)
                            xmlWriter5.WriteElementString("TRANS_DATE", dte)
                            xmlWriter5.WriteElementString("MARKDOWN", mkdn)
                            xmlWriter5.WriteElementString("DEPT", dept)
                            xmlWriter5.WriteElementString("CLASS", cls)
                            xmlWriter5.WriteElementString("BUYER", buyer)
                            xmlWriter5.WriteElementString("CUST_NO", cust)
                            xmlWriter5.WriteElementString("TKT_NO", tkt)
                            xmlWriter5.WriteElementString("EXTRACT_DATE", extrct)
                            xmlWriter5.WriteElementString("MKDN_REASON", reason)
                            xmlWriter5.WriteElementString("COUPON_CODE", coupon)
                            xmlWriter5.WriteEndElement()
                            w5 += 1
                        End If
                    End If
            End Select

        End While
        xmlReader.Close()
        xmlWriter1.WriteEndElement()
        xmlWriter2.WriteEndElement()
        xmlWriter3.WriteEndElement()
        xmlWriter4.WriteEndElement()
        xmlWriter5.WriteEndElement()
        xmlWriter1.WriteEndDocument()
        xmlWriter2.WriteEndDocument()
        xmlWriter3.WriteEndDocument()
        xmlWriter4.WriteEndDocument()
        xmlWriter5.WriteEndDocument()
        xmlWriter1.Close()
        xmlWriter2.Close()
        xmlWriter3.Close()
        xmlWriter4.Close()
        xmlWriter5.Close()
        txtprogress.Text = "read " & cnt & " sales records " & " wrote1 " & w1 & " wrote2 " & w2 & " wrote3 " & w3
        MsgBox("Sales.xml split into 4 seperate files. Select Sales1.xml for the first file to process.")
    End Sub

    Private Sub cboClient_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClient.SelectedIndexChanged
        Try
            thisClient = cboClient.SelectedItem
            txtClient.Text = thisClient
            If thisClient = "ADD NEW" Then
                lblClient.Visible = True
                txtClient.Visible = True
            Else
                Call Connect_Database()
                lblClient.Text = "Client"
                GroupBox5.Visible = True
            End If
            arguments = thisClient & ";" & server & ";" & dbName & ";" & exePath & ";" & xmlPath & ";" & sqlUserID & ";" &
                            sqlPassword & ";" & Date.Today & ";" & Date.Today & ";" & errorLog & ";" & Client.DoMarketing
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnForecast_Click(sender As Object, e As EventArgs) Handles btnForecast.Click
        Dim dbase, userid, password As String
        Dim strng() As String = Split(conString, ";")
        server = strng(0).Replace("Server=", "")
        dbase = strng(1).Replace("Initial Catalog=", "")
        If InStr(conString, "Integrated") > 0 Then
            userid = sqlUserID
            password = sqlPassword
        Else
            userid = strng(2).Replace("User Id=", "")
            password = strng(3).Replace("Password=", "")
        End If
        If System.IO.File.Exists(exePath & "ItemForecast.exe") Then
            Dim p As New ProcessStartInfo
            p.FileName = exePath & "\ItemForecast.exe"
            p.Arguments = thisClient & ";" & server & ";" & dbase & ";" & userid & ";" & password
            p.UseShellExecute = True
            p.WindowStyle = ProcessWindowStyle.Normal
            Dim proc As Process = Process.Start(p)
            proc.WaitForExit()
        Else
            MessageBox.Show("Move ItemForecast.exe to " & exePath & " and try again", "ERROR! EXE WAS NOT FOUND")
        End If
    End Sub

    Private Sub btnScore_Click(sender As Object, e As EventArgs) Handles btnScore.Click
        MessageBox.Show("Run the 'Module2ItemScoring.sql' script then hit OK to continue.", "PAUSE")
        MsgBox("Go into RCAdmin to complete Scoring Setup.")

    End Sub

    Private Sub btnSalesPlan_Click(sender As Object, e As EventArgs) Handles btnSalesPlan.Click
        BuyerPCT.Show()
    End Sub

    Private Sub Update_Process_Log(ByRef modul As String, ByRef process As String, ByRef m As String, ByRef stat As String)
        stopwatch.Stop()
        Dim ts As TimeSpan = stopwatch.Elapsed
        Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        Dim pgm As String = "RCSetup"
        con.Open()
        sql = "INSERT INTO Process_Log (Date, Program, Module, Process, Message, Status, Duration) " & _
            "SELECT '" & Date.Now & "','" & pgm & "','" & modul & "','" & process & "','" & m & "','" & stat & "','" & et & "'"
        cmd = New SqlCommand(sql, con)
        cmd.CommandTimeout = 120
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub btnCustomers_Click(sender As Object, e As EventArgs) Handles btnCustomers.Click
        txtprogress.Text = "Loading Customers.xml"
        Me.Refresh()
        If IsNothing(conString) Then
            Call Connect_Database()
            con = New SqlConnection(conString)
        End If
        ''con.Open()
        ''cmd = New SqlCommand("dbo.sp_RCSETUP_CreateCustomersTable", con)
        ''cmd.CommandType = CommandType.StoredProcedure
        ''cmd.ExecuteNonQuery()
        ''con.Close()
        Dim thePath As String = ""
        thePath = xmlPath & "\Customers.xml"
        If System.IO.File.Exists(thePath) Then
        Else
            MsgBox(thePath & " Not found. Copy it there and try again.")
            Exit Sub
        End If
        Dim tbl As DataTable = New DataTable
        Dim dset As DataSet = New DataSet
        Dim xmlFile As XmlReader
        Dim row As DataRow
        Dim id As String
        Dim dtTbl = New DataTable
        dset = New DataSet
        Dim column As New DataColumn
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "Cust_No"
        dtTbl.Columns.Add(column)
        Dim PrimaryKey2(1) As DataColumn
        PrimaryKey2(0) = dtTbl.Columns("Cust_No")
        dtTbl.PrimaryKey = PrimaryKey2
        dtTbl.Columns.Add("Name", GetType(System.String))
        dtTbl.Columns.Add("fName", GetType(System.String))
        dtTbl.Columns.Add("lName", GetType(System.String))
        dtTbl.Columns.Add("Addr1", GetType(System.String))
        dtTbl.Columns.Add("Addr2", GetType(System.String))
        dtTbl.Columns.Add("Addr3", GetType(System.String))
        dtTbl.Columns.Add("City", GetType(System.String))
        dtTbl.Columns.Add("State", GetType(System.String))
        dtTbl.Columns.Add("Zip", GetType(System.String))
        dtTbl.Columns.Add("Phone1", GetType(System.String))
        dtTbl.Columns.Add("Phone2", GetType(System.String))
        dtTbl.Columns.Add("Cell1", GetType(System.String))
        dtTbl.Columns.Add("Cell2", GetType(System.String))
        dtTbl.Columns.Add("eMail", GetType(System.String))
        dtTbl.Columns.Add("eMail2", GetType(System.String))
        dtTbl.Columns.Add("Type", GetType(System.String))
        dtTbl.Columns.Add("Balance", GetType(System.Decimal))
        dtTbl.Columns.Add("Loyalty", GetType(System.Int32))
        dtTbl.Columns.Add("okToMail", GetType(System.String))
        dtTbl.Columns.Add("okToEmail", GetType(System.String))
        dtTbl.Columns.Add("fDate", GetType(System.DateTime))
        dtTbl.Columns.Add("lDate", GetType(System.DateTime))
        dtTbl.Columns.Add("lUpdate", GetType(System.DateTime))
        dtTbl.Columns.Add("Spouse", GetType(System.String))
        dtTbl.Columns.Add("Birthday", GetType(System.DateTime))
        Dim custNo, name, fName, lName, Addr1, Addr2, Addr3, city, state, zip, spouse, phone1, phone2,
            email, type, oktomail, oktoemail, cell1, cell2 As String
        Dim cnt As Integer = 0
        Dim birthday, fDate, lDate, lUpdate As String
        Dim bal, points As String
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        dset.ReadXml(xmlFile)
        If ((dset.Tables.Count > 0) AndAlso dset.Tables(0).Rows.Count > 0) Then
            tbl = dset.Tables(0)
            For Each row In tbl.Rows
                cnt += 1
                If cnt Mod 1000 = 0 Then
                    txtprogress.Text = "Loaded " & cnt & " Customer records"
                    Me.Refresh()
                End If
                custNo = Trim(Microsoft.VisualBasic.Left(row("CUST_NO"), 20))
                name = Trim(Microsoft.VisualBasic.Left(row("NAME"), 50))
                fName = Trim(Microsoft.VisualBasic.Left(row("FIRST_NAME"), 30))
                lName = Trim(Microsoft.VisualBasic.Left(row("LAST_NAME"), 30))
                Addr1 = Trim(Microsoft.VisualBasic.Left(row("ADDRESS_1"), 40))
                Addr2 = Trim(Microsoft.VisualBasic.Left(row("ADDRESS_2"), 40))
                Addr3 = Trim(Microsoft.VisualBasic.Left(row("ADDRESS_3"), 40))
                city = Trim(Microsoft.VisualBasic.Left(row("CITY"), 30))
                state = Trim(Microsoft.VisualBasic.Left(row("STATE"), 10))
                zip = Trim(Microsoft.VisualBasic.Left(row("ZIP"), 15))
                spouse = Trim(Microsoft.VisualBasic.Left(row("SPOUSE"), 30))
                birthday = Trim(Microsoft.VisualBasic.Left(row("BIRTHDAY"), 20))
                phone1 = Trim(Microsoft.VisualBasic.Left(row("PHONE_1"), 15))
                phone2 = Trim(Microsoft.VisualBasic.Left(row("PHONE_2"), 15))
                email = Trim(Microsoft.VisualBasic.Left(row("EMAIL"), 50))
                type = Trim(Microsoft.VisualBasic.Left(row("TYPE"), 10))
                bal = Trim(Microsoft.VisualBasic.Left(row("BALANCE"), 20))
                fDate = Trim(Microsoft.VisualBasic.Left(row("FIRST_SALE_DATE"), 20))
                lDate = Trim(Microsoft.VisualBasic.Left(row("LAST_SALE_DATE"), 20))
                points = Trim(Microsoft.VisualBasic.Left(row("LOYALTY_POINTS"), 10))
                oktoemail = Trim(Microsoft.VisualBasic.Left(row("OK_TO_EMAIL"), 1))
                oktomail = Trim(Microsoft.VisualBasic.Left(row("OK_TO_MAIL"), 1))
                cell1 = Trim(Microsoft.VisualBasic.Left(row("CELL_1"), 15))
                cell2 = Trim(Microsoft.VisualBasic.Left(row("CELL_2"), 15))
                lUpdate = Trim(Microsoft.VisualBasic.Left(row("LAST_UPDATE"), 50))
                Dim row2 As DataRow = dtTbl.NewRow
                row2("Cust_No") = custNo
                row2("Name") = name
                row2("fName") = fName
                row2("lName") = lName
                row2("Addr1") = Addr1
                row2("Addr2") = Addr2
                row2("Addr3") = Addr3
                row2("City") = city
                row2("State") = state
                row2("Zip") = zip
                row2("Spouse") = spouse
                If Not IsDBNull(birthday) And IsNothing(birthday) Then
                    If birthday <> "" And birthday Like "##/##/##" Then
                        row2("Birthday") = CDate(birthday)
                    End If
                End If
                row2("Phone1") = phone1
                row2("Phone2") = phone2
                row2("eMail") = email
                row2("Type") = type
                If IsNumeric(bal) Then row2("Balance") = CDec(bal)
                If Not IsNothing(fDate) And fDate <> "" Then row2("fDate") = CDate(fDate)
                If Not IsNothing(lDate) And lDate <> "" Then row2("lDate") = CDate(lDate)
                If Not IsNothing(lUpdate) And lUpdate <> "" Then row2("lUpdate") = CDate(lUpdate)
                If IsNumeric(points) Then row2("Loyalty") = CInt(points)
                row2("OkToMail") = oktomail
                row2("OkToEmail") = oktoemail
                row2("Cell1") = cell1
                row2("Cell2") = cell2
                dtTbl.Rows.Add(row2)
            Next
            con.Open()
            sql = "DELETE FROM CUSTOMERS"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            txtprogress.Text = Format(cnt, "###,###,##0") & " Customers loaded."
            Me.Refresh()
            Dim connection As SqlConnection = New SqlConnection(conString)
            Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection)
            connection.Open()
            bulkCopy.DestinationTableName = "dbo.Customers"
            bulkCopy.BulkCopyTimeout = 120
            bulkCopy.WriteToServer(dtTbl)
            connection.Close()

        End If
    End Sub

    Private Sub DTP1_ValueChanged(sender As Object, e As EventArgs) Handles DTP1.ValueChanged
        Dim tempDate As Date = DTP1.Value
        Dim sdate As Date
        Call Connect_Database()
        con2.Open()
        sql = "SELECT sDate FROM Calendar WHERE '" & tempDate & "' BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            sdate = rdr("sDate")
        End While
        con2.Close()
        txtsDate.Text = sdate.ToString("MM/dd/yyyy")
        BuildStartDate = sdate
        Me.Refresh()
    End Sub

    Private Sub dtp_ValueChanged(sender As Object, e As EventArgs) Handles dtp.ValueChanged
        Dim tempDate As Date = dtp.Value
        Dim edate As Date
        Call Connect_Database()
        con2.Open()
        sql = "SELECT eDate FROM Calendar WHERE '" & tempDate & "' BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            edate = CDate(rdr("eDate"))
        End While
        txteDate.Text = edate.ToString("MM/dd/yyyy")
        con2.Close()
        Me.Refresh()
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        If txteDate.Text = Nothing Or txtsDate.Text = Nothing Then
            MsgBox("Select Inventory date and try again!!")
            Exit Sub
        End If
        BuildEndDate = CDate(txteDate.Text)
        invDateConfirmed = True
        btnAdj.Enabled = True
        GroupBox3.Visible = True
    End Sub

    Private Sub Log_Error(ByVal err As String)
        rcCon.Open()
        sql = "INSERT INTO ErrorLog(Client, Date, Error) " & _
            "SELECT '" & thisClient & "','" & Date.Now & "','" & err & "'"
        cmd = New SqlCommand(sql, rcCon)
        '' cmd.ExecuteNonQuery()
        rcCon.Close()
    End Sub

End Class

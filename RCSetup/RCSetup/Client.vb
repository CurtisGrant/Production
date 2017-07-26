Imports System
Imports System.Xml
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Globalization
Public Class Client
    Public Shared thisClient, conString, sql, dbName, servr As String
    Public Shared con As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Private Shared rcConnection As String
    Public Shared oTest As Object
    Public Shared serverTable As DataTable
    Public Shared DoForecasting As Boolean
    Public Shared DoScoring As Boolean
    Public Shared DoMarketing As Boolean
    Public Shared DoSalesPlan As Boolean
    Private Sub Client_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            txtClient.Text = DBAdmin.txtClient.Text
            txtStatus.Text = "Active"
            Call Edit_Client()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub New(ByVal server As String, ByVal database As String, client As String, ByVal RCConString As String, serverTable As DataTable)
        InitializeComponent()
        servr = server
        dbName = database
        thisClient = client
        txtClient.Text = client
        txtDatabase.Text = database
        rcConnection = RCConString
        If IsNothing(serverTable) Then
            cboServer.Items.Add(server)
            cboServer.SelectedIndex = 0
        Else
            For Each rw In serverTable.Rows
                cboServer.Items.Add(rw("ServerName"))
            Next
            Dim pos As Integer = cboServer.FindString(servr)
            cboServer.SelectedIndex = pos
        End If
    End Sub
    Private Sub Edit_Client()
        Try
            If IsNothing(thisClient) Then Exit Sub
            Dim server, user, pw, cString As String
            server = InitialSetup.txtServer.Text
            user = InitialSetup.txtUserId.Text
            pw = InitialSetup.txtPassword.Text
            ''cString = "Server=" & server & ";Initial Catalog=RCClient;User Id=" & user & ";Password=" & pw
            cString = "Server=" & server & ";Initial Catalog=RCClient;Integrated Security=True"
            Dim ccon As New SqlConnection(cString)
            ccon.Open()
            sql = "SELECT * FROM CLIENT_MASTER WHERE Client_ID = '" & thisClient & "'"
            cmd = New SqlCommand(sql, ccon)
            rdr = cmd.ExecuteReader
            While rdr.Read
                If Not IsNothing(rdr("Client_Id")) Then
                    txtClient.Text = rdr("Client_ID")
                    oTest = rdr("Name")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtName.Text = oTest Else txtName.Text = ""
                    oTest = rdr("Contact")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtContact.Text = oTest Else txtContact.Text = ""
                    oTest = rdr("Address1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtAddr1.Text = oTest Else txtAddr1.Text = ""
                    oTest = rdr("Address2")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtAddr2.Text = oTest Else txtAddr2.Text = ""
                    oTest = rdr("City")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtCity.Text = oTest Else txtCity.Text = ""
                    oTest = rdr("State")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtState.Text = oTest Else txtState.Text = ""
                    oTest = rdr("Zip")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtZip.Text = oTest Else txtZip.Text = ""
                    oTest = rdr("Phone1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtPhone1.Text = oTest Else txtPhone1.Text = ""
                    oTest = rdr("Phone2")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtPhone2.Text = oTest Else txtPhone2.Text = ""
                    oTest = rdr("email1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtemail1.Text = oTest Else txtemail1.Text = ""
                    oTest = rdr("email2")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtemail2.Text = oTest Else txtemail2.Text = ""
                    oTest = rdr("Database")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtDatabase.Text = oTest Else txtDatabase.Text = ""
                    oTest = rdr("URL")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtURL.Text = oTest Else txtURL.Text = ""
                    oTest = rdr("Server")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then servr = oTest Else servr = ""
                    oTest = rdr("SQLUserID")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtUserid.Text = oTest Else txtUserid.Text = ""
                    oTest = rdr("SQLPassword")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtPassword.Text = oTest Else txtPassword.Text = ""
                    ''oTest = rdr("SQLData")
                    ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtData.Text = oTest Else txtData.Text = ""
                    ''oTest = rdr("SQLLog")
                    ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtLog.Text = oTest Else txtLog.Text = ""
                    oTest = rdr("XMLs")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtPath.Text = oTest Else txtPath.Text = ""
                    oTest = rdr("Status")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then txtStatus.Text = oTest Else txtStatus.Text = ""
                    Me.Refresh()
                    dbName = txtDatabase.Text
                End If
            End While
            ccon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cboClient_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClient.SelectedIndexChanged
        Try
            thisClient = cboClient.SelectedItem
            Call Edit_Client()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Try
            thisClient = dbName
            DBAdmin.GroupBox5.Visible = True
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            If Not IsNothing(txtClient.Text) Then
                oTest = txtUserid.Text
                If IsNothing(txtUserid.Text) Or oTest = "" Then
                    MessageBox.Show("Please enter a SQL User Id and try again.", "ERROR!")
                    Exit Sub
                End If
                oTest = txtPassword.Text
                If IsNothing(txtPassword) Or oTest = "" Then
                    MessageBox.Show("Please enter as SQL Password and try again.", "ERROR!")
                    Exit Sub
                End If
                IsNothing(txtURL.Text)
                If IsNothing(oTest) Or oTest = "" Then
                    MessageBox.Show("Please enter or select the windows path to the XML files.", "ERROR!")
                    Exit Sub
                End If
                thisClient = txtClient.Text
                con = New SqlConnection(rcConnection)
                con.Open()
                Dim item4cast, marketing, salesplan As String
                If chkForecast.Checked Then item4cast = "Y" Else item4cast = Nothing
                If chkMarketing.Checked Then marketing = "Y" Else marketing = Nothing
                If chkSalesPlan.Checked Then salesplan = "Y" Else salesplan = Nothing
                sql = "IF NOT EXISTS (SELECT * FROM CLIENT_MASTER WHERE Client_Id = '" & thisClient & "') " & _
                    "INSERT INTO CLIENT_MASTER (Client_Id, Name, Contact, Address1, Address2, City, [State], Zip, Phone1, Phone2, email1, email2, " & _
                    "URL, [Server], [Database], SQLUserId, SQLPassword, XMLs, [Status], Item4Cast, Marketing, SalesPlan) " & _
                    "SELECT '" & thisClient & "', '" & txtName.Text & "', '" & txtContact.Text & "', '" & txtAddr1.Text & "', '" & txtAddr2.Text & "', '" & _
                    txtCity.Text & "', '" & txtState.Text & "', '" & txtZip.Text & "', '" & txtPhone1.Text & "', '" & txtPhone2.Text & "', '" & _
                    txtemail1.Text & "', '" & txtemail2.Text & "', '" & txtURL.Text & "', '" & servr & "', '" & txtDatabase.Text & "', '" & _
                    txtUserid.Text & "','" & txtPassword.Text & "','" & _
                    txtPath.Text & "', '" & txtStatus.Text & "', '" & item4cast & "', '" & _
                    marketing & "', '" & salesplan & "' " & _
                    "ELSE " & _
                    "UPDATE CLIENT_MASTER SET Name = '" & txtName.Text & "', Contact = '" & txtContact.Text & "', Address1 = '" & txtAddr1.Text & "', " & _
                    "Address2 = '" & txtAddr2.Text & "', City = '" & txtCity.Text & "', State = '" & txtState.Text & "', Zip = '" & txtZip.Text & "', " & _
                    "Phone1 = '" & txtPhone1.Text & "', Phone2 = '" & txtPhone2.Text & "', email1 = '" & txtemail1.Text & "', " & _
                    "SQLUserId = '" & txtUserid.Text & "', SQLPassword = '" & txtPassword.Text & "', " & _
                    "email2 = '" & txtemail2.Text & "', URL = '" & txtURL.Text & "', [Server] = '" & servr & "', [Database] = '" & _
                    txtDatabase.Text & "', XMLs = '" & txtPath.Text & "', " & _
                    "Status = '" & txtStatus.Text & "', Item4Cast = '" & item4cast & "', Marketing = '" & marketing & "', " & _
                    "SalesPlan = '" & salesplan & "' " & _
                    "WHERE Client_Id = '" & thisClient & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()

                Call Create_Database()

                MessageBox.Show("RUN THE APPROPRIATE SQL BUILD SCRIPTS TO CREATE TABLES, SPs AND FUNCTIONS", txtDatabase.Text & " BUILT!")
                ''conString = "Server=" & servr & ";Initial Catalog=" & txtDatabase.Text & ";User Id=" & txtUserid.Text & ";Password=" & txtPassword.Text
                conString = "Server=" & servr & ";Initial Catalog=" & txtDatabase.Text & ";Integrated Security=True"
                con = New SqlConnection(conString)
                con.Open()
                ''cmd = New SqlCommand("sp_RCSetup_CreateModule1Tables", con)
                ''cmd.CommandType = CommandType.StoredProcedure
                ''cmd.ExecuteNonQuery()
                sql = "CREATE NONCLUSTERED INDEX IX_DATE ON DAILY_TRANSACTION_LOG ([DATE])"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()

                MessageBox.Show("Click Exit to continue", "Module 1 Tables Created")
                ''txtClient.Text = Nothing
                ''txtName.Text = Nothing
                ''txtAddr1.Text = Nothing
                ''txtAddr2.Text = Nothing
                ''txtCity.Text = Nothing
                ''txtState.Text = Nothing
                ''txtZip.Text = Nothing
                ''txtPhone1.Text = Nothing
                ''txtPhone2.Text = Nothing
                ''txtemail1.Text = Nothing
                ''txtemail2.Text = Nothing
                ''txtDatabase.Text = Nothing
                ''txtUserid.Text = Nothing
                ''txtPassword.Text = Nothing
                ''txtData.Text = Nothing
                ''txtLog.Text = Nothing
                ''txtPath.Text = Nothing
                ''Me.Refresh()
                ' Me.Close()
            Else
                MessageBox.Show("Enter Client Id and try again.", "ERROR!")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cboServer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboServer.SelectedIndexChanged
        Try
            'If thisClient Is Nothing Or thisClient = "" Then
            '    Dim ans As String = InputBox("Enter name", "Enter an ID for the new client", Nothing)
            '    If ans Is Nothing OrElse ans = "" Then
            '        MsgBox("Name is required. Exiting subroutine!")
            '        Exit Sub
            '    End If
            '    thisClient = ans
            'End If
           

            servr = cboServer.SelectedItem
            Me.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Create_Database()
        conString = "Server=" & servr & ";Initial Catalog=Master;Integrated Security=True"
        Dim con As New SqlConnection(conString)
        dbName = txtDatabase.Text

        con.Open()
        sql = "IF NOT EXISTS (SELECT * FROM SYS.DATABASES WHERE NAME = '" & dbName & "') " & _
            "CREATE DATABASE " & dbName & ""
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        txtDatabase.Text = dbName

        con.Open()
        sql = "SELECT physical_name AS Path FROM sys.master_files"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr(0)
            If InStr(oTest, dbName & ".mdf") Then txtData.Text = rdr(0)
            If InStr(oTest, dbName & "_log.ldf") Then txtLog.Text = rdr(0)
            Me.Refresh()
        End While
        con.Close()
    End Sub

    'Private Sub txtClient_TextChanged(sender As Object, e As EventArgs) Handles txtClient.TextChanged
    '    Try
    '        thisClient = txtClient.Text
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub btnGetPath_Click_1(sender As Object, e As EventArgs) Handles btnGetPath.Click
        Try
            Dim xmlPath As String
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
            If (result = DialogResult.OK) Then
                xmlPath = FolderBrowserDialog1.SelectedPath
                ''  xmlPath &= "\" & thisClient
                If (Not System.IO.Directory.Exists(xmlPath)) Then
                    System.IO.Directory.CreateDirectory(xmlPath)
                    Threading.Thread.Sleep(1000)
                End If
                txtPath.Text = xmlPath
                Me.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub btnGetTemplates_Click(sender As Object, e As EventArgs) Handles btnGetTemplates.Click
        Try
            Dim Path As String
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
            If (result = DialogResult.OK) Then
                Path = FolderBrowserDialog1.SelectedPath
                If (Not System.IO.Directory.Exists(Path)) Then
                    System.IO.Directory.CreateDirectory(Path)
                    Threading.Thread.Sleep(1000)
                End If
                txtTemplates.Text = Path
                Me.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ''Private Sub btnErrorLog_Click(sender As Object, e As EventArgs) Handles btnErrorLog.Click
    ''    Try
    ''        Dim Path As String
    ''        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
    ''        If (result = DialogResult.OK) Then
    ''            Path = FolderBrowserDialog1.SelectedPath
    ''            If (Not System.IO.Directory.Exists(Path)) Then
    ''                System.IO.Directory.CreateDirectory(Path)
    ''                Threading.Thread.Sleep(1000)
    ''            End If
    ''            txtErrorLog.Text = Path
    ''            Me.Refresh()
    ''        End If

    ''    Catch ex As Exception
    ''        MsgBox(ex.Message)
    ''    End Try
    ''End Sub


    Private Sub btnReports_Click(sender As Object, e As EventArgs) Handles btnReports.Click
        Try
            Dim Path As String
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
            If (result = DialogResult.OK) Then
                Path = FolderBrowserDialog1.SelectedPath
                If (Not System.IO.Directory.Exists(Path)) Then
                    System.IO.Directory.CreateDirectory(Path)
                    Threading.Thread.Sleep(1000)
                End If
                txtReports.Text = Path
                Me.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub chkForecast_CheckedChanged(sender As Object, e As EventArgs) Handles chkForecast.CheckedChanged
        If chkForecast.Checked Then
            DoForecasting = True
            DBAdmin.btnForecast.Enabled = True
        Else : DoForecasting = False
        End If
    End Sub

    Private Sub chkScoring_CheckedChanged(sender As Object, e As EventArgs) Handles chkScoring.CheckedChanged
        If chkScoring.Checked Then
            DoScoring = True
            DBAdmin.btnScore.Enabled = True
        Else : DoScoring = False
        End If
    End Sub

    Private Sub chkMarketing_CheckedChanged(sender As Object, e As EventArgs) Handles chkMarketing.CheckedChanged
        If chkMarketing.Checked Then
            DoMarketing = True
            DBAdmin.btnCustomers.Enabled = True
        Else : DoMarketing = False
        End If
    End Sub

    Private Sub chkSalesPlan_CheckedChanged(sender As Object, e As EventArgs) Handles chkSalesPlan.CheckedChanged
        If chkSalesPlan.Checked Then
            DoSalesPlan = True
            DBAdmin.btnSalesPlan.Enabled = True
        Else : DoSalesPlan = False
        End If
    End Sub

End Class
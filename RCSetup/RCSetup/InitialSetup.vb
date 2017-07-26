Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net
Imports System.Xml
Public Class InitialSetup
    Private userId, server, password As String
    Private oTest As Object
    Public serverTable As DataTable
    Public RCClientServer, RCClientConString, exePath As String

    Private Sub InitialSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("Search network for SQL Servers")
        'txtprogress.Text = "Searching network for available servers."
        Me.Refresh()
        Dim theXMLfile As String = "c:/RetailClarity/RCCLIENT.xml"
        If My.Computer.FileSystem.FileExists(theXMLfile) Then
            Dim rcClientExists As Boolean = False
            Dim xmlReader As XmlTextReader = New XmlTextReader("c:/RetailClarity/RCCLIENT.xml")
            Dim exePath As String = ""
            Dim client As String = ""
            Dim errors As String = ""
            Dim exes As String = ""
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
                If fld = "USERID" Then userId = valu
                If fld = "PD" Then password = valu
                If fld = "CLIENTID" Then client = valu
                If fld = "ERRORPATH" Then errors = valu
                If fld = "EXEPATH" Then exes = valu
            End While
            xmlReader.Close()
            txtServer.Text = server
            txtUserId.Text = userId
            txtPassword.Text = password
            txtErrors.Text = errors
            txtExe.Text = exes
        End If
        '
        '   Comment line below for CURTIS-MOBILE only!
        '
        Call Find_Servers()
    End Sub

    Private Sub Find_Servers()
        Me.Refresh()

        serverTable = System.Data.Sql.SqlDataSourceEnumerator.Instance.GetDataSources()
        For Each rw In serverTable.Rows
            If rw("InstanceName").ToString = "" Then
                Me.cboServer.Items.Add(rw("ServerName"))
            Else
                Me.cboServer.Items.Add(rw("ServerName").ToString & "/" & rw("InstanceName").ToString)
            End If
        Next
    End Sub

    Private Sub cboServer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboServer.SelectedIndexChanged
        server = cboServer.SelectedItem
        txtServer.Text = server
        Dim ans As String = InputBox("Eneter SQL User Id for " & server)
        If Not IsNothing(ans) Then txtUserId.Text = ans
        ans = InputBox("Enter SQL Password for " & server)
        If Not IsNothing(ans) Then txtPassword.Text = ans
    End Sub

    Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click
        Try
            Dim oTest As Object
            Dim constring, sql As String
            Dim con As SqlConnection
            Dim cmd As SqlCommand
            Dim rdr As SqlDataReader
            Dim rcClientExists As Boolean = False
            ''If IsNothing(serverTable) Then

            ''End If
            ''For Each rw In serverTable.Rows
            ''    oTest = rw("ServerName")
            ''    If oTest = server Then
            ''        userId = txtUserId.Text
            ''        password = txtPassword.Text
            ''        ''constring = "Server=" & server & ";Initial Catalog=Master;User Id=" & userId & ";Password=" & password
            ''        constring = "Server=" & server & ";Initial Catalog=Master;Integrated Security=True"
            ''        con = New SqlConnection(constring)
            ''        con.Open()
            ''        sql = "SELECT name FROM Master.dbo.SysDatabases WHERE name = 'RCClient'"
            ''        cmd = New SqlCommand(sql, con)
            ''        rdr = cmd.ExecuteReader
            ''        While rdr.Read
            ''            If rdr("name") = "RCClient" Then rcClientExists = True
            ''        End While
            ''        con.Close()
            ''    End If
            ''Next

            RCClientServer = server
            ''  RCClientConString = "Server=" & txtServer.Text & ";Initial Catalog=RCClient;User Id=" & txtUserId.Text & ";Password=" & txtPassword.Text
            RCClientConString = "Server=" & server & ";Initial Catalog=RCClient;Integrated Security=True"
            exePath = txtExe.Text
            DBAdmin.exePath = exePath
            DBAdmin.Show()

        Catch ex As Exception
            MsgBox("ERROR LOGGING IN" & ex.Message)
        End Try
    End Sub

End Class
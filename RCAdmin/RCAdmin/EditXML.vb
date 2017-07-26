Imports System.Xml
Imports System.Data.SqlClient
Public Class EditXML
    Public Shared con As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared tbl As DataTable
    Public Shared wks As Integer
    Public Shared sDate, eDate, xDate As Date
    Public Shared oTest As Object
    Public Shared server, password, client, log, exe As String
    Public Shared database As String
    Public Shared changesMade As Boolean
    Private Sub EditXML_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            tbl = New DataTable
            tbl.Columns.Add("Client_Id", GetType(System.String))
            tbl.Columns.Add("Status", GetType(System.String))
            tbl.Columns.Add("Server", GetType(System.String))
            tbl.Columns.Add("Database", GetType(System.String))
            tbl.Columns.Add("XMLs", GetType(System.String))
            tbl.Columns.Add("Templates", GetType(System.String))
            tbl.Columns.Add("Reports", GetType(System.String))
            tbl.Columns.Add("SQLUserID", GetType(System.String))
            tbl.Columns.Add("SQLPassword", GetType(System.String))
            changesMade = False
            serverLabel.Text = MainMenu.serverLabel.Text
            Dim conString As String
            Dim path As String = "c:\RCCLIENT.xml"
            Dim xmlReader As XmlTextReader = New XmlTextReader(path)

            Dim fld As String = ""
            Dim val As String = ""
            While xmlReader.Read
                Select Case xmlReader.NodeType
                    Case XmlNodeType.Element
                        fld = xmlReader.Name
                    Case XmlNodeType.Text
                        val = xmlReader.Value
                    Case XmlNodeType.EndElement
                        If fld = "SERVER" Then
                            server = val
                            txtServer.Text = val
                        End If
                        If fld = "PD" Then
                            txtPassword.Text = val
                            password = val
                        End If
                        If fld = "ERRORPATH" Then
                            txtLog.Text = val
                            log = val
                        End If
                        If fld = "EXEPATH" Then
                            txtExe.Text = val
                            exe = val
                        End If
                        If fld = "CLIENT" Then
                            txtClient.Text = val
                            client = val
                        End If
                End Select
            End While
            xmlReader.Close()

            conString = MainMenu.masterConString
            con = New SqlConnection(conString)
            con.Open()
            Dim sql As String
            Sql = "SELECT Client_Id, Status, Server, Database, XMLs, Templates, Reports FROM Client_Master ORDER BY Client_Id"
            cmd = New SqlCommand(Sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                Dim row As DataRow = tbl.NewRow
                row("Server") = rdr("Server")
                row("Client_Id") = rdr("Client_Id")
                row("Status") = rdr("Status")
                row("Database") = rdr("Database")
                row("XMLs") = rdr("XMLs")
                row("Templates") = rdr("Templates")
                row("Reports") = rdr("Reports")
                row("SQLUserID") = rdr("SQLUserID")
                row("SQLPassword") = rdr("SQLPassword")
                tbl.Rows.Add(row)
            End While
            con.Close()
            dgv1.DataSource = tbl.DefaultView
            dgv1.AutoResizeColumns()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("c:\RCClient.xml")
        Dim node As XmlNode = xmlDoc.SelectSingleNode("/PARAMETERS/SERVER")
        If node IsNot Nothing Then
            node.ChildNodes(0).InnerText = server
            xmlDoc.Save("c:\RCClient.xml")
        End If
        node = xmlDoc.SelectSingleNode("/PARAMETERS/PD")
        If node IsNot Nothing Then
            node.ChildNodes(0).InnerText = password
            xmlDoc.Save("c:\RCClient.xml")
        End If
        node = xmlDoc.SelectSingleNode("/PARAMETERS/ERRORPATH")
        If node IsNot Nothing Then
            node.ChildNodes(1).InnerText = log
            xmlDoc.Save("c:\RCClient.xml")
        End If
        node = xmlDoc.SelectSingleNode("/PARAMETERS/EXEPATH")
        If node IsNot Nothing Then
            node.ChildNodes(2).InnerText = exe
            xmlDoc.Save("c:\RCClient.xml")
        End If
        node = xmlDoc.SelectSingleNode("/PARAMETERS/CLIENTID")
        If node IsNot Nothing Then
            node.ChildNodes(3).InnerText = client
            xmlDoc.Save("c:\RCClient.xml")
        End If
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                If changesMade Then
                    Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
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

    Private Sub txtServer_TextChanged(sender As Object, e As EventArgs) Handles txtServer.TextChanged
        server = txtServer.Text
    End Sub

    Private Sub txtPassword_TextChanged(sender As Object, e As EventArgs) Handles txtPassword.TextChanged
        password = txtPassword.Text
    End Sub

    Private Sub txtClient_TextChanged(sender As Object, e As EventArgs) Handles txtClient.TextChanged
        client = txtClient.Text
    End Sub

    Private Sub txtLog_TextChanged(sender As Object, e As EventArgs) Handles txtLog.TextChanged
        log = txtLog.Text
    End Sub

    Private Sub txtExe_TextChanged(sender As Object, e As EventArgs) Handles txtExe.TextChanged
        exe = txtExe.Text
    End Sub
End Class
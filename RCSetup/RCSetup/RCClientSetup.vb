Imports System.Data.SqlClient
Public Class RCClientSetup
    Private conString As String
    Private con As SqlConnection
    Private cmd As SqlCommand
    Private rdr As SqlDataReader
    Private sql As String
    Public Shared masterServer, masterUser, masterPassword As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        masterServer = txtServer.Text
        masterUser = txtUser.Text
        masterPassword = txtPassword.Text

        conString = "Server=" & masterServer & ";Initial Catalog=Master;User Id=" & masterUser & ";Password=" & masterPassword
        con = New SqlConnection(conString)
        con.Open()
        sql = "USE Master; " & _
                "IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'RCClient') " & _
                "CREATE DATABASE RCClient; "
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

        conString = "Server=" & masterServer & ";Initial Catalog=RCClient;User Id=" & masterUser & ";Password=" & masterPassword
        con = New SqlConnection(conString)
        con.Open()
        sql = "IF NOT EXISTS (SELECT * FROM Client_Master) " & _
                "CREATE TABLE [dbo].[Client_Master](" & _
                "[Client_Id] [varchar](30) NOT NULL," & _
                "[Status] [varchar](10) NULL," & _
                "[Last_XML_Update] [datetime] NULL," & _
                "[Last_Summary_Update] [datetime] NULL," & _
                "[Last_Scored] [datetime] NULL," & _
                "[Last_Forecast] [datetime] NULL," & _
                "[Server] [varchar](50) NULL," & _
                "[Database] [varchar](30) NULL," & _
                "[XMLs] [varchar](max) NULL," & _
                "[Templates] [varchar](max) NULL," & _
                "[Reports] [varchar](max) NULL," & _
                "[SQLUserID] [varchar](20) NULL," & _
                "[SQLPassword] [varchar](20) NULL," & _
                "[Name] [varchar](50) NULL," & _
                "[Contact] [varchar](30) NULL," & _
                "[Address1] [varchar](50) NULL," & _
                "[Address2] [varchar](50) NULL," & _
                "[City] [varchar](40) NULL," & _
                "[State] [varchar](2) NULL," & _
                "[Zip] [varchar](10) NULL," & _
                "[Phone1] [varchar](15) NULL," & _
                "[Phone2] [varchar](15) NULL," & _
                "[email1] [varchar](max) NULL," & _
                "[email2] [varchar](max) NULL," & _
                "[URL] [varchar](max) NULL, " & _
                "[Item4Cast] [varchar](1) NULL, " & _
                "[Marketing] [varchar](1) NULL, " & _
                "[SalesPlan] [varchar](1) NULL, " & _
                "CONSTRAINT [PK_Client_Master] PRIMARY KEY NONCLUSTERED " & _
                "([Client_Id] ASC) " & _
                "WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) " & _
                "ON [PRIMARY]"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub
End Class
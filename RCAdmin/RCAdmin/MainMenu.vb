Imports System.Data.SqlClient
Imports System.Xml

Public Class MainMenu
    Public Shared conString, masterConString As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared tableName As String
    Public Shared Client_Id As String
    Public Shared Server As String
    Public Shared sql As String
    Public Shared dBase As String
    Public Shared UserID As String
    Public Shared Password As String
    Public Shared exePath As String
    Public Shared Item4CastOK, MarketingOK, SalesPlanOK As Boolean

    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RetailClarity\RCCLIENT.xml")
        Dim passWord As String = ""
        Dim client As String = ""
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
        masterConString = "Server=" & Server & ";Initial Catalog=RCClient;Integrated Security=True"
        con = New SqlConnection(masterConString)
        con.Open()
        sql = "SELECT Client_Id FROM Client_Master ORDER BY Client_ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboClient.Items.Add(rdr("Client_Id"))
        End While
        con.Close()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnBuyers_Click(sender As Object, e As EventArgs) Handles btnBuyers.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "Buyers"
        ''TableMaintenance.Show()
        ''TableMaintenance = Nothing
        Dim oForm As TableMaintenance
        oForm = New TableMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnClasses_Click(sender As Object, e As EventArgs) Handles btnClasses.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "Classes"
        ''ClassMaintenance.Show()
        ''ClassMaintenance = Nothing
        Dim oForm As ClassMaintenance
        oForm = New ClassMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnDept_Click(sender As Object, e As EventArgs) Handles btnDept.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "Departments"
        ''TableMaintenance.Show()
        ''TableMaintenance = Nothing
        Dim oForm As TableMaintenance
        oForm = New TableMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnMkdn_Click(sender As Object, e As EventArgs) Handles btnMkdn.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        ''Dim oForm As MkdnMaintenance
        ''oForm = New MkdnMaintenance()
        ''oForm.Show()
        ''oForm = Nothing
        MarkdownPlan.Show()
    End Sub

    Private Sub btnUpdatePlan_Click_1(sender As Object, e As EventArgs) Handles btnUpdatePlan.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        If Not SalesPlanOK Then
            Dim message As String = "CONTACT RETAIL CLARITY"
            Dim caption As String = "ERROR! SALES PLANNING NOT SET UP"
            MessageBox.Show(message, caption)
            Exit Sub
        End If
        Dim oForm As PlanMaintenance
        oForm = New PlanMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnEditXML_Click(sender As Object, e As EventArgs)
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        Dim oForm As EditXML
        oForm = New EditXML()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnStore_Click(sender As Object, e As EventArgs) Handles btnStore.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "Stores"
        Dim oForm As TableMaintenance
        oForm = New TableMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnWeeksSupply_Click(sender As Object, e As EventArgs) Handles btnWeeksSupply.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        Dim oForm As WeeksOnHand
        oForm = New WeeksOnHand()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnAdjustDay_Click(sender As Object, e As EventArgs) Handles btnAdjustDay.Click
        Try
            If Len(Client_Id) = 0 Then
                MsgBox("Select a client and try again!")
                Exit Sub
            End If
            Dim oForm As Days2
            oForm = New Days2()
            oForm.Show()
            oForm = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBuyer_Click(sender As Object, e As EventArgs) Handles btnBuyer.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        lblProcessing.Text = "Computing Averages. Please Wait."
        lblProcessing.Visible = True
        Me.Refresh()
        Dim oForm As BuyerPct
        oForm = New BuyerPct()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnDayPct_Click(sender As Object, e As EventArgs) Handles btnDayPct.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        ''DayPct.Show()
        Dim oForm As DayPct
        oForm = New DayPct()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnClassPct_Click(sender As Object, e As EventArgs) Handles btnClassPct.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        Dim oForm As ClassPct
        oForm = New ClassPct()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnScore_Click(sender As Object, e As EventArgs) Handles btnScore.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        ''Dim oForm As Scoring
        ''oForm = New Scoring()
        Dim oForm As ScoringSetup
        oForm = New ScoringSetup()
        oForm.Show()
    End Sub

    Private Sub cboClient_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClient.SelectedIndexChanged
        Client_Id = cboClient.SelectedItem
        Dim mcon As New SqlConnection(masterConString)
        Dim oTest As Object
        Dim val As String
        mcon.Open()
        sql = "SELECT Server, [Database], SQLUserID, SQLPassword, Item4Cast, Marketing, SalesPlan FROM Client_Master " & _
            "WHERE Client_Id = '" & Client_Id & "'"
        cmd = New SqlCommand(sql, mcon)
        rdr = cmd.ExecuteReader
        While rdr.Read
            Server = rdr("Server")
            dBase = rdr("Database")
            UserID = rdr("SQLUserID")
            Password = rdr("SQLPassword")
            oTest = rdr("Item4Cast")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then val = CStr(oTest)
            If val = "Y" Then Item4CastOK = True Else Item4CastOK = False
            oTest = rdr("Marketing")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then val = CStr(oTest)
            If val = "Y" Then MarketingOK = True Else MarketingOK = False
            oTest = rdr("SalesPlan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then val = CStr(oTest)
            If val = "Y" Then SalesPlanOK = True Else SalesPlanOK = False
        End While
        mcon.Close()
        conString = "Server=" & Server & ";Initial Catalog=" & dBase & ";Integrated Security=True"
        serverLabel.Text = "Server: " & Server & " Database: " & dBase
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        con4 = New SqlConnection(conString)
        con5 = New SqlConnection(conString)
    End Sub

    Private Sub btnControls_Click(sender As Object, e As EventArgs) Handles btnControls.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        Dim oForm As ControlsAdmin
        oForm = New ControlsAdmin
        oForm.Show()
    End Sub

    Private Sub btnPLine_Click(sender As Object, e As EventArgs) Handles btnPLine.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "ProductLines"
        Dim oForm As TableMaintenance
        oForm = New TableMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnSeason_Click(sender As Object, e As EventArgs) Handles btnSeason.Click
        If Len(Client_Id) = 0 Then
            MsgBox("Select a client and try again!")
            Exit Sub
        End If
        tableName = "Seasons"
        Dim oForm As TableMaintenance
        oForm = New TableMaintenance()
        oForm.Show()
        oForm = Nothing
    End Sub

    Private Sub btnForecast_Click(sender As Object, e As EventArgs) Handles btnForecast.Click
        If Not Item4CastOK Then
            Dim message As String = "CONTACT RETAIL CLARITY"
            Dim caption As String = "ERROR! ITEM FORECASTING NOT SET UP"
            MessageBox.Show(message, caption)
            Exit Sub
        End If
        Dim dbase, userid, password As String
        Dim strng() As String = Split(conString, ";")
        Server = strng(0).Replace("Server=", "")
        dbase = strng(1).Replace("Initial Catalog=", "")
        userid = strng(2).Replace("User Id=", "")
        password = strng(3).Replace("Password=", "")
        Dim p As New ProcessStartInfo
        p.FileName = exePath & "\ItemForecast.exe"
        p.Arguments = Client_Id & ";" & Server & ";" & dbase & ";" & userid & ";" & password
        p.UseShellExecute = True
        p.WindowStyle = ProcessWindowStyle.Normal
        Dim proc As Process = Process.Start(p)
    End Sub
End Class

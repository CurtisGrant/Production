

Imports System.Data.SqlClient
Imports System.Xml
Module Module1
    Private con As SqlConnection
    Private cmd As SqlCommand
    Private rdr As SqlDataReader
    Private oTest As Object

    Sub Main()
        Dim thePath As String = "c:\RetailClarity\XMLs\TCM\Sales.xml"
        Dim conString As String = "Server=RetailClarity1;Initial Catalog=TCMarketing;Integrated Security=True"
        con = New SqlConnection(conString)
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim xmlFile As XmlReader
        Dim row As DataRow
        Dim id, sku, item, dim1, dim2, dim3, loc, store, sql As String
        Dim seq As Integer
        Dim qty, cost, retail, mkdn As Decimal
        Dim transdate As Date
        xmlFile = Xml.XmlReader.Create(thePath, New XmlReaderSettings())
        ds.ReadXml(xmlFile)                                                                  '  bulk insert XML into dataset
        If ((ds.Tables.Count > 0) AndAlso ds.Tables(0).Rows.Count > 0) Then
            dt = ds.Tables(0)
            con.Open()
            For Each row In dt.Rows
                oTest = row("SKU")
                oTest = row("TRANS_ID")
                If Not IsDBNull(oTest) Then id = oTest Else id = ""
                oTest = row("SEQ_NO").ToString
                If Not IsDBNull(oTest) Then seq = CInt(oTest) Else seq = 0
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
                oTest = Trim(row("LOCATION"))
                If Not IsDBNull(oTest) Then Loc = CStr(oTest)
                oTest = row("QTY")
                qty = 0
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then qty = oTest
                End If
                cost = 0
                oTest = row("COST")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then cost = oTest
                End If
                retail = 0
                oTest = row("RETAIL")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then retail = oTest
                End If
                oTest = row("TRANS_DATE")
                If oTest = "0" Then oTest = "1900-01-01 00:00:00"
                transdate = CDate(oTest)
                oTest = row("STR_ID")
                If IsDBNull(oTest) Then store = "UNKNOWN" Else store = CStr(oTest)
                oTest = row("MARKDOWN")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) And oTest <> "" Then mkdn = CDec(oTest) Else mkdn = 0
                sql = "INSERT INTO SalesXML(TRANS_ID, SEQ_NO, SKU, LOCATION, STORE, TRANS_DATE) " & _
                    "SELECT '" & id & "'," & seq & ",'" & sku & "','" & loc & "','" & store & "','" & transdate & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
            Next
            con.Close()
        End If
    End Sub

End Module

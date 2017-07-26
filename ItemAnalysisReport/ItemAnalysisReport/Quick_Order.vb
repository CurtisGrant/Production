Imports System
Imports System.Data
Imports System.Text
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Microsoft
Imports Microsoft.VisualBasic
Imports excel = Microsoft.Office.Interop.Excel           ' add .net reference under Projects
Imports System.ComponentModel
Imports System.Xml
Public Class Quick_Order
    Public Shared con As SqlConnection = Item_Analysis.con
    Public Shared thisUser As String = Environment.UserName.ToString
    Public Shared server, database, path, template, thisVendor, thisDateString, store, displayStores As String
    Public Shared displayTable As DataTable
    Public Shared editTable As DataTable
    Public Shared Table1 As DataTable
    Public Shared Table2 As DataTable
    Public Shared Table3 As DataTable
    Public Shared Store1_Has_Items As Boolean = False
    Public Shared Store2_Has_Items As Boolean = False

    Private Sub QuickOrder_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(100, 150)
    End Sub

    Friend Sub QuickOrder(ByVal Vndr, ByVal dte, ByVal selectStores, ByVal tbl)
        Try
            Dim sql As String
            Dim cmd As SqlCommand
            Dim rdr As SqlDataReader
            If btnMerged.Checked = True Then gbShipTo.Visible = True
            Dim oTest As Object
            thisVendor = Vndr
            txtVendor.Text = thisVendor.ToString
            thisDateString = dte
            displayTable = tbl

            displayStores = selectStores
            Table1 = New DataTable
            Table2 = New DataTable
            Dim vItem As String
            Dim r As Integer = displayTable.Rows.Count
            Dim c As Integer = displayTable.Columns.Count
            editTable = New DataTable
            editTable = displayTable.Clone
            Dim tblRow As DataRow

            Table1 = New DataTable
            Dim key1(1) As DataColumn
            Dim column1 As New DataColumn()
            column1.DataType = System.Type.GetType("System.String")
            column1.ColumnName = "key1"
            Table1.Columns.Add(column1)
            key1(0) = column1
            Table1.PrimaryKey = key1
            Table1.Columns.Add("Store", GetType(System.String))
            Table1.Columns.Add("VendorItem_No", GetType(System.String))
            Table1.Columns.Add("Our_No", GetType(System.String))
            Table1.Columns.Add("UOM", GetType(System.String))
            Table1.Columns.Add("Description", GetType(System.String))
            Table1.Columns.Add("Qty", GetType(System.Decimal))
            Table1.Columns.Add("Cost", GetType(System.Decimal))
            Table1.Columns.Add("Merged", GetType(System.String))
            Dim tbl1Row As DataRow

            Table2 = New DataTable
            Dim key2(1) As DataColumn
            Dim column2 As New DataColumn()
            column2.DataType = System.Type.GetType("System.String")
            column2.ColumnName = "key2"
            Table2.Columns.Add(column2)
            key2(0) = column2
            Table2.PrimaryKey = key2
            Table2.Columns.Add("Store", GetType(System.String))
            Table2.Columns.Add("VendorItem_No", GetType(System.String))
            Table2.Columns.Add("Our_No", GetType(System.String))
            Table2.Columns.Add("UOM", GetType(System.String))
            Table2.Columns.Add("Description", GetType(System.String))
            Table2.Columns.Add("Qty", GetType(System.Decimal))
            Table2.Columns.Add("Cost", GetType(System.Decimal))
            Table2.Columns.Add("Merged", GetType(System.String))
            Dim tbl2Row As DataRow
            Dim tCost As Decimal = 0
            Dim cost, qty As Decimal
            For Each row In displayTable.Rows
                oTest = row.item("ordQty")
                Dim i As Integer
                If Not IsDBNull(oTest) Then
                    If (Integer.TryParse(oTest, i)) Then
                        If Convert.ToInt32(oTest) > 0 Then
                            con.Open()
                            Dim itm As String = row.item(2)
                            Dim buyunit As Integer
                            sql = "SELECT Vend_Item_No, BuyUnit FROM Item_Master WHERE Item_No = '" & itm & "'"
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            vItem = itm
                            While rdr.Read
                                oTest = rdr(0)
                                If Not IsDBNull(oTest) Then vItem = rdr(0)
                                oTest = rdr("BuyUnit")
                                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyunit = oTest Else buyunit = 1
                            End While
                            con.Close()
                            qty = row.item("OrdQty")
                            cost = row.item("Cost")
                            tCost += (qty * cost) * buyunit
                            store = row.item(1)
                            tblRow = editTable.NewRow                                ' save editTable so it can be saved to xlsx later
                            For i = 0 To displayTable.Columns.Count - 1
                                tblRow.Item(i) = row.item(i)
                            Next
                            editTable.Rows.Add(tblRow)
                            If store = "2" Then                                       ' create a new row table2 row for NP qty
                                Store2_Has_Items = True
                                tbl2Row = Table2.NewRow()
                                tbl2Row.Item("key2") = row.item("Item_No")      ' our Item#
                                tbl2Row.Item("Store") = row.item("Str_Id")      ' Store
                                tbl2Row.Item("VendorItem_No") = vItem
                                tbl2Row.Item("Our_No") = row.item("Item_No")      ' Our Item#
                                tbl2Row.Item("UOM") = row.item("Pur_Unit")      ' UOM
                                tbl2Row.Item("Description") = row.item("Descr")      ' Description
                                tbl2Row.Item("Qty") = row.item("OrdQty")      ' Qty
                                tbl2Row.Item("Cost") = row.item("Cost") * CDec(buyunit)    ' Cost
                                tbl2Row.Item("Merged") = ""
                                Table2.Rows.Add(tbl2Row)
                            Else
                                tbl1Row = Table1.NewRow()
                                Store1_Has_Items = True
                                tbl1Row.Item("key1") = row.item("Item_No")      ' Our Item# was 2
                                tbl1Row.Item("Store") = row.item("Str_Id")      ' Store was 1
                                tbl1Row.Item("VendorItem_No") = vItem
                                tbl1Row.Item("Our_No") = row.item("Item_No")      ' Our Item# was 2
                                tbl1Row.Item("UOM") = row.item("Pur_Unit")      ' UOM was 4
                                tbl1Row.Item("Description") = row.item("Descr")         ' Qty was 3
                                tbl1Row.Item("Qty") = row.item("OrdQty")
                                tbl1Row.Item("Cost") = row.item("Cost") * CDec(buyunit)     ' Cost was 5
                                tbl1Row.Item("Merged") = ""
                                Table1.Rows.Add(tbl1Row)
                            End If
                        End If
                    End If
                End If
            Next
            If store = "2" Then
                tbl2Row = Table2.NewRow
                tbl2Row.Item("key2") = ""
                tbl2Row.Item("VendorItem_No") = " TOTAL"
                tbl2Row.Item("Cost") = tCost
                Table2.Rows.Add(tbl2Row)
            Else
                tbl1Row = Table1.NewRow
                tbl1Row.Item("key1") = ""
                tbl1Row.Item("VendorItem_No") = " TOTAL"
                tbl1Row.Item("Cost") = tCost
                Table1.Rows.Add(tbl1Row)
            End If
            dg7.DataSource = Table1.DefaultView
            '' dg7.DataSource = Table1.DefaultView
            dg7.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dg7.Columns(0).Visible = False                                                           ' KeyId
            dg7.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dg7.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dg7.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Merge_Qtys()
        Dim foundRow As DataRow
        Dim itm As String
        For Each row In Table1.Rows               ' merge quantities for items in both stores
            itm = row.item(0)
            foundRow = Table2.Rows.Find(itm)
            If foundRow IsNot Nothing Then
                foundRow.Item(8) = "Y"
                row.item(6) += foundRow.Item(6)
            End If
        Next
        Dim newRow As DataRow

        For Each row In Table2.Rows                ' add items for store 2 only
            If row.item(8) <> "Y" Then
                newRow = Table1.NewRow()
                For i = 0 To 8
                    newRow.Item(i) = row.item(i)
                Next
                Table1.Rows.Add(newRow)
            End If
        Next
        dg7.DataSource = Table1.DefaultView


    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        If chkAllocated.Checked = False Then
            If Store1_Has_Items = True And Store2_Has_Items = True Then
                MsgBox("Cant't have Non_Allocated and multiple stores in ItemAnalysis data.")
                Application.Exit()
            End If
        End If
        template = Item_Analysis.templatesFolder & "\QuickOrder_Cumming_Template.xlsx"
        If btnMerged.Checked = True Then
            Call Merge_Qtys()
        End If

        If Store1_Has_Items = True Then
            Call Create_Excel(template, Table1)                       ' changed from displayTable
        End If
        If Store2_Has_Items = True And btnSeparate.Checked = True Then
            template = Item_Analysis.templatesFolder & "\QuickOrder_NorthPoint_Template.xlsx"
            Call Create_Excel(template, Table2)
        End If
        If Store1_Has_Items = True Or Store2_Has_Items = True Then
            template = Item_Analysis.templatesFolder & "\QuickOrder_Import.xlsx"
            Call Create_Excel_Copy(template, editTable)
        End If
    End Sub

    Private Sub Create_Excel(template, tbl)
        Dim fileName As String
        If Item_Analysis.reportsFolder = "DropBox" Then
            fileName = "c:\users\" & thisUser & "\dropbox\" & thisUser & "\QuickOrd\QuickOrd_" & thisVendor & "_" & displayStores _
                                    & "_" & thisDateString
        Else
            Dim reportsDirectory = Item_Analysis.reportsFolder & "\QuickOrd\"
            If (Not System.IO.Directory.Exists(reportsDirectory)) Then
                System.IO.Directory.CreateDirectory(reportsDirectory)
            End If
            fileName = reportsDirectory & "QuickOrd_" & thisVendor & "_" & displayStores _
                                    & "_" & thisDateString
        End If
      
        If System.IO.File.Exists(fileName & ".xlsx") = True Then
            Dim test As String
            Dim i As Integer
            For i = 1 To 100
                test = fileName & "(" & i & ")"
                If Not System.IO.File.Exists(test & ".xlsx") Then
                    fileName = test
                    Exit For
                End If
            Next
        End If
        System.IO.File.Copy(template, fileName & ".xlsx")
        Dim xlApp As excel.Application
        Dim xlWorkbook As excel.Workbook
        Dim xlWorksheet As excel.Worksheet
        Dim oTest As Object
        xlApp = New excel.Application
        xlWorkbook = xlApp.Workbooks.Open(fileName & ".xlsx")
        xlWorksheet = xlWorkbook.ActiveSheet
        xlWorksheet.Cells(3, 3) = thisUser
        xlWorksheet.Cells(4, 3) = thisVendor
        Dim r As Integer = tbl.Rows.Count
        Dim c As Integer = tbl.Columns.Count
        Dim tcost As Decimal = 0
        Dim cost, qty As Decimal
        Dim arr(r + 1, c) As String
        r = 0
        c = 0
        For Each row In tbl.Rows
            oTest = row("Merged")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "Y" Then                                           ' skip merged items
                    cost = 0
                    qty = 0
                    oTest = row.item("VendorItem_No")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then arr(r, 0) = oTest
                    oTest = row.item("Our_No")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then arr(r, 1) = oTest
                    oTest = row.item("Description")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then arr(r, 2) = oTest
                    oTest = row.item("Qty")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        arr(r, 3) = oTest
                        qty = CDec(oTest)
                    End If
                    oTest = row.item("UOM")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then arr(r, 4) = oTest
                    oTest = row.item("Cost")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        arr(r, 5) = oTest
                        cost = CDec(oTest)
                    End If
                    tcost += (cost * qty)
                    r += 1
                End If
            End If
        Next
        arr(r, 5) = tcost
        Dim c1 As excel.Range = xlWorksheet.Cells(7, 1)
        Dim c2 As excel.Range = xlWorksheet.Cells(r + 7, 6)
        Dim range As excel.Range = xlWorksheet.Range(c1, c2)
        range.Value = arr
        xlWorkbook.Save()
        xlWorkbook.Close()
        xlWorksheet = Nothing
        xlWorkbook = Nothing
        xlApp.Quit()
        releaseobject(xlWorkbook)
        releaseobject(xlWorksheet)
        If chkOpenXL.Checked = True Then
            Process.Start("excel", "" & fileName & ".xlsx")
        End If
    End Sub

    Private Sub Create_Excel_Copy(template, tbl)
        Dim fileName As String
        If Item_Analysis.reportsFolder = "DropBox" Then
            fileName = "c:\users\" & thisUser & "\dropbox\" & thisUser & "\QuickOrd\Import\QuickOrd_" & thisVendor & "_" & displayStores _
                                    & "_" & thisDateString
        Else
            Dim reportsDirectory As String = Item_Analysis.reportsFolder & "\Import\"
            If (Not System.IO.Directory.Exists(reportsDirectory)) Then
                System.IO.Directory.CreateDirectory(reportsDirectory)
            End If
            fileName = reportsDirectory & "QuickOrd_" & thisVendor & "_" & displayStores & "_" & thisDateString
        End If
        
        If System.IO.File.Exists(fileName & ".xlsx") = True Then
            Dim test As String
            Dim i As Integer
            For i = 1 To 100
                test = fileName & "(" & i & ")"
                If Not System.IO.File.Exists(test & ".xlsx") Then
                    fileName = test
                    Exit For
                End If
            Next
        End If
        System.IO.File.Copy(template, fileName & ".xlsx")
        Dim xlApp2 As excel.Application
        Dim xlWorkbook2 As excel.Workbook
        Dim xlWorksheet2 As excel.Worksheet
        xlApp2 = New excel.Application
        xlWorkbook2 = xlApp2.Workbooks.Open(fileName & ".xlsx")
        xlWorksheet2 = xlWorkbook2.ActiveSheet
        xlWorksheet2.Cells(3, 3) = thisUser
        Dim r As Integer = tbl.Rows.Count
        Dim c As Integer = tbl.Columns.Count
        Dim arr2(r, c) As String
        Dim oTest As Object
        Dim sql, vItem As String



        r = 0
        For Each row In tbl.rows
            con.Open()
            Dim itm As String = row.item(2)
            Dim buyunit As Integer
            sql = "SELECT Vend_Item_No, BuyUnit FROM Item_Master WHERE Item_No = '" & itm & "'"
            Dim cmd As New SqlCommand(sql, con)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            vItem = itm
            While rdr.Read
                oTest = rdr(0)
                If Not IsDBNull(oTest) Then vItem = rdr(0)
                oTest = rdr("BuyUnit")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyunit = oTest Else buyunit = 1
            End While
            con.Close()



            For c = 1 To tbl.columns.count - 1
                oTest = row.item(c)
                If c = 16 Then oTest = oTest * CDec(buyunit)
                If Not IsDBNull(oTest) Then arr2(r, c - 1) = oTest
            Next
            r += 1
        Next
        Dim c1 As excel.Range = xlWorksheet2.Cells(2, 1)
        Dim c2 As excel.Range = xlWorksheet2.Cells(r + 1, c)
        Dim range As excel.Range = xlWorksheet2.Range(c1, c2)
        range.Value = arr2
        xlWorkbook2.Save()
        xlWorkbook2.Close()
        xlWorksheet2 = Nothing
        xlWorkbook2 = Nothing
        xlApp2.Quit()
        releaseobject(xlWorkbook2)
        releaseobject(xlWorksheet2)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()

    End Sub

    Private Sub chkAllocated_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllocated.CheckedChanged
        If chkAllocated.Checked = False Then
            btnMerged.Checked = False
            btnSeparate.Checked = False
        End If
    End Sub

    Private Sub btnSeparate_CheckedChanged(sender As Object, e As EventArgs) Handles btnSeparate.CheckedChanged
        If btnSeparate.Checked = True Then
            btnMerged.Checked = False
            gbShipTo.Visible = False
        Else
            btnMerged.Checked = True
        End If
    End Sub

    Private Sub btnMerged_CheckedChanged(sender As Object, e As EventArgs) Handles btnMerged.CheckedChanged
        If btnMerged.Checked = True Then
            btnSeparate.Checked = False
            gbShipTo.Visible = True
            Me.Refresh()
            If btnNP.Checked = True Then
                template = Item_Analysis.templatesFolder & "\QuickOrder_NorthPoint_Template.xlsx"
            End If
        Else
            btnSeparate.Checked = True
        End If
    End Sub

    Shared Sub releaseobject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub btnNP_CheckedChanged(sender As Object, e As EventArgs) Handles btnNP.CheckedChanged
        If btnNP.Checked = True Then
            template = Item_Analysis.templatesFolder & "\QuickOrder_NorthPoint_Template.xlsx"
        Else
            template = Item_Analysis.templatesFolder & "\QuickOrder_Cumming_Template.xlsx"
        End If
    End Sub
End Class
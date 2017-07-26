Imports System.Data.SqlClient
Imports System.Xml
Imports excel = Microsoft.Office.Interop.Excel           ' add .net reference under Projects
Imports Microsoft.VisualBasic
Public Class Preview
    Public Shared thisUser As String = Environment.UserName.ToString
    Public Shared displayTable, thedataTable As DataTable
    Public Shared oTest As Object
    Public Shared selectedItem, theVendor, vendorName As String
    Public Shared searchString As String = Nothing
    Public Shared searchColumn As String = Nothing
    Public Shared columnIdx, rowIdx, selectedRow, selectedColumn, lastRow, currentRow As Integer
    Public Shared totOH, totOnPO, tot4Cast, totRecv, totSold, totd2Sold, totGMROI As Integer
    Public Shared searchTbl As DataTable

    Private Sub Preview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(50, 50)
        theVendor = Item_Analysis.vendorId
        vendorName = Item_Analysis.vendorName
        Item_Analysis.lblProcessing.Visible = False
        searchTbl = New DataTable
        searchTbl.Columns.Add("LineNo")
    End Sub
    Friend Sub ShowResults(ByVal tbl As DataTable, ByVal origTable As DataTable, ByVal ParamsTxt As String)
        thedataTable = origTable
        Me.Location = New Point(50, 50)
        theVendor = Item_Analysis.vendorId
        vendorName = Item_Analysis.vendorName
        Item_Analysis.lblProcessing.Visible = False
        searchTbl = New DataTable
        searchTbl.Columns.Add("LineNo")
        Dim cnt As Integer = 0
        displayTable = tbl
        searchTbl.Rows.Clear()
        txtParams.Text = ParamsTxt
        Dim clmName As String
        dgv1.DataSource = displayTable.DefaultView
        dgv1.Rows(dgv1.Rows.Count - 1).ReadOnly = True
        dgv1.Visible = True
        dgv1.Columns(0).Visible = False                                                           ' KeyId
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        For i As Integer = 0 To dgv1.Columns.Count - 1
            clmName = dgv1.Columns(i).Name
            If i < 13 Or i > 14 Then dgv1.Columns(i).ReadOnly = True
            If i > 3 And i < 19 Then
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                If i < 8 Or i > 9 Then dgv1.Columns(i).DefaultCellStyle.Format = "N0"
                If i = 4 Then dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
            If i > 17 And i < 23 Then
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Next
        dgv1.Columns(9).DefaultCellStyle.Font = New Font("Arial", 2, FontStyle.Regular)
        dgv1.Columns(9).DefaultCellStyle.Format = "N1"
        dgv1.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(16).DefaultCellStyle.Format = "N2"
        dgv1.Columns(17).DefaultCellStyle.Format = "N2"
        dgv1.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv1.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv1.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv1.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv1.Columns("OrdQty").DefaultCellStyle.BackColor = Color.Cornsilk
        dgv1.Columns("SC").DefaultCellStyle.BackColor = Color.Cornsilk
        txtOH.Text = Format(totOH, "###,###,##0")
        txtOnPO.Text = Format(totOnPO, "###,###,##0")
        txt4Cast.Text = Format(tot4Cast, "###,###,##0")
        txtRecv.Text = Format(totRecv, "###,###,##0")
        txtSold.Text = Format(totSold, "###,###,##0")
        txtd2Sold.Text = Format(totd2Sold, "###,###,##0")
        txtGMROI.Text = Format(totGMROI, "###,###,###")
    End Sub

    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        rowIdx = ht.RowIndex
        columnIdx = ht.ColumnIndex
        Dim str, hdr As String
        If rowIdx = 0 And columnIdx > 0 Then
            hdr = dgv1.Columns(columnIdx).HeaderText
            Select Case hdr
                Case "IF"
                    str = "Red: >50% above 4Cast Yellow: <75% of 4Cast"
                Case "4Cast"
                    str = "Blank: <3 wks hist Gray: <10 wks hist"
                Case "OnPO"
                    str = "Qty on PO within coverage prd"
                Case "Recv"
                    str = "Qty recv within coverage prd"
                Case "Sold"
                    str = "Qty sold within sales prd"
                Case "GMROI"
                    str = "Blank: not sold Gray:<10 wks Hist Red:<200 Yellow:<400"
                Case "MarkUp"
                    str = "Initial markup %"
            End Select
            'If columnIdx = 5 Then
            With dgv1.Rows(0).Cells(columnIdx)
                .ToolTipText = str
            End With
            ' End If
        End If
    End Sub

    Private Sub dgv11_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgv1.CellFormatting
        rowIdx = e.RowIndex
        columnIdx = e.ColumnIndex
        If dgv1.Columns(e.ColumnIndex).Name.Equals("GMROI") Then
            oTest = e.Value
            If IsNumeric(oTest) Then
                If CInt(e.Value) < 200 And CInt(e.Value) > -99999 Then
                    e.CellStyle.BackColor = Color.Red
                    ' e.CellStyle.SelectionBackColor = Color.DarkRed
                End If
                If CInt(e.Value) >= 200 And CInt(e.Value) <= 400 Then
                    e.CellStyle.BackColor = Color.Yellow
                End If
            End If
            oTest = dgv1.Rows(e.RowIndex).Cells("MaxWksOH").Value
            If IsNumeric(oTest) Then
                If CInt(oTest) < 10 Then
                    e.CellStyle.ForeColor = Color.Gray
                End If
            End If
            If IsNumeric(e.Value) Then
                If e.Value = -999999 Then e.CellStyle.ForeColor = Color.White
            End If
        End If
        If dgv1.Columns(e.ColumnIndex).Name.Equals("IF") Then
            oTest = e.Value
            If IsNumeric(oTest) Then
                oTest = dgv1.Rows(rowIdx).Cells("4Cast").Value
                If CDec(oTest) <> -999999 Then
                    If CDec(e.Value) >= 0 And CDec(e.Value) < 0.75 Then
                        e.CellStyle.BackColor = Color.Yellow
                    End If
                    If CDec(e.Value) > 1.5 Then
                        e.CellStyle.BackColor = Color.Red
                    End If
                End If
            End If
        End If
        If dgv1.Columns(e.ColumnIndex).Name.Equals("4Cast") Then
            oTest = dgv1.Rows(e.RowIndex).Cells("MaxWksOH").Value
            If IsNumeric(oTest) Then
                If CInt(oTest) > 2 And CInt(oTest) < 10 Then
                    e.CellStyle.ForeColor = Color.Gray
                End If
            End If
            If e.Value = -999999 Then
                e.CellStyle.ForeColor = Color.White
            End If
        End If
        Dim nam As String = dgv1.Columns(e.ColumnIndex).Name
        If nam = "Sold" Or nam = "Recv" Or nam = "OnPO" Then
            e.CellStyle.ForeColor = Color.Blue
            e.CellStyle.Font = New Font("Arial", 9.5, FontStyle.Underline)
        End If
        If dgv1.Columns(e.ColumnIndex).Name.Equals("OrdQty") Then
            oTest = dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then
                    If Not IsNumeric(oTest) Then
                        dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
                        MessageBox.Show("Entry was not numeric.", "ERROR!")
                        dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim openAnswer As Boolean = False
        Dim ans As DialogResult = MsgBox("Open Spreadsheet when processing complete?", MsgBoxStyle.YesNo)
        If ans = Windows.Forms.DialogResult.Yes Then
            openAnswer = True
        End If
        Call Dump2Excel(openAnswer)
    End Sub

    Public Sub Dump2Excel(openAnswer)
        Try
            ''lblProcessing.Visible = True
            ''Me.Refresh()
            Dim selectedstores As String
            Dim oTest As Object
            If Item_Analysis.chkCombined.Checked = True Then selectedstores = "0"
            txtParams.Text = Item_Analysis.vendorId & " Loc(s): " & Item_Analysis.displayStores & " Coverage: " &
                Item_Analysis.coverageSdate & " - " & Item_Analysis.coverageEdate &
                "  Sales: " & Item_Analysis.salesSdate & " - " & Item_Analysis.salesEdate & "  "
            txtParams.Text &= "Min GMROI: " & Item_Analysis.minGMROI & " Min Sold: " & Item_Analysis.minQtySold &
                " Min Weeks OnHand: " & Item_Analysis.MinWksOH
            Dim directory As String = Item_Analysis.templatesFolder
            Dim fileToCopy As String = directory & "\ItemAnalysisTemplate.xlsx"
            Dim fileName As String
            If Item_Analysis.reportsFolder = "DropBox" Then
                fileName = "c:\users\" & thisUser & "\dropbox\" & thisUser & "\TandG\" _
                                        & theVendor & "_" & Item_Analysis.displayStores & "_" & Item_Analysis.thisDateString
            Else
                Dim reportsDirectory As String = Item_Analysis.reportsFolder & "\TandG\"
                If (Not System.IO.Directory.Exists(reportsDirectory)) Then
                    System.IO.Directory.CreateDirectory(reportsDirectory)
                End If
                fileName = reportsDirectory & theVendor & "_" & Item_Analysis.displayStores & "_" & Item_Analysis.thisDateString
            End If
             
            If System.IO.File.Exists(fileName & ".xlsx") = True Then
                'System.IO.File.Delete(fileName)

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
            System.IO.File.Copy(fileToCopy, fileName & ".xlsx")
            Dim xlApp As excel.Application
            Dim xlWorkbook As excel.Workbook
            Dim xlWorksheet As excel.Worksheet
            xlApp = New excel.Application
            xlWorkbook = xlApp.Workbooks.Open(fileName & ".xlsx")
            xlWorksheet = xlWorkbook.ActiveSheet
            Dim xlRow As Integer = 1
            Dim xlCol As Integer = 1

            xlWorksheet.Cells(2, 1) = 0
            xlWorksheet.Cells(2, 2) = 0
            xlWorksheet.Cells(2, 4) = theVendor
            xlWorksheet.Cells(2, 3) = vendorName
            xlWorksheet.Cells(2, 8) = "Coverage # of Weeks: " & Item_Analysis.txtCweeks.Text
            xlWorksheet.Cells(2, 10) = Item_Analysis.cboCfrom.SelectedItem & " - " & Item_Analysis.cboCthru.SelectedItem
            xlWorksheet.Cells(2, 14) = "Sales1 Weeks = " & Item_Analysis.txtSweeks.Text
            xlWorksheet.Cells(2, 16) = Item_Analysis.cboSfrom.SelectedItem & " - " & Item_Analysis.cboSthru.SelectedItem
            xlWorksheet.Cells(2, 19) = "Sales2 Weeks = " & Item_Analysis.txtS2Weeks.Text
            xlWorksheet.Cells(2, 21) = Item_Analysis.cboS2from.SelectedItem & " - " & Item_Analysis.cboS2thru.SelectedItem
            xlRow = 3
            Dim r As Integer = displayTable.Rows.Count
            Dim c As Integer = displayTable.Columns.Count
            Dim arr(r, c) As String
            r = 0
            c = 0
            For Each row In displayTable.Rows
                For c = 1 To displayTable.Columns.Count - 1
                    oTest = row.item(c)
                    If Not IsDBNull(oTest) Then
                        arr(r, c - 1) = row.item(c)
                        If c > 4 And c < 13 Then
                            arr(r, c - 1) = Format(row.item(c), "###,###")
                        End If
                        If c = 8 Or c = 9 Then arr(r, c - 1) = Format(row.item(c), "###,###.0")
                        If c = 16 Or c = 17 Then arr(r, c - 1) = Format(row.item(c), "###,###.00")
                        If c = 26 Then arr(r, c - 1) = Format(row.item(c), "###,###")
                    End If
                Next
                r += 1
            Next
            oTest = c
            oTest = r
            Dim c1 As excel.Range = xlWorksheet.Cells(3, 1)
            Dim c2 As excel.Range = xlWorksheet.Cells(r + 2, c)
            Dim range As excel.Range = xlWorksheet.Range(c1, c2)
            range.Value = arr
            txtParams.Text = fileName & " ready."
            Me.Refresh()
            xlWorkbook.Save()
            xlWorkbook.Close()
            xlWorksheet = Nothing
            xlWorkbook = Nothing
            xlApp.Quit()
            releaseobject(xlWorkbook)
            releaseobject(xlWorksheet)

            If openAnswer = True Then
                Process.Start("excel", "" & fileName & ".xlsx")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
        Item_Analysis.Show()
    End Sub

    Private Sub btnOrder_Click(sender As Object, e As EventArgs) Handles btnOrder.Click
        Dim QuickForm As New Quick_Order
        QuickForm.QuickOrder(theVendor, Item_Analysis.thisDateString, Item_Analysis.displayStores, displayTable)
        QuickForm.Show()

    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim lastRow As Integer = displayTable.Rows.Count - 1
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If rowIdx >= 0 And rowIdx <= lastRow Then
                    If columnIdx = 7 Or columnIdx = 10 Then
                        selectedItem = dgv1.Rows(rowIdx).Cells(2).Value
                        'lblProcessing.Visible = True
                        OpenPOs.Show()
                    End If
                    If columnIdx = 11 Then
                        selectedItem = dgv1.Rows(rowIdx).Cells(2).Value
                        'lblProcessing.Visible = True
                        SalesHistory.Show()
                    End If

                End If
        End Select
    End Sub

    Private Sub dgv1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles dgv1.KeyDown
        If (e.KeyCode = Keys.F AndAlso e.Modifiers = Keys.Control) Then
            searchColumn = dgv1.CurrentCell.OwningColumn.Name
            If searchColumn <> "Item_No" And searchColumn <> "Descr" Then Exit Sub
            searchString = UCase(InputBox("ENTER VALUE", "SEARCH " & searchColumn & " ", "Enter Search Criteria Here"))
            txtSearch.Text = searchString
            txtSearch.Visible = True
            Call Flag_Hits()
        End If
        If e.KeyCode = Keys.Down AndAlso pbClear.Visible = True Then
            e.Handled = True
            Call Next_Row()
        End If
        If e.KeyCode = Keys.Up AndAlso pbClear.Visible = True Then
            e.Handled = True
            Call Previous_Row()
        End If
    End Sub

    'Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
    '    searchColumn = dgv1.CurrentCell.OwningColumn.Name
    '    If searchColumn = "OrdQty" Then
    '        oTest = dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    '        If Not IsNumeric(oTest) Then dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
    '    End If
    'End Sub
    Private Sub Flag_Hits()
        If Not IsNothing(searchString) Then
            Call Clear_Search()
            Dim columns As DataColumnCollection = displayTable.Columns
            Dim row As DataRow
            Dim cnt As Integer = 0
            searchTbl.Rows.Clear()
            If columns.Contains(searchColumn) Then columnIdx = columns.IndexOf(searchColumn)
            For i As Integer = 0 To dgv1.Rows.Count - 1
                dgv1.Rows(i).Cells(columnIdx).Style.BackColor = Color.White
                If Not IsNothing(searchString) And InStr(UCase(dgv1.Rows(i).Cells(columnIdx).Value), searchString) Then
                    dgv1.Rows(i).Cells(columnIdx).Style.BackColor = Color.Salmon
                    row = searchTbl.NewRow
                    row(0) = i
                    searchTbl.Rows.Add(row)
                    LeftArrow.Visible = True
                    RightArrow.Visible = True
                    pbClear.Visible = True
                    cnt += 1
                End If
            Next
            If cnt = 0 Then Exit Sub
            row = searchTbl.Rows(0)
            selectedRow = row(0)
            selectedColumn = columnIdx
            currentRow = 0
            lastRow = searchTbl.Rows.Count - 1
            dgv1.CurrentCell = dgv1.Rows(selectedRow).Cells(columnIdx)
        End If
    End Sub


    Private Sub LeftArrow_Click(sender As Object, e As EventArgs) Handles LeftArrow.Click
        Call Previous_row()
    End Sub

    Private Sub Previous_Row()
        Dim rw As DataRow
        If currentRow > 0 Then
            currentRow -= 1
            rw = searchTbl.Rows(currentRow)
            selectedRow = rw(0)

            dgv1.CurrentCell = dgv1.Rows(selectedRow).Cells(selectedColumn)
        End If
    End Sub

    Private Sub RightArrow_Click(sender As Object, e As EventArgs) Handles RightArrow.Click
        Call Next_Row()
    End Sub

    Private Sub Next_Row()
        Dim rw As DataRow
        If currentRow < lastRow Then
            currentRow += 1
            rw = searchTbl.Rows(currentRow)
            selectedRow = rw(0)

            dgv1.CurrentCell = dgv1.Rows(selectedRow).Cells(selectedColumn)
        End If
    End Sub

    Private Sub pbClear_Click(sender As Object, e As EventArgs) Handles pbClear.Click
        Call Clear_Search()
    End Sub

    Private Sub Clear_Search()
        If selectedColumn > 0 Then
            txtSearch.Visible = False
            LeftArrow.Visible = False
            RightArrow.Visible = False
            pbClear.Visible = False
            For i As Integer = 0 To dgv1.Rows.Count - 1
                dgv1.Rows(i).Cells(selectedColumn).Style.BackColor = Color.White
            Next
        End If
    End Sub
End Class
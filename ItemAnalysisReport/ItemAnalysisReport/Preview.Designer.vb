<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Preview
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Preview))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOrder = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.txtParams = New System.Windows.Forms.TextBox()
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.txtOH = New System.Windows.Forms.TextBox()
        Me.txtOnPO = New System.Windows.Forms.TextBox()
        Me.txtRecv = New System.Windows.Forms.TextBox()
        Me.txtSold = New System.Windows.Forms.TextBox()
        Me.txtd2Sold = New System.Windows.Forms.TextBox()
        Me.txt4Cast = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.LeftArrow = New System.Windows.Forms.PictureBox()
        Me.RightArrow = New System.Windows.Forms.PictureBox()
        Me.pbClear = New System.Windows.Forms.PictureBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtGMROI = New System.Windows.Forms.TextBox()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LeftArrow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RightArrow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbClear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(1095, 675)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(128, 28)
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOrder
        '
        Me.btnOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOrder.Location = New System.Drawing.Point(564, 675)
        Me.btnOrder.Name = "btnOrder"
        Me.btnOrder.Size = New System.Drawing.Size(128, 28)
        Me.btnOrder.TabIndex = 8
        Me.btnOrder.Text = "Quick Order"
        Me.btnOrder.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.Location = New System.Drawing.Point(13, 675)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(128, 28)
        Me.btnExport.TabIndex = 0
        Me.btnExport.Text = "Export to Excel"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.AllowUserToOrderColumns = True
        Me.dgv1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv1.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgv1.Location = New System.Drawing.Point(12, 39)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(1291, 574)
        Me.dgv1.TabIndex = 6
        '
        'txtParams
        '
        Me.txtParams.Location = New System.Drawing.Point(4, 11)
        Me.txtParams.Name = "txtParams"
        Me.txtParams.Size = New System.Drawing.Size(1291, 20)
        Me.txtParams.TabIndex = 5
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 28.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.Location = New System.Drawing.Point(243, 209)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(434, 44)
        Me.lblProcessing.TabIndex = 10
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'txtOH
        '
        Me.txtOH.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOH.Location = New System.Drawing.Point(577, 637)
        Me.txtOH.Name = "txtOH"
        Me.txtOH.Size = New System.Drawing.Size(60, 20)
        Me.txtOH.TabIndex = 11
        Me.txtOH.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOnPO
        '
        Me.txtOnPO.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOnPO.Location = New System.Drawing.Point(643, 637)
        Me.txtOnPO.Name = "txtOnPO"
        Me.txtOnPO.Size = New System.Drawing.Size(60, 20)
        Me.txtOnPO.TabIndex = 12
        Me.txtOnPO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtRecv
        '
        Me.txtRecv.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecv.Location = New System.Drawing.Point(774, 636)
        Me.txtRecv.Name = "txtRecv"
        Me.txtRecv.Size = New System.Drawing.Size(60, 20)
        Me.txtRecv.TabIndex = 13
        Me.txtRecv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSold
        '
        Me.txtSold.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSold.Location = New System.Drawing.Point(839, 636)
        Me.txtSold.Name = "txtSold"
        Me.txtSold.Size = New System.Drawing.Size(60, 20)
        Me.txtSold.TabIndex = 14
        Me.txtSold.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtd2Sold
        '
        Me.txtd2Sold.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtd2Sold.Location = New System.Drawing.Point(905, 636)
        Me.txtd2Sold.Name = "txtd2Sold"
        Me.txtd2Sold.Size = New System.Drawing.Size(60, 20)
        Me.txtd2Sold.TabIndex = 15
        Me.txtd2Sold.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txt4Cast
        '
        Me.txt4Cast.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt4Cast.Location = New System.Drawing.Point(709, 637)
        Me.txt4Cast.Name = "txt4Cast"
        Me.txt4Cast.Size = New System.Drawing.Size(60, 20)
        Me.txt4Cast.TabIndex = 16
        Me.txt4Cast.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(614, 618)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "OH"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(667, 618)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "OnPO"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(735, 618)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "4Cast"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(801, 618)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Recv"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(871, 618)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(28, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Sold"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(925, 618)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 13)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "d2Sold"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(454, 634)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(42, 13)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Totals"
        '
        'txtSearch
        '
        Me.txtSearch.BackColor = System.Drawing.Color.Salmon
        Me.txtSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(1115, 12)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(112, 20)
        Me.txtSearch.TabIndex = 25
        Me.txtSearch.Visible = False
        '
        'LeftArrow
        '
        Me.LeftArrow.Image = CType(resources.GetObject("LeftArrow.Image"), System.Drawing.Image)
        Me.LeftArrow.Location = New System.Drawing.Point(1233, 12)
        Me.LeftArrow.Name = "LeftArrow"
        Me.LeftArrow.Size = New System.Drawing.Size(22, 20)
        Me.LeftArrow.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.LeftArrow.TabIndex = 26
        Me.LeftArrow.TabStop = False
        Me.LeftArrow.Visible = False
        '
        'RightArrow
        '
        Me.RightArrow.Image = CType(resources.GetObject("RightArrow.Image"), System.Drawing.Image)
        Me.RightArrow.Location = New System.Drawing.Point(1251, 12)
        Me.RightArrow.Name = "RightArrow"
        Me.RightArrow.Size = New System.Drawing.Size(22, 20)
        Me.RightArrow.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.RightArrow.TabIndex = 27
        Me.RightArrow.TabStop = False
        Me.RightArrow.Visible = False
        '
        'pbClear
        '
        Me.pbClear.Image = CType(resources.GetObject("pbClear.Image"), System.Drawing.Image)
        Me.pbClear.Location = New System.Drawing.Point(1275, 14)
        Me.pbClear.Name = "pbClear"
        Me.pbClear.Size = New System.Drawing.Size(15, 15)
        Me.pbClear.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbClear.TabIndex = 28
        Me.pbClear.TabStop = False
        Me.pbClear.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(526, 618)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(43, 13)
        Me.Label8.TabIndex = 30
        Me.Label8.Text = "GMROI"
        '
        'txtGMROI
        '
        Me.txtGMROI.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGMROI.Location = New System.Drawing.Point(511, 637)
        Me.txtGMROI.Name = "txtGMROI"
        Me.txtGMROI.Size = New System.Drawing.Size(60, 20)
        Me.txtGMROI.TabIndex = 29
        Me.txtGMROI.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Preview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1316, 713)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtGMROI)
        Me.Controls.Add(Me.pbClear)
        Me.Controls.Add(Me.RightArrow)
        Me.Controls.Add(Me.LeftArrow)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt4Cast)
        Me.Controls.Add(Me.txtd2Sold)
        Me.Controls.Add(Me.txtSold)
        Me.Controls.Add(Me.txtRecv)
        Me.Controls.Add(Me.txtOnPO)
        Me.Controls.Add(Me.txtOH)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnOrder)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.dgv1)
        Me.Controls.Add(Me.txtParams)
        Me.Name = "Preview"
        Me.Text = "Preview"
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LeftArrow, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RightArrow, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbClear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnOrder As System.Windows.Forms.Button
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtParams As System.Windows.Forms.TextBox
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents txtOH As System.Windows.Forms.TextBox
    Friend WithEvents txtOnPO As System.Windows.Forms.TextBox
    Friend WithEvents txtRecv As System.Windows.Forms.TextBox
    Friend WithEvents txtSold As System.Windows.Forms.TextBox
    Friend WithEvents txtd2Sold As System.Windows.Forms.TextBox
    Friend WithEvents txt4Cast As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents LeftArrow As System.Windows.Forms.PictureBox
    Friend WithEvents RightArrow As System.Windows.Forms.PictureBox
    Friend WithEvents pbClear As System.Windows.Forms.PictureBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtGMROI As System.Windows.Forms.TextBox
End Class

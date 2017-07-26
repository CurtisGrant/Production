<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Quick_Order
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.chkOpenXL = New System.Windows.Forms.CheckBox()
        Me.gbShipTo = New System.Windows.Forms.GroupBox()
        Me.btnNP = New System.Windows.Forms.RadioButton()
        Me.btnLP = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnMerged = New System.Windows.Forms.RadioButton()
        Me.btnSeparate = New System.Windows.Forms.RadioButton()
        Me.chkAllocated = New System.Windows.Forms.CheckBox()
        Me.dg7 = New System.Windows.Forms.DataGridView()
        Me.txtVendor = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.gbShipTo.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dg7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkOpenXL
        '
        Me.chkOpenXL.AutoSize = True
        Me.chkOpenXL.Checked = True
        Me.chkOpenXL.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOpenXL.Location = New System.Drawing.Point(19, 34)
        Me.chkOpenXL.Name = "chkOpenXL"
        Me.chkOpenXL.Size = New System.Drawing.Size(156, 17)
        Me.chkOpenXL.TabIndex = 16
        Me.chkOpenXL.Text = "Open XLSX When Finished"
        Me.chkOpenXL.UseVisualStyleBackColor = True
        '
        'gbShipTo
        '
        Me.gbShipTo.Controls.Add(Me.btnNP)
        Me.gbShipTo.Controls.Add(Me.btnLP)
        Me.gbShipTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbShipTo.Location = New System.Drawing.Point(12, 116)
        Me.gbShipTo.Name = "gbShipTo"
        Me.gbShipTo.Size = New System.Drawing.Size(252, 58)
        Me.gbShipTo.TabIndex = 15
        Me.gbShipTo.TabStop = False
        Me.gbShipTo.Text = "Ship to:"
        Me.gbShipTo.Visible = False
        '
        'btnNP
        '
        Me.btnNP.AutoSize = True
        Me.btnNP.Location = New System.Drawing.Point(105, 19)
        Me.btnNP.Name = "btnNP"
        Me.btnNP.Size = New System.Drawing.Size(91, 20)
        Me.btnNP.TabIndex = 10
        Me.btnNP.Text = "North Point"
        Me.btnNP.UseVisualStyleBackColor = True
        '
        'btnLP
        '
        Me.btnLP.AutoSize = True
        Me.btnLP.Checked = True
        Me.btnLP.Location = New System.Drawing.Point(7, 19)
        Me.btnLP.Name = "btnLP"
        Me.btnLP.Size = New System.Drawing.Size(82, 20)
        Me.btnLP.TabIndex = 10
        Me.btnLP.TabStop = True
        Me.btnLP.Text = "Cumming"
        Me.btnLP.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnMerged)
        Me.GroupBox2.Controls.Add(Me.btnSeparate)
        Me.GroupBox2.Controls.Add(Me.chkAllocated)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 57)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(252, 53)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        '
        'btnMerged
        '
        Me.btnMerged.AutoSize = True
        Me.btnMerged.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMerged.Location = New System.Drawing.Point(182, 16)
        Me.btnMerged.Name = "btnMerged"
        Me.btnMerged.Size = New System.Drawing.Size(73, 20)
        Me.btnMerged.TabIndex = 11
        Me.btnMerged.Text = "Merged"
        Me.btnMerged.UseVisualStyleBackColor = True
        '
        'btnSeparate
        '
        Me.btnSeparate.AutoSize = True
        Me.btnSeparate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSeparate.Location = New System.Drawing.Point(93, 15)
        Me.btnSeparate.Name = "btnSeparate"
        Me.btnSeparate.Size = New System.Drawing.Size(82, 20)
        Me.btnSeparate.TabIndex = 10
        Me.btnSeparate.Text = "Separate"
        Me.btnSeparate.UseVisualStyleBackColor = True
        '
        'chkAllocated
        '
        Me.chkAllocated.AutoSize = True
        Me.chkAllocated.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAllocated.Location = New System.Drawing.Point(7, 16)
        Me.chkAllocated.Name = "chkAllocated"
        Me.chkAllocated.Size = New System.Drawing.Size(84, 20)
        Me.chkAllocated.TabIndex = 9
        Me.chkAllocated.Text = "Allocated"
        Me.chkAllocated.UseVisualStyleBackColor = True
        '
        'dg7
        '
        Me.dg7.AllowUserToAddRows = False
        Me.dg7.AllowUserToDeleteRows = False
        Me.dg7.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dg7.DefaultCellStyle = DataGridViewCellStyle2
        Me.dg7.Enabled = False
        Me.dg7.Location = New System.Drawing.Point(19, 202)
        Me.dg7.Name = "dg7"
        Me.dg7.Size = New System.Drawing.Size(1289, 346)
        Me.dg7.TabIndex = 13
        '
        'txtVendor
        '
        Me.txtVendor.Location = New System.Drawing.Point(261, 8)
        Me.txtVendor.Name = "txtVendor"
        Me.txtVendor.Size = New System.Drawing.Size(135, 20)
        Me.txtVendor.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(25, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(239, 15)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Create Quick Order Spreadsheet for:"
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(308, 571)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 18
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRun.Location = New System.Drawing.Point(19, 571)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 17
        Me.btnRun.Text = "Create XLS"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'Quick_Order
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1319, 606)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.chkOpenXL)
        Me.Controls.Add(Me.gbShipTo)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.dg7)
        Me.Controls.Add(Me.txtVendor)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Quick_Order"
        Me.Text = "QuickOrder"
        Me.gbShipTo.ResumeLayout(False)
        Me.gbShipTo.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dg7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkOpenXL As System.Windows.Forms.CheckBox
    Friend WithEvents gbShipTo As System.Windows.Forms.GroupBox
    Friend WithEvents btnNP As System.Windows.Forms.RadioButton
    Friend WithEvents btnLP As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnMerged As System.Windows.Forms.RadioButton
    Friend WithEvents btnSeparate As System.Windows.Forms.RadioButton
    Friend WithEvents chkAllocated As System.Windows.Forms.CheckBox
    Friend WithEvents dg7 As System.Windows.Forms.DataGridView
    Friend WithEvents txtVendor As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnRun As System.Windows.Forms.Button
End Class

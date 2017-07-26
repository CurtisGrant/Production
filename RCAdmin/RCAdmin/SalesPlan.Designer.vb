<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlanMaintenance
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.btnCreatePlan = New System.Windows.Forms.Button()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.lblSaving = New System.Windows.Forms.Label()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cboRnd = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdoPct = New System.Windows.Forms.RadioButton()
        Me.rdoAmt = New System.Windows.Forms.RadioButton()
        Me.btnSaveAs = New System.Windows.Forms.Button()
        Me.btnActivate = New System.Windows.Forms.Button()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtLastUpdate = New System.Windows.Forms.TextBox()
        Me.txtType = New System.Windows.Forms.TextBox()
        Me.cboPlan = New System.Windows.Forms.ComboBox()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gb1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.BackColor = System.Drawing.SystemColors.Info
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.lblProcessing.Location = New System.Drawing.Point(91, 268)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(731, 73)
        Me.lblProcessing.TabIndex = 61
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'btnCreatePlan
        '
        Me.btnCreatePlan.Location = New System.Drawing.Point(874, 35)
        Me.btnCreatePlan.Name = "btnCreatePlan"
        Me.btnCreatePlan.Size = New System.Drawing.Size(176, 23)
        Me.btnCreatePlan.TabIndex = 60
        Me.btnCreatePlan.Text = "CREATE PLAN FOR NEW YEAR"
        Me.btnCreatePlan.UseVisualStyleBackColor = True
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(19, 571)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(77, 13)
        Me.serverLabel.TabIndex = 59
        Me.serverLabel.Text = "Connected to: "
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(981, 552)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(78, 24)
        Me.btnDelete.TabIndex = 58
        Me.btnDelete.Text = "Delete Plan"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'lblSaving
        '
        Me.lblSaving.AutoSize = True
        Me.lblSaving.Location = New System.Drawing.Point(467, 552)
        Me.lblSaving.Name = "lblSaving"
        Me.lblSaving.Size = New System.Drawing.Size(64, 13)
        Me.lblSaving.TabIndex = 53
        Me.lblSaving.Text = "Saving Plan"
        Me.lblSaving.Visible = False
        '
        'pb1
        '
        Me.pb1.BackColor = System.Drawing.Color.Green
        Me.pb1.Location = New System.Drawing.Point(537, 546)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(246, 23)
        Me.pb1.TabIndex = 52
        Me.pb1.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboRnd)
        Me.GroupBox2.Location = New System.Drawing.Point(703, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(116, 78)
        Me.GroupBox2.TabIndex = 57
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Round to Nearest:"
        '
        'cboRnd
        '
        Me.cboRnd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRnd.FormattingEnabled = True
        Me.cboRnd.Location = New System.Drawing.Point(18, 28)
        Me.cboRnd.Name = "cboRnd"
        Me.cboRnd.Size = New System.Drawing.Size(44, 21)
        Me.cboRnd.TabIndex = 32
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoPct)
        Me.GroupBox1.Controls.Add(Me.rdoAmt)
        Me.GroupBox1.Location = New System.Drawing.Point(536, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(158, 78)
        Me.GroupBox1.TabIndex = 56
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter Plan Adjustments in:"
        '
        'rdoPct
        '
        Me.rdoPct.AutoSize = True
        Me.rdoPct.Location = New System.Drawing.Point(7, 39)
        Me.rdoPct.Name = "rdoPct"
        Me.rdoPct.Size = New System.Drawing.Size(85, 17)
        Me.rdoPct.TabIndex = 43
        Me.rdoPct.TabStop = True
        Me.rdoPct.Text = "Percentages"
        Me.rdoPct.UseVisualStyleBackColor = True
        '
        'rdoAmt
        '
        Me.rdoAmt.AutoSize = True
        Me.rdoAmt.Location = New System.Drawing.Point(7, 22)
        Me.rdoAmt.Name = "rdoAmt"
        Me.rdoAmt.Size = New System.Drawing.Size(66, 17)
        Me.rdoAmt.TabIndex = 0
        Me.rdoAmt.TabStop = True
        Me.rdoAmt.Text = "Amounts"
        Me.rdoAmt.UseVisualStyleBackColor = True
        '
        'btnSaveAs
        '
        Me.btnSaveAs.Location = New System.Drawing.Point(131, 546)
        Me.btnSaveAs.Name = "btnSaveAs"
        Me.btnSaveAs.Size = New System.Drawing.Size(78, 24)
        Me.btnSaveAs.TabIndex = 55
        Me.btnSaveAs.Text = "Save As"
        Me.btnSaveAs.UseVisualStyleBackColor = True
        '
        'btnActivate
        '
        Me.btnActivate.Location = New System.Drawing.Point(345, 546)
        Me.btnActivate.Name = "btnActivate"
        Me.btnActivate.Size = New System.Drawing.Size(114, 24)
        Me.btnActivate.TabIndex = 54
        Me.btnActivate.Text = "Activate Draft Plan"
        Me.btnActivate.UseVisualStyleBackColor = True
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.Label1)
        Me.gb1.Controls.Add(Me.cboDept)
        Me.gb1.Controls.Add(Me.Label2)
        Me.gb1.Controls.Add(Me.Label5)
        Me.gb1.Controls.Add(Me.txtLastUpdate)
        Me.gb1.Controls.Add(Me.txtType)
        Me.gb1.Controls.Add(Me.cboPlan)
        Me.gb1.Controls.Add(Me.cboStore)
        Me.gb1.Controls.Add(Me.Label3)
        Me.gb1.Location = New System.Drawing.Point(22, 6)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(508, 78)
        Me.gb1.TabIndex = 51
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Plan"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(161, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Status"
        '
        'cboDept
        '
        Me.cboDept.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDept.Enabled = False
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(421, 30)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(81, 21)
        Me.cboDept.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(233, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Last Update"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(421, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Dept"
        '
        'txtLastUpdate
        '
        Me.txtLastUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastUpdate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLastUpdate.Location = New System.Drawing.Point(233, 29)
        Me.txtLastUpdate.Name = "txtLastUpdate"
        Me.txtLastUpdate.Size = New System.Drawing.Size(66, 20)
        Me.txtLastUpdate.TabIndex = 29
        '
        'txtType
        '
        Me.txtType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtType.Location = New System.Drawing.Point(161, 29)
        Me.txtType.Name = "txtType"
        Me.txtType.Size = New System.Drawing.Size(66, 20)
        Me.txtType.TabIndex = 27
        '
        'cboPlan
        '
        Me.cboPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPlan.FormattingEnabled = True
        Me.cboPlan.Location = New System.Drawing.Point(6, 29)
        Me.cboPlan.Name = "cboPlan"
        Me.cboPlan.Size = New System.Drawing.Size(129, 21)
        Me.cboPlan.TabIndex = 14
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.Enabled = False
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(334, 30)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(82, 21)
        Me.cboStore.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(334, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Store"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(18, 545)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(78, 24)
        Me.btnSave.TabIndex = 50
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv1.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgv1.Location = New System.Drawing.Point(19, 92)
        Me.dgv1.Name = "dgv1"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgv1.Size = New System.Drawing.Size(1042, 448)
        Me.dgv1.TabIndex = 49
        '
        'PlanMaintenance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1079, 590)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.btnCreatePlan)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.lblSaving)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnSaveAs)
        Me.Controls.Add(Me.btnActivate)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "PlanMaintenance"
        Me.Text = "Modify Sales Plan"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents btnCreatePlan As System.Windows.Forms.Button
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents lblSaving As System.Windows.Forms.Label
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cboRnd As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoPct As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAmt As System.Windows.Forms.RadioButton
    Friend WithEvents btnSaveAs As System.Windows.Forms.Button
    Friend WithEvents btnActivate As System.Windows.Forms.Button
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboDept As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLastUpdate As System.Windows.Forms.TextBox
    Friend WithEvents txtType As System.Windows.Forms.TextBox
    Friend WithEvents cboPlan As System.Windows.Forms.ComboBox
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
End Class

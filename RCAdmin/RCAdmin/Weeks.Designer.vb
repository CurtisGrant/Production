<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Weeks
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
        Me.lblPromotions = New System.Windows.Forms.Label()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.dgv2 = New System.Windows.Forms.DataGridView()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.txtPlan = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtLastUpdate = New System.Windows.Forms.TextBox()
        Me.txtType = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cboRnd = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdoPct = New System.Windows.Forms.RadioButton()
        Me.rdoAmt = New System.Windows.Forms.RadioButton()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.txtStore = New System.Windows.Forms.TextBox()
        Me.txtDept = New System.Windows.Forms.TextBox()
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gb1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblPromotions
        '
        Me.lblPromotions.AutoSize = True
        Me.lblPromotions.Location = New System.Drawing.Point(22, 272)
        Me.lblPromotions.Name = "lblPromotions"
        Me.lblPromotions.Size = New System.Drawing.Size(59, 13)
        Me.lblPromotions.TabIndex = 55
        Me.lblPromotions.Text = "Promotions"
        Me.lblPromotions.Visible = False
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(18, 473)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 54
        Me.serverLabel.Text = "server"
        '
        'dgv2
        '
        Me.dgv2.AllowUserToAddRows = False
        Me.dgv2.AllowUserToDeleteRows = False
        Me.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgv2.Enabled = False
        Me.dgv2.Location = New System.Drawing.Point(18, 291)
        Me.dgv2.Name = "dgv2"
        Me.dgv2.Size = New System.Drawing.Size(1042, 150)
        Me.dgv2.TabIndex = 53
        Me.dgv2.Visible = False
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.txtDept)
        Me.gb1.Controls.Add(Me.txtStore)
        Me.gb1.Controls.Add(Me.txtPlan)
        Me.gb1.Controls.Add(Me.Label2)
        Me.gb1.Controls.Add(Me.Label5)
        Me.gb1.Controls.Add(Me.txtLastUpdate)
        Me.gb1.Controls.Add(Me.txtType)
        Me.gb1.Controls.Add(Me.Label3)
        Me.gb1.Location = New System.Drawing.Point(21, 7)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(444, 78)
        Me.gb1.TabIndex = 52
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Plan"
        '
        'txtPlan
        '
        Me.txtPlan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPlan.Location = New System.Drawing.Point(18, 29)
        Me.txtPlan.Name = "txtPlan"
        Me.txtPlan.ReadOnly = True
        Me.txtPlan.Size = New System.Drawing.Size(137, 20)
        Me.txtPlan.TabIndex = 31
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
        Me.Label5.Location = New System.Drawing.Point(385, 16)
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
        Me.txtLastUpdate.ReadOnly = True
        Me.txtLastUpdate.Size = New System.Drawing.Size(66, 20)
        Me.txtLastUpdate.TabIndex = 29
        '
        'txtType
        '
        Me.txtType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtType.HideSelection = False
        Me.txtType.Location = New System.Drawing.Point(161, 29)
        Me.txtType.Name = "txtType"
        Me.txtType.ReadOnly = True
        Me.txtType.Size = New System.Drawing.Size(66, 20)
        Me.txtType.TabIndex = 27
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboRnd)
        Me.GroupBox2.Location = New System.Drawing.Point(832, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(116, 78)
        Me.GroupBox2.TabIndex = 51
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Round to Nearest:"
        '
        'cboRnd
        '
        Me.cboRnd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRnd.FormattingEnabled = True
        Me.cboRnd.Location = New System.Drawing.Point(16, 19)
        Me.cboRnd.Name = "cboRnd"
        Me.cboRnd.Size = New System.Drawing.Size(44, 21)
        Me.cboRnd.TabIndex = 31
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoPct)
        Me.GroupBox1.Controls.Add(Me.rdoAmt)
        Me.GroupBox1.Location = New System.Drawing.Point(636, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(158, 78)
        Me.GroupBox1.TabIndex = 50
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
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(18, 447)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(97, 23)
        Me.btnSave.TabIndex = 49
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.Location = New System.Drawing.Point(18, 93)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(1042, 168)
        Me.dgv1.TabIndex = 48
        '
        'txtStore
        '
        Me.txtStore.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStore.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStore.Location = New System.Drawing.Point(326, 29)
        Me.txtStore.Name = "txtStore"
        Me.txtStore.ReadOnly = True
        Me.txtStore.Size = New System.Drawing.Size(53, 20)
        Me.txtStore.TabIndex = 56
        '
        'txtDept
        '
        Me.txtDept.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDept.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDept.Location = New System.Drawing.Point(385, 29)
        Me.txtDept.Name = "txtDept"
        Me.txtDept.ReadOnly = True
        Me.txtDept.Size = New System.Drawing.Size(53, 20)
        Me.txtDept.TabIndex = 57
        '
        'Weeks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1079, 492)
        Me.Controls.Add(Me.lblPromotions)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.dgv2)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "Weeks"
        Me.Text = "Modify Week Plan"
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPromotions As System.Windows.Forms.Label
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents dgv2 As System.Windows.Forms.DataGridView
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPlan As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLastUpdate As System.Windows.Forms.TextBox
    Friend WithEvents txtType As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cboRnd As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoPct As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAmt As System.Windows.Forms.RadioButton
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtDept As System.Windows.Forms.TextBox
    Friend WithEvents txtStore As System.Windows.Forms.TextBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PercentageMaintenance
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.rdoBuyer = New System.Windows.Forms.RadioButton()
        Me.rdoDay = New System.Windows.Forms.RadioButton()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.gb1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnEdit)
        Me.GroupBox1.Controls.Add(Me.rdoBuyer)
        Me.GroupBox1.Controls.Add(Me.rdoDay)
        Me.GroupBox1.Location = New System.Drawing.Point(474, 22)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 78)
        Me.GroupBox1.TabIndex = 51
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "By"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(129, 16)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(65, 49)
        Me.btnEdit.TabIndex = 33
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'rdoBuyer
        '
        Me.rdoBuyer.AutoSize = True
        Me.rdoBuyer.Checked = True
        Me.rdoBuyer.Location = New System.Drawing.Point(15, 19)
        Me.rdoBuyer.Name = "rdoBuyer"
        Me.rdoBuyer.Size = New System.Drawing.Size(52, 17)
        Me.rdoBuyer.TabIndex = 30
        Me.rdoBuyer.TabStop = True
        Me.rdoBuyer.Text = "Buyer"
        Me.rdoBuyer.UseVisualStyleBackColor = True
        '
        'rdoDay
        '
        Me.rdoDay.AutoSize = True
        Me.rdoDay.Location = New System.Drawing.Point(15, 48)
        Me.rdoDay.Name = "rdoDay"
        Me.rdoDay.Size = New System.Drawing.Size(88, 17)
        Me.rdoDay.TabIndex = 32
        Me.rdoDay.Text = "Day of Week"
        Me.rdoDay.UseVisualStyleBackColor = True
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.Label1)
        Me.gb1.Controls.Add(Me.cboYear)
        Me.gb1.Controls.Add(Me.cboDept)
        Me.gb1.Controls.Add(Me.Label5)
        Me.gb1.Controls.Add(Me.cboStore)
        Me.gb1.Controls.Add(Me.Label3)
        Me.gb1.Location = New System.Drawing.Point(13, 22)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(387, 78)
        Me.gb1.TabIndex = 50
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Parameters"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Year"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(7, 30)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(58, 21)
        Me.cboYear.TabIndex = 11
        '
        'cboDept
        '
        Me.cboDept.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDept.Enabled = False
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(122, 30)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(44, 21)
        Me.cboDept.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(122, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Dept"
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.Enabled = False
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(71, 30)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(44, 21)
        Me.cboStore.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(71, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Store"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(795, 548)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 49
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        Me.btnExit.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(301, 548)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 48
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        Me.btnSave.Visible = False
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.Location = New System.Drawing.Point(13, 143)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(989, 384)
        Me.dgv1.TabIndex = 47
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(13, 573)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 52
        Me.serverLabel.Text = "server"
        '
        'PercentageMaintenance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1014, 592)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "PercentageMaintenance"
        Me.Text = "PercentageMaintenance"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents rdoBuyer As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDay As System.Windows.Forms.RadioButton
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboDept As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents serverLabel As System.Windows.Forms.Label
End Class

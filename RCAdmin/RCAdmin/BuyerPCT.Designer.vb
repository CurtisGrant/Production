<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BuyerPct
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
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboPeriod = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.gb1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.Label2)
        Me.gb1.Controls.Add(Me.cboPeriod)
        Me.gb1.Controls.Add(Me.Label1)
        Me.gb1.Controls.Add(Me.cboYear)
        Me.gb1.Controls.Add(Me.cboDept)
        Me.gb1.Controls.Add(Me.Label5)
        Me.gb1.Controls.Add(Me.cboStore)
        Me.gb1.Controls.Add(Me.Label3)
        Me.gb1.Location = New System.Drawing.Point(27, 12)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(335, 78)
        Me.gb1.TabIndex = 53
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Parameters"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(83, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Period"
        '
        'cboPeriod
        '
        Me.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriod.FormattingEnabled = True
        Me.cboPeriod.Location = New System.Drawing.Point(86, 41)
        Me.cboPeriod.Name = "cboPeriod"
        Me.cboPeriod.Size = New System.Drawing.Size(54, 21)
        Me.cboPeriod.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Year"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(21, 41)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(54, 21)
        Me.cboYear.TabIndex = 11
        '
        'cboDept
        '
        Me.cboDept.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(210, 41)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(119, 21)
        Me.cboDept.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(210, 25)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Dept"
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(148, 41)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(54, 21)
        Me.cboStore.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(148, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Store"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(27, 523)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 52
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.BackgroundColor = System.Drawing.SystemColors.ControlDark
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.GridColor = System.Drawing.SystemColors.ControlLight
        Me.dgv1.Location = New System.Drawing.Point(27, 133)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(989, 384)
        Me.dgv1.TabIndex = 51
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(13, 623)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 54
        Me.serverLabel.Text = "server"
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 28.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.Location = New System.Drawing.Point(170, 306)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(458, 44)
        Me.lblProcessing.TabIndex = 55
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'BuyerPct
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1046, 648)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "BuyerPct"
        Me.Text = "Modify Buyer Percentage"
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboDept As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Days2
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.lblPrdWk = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboDate = New System.Windows.Forms.ComboBox()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.dgv3 = New System.Windows.Forms.DataGridView()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgv3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(322, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(275, 26)
        Me.Label1.TabIndex = 77
        Me.Label1.Text = "Note: Click on %WK to make changes at the store level. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Click on any departmentt" & _
    "o change its day plan amount."
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.BackColor = System.Drawing.SystemColors.Info
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.lblProcessing.Location = New System.Drawing.Point(32, 246)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(731, 73)
        Me.lblProcessing.TabIndex = 76
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(12, 572)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 75
        Me.serverLabel.Text = "server"
        '
        'lblPrdWk
        '
        Me.lblPrdWk.AutoSize = True
        Me.lblPrdWk.Location = New System.Drawing.Point(10, 120)
        Me.lblPrdWk.Name = "lblPrdWk"
        Me.lblPrdWk.Size = New System.Drawing.Size(69, 13)
        Me.lblPrdWk.TabIndex = 74
        Me.lblPrdWk.Text = "Period Week"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cboDate)
        Me.GroupBox1.Controls.Add(Me.cboYear)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cboStore)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(242, 78)
        Me.GroupBox1.TabIndex = 73
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Parameters"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(5, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 13)
        Me.Label6.TabIndex = 33
        Me.Label6.Text = "Year"
        '
        'cboDate
        '
        Me.cboDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDate.FormattingEnabled = True
        Me.cboDate.Location = New System.Drawing.Point(133, 37)
        Me.cboDate.Name = "cboDate"
        Me.cboDate.Size = New System.Drawing.Size(91, 21)
        Me.cboDate.TabIndex = 66
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(6, 38)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(63, 21)
        Me.cboYear.TabIndex = 32
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(133, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 13)
        Me.Label5.TabIndex = 65
        Me.Label5.Text = "Week Beginning"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(75, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Store"
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(75, 38)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(44, 21)
        Me.cboStore.TabIndex = 0
        '
        'dgv3
        '
        Me.dgv3.AllowUserToAddRows = False
        Me.dgv3.AllowUserToDeleteRows = False
        Me.dgv3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv3.Enabled = False
        Me.dgv3.Location = New System.Drawing.Point(13, 454)
        Me.dgv3.Name = "dgv3"
        Me.dgv3.Size = New System.Drawing.Size(795, 105)
        Me.dgv3.TabIndex = 72
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.Location = New System.Drawing.Point(13, 139)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.ReadOnly = True
        Me.dgv1.Size = New System.Drawing.Size(795, 309)
        Me.dgv1.TabIndex = 71
        '
        'Days2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(819, 595)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.lblPrdWk)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgv3)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "Days2"
        Me.Text = "Day Sales by Week"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgv3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents lblPrdWk As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboDate As System.Windows.Forms.ComboBox
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
    Friend WithEvents dgv3 As System.Windows.Forms.DataGridView
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
End Class

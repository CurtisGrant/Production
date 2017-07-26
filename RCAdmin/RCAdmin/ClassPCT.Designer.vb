<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClassPct
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
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboBuyer = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboDept = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboStr = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboPrd = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboYr = New System.Windows.Forms.ComboBox()
        Me.dgv2 = New System.Windows.Forms.DataGridView()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.BackColor = System.Drawing.SystemColors.Info
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.lblProcessing.Location = New System.Drawing.Point(86, 263)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(731, 73)
        Me.lblProcessing.TabIndex = 49
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(14, 517)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 48
        Me.serverLabel.Text = "server"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(13, 489)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(97, 23)
        Me.btnSave.TabIndex = 47
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.Location = New System.Drawing.Point(13, 125)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(913, 294)
        Me.dgv1.TabIndex = 46
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.cboBuyer)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cboDept)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cboStr)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cboPrd)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cboYr)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(406, 96)
        Me.GroupBox1.TabIndex = 45
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Saved Plan"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(298, 38)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 13)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Buyer"
        '
        'cboBuyer
        '
        Me.cboBuyer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBuyer.FormattingEnabled = True
        Me.cboBuyer.Location = New System.Drawing.Point(296, 55)
        Me.cboBuyer.Name = "cboBuyer"
        Me.cboBuyer.Size = New System.Drawing.Size(104, 21)
        Me.cboBuyer.TabIndex = 34
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(174, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Dept"
        '
        'cboDept
        '
        Me.cboDept.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDept.FormattingEnabled = True
        Me.cboDept.Location = New System.Drawing.Point(175, 55)
        Me.cboDept.Name = "cboDept"
        Me.cboDept.Size = New System.Drawing.Size(118, 21)
        Me.cboDept.TabIndex = 27
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 38)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Store"
        '
        'cboStr
        '
        Me.cboStr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStr.FormattingEnabled = True
        Me.cboStr.Location = New System.Drawing.Point(10, 55)
        Me.cboStr.Name = "cboStr"
        Me.cboStr.Size = New System.Drawing.Size(53, 21)
        Me.cboStr.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(117, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Period"
        '
        'cboPrd
        '
        Me.cboPrd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrd.FormattingEnabled = True
        Me.cboPrd.Location = New System.Drawing.Point(120, 55)
        Me.cboPrd.Name = "cboPrd"
        Me.cboPrd.Size = New System.Drawing.Size(53, 21)
        Me.cboPrd.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(62, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Year"
        '
        'cboYr
        '
        Me.cboYr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYr.FormattingEnabled = True
        Me.cboYr.Location = New System.Drawing.Point(65, 55)
        Me.cboYr.Name = "cboYr"
        Me.cboYr.Size = New System.Drawing.Size(53, 21)
        Me.cboYr.TabIndex = 24
        '
        'dgv2
        '
        Me.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv2.Location = New System.Drawing.Point(13, 425)
        Me.dgv2.Name = "dgv2"
        Me.dgv2.Size = New System.Drawing.Size(914, 46)
        Me.dgv2.TabIndex = 50
        '
        'ClassPct
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(939, 539)
        Me.Controls.Add(Me.dgv2)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ClassPct"
        Me.Text = "Modify Class Percentage"
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgv2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboBuyer As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboDept As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboStr As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboPrd As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboYr As System.Windows.Forms.ComboBox
    Friend WithEvents dgv2 As System.Windows.Forms.DataGridView
End Class

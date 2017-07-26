<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DayPct
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
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
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
        Me.lblProcessing.Location = New System.Drawing.Point(26, 318)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(731, 73)
        Me.lblProcessing.TabIndex = 66
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(26, 615)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 65
        Me.serverLabel.Text = "server"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(20, 111)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(492, 32)
        Me.Label5.TabIndex = 64
        Me.Label5.Text = "WARNING! This interface will change each day of week within the selected period." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Use Modify Day Sales Plan to Change the Sales Plan for a Specific Week."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(23, 161)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(188, 16)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Planned Day Percent of Week"
        '
        'gb1
        '
        Me.gb1.Controls.Add(Me.Label1)
        Me.gb1.Controls.Add(Me.cboYear)
        Me.gb1.Controls.Add(Me.cboStore)
        Me.gb1.Controls.Add(Me.Label3)
        Me.gb1.Location = New System.Drawing.Point(20, 12)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(167, 78)
        Me.gb1.TabIndex = 62
        Me.gb1.TabStop = False
        Me.gb1.Text = "Select Parameters"
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
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(91, 41)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(54, 21)
        Me.cboStore.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(91, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Store"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(20, 579)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 61
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv1.Location = New System.Drawing.Point(20, 180)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.Size = New System.Drawing.Size(503, 384)
        Me.dgv1.TabIndex = 60
        '
        'DayPct
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(776, 640)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.dgv1)
        Me.Name = "DayPct"
        Me.Text = "Day Sales Percent by Period"
        Me.gb1.ResumeLayout(False)
        Me.gb1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
End Class

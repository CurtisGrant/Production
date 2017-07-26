<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreatePlan
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cboRnd = New System.Windows.Forms.ComboBox()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.lblSaving = New System.Windows.Forms.Label()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboStore = New System.Windows.Forms.ComboBox()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.BackColor = System.Drawing.SystemColors.Info
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.lblProcessing.Location = New System.Drawing.Point(60, 135)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(364, 37)
        Me.lblProcessing.TabIndex = 57
        Me.lblProcessing.Text = "Processing. Please wait."
        Me.lblProcessing.Visible = False
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(12, 330)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 56
        Me.serverLabel.Text = "server"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboRnd)
        Me.GroupBox2.Location = New System.Drawing.Point(21, 107)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(121, 78)
        Me.GroupBox2.TabIndex = 55
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Round Week Plan to Nearest:"
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
        'pb1
        '
        Me.pb1.Location = New System.Drawing.Point(67, 270)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(202, 23)
        Me.pb1.TabIndex = 54
        Me.pb1.Visible = False
        '
        'lblSaving
        '
        Me.lblSaving.AutoSize = True
        Me.lblSaving.Location = New System.Drawing.Point(21, 270)
        Me.lblSaving.Name = "lblSaving"
        Me.lblSaving.Size = New System.Drawing.Size(40, 13)
        Me.lblSaving.TabIndex = 53
        Me.lblSaving.Text = "Saving"
        Me.lblSaving.Visible = False
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(24, 212)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(118, 23)
        Me.btnCreate.TabIndex = 52
        Me.btnCreate.Text = "Create New Plan"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Select Year"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "Select Store"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(21, 71)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(121, 21)
        Me.cboYear.TabIndex = 49
        '
        'cboStore
        '
        Me.cboStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStore.FormattingEnabled = True
        Me.cboStore.Location = New System.Drawing.Point(24, 29)
        Me.cboStore.Name = "cboStore"
        Me.cboStore.Size = New System.Drawing.Size(121, 21)
        Me.cboStore.TabIndex = 48
        '
        'CreatePlan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(437, 356)
        Me.Controls.Add(Me.lblProcessing)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.lblSaving)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboYear)
        Me.Controls.Add(Me.cboStore)
        Me.Name = "CreatePlan"
        Me.Text = "Select Parameters"
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblProcessing As System.Windows.Forms.Label
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cboRnd As System.Windows.Forms.ComboBox
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblSaving As System.Windows.Forms.Label
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboStore As System.Windows.Forms.ComboBox
End Class

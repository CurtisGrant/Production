<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Calendar
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnPeriod = New System.Windows.Forms.Button()
        Me.btnWeek = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtYrEnd = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtYrBegin = New System.Windows.Forms.TextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabQuarters = New System.Windows.Forms.TabPage()
        Me.TabPeriods = New System.Windows.Forms.TabPage()
        Me.TabWeeks = New System.Windows.Forms.TabPage()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabQuarters.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(478, 561)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(909, 561)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnPeriod
        '
        Me.btnPeriod.Location = New System.Drawing.Point(316, 8)
        Me.btnPeriod.Name = "btnPeriod"
        Me.btnPeriod.Size = New System.Drawing.Size(96, 23)
        Me.btnPeriod.TabIndex = 4
        Me.btnPeriod.Text = "Define Periods"
        Me.btnPeriod.UseVisualStyleBackColor = True
        '
        'btnWeek
        '
        Me.btnWeek.Location = New System.Drawing.Point(316, 44)
        Me.btnWeek.Name = "btnWeek"
        Me.btnWeek.Size = New System.Drawing.Size(97, 23)
        Me.btnWeek.TabIndex = 5
        Me.btnWeek.Text = "Define Weeks"
        Me.btnWeek.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboYear)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtYrEnd)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtYrBegin)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(247, 79)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Business Year"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(145, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(91, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Ends  yyyy-mm-dd"
        '
        'txtYrEnd
        '
        Me.txtYrEnd.Location = New System.Drawing.Point(141, 46)
        Me.txtYrEnd.Name = "txtYrEnd"
        Me.txtYrEnd.Size = New System.Drawing.Size(100, 20)
        Me.txtYrEnd.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Begins  yyyy-mm-dd"
        '
        'txtYrBegin
        '
        Me.txtYrBegin.Location = New System.Drawing.Point(7, 45)
        Me.txtYrBegin.Name = "txtYrBegin"
        Me.txtYrBegin.Size = New System.Drawing.Size(100, 20)
        Me.txtYrBegin.TabIndex = 9
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabQuarters)
        Me.TabControl1.Controls.Add(Me.TabPeriods)
        Me.TabControl1.Controls.Add(Me.TabWeeks)
        Me.TabControl1.Location = New System.Drawing.Point(12, 103)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(460, 513)
        Me.TabControl1.TabIndex = 10
        '
        'TabQuarters
        '
        Me.TabQuarters.Controls.Add(Me.DateTimePicker1)
        Me.TabQuarters.Controls.Add(Me.TextBox1)
        Me.TabQuarters.Location = New System.Drawing.Point(4, 22)
        Me.TabQuarters.Name = "TabQuarters"
        Me.TabQuarters.Padding = New System.Windows.Forms.Padding(3)
        Me.TabQuarters.Size = New System.Drawing.Size(452, 487)
        Me.TabQuarters.TabIndex = 0
        Me.TabQuarters.Text = "Quarters"
        Me.TabQuarters.UseVisualStyleBackColor = True
        '
        'TabPeriods
        '
        Me.TabPeriods.Location = New System.Drawing.Point(4, 22)
        Me.TabPeriods.Name = "TabPeriods"
        Me.TabPeriods.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPeriods.Size = New System.Drawing.Size(452, 487)
        Me.TabPeriods.TabIndex = 1
        Me.TabPeriods.Text = "Periods"
        Me.TabPeriods.UseVisualStyleBackColor = True
        '
        'TabWeeks
        '
        Me.TabWeeks.Location = New System.Drawing.Point(4, 22)
        Me.TabWeeks.Name = "TabWeeks"
        Me.TabWeeks.Size = New System.Drawing.Size(452, 487)
        Me.TabWeeks.TabIndex = 2
        Me.TabWeeks.Text = "Weeks"
        Me.TabWeeks.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(74, 18)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 0
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "MM/dd/yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(341, 17)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(97, 20)
        Me.DateTimePicker1.TabIndex = 1
        '
        'cboYear
        '
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(98, 10)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(57, 21)
        Me.cboYear.TabIndex = 13
        '
        'Calendar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(996, 637)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnWeek)
        Me.Controls.Add(Me.btnPeriod)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSave)
        Me.Name = "Calendar"
        Me.Text = "Calendar"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabQuarters.ResumeLayout(False)
        Me.TabQuarters.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnPeriod As System.Windows.Forms.Button
    Friend WithEvents btnWeek As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtYrEnd As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtYrBegin As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabQuarters As System.Windows.Forms.TabPage
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TabPeriods As System.Windows.Forms.TabPage
    Friend WithEvents TabWeeks As System.Windows.Forms.TabPage
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
End Class

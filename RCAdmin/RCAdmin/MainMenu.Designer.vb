<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
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
        Me.btnControls = New System.Windows.Forms.Button()
        Me.btnDayPct = New System.Windows.Forms.Button()
        Me.btnAdjustDay = New System.Windows.Forms.Button()
        Me.btnBuyer = New System.Windows.Forms.Button()
        Me.btnClassPct = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cboClient = New System.Windows.Forms.ComboBox()
        Me.btnScore = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblProcessing = New System.Windows.Forms.Label()
        Me.btnUpdatePlan = New System.Windows.Forms.Button()
        Me.btnWeeksSupply = New System.Windows.Forms.Button()
        Me.btnMkdn = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSeason = New System.Windows.Forms.Button()
        Me.btnPLine = New System.Windows.Forms.Button()
        Me.btnBuyers = New System.Windows.Forms.Button()
        Me.btnStore = New System.Windows.Forms.Button()
        Me.btnClasses = New System.Windows.Forms.Button()
        Me.btnDept = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.serverLabel = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnForecast = New System.Windows.Forms.Button()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnControls
        '
        Me.btnControls.Location = New System.Drawing.Point(159, 19)
        Me.btnControls.Name = "btnControls"
        Me.btnControls.Size = New System.Drawing.Size(110, 30)
        Me.btnControls.TabIndex = 42
        Me.btnControls.Text = "Controls"
        Me.btnControls.UseVisualStyleBackColor = True
        '
        'btnDayPct
        '
        Me.btnDayPct.Location = New System.Drawing.Point(19, 139)
        Me.btnDayPct.Name = "btnDayPct"
        Me.btnDayPct.Size = New System.Drawing.Size(180, 30)
        Me.btnDayPct.TabIndex = 27
        Me.btnDayPct.Text = "Day Sales Percent by Period"
        Me.btnDayPct.UseVisualStyleBackColor = True
        '
        'btnAdjustDay
        '
        Me.btnAdjustDay.Location = New System.Drawing.Point(232, 139)
        Me.btnAdjustDay.Name = "btnAdjustDay"
        Me.btnAdjustDay.Size = New System.Drawing.Size(180, 30)
        Me.btnAdjustDay.TabIndex = 24
        Me.btnAdjustDay.Text = "Modify Daily Sales Plan"
        Me.btnAdjustDay.UseVisualStyleBackColor = True
        '
        'btnBuyer
        '
        Me.btnBuyer.Location = New System.Drawing.Point(19, 67)
        Me.btnBuyer.Name = "btnBuyer"
        Me.btnBuyer.Size = New System.Drawing.Size(180, 30)
        Me.btnBuyer.TabIndex = 26
        Me.btnBuyer.Text = "Buyer Percent"
        Me.btnBuyer.UseVisualStyleBackColor = True
        '
        'btnClassPct
        '
        Me.btnClassPct.Location = New System.Drawing.Point(233, 67)
        Me.btnClassPct.Name = "btnClassPct"
        Me.btnClassPct.Size = New System.Drawing.Size(180, 30)
        Me.btnClassPct.TabIndex = 28
        Me.btnClassPct.Text = "Class Percent"
        Me.btnClassPct.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cboClient)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(430, 76)
        Me.GroupBox3.TabIndex = 39
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Select Client"
        '
        'cboClient
        '
        Me.cboClient.FormattingEnabled = True
        Me.cboClient.Location = New System.Drawing.Point(19, 27)
        Me.cboClient.Name = "cboClient"
        Me.cboClient.Size = New System.Drawing.Size(121, 21)
        Me.cboClient.TabIndex = 0
        '
        'btnScore
        '
        Me.btnScore.Location = New System.Drawing.Point(19, 29)
        Me.btnScore.Name = "btnScore"
        Me.btnScore.Size = New System.Drawing.Size(180, 30)
        Me.btnScore.TabIndex = 38
        Me.btnScore.Text = "Item Scoring"
        Me.btnScore.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblProcessing)
        Me.GroupBox2.Controls.Add(Me.btnUpdatePlan)
        Me.GroupBox2.Controls.Add(Me.btnDayPct)
        Me.GroupBox2.Controls.Add(Me.btnAdjustDay)
        Me.GroupBox2.Controls.Add(Me.btnBuyer)
        Me.GroupBox2.Controls.Add(Me.btnClassPct)
        Me.GroupBox2.Controls.Add(Me.btnWeeksSupply)
        Me.GroupBox2.Controls.Add(Me.btnMkdn)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 243)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(430, 187)
        Me.GroupBox2.TabIndex = 37
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Sales and Inventory Planning"
        '
        'lblProcessing
        '
        Me.lblProcessing.AutoSize = True
        Me.lblProcessing.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcessing.Location = New System.Drawing.Point(6, 67)
        Me.lblProcessing.Name = "lblProcessing"
        Me.lblProcessing.Size = New System.Drawing.Size(86, 29)
        Me.lblProcessing.TabIndex = 46
        Me.lblProcessing.Text = "Label1"
        Me.lblProcessing.Visible = False
        '
        'btnUpdatePlan
        '
        Me.btnUpdatePlan.Location = New System.Drawing.Point(19, 31)
        Me.btnUpdatePlan.Name = "btnUpdatePlan"
        Me.btnUpdatePlan.Size = New System.Drawing.Size(393, 30)
        Me.btnUpdatePlan.TabIndex = 22
        Me.btnUpdatePlan.Text = "Sales Plan"
        Me.btnUpdatePlan.UseVisualStyleBackColor = True
        '
        'btnWeeksSupply
        '
        Me.btnWeeksSupply.Location = New System.Drawing.Point(232, 103)
        Me.btnWeeksSupply.Name = "btnWeeksSupply"
        Me.btnWeeksSupply.Size = New System.Drawing.Size(180, 30)
        Me.btnWeeksSupply.TabIndex = 23
        Me.btnWeeksSupply.Text = "Plan Weeks On Hand"
        Me.btnWeeksSupply.UseVisualStyleBackColor = True
        '
        'btnMkdn
        '
        Me.btnMkdn.Location = New System.Drawing.Point(19, 103)
        Me.btnMkdn.Name = "btnMkdn"
        Me.btnMkdn.Size = New System.Drawing.Size(180, 30)
        Me.btnMkdn.TabIndex = 21
        Me.btnMkdn.Text = "Markdown Plan Percentages"
        Me.btnMkdn.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSeason)
        Me.GroupBox1.Controls.Add(Me.btnControls)
        Me.GroupBox1.Controls.Add(Me.btnPLine)
        Me.GroupBox1.Controls.Add(Me.btnBuyers)
        Me.GroupBox1.Controls.Add(Me.btnStore)
        Me.GroupBox1.Controls.Add(Me.btnClasses)
        Me.GroupBox1.Controls.Add(Me.btnDept)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 94)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(430, 143)
        Me.GroupBox1.TabIndex = 36
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Edit Tables"
        '
        'btnSeason
        '
        Me.btnSeason.Location = New System.Drawing.Point(302, 55)
        Me.btnSeason.Name = "btnSeason"
        Me.btnSeason.Size = New System.Drawing.Size(110, 30)
        Me.btnSeason.TabIndex = 25
        Me.btnSeason.Text = "Seasons"
        Me.btnSeason.UseVisualStyleBackColor = True
        '
        'btnPLine
        '
        Me.btnPLine.Location = New System.Drawing.Point(159, 55)
        Me.btnPLine.Name = "btnPLine"
        Me.btnPLine.Size = New System.Drawing.Size(110, 30)
        Me.btnPLine.TabIndex = 24
        Me.btnPLine.Text = "Product Lines"
        Me.btnPLine.UseVisualStyleBackColor = True
        '
        'btnBuyers
        '
        Me.btnBuyers.Location = New System.Drawing.Point(19, 19)
        Me.btnBuyers.Name = "btnBuyers"
        Me.btnBuyers.Size = New System.Drawing.Size(110, 30)
        Me.btnBuyers.TabIndex = 18
        Me.btnBuyers.Text = " Buyers"
        Me.btnBuyers.UseVisualStyleBackColor = True
        '
        'btnStore
        '
        Me.btnStore.Location = New System.Drawing.Point(19, 91)
        Me.btnStore.Name = "btnStore"
        Me.btnStore.Size = New System.Drawing.Size(110, 30)
        Me.btnStore.TabIndex = 23
        Me.btnStore.Text = "Stores"
        Me.btnStore.UseVisualStyleBackColor = True
        '
        'btnClasses
        '
        Me.btnClasses.Location = New System.Drawing.Point(302, 19)
        Me.btnClasses.Name = "btnClasses"
        Me.btnClasses.Size = New System.Drawing.Size(110, 30)
        Me.btnClasses.TabIndex = 19
        Me.btnClasses.Text = "Classes"
        Me.btnClasses.UseVisualStyleBackColor = True
        '
        'btnDept
        '
        Me.btnDept.Location = New System.Drawing.Point(19, 55)
        Me.btnDept.Name = "btnDept"
        Me.btnDept.Size = New System.Drawing.Size(110, 30)
        Me.btnDept.TabIndex = 20
        Me.btnDept.Text = "Departments"
        Me.btnDept.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(516, 422)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(148, 50)
        Me.btnExit.TabIndex = 35
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'serverLabel
        '
        Me.serverLabel.AutoSize = True
        Me.serverLabel.Location = New System.Drawing.Point(12, 609)
        Me.serverLabel.Name = "serverLabel"
        Me.serverLabel.Size = New System.Drawing.Size(36, 13)
        Me.serverLabel.TabIndex = 43
        Me.serverLabel.Text = "server"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnForecast)
        Me.GroupBox5.Controls.Add(Me.btnScore)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 465)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(430, 112)
        Me.GroupBox5.TabIndex = 45
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Miscellaneous"
        '
        'btnForecast
        '
        Me.btnForecast.Location = New System.Drawing.Point(233, 29)
        Me.btnForecast.Name = "btnForecast"
        Me.btnForecast.Size = New System.Drawing.Size(180, 30)
        Me.btnForecast.TabIndex = 39
        Me.btnForecast.Text = "Item Forecasting"
        Me.btnForecast.UseVisualStyleBackColor = True
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(999, 629)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.serverLabel)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "MainMenu"
        Me.Text = "Retail Clarity Administration"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnControls As System.Windows.Forms.Button
    Friend WithEvents btnDayPct As System.Windows.Forms.Button
    Friend WithEvents btnAdjustDay As System.Windows.Forms.Button
    Friend WithEvents btnBuyer As System.Windows.Forms.Button
    Friend WithEvents btnClassPct As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cboClient As System.Windows.Forms.ComboBox
    Friend WithEvents btnScore As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnWeeksSupply As System.Windows.Forms.Button
    Friend WithEvents btnUpdatePlan As System.Windows.Forms.Button
    Friend WithEvents btnMkdn As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBuyers As System.Windows.Forms.Button
    Friend WithEvents btnStore As System.Windows.Forms.Button
    Friend WithEvents btnClasses As System.Windows.Forms.Button
    Friend WithEvents btnDept As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents serverLabel As System.Windows.Forms.Label
    Friend WithEvents btnSeason As System.Windows.Forms.Button
    Friend WithEvents btnPLine As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnForecast As System.Windows.Forms.Button
    Friend WithEvents lblProcessing As System.Windows.Forms.Label

End Class

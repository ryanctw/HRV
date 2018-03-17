<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.gbTop = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnStartTest = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnDel = New System.Windows.Forms.Button()
        Me.rbF = New System.Windows.Forms.RadioButton()
        Me.rbM = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpBrith = New System.Windows.Forms.DateTimePicker()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPid = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.gb3 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbLogin_Name = New System.Windows.Forms.Label()
        Me.btnGetReport = New System.Windows.Forms.Button()
        Me.cbCondition = New System.Windows.Forms.ComboBox()
        Me.dgvData = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtKeyWord = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pbLogo = New System.Windows.Forms.PictureBox()
        Me.lbTitle = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tmr_logout = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label()
        Me.gbTop.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.gb3.SuspendLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gbTop
        '
        Me.gbTop.BackColor = System.Drawing.Color.LightSteelBlue
        Me.gbTop.Controls.Add(Me.GroupBox2)
        Me.gbTop.Controls.Add(Me.gb3)
        Me.gbTop.Location = New System.Drawing.Point(13, 25)
        Me.gbTop.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.gbTop.Name = "gbTop"
        Me.gbTop.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.gbTop.Size = New System.Drawing.Size(976, 560)
        Me.gbTop.TabIndex = 4
        Me.gbTop.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.LightSteelBlue
        Me.GroupBox2.Controls.Add(Me.btnStartTest)
        Me.GroupBox2.Controls.Add(Me.btnEdit)
        Me.GroupBox2.Controls.Add(Me.btnSave)
        Me.GroupBox2.Controls.Add(Me.btnNew)
        Me.GroupBox2.Controls.Add(Me.btnDel)
        Me.GroupBox2.Controls.Add(Me.rbF)
        Me.GroupBox2.Controls.Add(Me.rbM)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.dtpBrith)
        Me.GroupBox2.Controls.Add(Me.txtName)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtPid)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.GroupBox2.Location = New System.Drawing.Point(18, 21)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GroupBox2.Size = New System.Drawing.Size(941, 121)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "檢測人"
        '
        'btnStartTest
        '
        Me.btnStartTest.Image = CType(resources.GetObject("btnStartTest.Image"), System.Drawing.Image)
        Me.btnStartTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnStartTest.Location = New System.Drawing.Point(608, 30)
        Me.btnStartTest.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnStartTest.Name = "btnStartTest"
        Me.btnStartTest.Size = New System.Drawing.Size(289, 75)
        Me.btnStartTest.TabIndex = 10
        Me.btnStartTest.Text = "                 開始檢測"
        Me.btnStartTest.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(398, 27)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(76, 32)
        Me.btnEdit.TabIndex = 8
        Me.btnEdit.Text = "修改"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(478, 27)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 32)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "儲存"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(236, 27)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(76, 32)
        Me.btnNew.TabIndex = 6
        Me.btnNew.Text = "新增"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDel
        '
        Me.btnDel.Location = New System.Drawing.Point(318, 27)
        Me.btnDel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(76, 32)
        Me.btnDel.TabIndex = 7
        Me.btnDel.Text = "刪除"
        Me.btnDel.UseVisualStyleBackColor = True
        '
        'rbF
        '
        Me.rbF.AutoSize = True
        Me.rbF.Location = New System.Drawing.Point(512, 77)
        Me.rbF.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rbF.Name = "rbF"
        Me.rbF.Size = New System.Drawing.Size(42, 22)
        Me.rbF.TabIndex = 5
        Me.rbF.Text = "女"
        Me.rbF.UseVisualStyleBackColor = True
        '
        'rbM
        '
        Me.rbM.AutoSize = True
        Me.rbM.Location = New System.Drawing.Point(473, 78)
        Me.rbM.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rbM.Name = "rbM"
        Me.rbM.Size = New System.Drawing.Size(42, 22)
        Me.rbM.TabIndex = 4
        Me.rbM.Text = "男"
        Me.rbM.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(434, 79)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "性別"
        '
        'dtpBrith
        '
        Me.dtpBrith.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.dtpBrith.Location = New System.Drawing.Point(271, 73)
        Me.dtpBrith.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.dtpBrith.Name = "dtpBrith"
        Me.dtpBrith.Size = New System.Drawing.Size(148, 26)
        Me.dtpBrith.TabIndex = 3
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(103, 73)
        Me.txtName.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(111, 26)
        Me.txtName.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(233, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 18)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "生日"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(63, 76)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 18)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "姓名"
        '
        'txtPid
        '
        Me.txtPid.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.txtPid.Location = New System.Drawing.Point(103, 30)
        Me.txtPid.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtPid.Name = "txtPid"
        Me.txtPid.Size = New System.Drawing.Size(111, 26)
        Me.txtPid.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(63, 33)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 18)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "編號"
        '
        'gb3
        '
        Me.gb3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.gb3.Controls.Add(Me.Label1)
        Me.gb3.Controls.Add(Me.lbLogin_Name)
        Me.gb3.Controls.Add(Me.btnGetReport)
        Me.gb3.Controls.Add(Me.cbCondition)
        Me.gb3.Controls.Add(Me.dgvData)
        Me.gb3.Controls.Add(Me.txtKeyWord)
        Me.gb3.Controls.Add(Me.Label2)
        Me.gb3.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.gb3.Location = New System.Drawing.Point(18, 159)
        Me.gb3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.gb3.Name = "gb3"
        Me.gb3.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.gb3.Size = New System.Drawing.Size(941, 392)
        Me.gb3.TabIndex = 2
        Me.gb3.TabStop = False
        Me.gb3.Text = "檢測紀錄"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 18)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "操作人員 : "
        '
        'lbLogin_Name
        '
        Me.lbLogin_Name.AutoSize = True
        Me.lbLogin_Name.Location = New System.Drawing.Point(96, 38)
        Me.lbLogin_Name.Name = "lbLogin_Name"
        Me.lbLogin_Name.Size = New System.Drawing.Size(38, 18)
        Me.lbLogin_Name.TabIndex = 15
        Me.lbLogin_Name.Text = "OOO"
        '
        'btnGetReport
        '
        Me.btnGetReport.Location = New System.Drawing.Point(790, 336)
        Me.btnGetReport.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnGetReport.Name = "btnGetReport"
        Me.btnGetReport.Size = New System.Drawing.Size(136, 41)
        Me.btnGetReport.TabIndex = 14
        Me.btnGetReport.Text = "查詢紀錄"
        Me.btnGetReport.UseVisualStyleBackColor = True
        '
        'cbCondition
        '
        Me.cbCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbCondition.DropDownWidth = 78
        Me.cbCondition.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.cbCondition.FormattingEnabled = True
        Me.cbCondition.IntegralHeight = False
        Me.cbCondition.ItemHeight = 18
        Me.cbCondition.Items.AddRange(New Object() {"編號", "姓名", "生日"})
        Me.cbCondition.Location = New System.Drawing.Point(687, 35)
        Me.cbCondition.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cbCondition.Name = "cbCondition"
        Me.cbCondition.Size = New System.Drawing.Size(76, 26)
        Me.cbCondition.TabIndex = 11
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToResizeColumns = False
        Me.dgvData.AllowUserToResizeRows = False
        Me.dgvData.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvData.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5})
        Me.dgvData.Location = New System.Drawing.Point(16, 75)
        Me.dgvData.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.dgvData.MultiSelect = False
        Me.dgvData.Name = "dgvData"
        Me.dgvData.ReadOnly = True
        Me.dgvData.RowHeadersVisible = False
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgvData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvData.Size = New System.Drawing.Size(910, 245)
        Me.dgvData.TabIndex = 13
        '
        'Column1
        '
        Me.Column1.HeaderText = "ID"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Visible = False
        Me.Column1.Width = 200
        '
        'Column2
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column2.HeaderText = "編號"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 240
        '
        'Column3
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column3.HeaderText = "姓名"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 240
        '
        'Column4
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column4.HeaderText = "生日"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 230
        '
        'Column5
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column5.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column5.HeaderText = "性別"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        Me.Column5.Width = 200
        '
        'txtKeyWord
        '
        Me.txtKeyWord.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.txtKeyWord.Location = New System.Drawing.Point(771, 34)
        Me.txtKeyWord.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtKeyWord.Name = "txtKeyWord"
        Me.txtKeyWord.Size = New System.Drawing.Size(155, 26)
        Me.txtKeyWord.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.Label2.Location = New System.Drawing.Point(617, 38)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "查詢條件"
        '
        'pbLogo
        '
        Me.pbLogo.Image = CType(resources.GetObject("pbLogo.Image"), System.Drawing.Image)
        Me.pbLogo.Location = New System.Drawing.Point(188, 119)
        Me.pbLogo.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pbLogo.Name = "pbLogo"
        Me.pbLogo.Size = New System.Drawing.Size(701, 351)
        Me.pbLogo.TabIndex = 5
        Me.pbLogo.TabStop = False
        '
        'lbTitle
        '
        Me.lbTitle.AutoSize = True
        Me.lbTitle.Font = New System.Drawing.Font("Comic Sans MS", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTitle.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.lbTitle.Location = New System.Drawing.Point(200, 154)
        Me.lbTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbTitle.Name = "lbTitle"
        Me.lbTitle.Size = New System.Drawing.Size(235, 52)
        Me.lbTitle.TabIndex = 6
        Me.lbTitle.Text = "HRVapp v1.0"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Comic Sans MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.Label7.Location = New System.Drawing.Point(234, 206)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(155, 23)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Your health partner"
        '
        'tmr_logout
        '
        Me.tmr_logout.Interval = 1000
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label8.Location = New System.Drawing.Point(664, 554)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(305, 21)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "旺三豐生技股份有限公司版權所有©2016"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(999, 597)
        Me.Controls.Add(Me.gbTop)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lbTitle)
        Me.Controls.Add(Me.pbLogo)
        Me.Controls.Add(Me.Label8)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximumSize = New System.Drawing.Size(1015, 635)
        Me.MinimumSize = New System.Drawing.Size(1015, 635)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HRVapp"
        Me.gbTop.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gb3.ResumeLayout(False)
        Me.gb3.PerformLayout()
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gbTop As System.Windows.Forms.GroupBox
    Friend WithEvents gb3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtKeyWord As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbF As System.Windows.Forms.RadioButton
    Friend WithEvents rbM As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpBrith As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPid As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbCondition As System.Windows.Forms.ComboBox
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents btnGetReport As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents btnStartTest As System.Windows.Forms.Button
    Friend WithEvents pbLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lbTitle As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents tmr_logout As System.Windows.Forms.Timer
    Friend WithEvents lbLogin_Name As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label

End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDoHRV
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
        Me.gbName = New System.Windows.Forms.GroupBox()
        Me.labTesing = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.web_msg = New System.Windows.Forms.WebBrowser()
        Me.gbName.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbName
        '
        Me.gbName.BackColor = System.Drawing.Color.LightSteelBlue
        Me.gbName.Controls.Add(Me.labTesing)
        Me.gbName.Controls.Add(Me.btnCancel)
        Me.gbName.Location = New System.Drawing.Point(22, 19)
        Me.gbName.Margin = New System.Windows.Forms.Padding(4)
        Me.gbName.Name = "gbName"
        Me.gbName.Padding = New System.Windows.Forms.Padding(4)
        Me.gbName.Size = New System.Drawing.Size(739, 122)
        Me.gbName.TabIndex = 0
        Me.gbName.TabStop = False
        '
        'labTesing
        '
        Me.labTesing.AutoSize = True
        Me.labTesing.Location = New System.Drawing.Point(36, 56)
        Me.labTesing.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.labTesing.Name = "labTesing"
        Me.labTesing.Size = New System.Drawing.Size(65, 18)
        Me.labTesing.TabIndex = 8
        Me.labTesing.Text = "檢測中..."
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(633, 47)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 37)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "取消"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.web_msg)
        Me.GroupBox2.Location = New System.Drawing.Point(19, 316)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(10)
        Me.GroupBox2.Size = New System.Drawing.Size(698, 333)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'web_msg
        '
        Me.web_msg.Location = New System.Drawing.Point(13, 235)
        Me.web_msg.Name = "web_msg"
        Me.web_msg.Size = New System.Drawing.Size(672, 284)
        Me.web_msg.TabIndex = 1
        '
        'frmDoHRV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 162)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.gbName)
        Me.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximumSize = New System.Drawing.Size(800, 200)
        Me.MinimumSize = New System.Drawing.Size(800, 200)
        Me.Name = "frmDoHRV"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HRVapp"
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents labTesing As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents web_msg As System.Windows.Forms.WebBrowser
End Class

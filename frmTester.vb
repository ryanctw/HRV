Imports System.Data.OleDb

Public Class frmTester

    Dim btn_status As String
    Public current_tester_id As Integer = -1

    Private Sub frmTester_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        frmMain.Enabled = True

    End Sub 'Form1_Closing

    Private Sub frmTester_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = frmMain.HRVapp

        LoadTesterRecord()

        Set_Default()

    End Sub

    Public Sub LoadTesterRecord()

        Dim OleDBC As New OleDbCommand
        Dim OleDBDR As OleDbDataReader
        Dim c As Integer
        c = 0

        OleDBC = frmMain.myHrv.HRV_Get_tTester

        OleDBDR = OleDBC.ExecuteReader
        dgvData.Rows.Clear()
        If OleDBDR.HasRows Then
            While OleDBDR.Read
                dgvData.Rows.Add()

                dgvData.Item(0, c).Value = OleDBDR.Item(0)
                dgvData.Item(1, c).Value = OleDBDR.Item(1)
                dgvData.Item(2, c).Value = OleDBDR.Item(2)
                dgvData.Item(3, c).Value = OleDBDR.Item(3)

                c = c + 1
            End While
        Else
            'btnGetReport.Text = "查詢紀錄 ( 0 筆 )"
        End If
    End Sub

    Private Sub Set_Default()

        btnNew.Text = "新增"
        btnNew.Visible = True
        btnDel.Visible = False
        btnEdit.Visible = False
        btnSave.Visible = False

        txtLogin_Id.Text = ""
        txtLogin_Id.Enabled = False
        txtName.Text = ""
        txtName.Enabled = False
        txtUnit.Text = ""
        txtUnit.Enabled = False

        gb3.Visible = True

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click

        If btnNew.Text = "新增" Then

            btn_status = "新增"

            txtName.Enabled = True
            txtName.Text = ""
            txtName.Select()
            txtLogin_Id.Enabled = True
            txtLogin_Id.Text = ""
            txtUnit.Enabled = True
            txtUnit.Text = ""

            btnNew.Text = "取消"

            btnSave.Visible = True
            btnEdit.Visible = False
            btnDel.Visible = False
            gb3.Visible = False
        Else
            btnNew.Text = "新增"
            Set_Default()
        End If

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        If txtLogin_Id.Text = "系統管理者" Then
            MsgBox("無法修改 [系統管理者]", MsgBoxStyle.Critical)
            Exit Sub
        End If

        If btnEdit.Text = "修改" Then

            btn_status = "修改"

            txtName.Enabled = True
            txtName.Select()
            txtLogin_Id.Enabled = True
            btnEdit.Text = "取消"
            txtUnit.Enabled = True

            btnNew.Visible = False
            btnDel.Visible = False
            btnSave.Visible = True

            gb3.Visible = False
        Else
            btnEdit.Text = "修改"
            Set_Default()
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Try

            Console.WriteLine(btn_status)

            If txtName.Text = "" Then
                MsgBox("錯誤 : [姓名] 欄位請勿空白", MsgBoxStyle.Critical)
                txtName.Select()
                Exit Sub
            End If

            If txtLogin_Id.Text = "" Then
                MsgBox("錯誤 : [登入帳號] 欄位請勿空白", MsgBoxStyle.Critical)
                txtLogin_Id.Select()
                Exit Sub
            End If

            If txtUnit.Text = "" Then
                txtUnit.Text = "---"
            End If

            If btn_status = "新增" Then
                'Console.WriteLine(dtpBrith.Value.ToShortDateString)
                If frmMain.myHrv.HRV_NEW_TESTER(txtName.Text, txtLogin_Id.Text, txtLogin_Id.Text, "登入帳號", txtUnit.Text) = True Then
                    MsgBox("資料已新增", MsgBoxStyle.Information)
                    LoadTesterRecord()
                    Set_Default()
                End If
                txtLogin_Id.Select()
            End If

            If btn_status = "修改" Then
                'Console.WriteLine(dtpBrith.Value.ToShortDateString)
                Console.WriteLine("修改 " & current_tester_id)
                If frmMain.myHrv.HRV_UPDATE_TESTER(current_tester_id, txtName.Text, txtLogin_Id.Text, txtUnit.Text) = True Then
                    MsgBox("資料已修改", MsgBoxStyle.Information)
                    LoadTesterRecord()
                    Set_Default()
                End If
            End If

        Catch Ex As Exception
            MsgBox(btn_status & "錯誤 : " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click

        Console.WriteLine("刪除 " & current_tester_id)

        If txtName.Text = "系統管理者" Then
            MsgBox("無法刪除 [系統管理者]", MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim result As Integer = MessageBox.Show("確定刪除 " & txtLogin_Id.Text & " ?", frmMain.HRVapp, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            If frmMain.myHrv.HRV_DEL_TESTER(current_tester_id) = True Then
                MsgBox("資料已刪除", MsgBoxStyle.Information)
                LoadTesterRecord()
                Set_Default()
            End If
        End If

    End Sub

    Private Sub dgvData_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvData.CellMouseClick

        If e.RowIndex < 0 Then
            Exit Sub
        End If

        current_tester_id = dgvData.Item(0, e.RowIndex).Value
        txtName.Text = dgvData.Item(1, e.RowIndex).Value
        txtLogin_Id.Text = dgvData.Item(2, e.RowIndex).Value
        txtUnit.Text = dgvData.Item(3, e.RowIndex).Value

        If txtName.Text = "系統管理者" Then
            btnEdit.Visible = False
            btnDel.Visible = False
        Else
            btnEdit.Text = "修改"
            btnDel.Text = "刪除"
            btnEdit.Visible = True
            btnDel.Visible = True
        End If
       

    End Sub
End Class


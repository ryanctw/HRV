Public Class frmPwd

    Private Sub frmPwd_Closing(sender As Object, e As EventArgs) Handles MyBase.Load

        frmMain.Enabled = True

    End Sub
    Private Sub frmPwd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = frmMain.HRVapp
        lbLogin_Id.Text = frmMain.current_login_id

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        frmMain.Enabled = True
        Me.Close()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        If txtPWD_1.Text.ToUpper <> txtPWD_2.Text.ToUpper Then
            MsgBox("密碼輸入錯誤", MsgBoxStyle.Critical)
            txtPWD_1.Select()
            Exit Sub
        End If

        'Console.WriteLine(lbLogin_Id.Text)
        'Console.WriteLine(frmMain.current_login_id)
        If frmMain.myHrv.HRV_UPDATE_TESTER_PWD(lbLogin_Id.Text, txtPWD_1.Text) = True Then
            MsgBox("密碼變更完成", MsgBoxStyle.Information)
            Me.Close()
        End If

    End Sub
End Class
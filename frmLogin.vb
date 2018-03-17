Public Class frmLogin

    Private Sub frmLogin_Closing(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtID.Select()

    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        Login()

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        frmMain.Enabled = True
        Me.Close()

    End Sub

    Private Sub txtPWD_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPWD.KeyDown

        'frmMain.show_msg(e.KeyCode)

        If e.KeyCode = Keys.Enter Then
            Login()
        End If

    End Sub

    Private Function Login() As Boolean

        If txtID.Text.Length = 0 Then
            MessageBox.Show("帳號請勿空白", "登入", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtID.Select()
        ElseIf txtPWD.Text.Length = 0 Then
            MessageBox.Show("密碼請勿空白", "登入", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPWD.Select()
        Else
            Dim login_name = frmMain.myHrv.HRV_LOGIN(txtID.Text, txtPWD.Text)

            If login_name.Length > 0 Then

                Dim login_unit = frmMain.myHrv.HRV_LOGIN_UNIT(txtID.Text, txtPWD.Text)

                frmMain.tmr_logout.Interval = 1000
                frmMain.logout_timer = 0
                'frmMain.logout_timer_max = 60 * 10 '10 Min
                'frmMain.logout_timer_max = 60 * 1 'For test only
                frmMain.tmr_logout.Enabled = True
                frmMain.gbTop.Visible = True
                frmMain.mConfig.Visible = True
                frmMain.Enabled = True

                frmMain.current_login_name = login_name
                frmMain.current_login_id = txtID.Text
                If login_unit = "---" Then
                    login_unit = ""
                End If
                frmMain.current_login_unit = login_unit
                frmMain.lbLogin_Name.Text = login_name & "      檢測單位 : " & login_unit

                If txtID.Text.ToUpper = "ADMIN" Then
                    frmMain.mConfig_Com.Visible = True
                    frmMain.mConfig_Backup.Visible = True
                    frmMain.mConfig_TESTER.Visible = True
                    frmMain.mConfig_COUNT.Visible = True
                Else
                    frmMain.mConfig_Com.Visible = True
                    frmMain.mConfig_Backup.Visible = False
                    frmMain.mConfig_TESTER.Visible = False
                    frmMain.mConfig_COUNT.Visible = False
                End If

                frmMain.mSystem_Login.Text = "登出"
                Me.Close()
            Else
                MessageBox.Show("帳號或密碼錯誤", "登入", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        Return 0

    End Function

End Class
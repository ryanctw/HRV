Imports libhrv
Imports System.Threading
Imports System.IO

Public Class frmDoHRV

    'Dim mdb_FilePath As String = Application.StartupPath
    'Dim myHrv As New Hrvlib
    Dim STOP_FLAG As Integer = 0

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        If MessageBox.Show("確定取消?", frmMain.HRVapp, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            STOP_FLAG = 1
            frmMain.myHrv.HRV_PROCESS_OUT()
            frmMain.myHrv.HRV_CLEAN_FILES()
            'frmMain.myHrv.HRV_COM_DISABLE()
            frmMain.btnStartTest.Text = "                 開始檢測"
            'frmMain.Enabled = True
            frmMain.gbTop.Enabled = True
            frmMain.logout_timer = 0
            frmMain.tmr_logout.Enabled = True
            'Me.Dispose()
            'Exit Sub
            Me.Close()

        End If


    End Sub

    Private Sub frmDoHRV_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.Text = frmMain.HRVapp & " (" & frmMain.current_com_port & ")"
        gbName.Text = frmMain.current_name & " (" & frmMain.current_pid & ")"
        Dim disply_string As String = "檢測中"
        Dim disply_dot As String = " . "
        Dim img_index As Integer = 0

        'Dim web_name As String
        'Dim web_prefix As String = "AD"
        'Dim web_index As Integer = 1

        Dim GO_NEXT As Integer = 1

        Try

            '----------------------------------------------------------------------
            ' 工商服務
            '----------------------------------------------------------------------
            web_msg.Visible = True
            'web_name = "index.html"
            'web_msg.Navigate(frmMain.app_FilePath & "\" & web_name)
            While (1)
                'web_name = web_prefix & web_index & ".html"
                'web_msg.Navigate(frmMain.app_FilePath & "\" & web_name)
                For i As Integer = 0 To 20
                    disply_string = disply_string + disply_dot
                    labTesing.Text = disply_string
                    Thread.Sleep(200)
                    Application.DoEvents()
                Next

                If STOP_FLAG = 1 Then
                    Console.WriteLine("frmDoHRV_Shown Exit Sub")
                    Exit Sub
                End If
                '----------------------------------------------------------------------
                ' 跑馬燈
                '----------------------------------------------------------------------
                disply_string = "檢測中"
                disply_dot = " . "

                'web_index = web_index + 1
                'web_name = web_prefix & web_index & "\" & ".html"
                'web_msg.Dispose()
                'web_msg.Navigate(frmMain.app_FilePath & web_name)
                'If web_index > 3 Then
                ' web_index = 1
                ' End If

                '----------------------------------------------------------------------
                If frmMain.myHrv.HRV_CHECK_DATA() = True Then
                    'MsgBox("檢測完成", MsgBoxStyle.Information)
                    labTesing.Text = "資料接收中......"
                    frmMain.myHrv.HRV_PROCESS_OUT()
                    web_msg.Visible = False

                    If frmMain.myHrv.HRV_DATA_PARSER(frmMain.current_user_id, frmMain.current_login_name, frmMain.current_login_unit) <> True Then
                        'End
                        GO_NEXT = 0
                        'MsgBox("錯誤 : HRV資料異常,請重新檢測.", MsgBoxStyle.Critical)
                    Else
                        GO_NEXT = 1
                        'frmMain.btnStartTest.Text = "                 開始檢測"
                        'frmMain.Enabled = True
                        'frmMain.logout_timer = 0
                        'frmMain.tmr_logout.Enabled = True
                    End If

                    Exit While

                End If

                GO_NEXT = 0

            End While

            'frmMain.myHrv.HRV_PROCESS_OUT()
            'Me.Dispose()
            'Me.Close()

        Catch ex As Exception
            'frmMain.myHrv.HRV_PROCESS_OUT()
            'Me.Dispose()
            'Me.Close()

        End Try

        frmMain.myHrv.HRV_PROCESS_OUT()
        'GO_NEXT = 1 'DEBUG ONLY

        If GO_NEXT = 1 Then

            Dim result As String = MsgBox("檢測完成,是否讀取檢測報告?", MsgBoxStyle.YesNo, frmMain.HRVapp)

            If result = vbYes Then

                If frmMain.myHrv.HRV_PROCESS_DATA(frmMain.current_user_id, frmMain.current_birth, frmMain.current_sex) = True Then

                    frmDataList.Make_Report()

                    Dim work_file_name As String = Now().ToString("yyyyMMddHHmmss")
                    Debug.WriteLine(work_file_name)
                    'Dim path As String = Directory.GetCurrentDirectory() & "\data\"
                    Dim path As String = Application.StartupPath() & "\data\"
                    Console.WriteLine(path)

                    'If File.Exists(path & "index.pdf") Then
                    '    File.Delete(path & "index.pdf")
                    'End If

                    Dim pHelp As New ProcessStartInfo
                    pHelp.FileName = path & frmDataList.HTML2PDF
                    'pHelp.Arguments = path & "report.html " & path & work_file_name & ".pdf"
                    pHelp.Arguments = path & "report_page_1.html " & path & "report_page_2.html " & path & "report_page_3.html " & path & "report_page_4.html " & path & work_file_name & ".pdf"
                    'pHelp.FileName = path & "Html2PDF.bat"
                    pHelp.UseShellExecute = True
                    pHelp.WindowStyle = ProcessWindowStyle.Hidden
                    Dim proc As Process = Process.Start(pHelp)
                    proc.Dispose()

                    While (1)
                        If File.Exists(path & work_file_name & ".pdf") Then
                            Exit While
                        End If
                        Thread.Sleep(200)
                        Application.DoEvents()
                    End While

                    MsgBox("檢測報告轉檔成功." & _
                           path & work_file_name & ".pdf", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, frmMain.HRVapp)

                    Dim pdf_proc As New Process
                    'Set the process details
                    With pdf_proc.StartInfo
                        'Set the information for the file to launch
                        '.Verb = "print"
                        .FileName = path & work_file_name & ".pdf"
                        .UseShellExecute = True
                    End With
                    'Open the file
                    pdf_proc.Start()
                    pdf_proc.Dispose()

                    'Me.Close() 'Current Form Closed

                Else
                    MsgBox("錯誤 : HRV資料異常,請重新檢測.", MsgBoxStyle.Critical)

                End If


            Else

                'Me.Close()
            End If

        Else
            'Me.Dispose()
            MsgBox("錯誤 : HRV資料異常,請重新檢測.", MsgBoxStyle.Critical)
            'Me.Close()
        End If

        frmMain.btnStartTest.Text = "                 開始檢測"
        'frmMain.Enabled = True
        frmMain.gbTop.Enabled = True
        frmMain.logout_timer = 0
        frmMain.tmr_logout.Enabled = True
        Me.Close()

        'myHrv.COM_STATUS = False
        'frmMain.myHrv.HRV_COM_COUNT = 0

        'While (1)
        '    'If myHrv.COM_STATUS = True Then
        '    '    Console.WriteLine("++++++++++++++++++")
        '    '    Exit While
        '    'Else
        '    '    Console.WriteLine("--------------------")
        '    'End If
        '    Thread.Sleep(200)
        '    'Console.WriteLine("xxxxxxxxxxxxxxxxxxxxxxx " & frmMian.myHrv.HRV_Get_Count)
        '    labTesing.Text = frmMain.myHrv.HRV_COM_COUNT.ToString
        '    If frmMain.myHrv.HRV_COM_COUNT > (frmMain.myHrv.SHARE_LEN - 1) Then
        '        Exit While
        '    End If
        '    Application.DoEvents()
        'End While

        'frmMain.myHrv.HRV_COM_DISABLE()
        'frmMain.btnStartTest.Text = "                 開始檢測"
        'frmMain.Enabled = True
        'frmMain.logout_timer = 0
        'frmMain.tmr_logout.Enabled = True
        'Me.Dispose()

    End Sub
   
End Class
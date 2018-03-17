Imports System.Data.OleDb

Public Class frmCount

    Private Sub frmCount_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmMain.Enabled = True

    End Sub

    Private Sub frmCount_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = frmMain.HRVapp

        LoadDataRecord()

    End Sub

    Private Sub LoadDataRecord()

        Dim OleDBC As New OleDbCommand
        Dim OleDBDR As OleDbDataReader

        Dim YYYYMMDD As New ArrayList()
        Dim tmp As String = ""
        Dim i As Integer

        YYYYMMDD.Clear()

        dgvData.Rows.Clear()

        Try

            OleDBC = frmMain.myHrv.HRV_Get_tData_Date()
            OleDBDR = OleDBC.ExecuteReader

            If OleDBDR.HasRows Then
                While OleDBDR.Read

                    Console.WriteLine(OleDBDR.Item(0))
                    If tmp = OleDBDR.Item(0).ToString Then

                    Else
                        YYYYMMDD.Add(OleDBDR.Item(0))
                        tmp = OleDBDR.Item(0).ToString
                    End If

                End While

                'For i = 0 To YYYYMMDD.Count - 1
                '    Console.WriteLine("+ " & YYYYMMDD(i))
                'Next                
            Else
                'btnGetReport.Text = "查詢紀錄 ( 0 筆 )"
            End If
            OleDBDR.Close()
            OleDBC.Dispose()

            '=======================================================

            Dim man As Integer = 0
            Dim all_test As Integer = 0

            For i = 0 To YYYYMMDD.Count - 1
                'Console.WriteLine("+ " & YYYYMMDD(i))

                OleDBC = frmMain.myHrv.HRV_Get_tData_ALL_TEST(YYYYMMDD(i).ToString)
                OleDBDR = OleDBC.ExecuteReader

                If OleDBDR.HasRows Then
                    While OleDBDR.Read

                        all_test = OleDBDR.Item(0)
                        'Console.WriteLine(YYYYMMDD(i).ToString & " 檢測總次數: " & all_test)

                    End While
                Else
                    'btnGetReport.Text = "查詢紀錄 ( 0 筆 )"
                End If
                OleDBDR.Close()
                OleDBC.Dispose()
                '========================================================================

                OleDBC = frmMain.myHrv.HRV_Get_tData_UNIT_TEST(YYYYMMDD(i).ToString)
                OleDBDR = OleDBC.ExecuteReader

                man = 0
                If OleDBDR.HasRows Then
                    While OleDBDR.Read

                        man = man + 1

                    End While

                    'Console.WriteLine(YYYYMMDD(i).ToString & " 受測人數: " & man)
                Else
                    'btnGetReport.Text = "查詢紀錄 ( 0 筆 )"
                End If

                OleDBDR.Close()
                OleDBC.Dispose()

                Console.WriteLine(YYYYMMDD(i).ToString & " 檢測總次數: " & all_test & " 受測人數: " & man)

                dgvData.Rows.Add()
                dgvData.Item(0, i).Value = YYYYMMDD(i).ToString
                dgvData.Item(1, i).Value = all_test
                dgvData.Item(2, i).Value = man

            Next



        Catch ex As Exception
            MsgBox("LoadUserRecord " & ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub
End Class
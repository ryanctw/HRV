Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports System.Text

Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmDataList

    Public HTML2PDF As String = "gtool.exe"
    Dim gbTitle As String = "檢測紀錄"
    Dim REPORT_ARRAY_id As New ArrayList()

    'Dim jpg_width As Integer = 374
    'Dim jpg_height As Integer = 334
    Dim jpg_width As Integer = 400
    Dim jpg_height As Integer = 350


    Private Sub frmDataList_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = frmMain.HRVapp
        GroupBox1.Text = gbTitle & " : " & frmMain.current_name & " "

        LoadDataRecord()

    End Sub

    Public Sub LoadDataRecord()

        Dim OleDBC As New OleDbCommand
        Dim OleDBDR As OleDbDataReader
        Dim c As Integer
        c = 0

        OleDBC = frmMain.myHrv.HRV_Get_tData(frmMain.current_user_id)

        OleDBDR = OleDBC.ExecuteReader
        dgvData.Rows.Clear()

        If OleDBDR.HasRows Then
            While OleDBDR.Read
                dgvData.Rows.Add()

                'data_id
                dgvData.Rows(c).Cells(0).Value = False
                'Console.WriteLine("data_id " & OleDBDR.Item(0))

                '_id
                dgvData.Item(1, c).Value = OleDBDR.Item(0)

                'test_dt
                dgvData.Item(2, c).Value = OleDBDR.Item(2).ToString

                c = c + 1
            End While


            'Console.WriteLine("c " & c)

            If c = 0 Then
                btnDel.Enabled = False
                btnExport.Enabled = False
            Else
                btnDel.Enabled = True
                btnExport.Enabled = True
            End If

            If c > 1 Then
                btnMultiReport.Enabled = True
            Else
                btnMultiReport.Enabled = False
            End If

            'OleDBDR.Close()
            'OleDBC.Dispose()
            frmMain.gbTop.Enabled = False

        Else
            btnReport.Enabled = False
            btnMultiReport.Enabled = False

            MsgBox("查無 " & frmMain.current_name & " 檢測資料", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            frmMain.Enabled = True
            Me.Dispose()
            'Returns
        End If

    End Sub

    Public Sub LoadDataRecord_By_Date(ByVal from_date As String, ByVal to_date As String)

        Dim OleDBC As New OleDbCommand
        Dim OleDBDR As OleDbDataReader
        Dim c As Integer
        c = 0

        Try

            OleDBC = frmMain.myHrv.HRV_Get_tData_By_Date(frmMain.current_user_id, from_date, to_date)

            OleDBDR = OleDBC.ExecuteReader
            dgvData.Rows.Clear()

            If OleDBDR.HasRows Then
                While OleDBDR.Read
                    dgvData.Rows.Add()

                    'data_id
                    dgvData.Rows(c).Cells(0).Value = True
                    'Console.WriteLine("data_id " & OleDBDR.Item(0))

                    '_id
                    dgvData.Item(1, c).Value = OleDBDR.Item(0)

                    'test_dt
                    dgvData.Item(2, c).Value = OleDBDR.Item(2).ToString

                    c = c + 1
                End While


                'Console.WriteLine("c " & c)

                If c = 0 Then
                    btnDel.Enabled = False
                    btnExport.Enabled = False
                Else
                    btnDel.Enabled = True
                    btnExport.Enabled = True
                End If

                If c > 1 Then
                    btnMultiReport.Enabled = True
                Else
                    btnMultiReport.Enabled = False
                End If

                'OleDBDR.Close()
                'OleDBC.Dispose()
                frmMain.gbTop.Enabled = False

            Else
                btnReport.Enabled = False
                btnMultiReport.Enabled = False

                MsgBox("查無 " & frmMain.current_name & " 檢測資料", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                frmMain.gbTop.Enabled = True
                frmMain.Enabled = True
                Me.Dispose()
                'Returns
            End If

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        'btnReport.Enabled = True
        'btnMultiReport.Enabled = True
        'btnExit.Enabled = True
        frmMain.gbTop.Enabled = True

        'frmMain.Enabled = True
        Me.Dispose()

    End Sub

    'Private Sub dgvData_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellValueChanged
    Private Sub dgvData_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvData.CellMouseClick

        'Console.WriteLine("R " & e.RowIndex)
        'Console.WriteLine("C " & e.ColumnIndex)

        ''Dim destination_abroad As Boolean = dgvData.CurrentCell.EditedFormattedValue
        'Dim destination_abroad As Boolean = dgvData.CurrentCell.EditedFormattedValue

        'If destination_abroad = True Then
        '    Console.WriteLine("Un-Check")

        'Else
        '    Console.WriteLine("Check")
        'End If

    End Sub


    'Private Sub dgvData_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvData.CellMouseClick
    '    'Console.WriteLine(dgvData.CurrentCell.OwningColumn.Name)
    '    If dgvData.CurrentCell. Then
    '        Console.WriteLine("T")
    '    Else
    '        Console.WriteLine("F")
    '    End If
    '    If String.Compare(dgvData.CurrentCell.OwningColumn.Name, "Colume1") Then
    '        Dim checkBoxStatus As Boolean = Convert.ToBoolean(dgvData.CurrentCell.EditedFormattedValue)
    '        'checkBoxStatus gives you whether checkbox cell value of selected row for the "Colume1 " column value is checked or not.
    '        If (checkBoxStatus) Then
    '            'write your code
    '            Console.WriteLine("1")
    '        Else
    '            'write your code
    '            Console.WriteLine("2")
    '        End If
    '    End If

    'End Sub

    Public Sub Make_Report()

        frmProgress.ProgressBar.Value = 20
        Report_Page_1()

        frmProgress.ProgressBar.Value = 40
        Report_Page_2()

        frmProgress.ProgressBar.Value = 60
        Report_Page_3()

        frmProgress.ProgressBar.Value = 80
        Report_Page_4()

    End Sub

    Private Sub Report_Page_1()

        'Dim ID As String = "B111222333"
        'Dim TESTER As String = "好厲害"
        'Dim USER As String = "林持人"
        'Dim SEX As String = "男"
        'Dim BIRTH As String = "1977-01-01"
        'Dim TEST_TIME As String = "10:10:10"
        'Dim TEST_DATE As String = "2015-01-01"
        'Dim AGE As String = "40"

        'Response.Write("" + Environment.NewLine)

        Dim FILE_PATH As String = Application.StartupPath & "\data\report_page_1.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)


        'http://www.webtech.tw/info.php?tid=CSS_DIV_%E4%B8%A6%E6%8E%92%E8%AA%9E%E6%B3%95
        '#DIV1{
        'width:200px;　<!--DIV區塊寬度-->
        'line-height:50px;　//DIV區塊高度
        'padding:20px;　//DIV區塊內距，參閱：CSS padding 內距。
        'border:2px blue solid;　//DIV區塊邊框，參閱：CSS border 邊框設計。
        'margin-right:10px;　//靠右外距，參閱：CSS margin 邊界使用介紹範例教學。
        'float:left;
        '}

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write(".table-c table{padding:10px; border:0px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 100%; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)

        Response.Write("table.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("tr.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("td.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th, td {padding: 3px;}" + Environment.NewLine)

        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;margin-top:0px; float:left;font-size: 15px;}" + Environment.NewLine)
        Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;margin-left:20px;margin-top:0px;}" + Environment.NewLine)
        'Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;float:left;font-size: 20px;}" + Environment.NewLine)
        'Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;}" + Environment.NewLine)
        Response.Write("#HR{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)

        Response.Write("#wrap {width:auto;margin:0 auto;overflow:hidden;font-size:130%;font-weight: bold;} " + Environment.NewLine)
        Response.Write("#sideColumn {background-color:none;width:70%;float:left;}" + Environment.NewLine)
        Response.Write("#mainColumn {background-color:none;float:left;width:30%;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        'Response.Write("<div align=""left"" class=""table-a"">" + Environment.NewLine)
        'Response.Write("<h2>基本資料 受檢者ID：" & frmMain.current_pid & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;操作者：OOO </h2>" + Environment.NewLine)
        'Response.Write("<h2>基本資料 受檢者ID：" & frmMain.current_pid & "</h2>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        If frmMain.myHrv.UNIT.Length > 0 Then
            Response.Write("<h1 align=""center"">檢測單位 : " & frmMain.myHrv.UNIT & " 自律神經檢測報告</h1>" + Environment.NewLine)
        Else
            Response.Write("<h1 align=""center"">自律神經檢測報告</h1>" + Environment.NewLine)
        End If

        Response.Write("<div id=""wrap"">" + Environment.NewLine)
        Response.Write("<div id=""sideColumn"">基本資料 受檢者ID：" & frmMain.current_pid & "</div> " + Environment.NewLine)
        'Response.Write("<div id=""mainColumn"">操作人員：" & frmMain.current_login_name & "</div> " + Environment.NewLine)
        Response.Write("<div id=""mainColumn"">操作人員：" & frmMain.myHrv.TESTER & "</div> " + Environment.NewLine)
        Response.Write("</div> " + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div style=""font-size: 18px; margin: 0px; padding: 0px; width: auto; overflow: hidden;"">" + Environment.NewLine)
        Response.Write("<center>" + Environment.NewLine)
        Response.Write("<table border:0px solid #ccc style=""width:100%"">" + Environment.NewLine)
        Response.Write("<tbody>" + Environment.NewLine)
        Response.Write("<tr bgcolor=""#78C0D4"">" + Environment.NewLine)
        Response.Write("<td style=""padding:6px;"" >姓名：" & frmMain.current_name & "</td>" + Environment.NewLine)
        Response.Write("<td >性別：" & frmMain.current_sex & "</td>" + Environment.NewLine)
        Response.Write("<td>出生日期：" & frmMain.current_birth & "</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr bgcolor=""#d2eaf1"">" + Environment.NewLine)
        Response.Write("<td style=""padding:6px;"">受測時間：" & frmMain.myHrv.test_time & "</td>" + Environment.NewLine)
        Response.Write("<td>受測日期：" & frmMain.myHrv.test_date & "</td>" + Environment.NewLine)
        Response.Write("<td>年齡：" & frmMain.myHrv.age & " 歲</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</tbody>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</center>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div align=""center"" class=""table-b"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p"">自律神經檢測報告說明</p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td><br>一、自律神經檢測透過心電圖紀錄分析來檢測我們心率頻率的高頻、低頻，以及透過資料庫的分析，來呈現下列報告資料，透過檢測能得知自律神經是否異常或是偏向。</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td><br>二、檢測數據呈現受測者年齡的標準範圍，同時經過分析比較呈現於下列報告的圖形，請檢閱下列報告前，再次確認個人資料無誤。</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td><br>三、各項檢查分析內容為自律神經資料庫檢測結果，生理指數經過精密數值分析，如有異常建議與醫療診所安排諮詢，守護個人健康、遠離疾病。</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div class=""table-a"">" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div align=""center"" class=""table-c"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>檢測數據:</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<table align=""center"" style=""width:100%"" width=""100"">" + Environment.NewLine)
        Response.Write("<tbody>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""180"">" + Environment.NewLine)
        Response.Write("<p><strong>項目</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""161"">" + Environment.NewLine)
        Response.Write("<p><strong>檢測數值&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""50"">" + Environment.NewLine)
        Response.Write("<p><strong>單位</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""190"">" + Environment.NewLine)
        Response.Write("<p><strong>結果說明</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""100"">" + Environment.NewLine)
        Response.Write("<p><strong>標準範圍</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>心跳速率(Heart rate)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.HR & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>次/分鐘</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        If frmMain.myHrv.HR_LEVEL = "正常" Then
            Response.Write("<p>" & frmMain.myHrv.HR_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.HR_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>60 - 100(次/分鐘)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>NN</strong><strong>間距標準差(SDNN)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.SD & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>ms</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""141"">" + Environment.NewLine)

        If frmMain.myHrv.SDNN_LEVEL = "正常" Then
            Response.Write("<p>" & frmMain.myHrv.SDNN_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.SDNN_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.SDNN_LOWER_ & " - " & frmMain.myHrv.SDNN_UPPER_ & "  (ms)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經年齡(ANS Age)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.age_MAX & "-" & frmMain.myHrv.age_MIN & "</p>" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.ANS_AGE & "</p>" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.LF_X_UP & " - " & frmMain.myHrv.LF_X & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>歲</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        'If frmMain.myHrv.ANS_LEVEL = "自律神經年齡相當同年齡者" Then
        Response.Write("<p>" & frmMain.myHrv.ANS_LEVEL & "</p>" + Environment.NewLine)
        'Else
        'Response.Write("<p><font color=""red"">" & frmMain.myHrv.ANS_LEVEL & "</font></p>" + Environment.NewLine)
        'End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.age & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td rowspan=""2"" class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經總功能(ANS)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.ln_LF & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>ms&sup2;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""141"">" + Environment.NewLine)
        Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.LF_LOWER_ & " - " & frmMain.myHrv.LF_UPPER_ & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.ANS_SD & "&nbsp;&nbsp;&nbsp;標準差(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        If frmMain.myHrv.ANS_SD_LEVEL = "正常" Then
            Response.Write("<p>" & frmMain.myHrv.ANS_SD_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.ANS_SD_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>-1.5(&delta;) - 1.5(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        '=========================================================================================================================
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td rowspan=""2"" class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>交感神經功能(SYM)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.SYM & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>%</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""141"">" + Environment.NewLine)
        Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.LF_P_LOWER_ & " - " & frmMain.myHrv.LF_P_UPPER_ & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.SYM_SD & "&nbsp;&nbsp;&nbsp;標準差(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        If frmMain.myHrv.SYM_LEVEL = "正常" Then
            Response.Write("<p>" & frmMain.myHrv.SYM_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.SYM_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>-1.5(&delta;) - 1.5(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        '=========================================================================================================================
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td rowspan=""2"" class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>副交感神經功能(VAG)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.ln_HF & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>ms&sup2;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""141"">" + Environment.NewLine)
        Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.HF_LOWER_ & " - " & frmMain.myHrv.HF_UPPER_ & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.ln_HF_SD & " &nbsp;&nbsp;&nbsp;標準差(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        If frmMain.myHrv.ln_HF_LEVEL = "正常" Then
            Response.Write("<p>" & frmMain.myHrv.ln_HF_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.ln_HF_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>-1.5(&delta;) - 1.5(&delta;)</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        '=========================================================================================================================
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title"">" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經偏向(Balance)</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>壓力指標分析</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.Blance & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""141"">" + Environment.NewLine)

        'If frmMain.myHrv.Blance_LEVEL = "自律神經平衡" Then
        If frmMain.myHrv.Blance_LEVEL = "身心正常" Then '20160614
            Response.Write("<p>" & frmMain.myHrv.Blance_LEVEL & "</p>" + Environment.NewLine)
        Else
            Response.Write("<p><font color=""red"">" & frmMain.myHrv.Blance_LEVEL & "</font></p>" + Environment.NewLine)
        End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_1"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>-0.8 - 0.8</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        '=========================================================================================================================
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""title"">" + Environment.NewLine)
        Response.Write("<p><strong>交感調控(SYM Modulation)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""161"">" + Environment.NewLine)
        Response.Write("<p>" & frmMain.myHrv.SYM_Modulation & "</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""75"">" + Environment.NewLine)
        Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""141"">" + Environment.NewLine)

        'If frmMain.myHrv.SYM_Modulation_LEVEL = "平衡" Then
        Response.Write("<p>" & frmMain.myHrv.SYM_Modulation_LEVEL & "</p>" + Environment.NewLine)
        'Else
        'Response.Write("<p><font color=""red"">" & frmMain.myHrv.SYM_Modulation_LEVEL & "</font></p>" + Environment.NewLine)
        'End If

        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td class=""title_2"" width=""142"">" + Environment.NewLine)
        Response.Write("<p>0.6 - 1.5</p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</tbody>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<div align=""right"">檢測報告第一頁/共四頁</div>" + Environment.NewLine)
        Response.Write("<div align=""right"">檢測報告第1頁</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()

    End Sub

    Private Sub Report_Page_2()

        Dim FILE_PATH As String = Application.StartupPath & "\data\report_page_2.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim FigC(5, 5) As String

        FigC(4, 4) = "FigC_44.raw"
        FigC(4, 3) = "FigC_43.raw"
        FigC(4, 2) = "FigC_42.raw"
        FigC(4, 1) = "FigC_41.raw"
        FigC(4, 0) = "FigC_40.raw"
        FigC(3, 4) = "FigC_34.raw"
        FigC(3, 3) = "FigC_33.raw"
        FigC(3, 2) = "FigC_32.raw"
        FigC(3, 1) = "FigC_31.raw"
        FigC(3, 0) = "FigC_30.raw"
        FigC(2, 4) = "FigC_24.raw"
        FigC(2, 3) = "FigC_23.raw"
        FigC(2, 2) = "FigC_22.raw"
        FigC(2, 1) = "FigC_21.raw"
        FigC(2, 0) = "FigC_20.raw"
        FigC(1, 4) = "FigC_14.raw"
        FigC(1, 3) = "FigC_13.raw"
        FigC(1, 2) = "FigC_12.raw"
        FigC(1, 1) = "FigC_11.raw"
        FigC(1, 0) = "FigC_10.raw"
        FigC(0, 4) = "FigC_04.raw"
        FigC(0, 3) = "FigC_03.raw"
        FigC(0, 2) = "FigC_02.raw"
        FigC(0, 1) = "FigC_01.raw"
        FigC(0, 0) = "FigC_00.raw"

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)


        'http://www.webtech.tw/info.php?tid=CSS_DIV_%E4%B8%A6%E6%8E%92%E8%AA%9E%E6%B3%95
        '#DIV1{
        'width:200px;　<!--DIV區塊寬度-->
        'line-height:50px;　//DIV區塊高度
        'padding:20px;　//DIV區塊內距，參閱：CSS padding 內距。
        'border:2px blue solid;　//DIV區塊邊框，參閱：CSS border 邊框設計。
        'margin-right:10px;　//靠右外距，參閱：CSS margin 邊界使用介紹範例教學。
        'float:left;
        '}

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write(".table-c table{padding:10px; border:0px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 100%; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)

        Response.Write("table.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("tr.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("td.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th, td {padding: 3px;}" + Environment.NewLine)

        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;margin-top:0px; float:left;font-size: 15px;}" + Environment.NewLine)
        Response.Write("#DIV2{width:350px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;margin-left:20px;margin-top:0px;}" + Environment.NewLine)
        'Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;float:left;font-size: 20px;}" + Environment.NewLine)
        'Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;}" + Environment.NewLine)
        Response.Write("#HR{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)

        Response.Write("#left{float:left; width:20%; height:180px; background:auto;margin-top:30px;}" + Environment.NewLine)
        Response.Write("#center{float:right;width:20%;height:180px;background:auto;text-align: left;margin-top:40px; font-size: 150%;}" + Environment.NewLine)
        Response.Write("#right{float:right;width:60%;height:180px; background:auto;text-align: left;margin-top:10px;}" + Environment.NewLine)
        Response.Write("#sideColumn {background-color:none; height:220px; width:40%;float:left; margin-top:80px;border-radius: 25px; border: 2px solid #73AD21; padding: 20px;margin-left:20px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)

        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>A. 心跳分析</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div id=""DIV1"">" + Environment.NewLine)
        'Response.Write("<img src=""heart.png"" alt="""" style=""width:128px;height:128px;"">" + Environment.NewLine)
        'Response.Write("<strong>平均心跳率</strong>	" + Environment.NewLine)
        'Response.Write("<strong><u>" & frmMain.myHrv.HR & "</u>  次/分鐘</strong>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<div id=""DIV2"">" + Environment.NewLine)
        'Response.Write("<p><strong>受測者身體狀況 : </strong></p>" + Environment.NewLine)
        ''Response.Write("<p>1. 平日有大量運動習慣或服用藥物。</p>" + Environment.NewLine)
        ''Response.Write("<p>2. 受測者可能有心臟相關的病情，例如身體老化、病態竇房結綜合症等。</p>" + Environment.NewLine)
        ''Response.Write("<p>3. 身體其他相關醫療狀況導致，請向相關醫生洽詢。</p>" + Environment.NewLine)
        'If frmMain.myHrv.HR < 50 Then
        '    Response.Write("<p>1. 平日有大量運動習慣或服用藥物。</p>" + Environment.NewLine)
        '    Response.Write("<p>2. 受測者可能有心臟相關的病情，例如身體老化、病態竇房結綜合症等。</p>" + Environment.NewLine)
        '    Response.Write("<p>3. 身體其他相關醫療狀況導致，請向相關醫生洽詢。</p>" + Environment.NewLine)

        'ElseIf 50 <= frmMain.myHrv.HR And frmMain.myHrv.HR <= 100 Then
        '    Response.Write("<p>1. 保持規律作息、飲食正常、規律運動、維持身體健康生理機能。</p>" + Environment.NewLine)
        'ElseIf frmMain.myHrv.HR > 100 Then
        '    Response.Write("<p>1. 情緒壓力過大或大量飲酒或飲用咖啡因飲料。</p>" + Environment.NewLine)
        '    Response.Write("<p>2. 受測者可能有與心臟相關的病情，例如高血壓、冠狀動脈疾病等。</p>" + Environment.NewLine)
        '    Response.Write("<p>3.身體其他相關醫療狀況導致，請向相關醫生洽詢。</p>" + Environment.NewLine)
        'End If
        'Response.Write("</div>" + Environment.NewLine)
        '=========================================================================================================================================
        Response.Write("<div id=""left"" ><img src=""heart.png"" alt="""" style=""width:128px;height:128px;""></div> " + Environment.NewLine)
        Response.Write("<div id=""right"">" + Environment.NewLine)
        Response.Write("<p><strong>受測者身體狀況 : </strong></p>" + Environment.NewLine)
        If frmMain.myHrv.HR < 50 Then
            Response.Write("<p>1. 平日有大量運動習慣或服用藥物。</p>" + Environment.NewLine)
            Response.Write("<p>2. 受測者可能有心臟相關的病情，例如身體老化、病態竇房結綜合症等。</p>" + Environment.NewLine)
            Response.Write("<p>3. 身體其他相關醫療狀況導致，請向相關醫生洽詢。</p>" + Environment.NewLine)
        ElseIf 50 <= frmMain.myHrv.HR And frmMain.myHrv.HR <= 100 Then
            Response.Write("<p>1. 保持規律作息、飲食正常、規律運動、維持身體健康生理機能。</p>" + Environment.NewLine)
        ElseIf frmMain.myHrv.HR > 100 Then
            Response.Write("<p>1. 情緒壓力過大或大量飲酒或飲用咖啡因飲料。</p>" + Environment.NewLine)
            Response.Write("<p>2. 受測者可能有與心臟相關的病情，例如高血壓、冠狀動脈疾病等。</p>" + Environment.NewLine)
            Response.Write("<p>3.身體其他相關醫療狀況導致，請向相關醫生洽詢。</p>" + Environment.NewLine)
        End If
        Response.Write("</div> " + Environment.NewLine)
        Response.Write("<div id=""center"" style="""">" + Environment.NewLine)
        Response.Write("<strong>平均心跳率</strong></p>" + Environment.NewLine)
        Response.Write("<strong><u>&nbsp;&nbsp;&nbsp;&nbsp;" & frmMain.myHrv.HR & "&nbsp;&nbsp;&nbsp;</u> 次/分鐘</strong>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div style=""clear:both;""></div>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div style=""clear:both;""></div>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        '=========================================================================================================================================
        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>B. 自律神經活性年齡</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div id=""DIV1"">" + Environment.NewLine)
        Response.Write("<div id=""sidebar2"">" + Environment.NewLine)
        'Response.Write("<img src=""FigB.jpg"" alt=""自律神經年齡FigB.jpg"" style=""width:350px;height:300px;"">" + Environment.NewLine)
        'Response.Write("<img src=" & frmMain.myHrv.LF_X_UP & " - " & frmMain.myHrv.LF_X & ".raw alt=""自律神經年齡FigB.jpg"" style=""width:350px;height:300px;"">" + Environment.NewLine)
        Response.Write("<img src=""../raw/3742.raw"" alt=""自律神經年齡FigB.jpg"" style=""width:350px;height:300px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div id=""DIV2"">	" + Environment.NewLine)
        'Response.Write("<td>自律神經年齡與實際年齡差距 " & Math.Abs((frmMain.myHrv.ANS_AGE - frmMain.myHrv.age)) & " 歲(" & frmMain.myHrv.ANS_AGE & " - " & frmMain.myHrv.age & ")</td>" + Environment.NewLine)
        Response.Write("<p>生理年齡差距:</p>" + Environment.NewLine)
        Response.Write("<p>若與實際年齡比較小，生理功能處於激發狀態或健康狀態維持較佳。</p>" + Environment.NewLine)
        Response.Write("<p>若與實際年齡比較大，可能可能是因過度疲累、失眠、疾病等因素而生理功能低下。</p>" + Environment.NewLine)
        Response.Write("<p>若與實際年齡大15歲以上，可能生理功能異常，需注意慢性病變的可能。</p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div style=""clear:both;""></div><!--這是用來清除上方的浮動效果-->" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        'Response.Write("<p><strong><font color=""red"">C. 自律神經數值與活性分析 交感神經:" & frmMain.myHrv.C_SYM_LEVEL & "(" & frmMain.myHrv.C_SYM & ")" & " 副交感神經:" & frmMain.myHrv.C_VGA_LEVEL & "(" & frmMain.myHrv.C_VGA & ")" & "</font></strong></p>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>C. 自律神經數值與活性分析 交感神經:" & frmMain.myHrv.C_SYM_LEVEL & "(" & frmMain.myHrv.C_SYM & ")" & " 副交感神經:" & frmMain.myHrv.C_VGA_LEVEL & "(" & frmMain.myHrv.C_VGA & ")" & "</font></strong></p>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>C. 自律神經數值與活性分析</font></strong></p>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>C. 自律神經活性分析</font></strong></p>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>C. 自律神經活性分析 交感神經:" & frmMain.myHrv.C_SYM_LEVEL & "(" & frmMain.myHrv.C_SYM & ")" & " 副交感神經:" & frmMain.myHrv.C_VGA_LEVEL & "(" & frmMain.myHrv.C_VGA & ")" & "</font></strong></p>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div id=""DIV1"">" + Environment.NewLine)
        'Response.Write("<img src=""FigC.jpg"" alt=""自律神經活性分析FigC.jpg"" style=""width:350px;height:300px;"">" + Environment.NewLine)
        Response.Write("<img border=""1"" src=../raw/" & FigC(frmMain.myHrv.C_SYM_LEVEL_int, frmMain.myHrv.C_VGA_LEVEL_int) & " alt=""自律神經活性分析FigC.jpg"" style=""width:240px;height:340px;padding:10px;"">" + Environment.NewLine)
        'Response.Write("<div><BR></div>" + Environment.NewLine)
        'Response.Write("<div>自律神經活性分析</div>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div id=""sideColumn"">" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經平衡分析</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>自律神經分析描述</strong></p>" + Environment.NewLine)
        Response.Write("<p>" + frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) + "</p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)


        ''Response.Write("<div id=""DIV1"">" + Environment.NewLine)
        'Response.Write("<img src=""FigC1.jpg"" alt=""自律神經比率FigC1.jpg"" style=""width:410px;height:360px;"">" + Environment.NewLine)
        ''Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div style=""clear:both;""></div><!--這是用來清除上方的浮動效果-->" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<div align=""right"">檢測報告第二頁/共四頁</div>" + Environment.NewLine)
        Response.Write("<div align=""right"">檢測報告第2頁</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()

    End Sub

    Private Sub Report_Page_3()

        Dim FILE_PATH As String = Application.StartupPath & "\data\report_page_3.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)


        'http://www.webtech.tw/info.php?tid=CSS_DIV_%E4%B8%A6%E6%8E%92%E8%AA%9E%E6%B3%95
        '#DIV1{
        'width:200px;　<!--DIV區塊寬度-->
        'line-height:50px;　//DIV區塊高度
        'padding:20px;　//DIV區塊內距，參閱：CSS padding 內距。
        'border:2px blue solid;　//DIV區塊邊框，參閱：CSS border 邊框設計。
        'margin-right:10px;　//靠右外距，參閱：CSS margin 邊界使用介紹範例教學。
        'float:left;
        '}

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write(".table-c table{padding:10px; border:0px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 100%; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)

        Response.Write("table.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("tr.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("td.FFF{border: 1px solid black; border-collapse: collapse;}" + Environment.NewLine)
        Response.Write("th, td {padding: 3px;}" + Environment.NewLine)

        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;margin-top:0px; float:left;font-size: 15px;}" + Environment.NewLine)
        Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;margin-left:20px;margin-top:0px;}" + Environment.NewLine)
        'Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;float:left;font-size: 20px;}" + Environment.NewLine)
        'Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;}" + Environment.NewLine)
        Response.Write("#HR{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)

        Response.Write("#wrap {width:auto;margin:20px;overflow:hidden;font-size:100%;font-weight: bold;text-align:left;}" + Environment.NewLine)
        Response.Write("#sideColumn {background-color:none; height:180px; width:40%;float:left; margin-top:20px;border-radius: 25px; border: 2px solid #73AD21; padding: 20px;margin-left:20px;}" + Environment.NewLine)
        Response.Write("#mainColumn {background-color:none; height:180px; float:left;width:40%; margin-top:10px; padding-left: 50px;border-radius: 25px; border: 2px solid #73AD21; padding: 20px; margin-left:30px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        'Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        'Response.Write("<tr>" + Environment.NewLine)
        ''Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 自律神經分布圖  交感:-5.53 副交感:5.47; 副交感神經過強，受測者可能有長期有過於疲勞、壓力過大或免疫發炎反應，導致交感神經活力不足。</strong></p></td>" + Environment.NewLine)
        ''Response.Write("<td class=""table-c""><p class=""style-p""><strong><font color=""red"">D.[自律神經分布圖]  交感:" & frmMain.myHrv.C_SYM & " 副交感:" & frmMain.myHrv.C_VGA & "; " & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "</font></strong></p></td>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 自律神經分布圖  交感:" & frmMain.myHrv.C_SYM & " 副交感:" & frmMain.myHrv.C_VGA & "; " & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "</strong></p></td>" + Environment.NewLine)

        'Response.Write("</tr>" + Environment.NewLine)
        'Response.Write("</table>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div id=""DIV1"">" + Environment.NewLine)
        'Response.Write("<img src=""FigD.jpg"" alt=""自律神經布圖FigD.jpg"" style=""width:400px;height:400px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<div id=""DIV2"">" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經平衡分析</strong></p>" + Environment.NewLine)
        ''Response.Write("<p>副交感神經過強，受測者可能有長期有過於疲勞、壓力過大或免疫發炎反應，導致交感神經活力不足。</p>" + Environment.NewLine)
        'Response.Write("<p>" + frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) + "</p>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<div style=""clear:both;""></div><!--這是用來清除上方的浮動效果-->" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 自律神經分布圖  交感:" & frmMain.myHrv.C_SYM & " 副交感:" & frmMain.myHrv.C_VGA & "; " & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "</strong></p></td>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 自律神經的平衡分析  交感:" & frmMain.myHrv.C_SYM & " 副交感:" & frmMain.myHrv.C_VGA & "; " & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "</strong></p></td>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 自律神經的平衡分析</strong></p></td>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>D. 心率變異分析數值</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div id=""wrap"">" + Environment.NewLine)

        'Response.Write("<div id=""sideColumn"">" + Environment.NewLine)
        ''Response.Write("<p><strong>自律神經平衡分析</strong></p>" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經分析描述</strong></p>" + Environment.NewLine)
        'Response.Write("<p>" + frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) + "</p>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div id=""mainColumn"">" + Environment.NewLine)
        'Response.Write("<p>心率變異頻譜分析數值</p>" + Environment.NewLine)
        Response.Write("<p>HR : " & frmMain.myHrv.HR & "</p>" + Environment.NewLine)
        Response.Write("<p>SD : " & frmMain.myHrv.SD & "</p>" + Environment.NewLine)
        Response.Write("<p>HF : " & frmMain.myHrv.ln_HF & "</p>" + Environment.NewLine)
        Response.Write("<p>LF : " & frmMain.myHrv.ln_LF & "</p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div id=""mainColumn"">" + Environment.NewLine)
        'Response.Write("<p>心率變異頻譜分析數值</p>" + Environment.NewLine)
        'Response.Write("<p>TP : " & frmMain.myHrv.TP & "</p>" + Environment.NewLine)
        'Response.Write("<p>LF : " & frmMain.myHrv.LF & "</p>" + Environment.NewLine)
        'Response.Write("<p>HF : " & frmMain.myHrv.HF & "</p>" + Environment.NewLine)
        'Response.Write("<p>VL : " & frmMain.myHrv.VL & "</p>" + Environment.NewLine)
        Response.Write("<p>VL : " & frmMain.myHrv.ln_VL & "</p>" + Environment.NewLine)
        Response.Write("<p>LF% : " & frmMain.myHrv.sym_percent & "</p>" + Environment.NewLine)
        Response.Write("<p>HF% : " & frmMain.myHrv.vga_percent & "</p>" + Environment.NewLine)
        Response.Write("<p>LF/HF : " & frmMain.myHrv.Blance & "</p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<div style="clear:both;"></div><!--這是用來清除上方的浮動效果-->" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div>" + Environment.NewLine)
        'Response.Write("<p><strong><font color=""red"">D.[自律神經分布圖]  交感:" & frmMain.myHrv.C_SYM & " 副交感:" & frmMain.myHrv.C_VGA & "; " & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "</font></strong></p>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""FigD.jpg"" alt=""自律神經布圖FigD.jpg"" style=""width:400px;height:400px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Dim FigX As String = ""

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>E. 自律神經偏向 1.99 嚴重偏向交感</strong></p></td>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>E. 自律神經偏向 " & frmMain.myHrv.E_VALUE & "" & frmMain.myHrv.E_VALUE_LEVEL & "</strong></p></td>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>E. 自律神經偏向</strong></p></td>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>E. 壓力指標分析</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)


        'If frmMain.myHrv.E_VALUE > -0.8 And frmMain.myHrv.E_VALUE <= 0 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd0.raw"
        'ElseIf frmMain.myHrv.E_VALUE > -1.5 And frmMain.myHrv.E_VALUE <= 0.8 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-08.raw"
        'ElseIf frmMain.myHrv.E_VALUE < -1.5 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-18.raw"
        'End If

        'MsgBox(frmMain.myHrv.E_VALUE)

        If (frmMain.myHrv.E_VALUE > 0) Then
            If frmMain.myHrv.E_VALUE >= 0 And frmMain.myHrv.E_VALUE < 0.8 Then
                FigX = frmMain.app_FilePath & "\raw\figd0.raw"
            ElseIf frmMain.myHrv.E_VALUE < 1 And frmMain.myHrv.E_VALUE >= 0.8 Then
                FigX = frmMain.app_FilePath & "\raw\figd08.raw"
            ElseIf frmMain.myHrv.E_VALUE < 1.2 And frmMain.myHrv.E_VALUE >= 1 Then
                FigX = frmMain.app_FilePath & "\raw\figd1.raw"
            ElseIf frmMain.myHrv.E_VALUE < 1.5 And frmMain.myHrv.E_VALUE >= 1.2 Then
                FigX = frmMain.app_FilePath & "\raw\figd12.raw"
            ElseIf frmMain.myHrv.E_VALUE <= 1.5 Then
                FigX = frmMain.app_FilePath & "\raw\figd15.raw"
            ElseIf frmMain.myHrv.E_VALUE > 1.5 Then
                FigX = frmMain.app_FilePath & "\raw\figd18.raw"
            End If
        Else
            If frmMain.myHrv.E_VALUE <= 0 And frmMain.myHrv.E_VALUE > -0.8 Then
                FigX = frmMain.app_FilePath & "\raw\figd0.raw"
            ElseIf frmMain.myHrv.E_VALUE > -1 And frmMain.myHrv.E_VALUE <= -0.8 Then
                FigX = frmMain.app_FilePath & "\raw\figd-08.raw"
            ElseIf frmMain.myHrv.E_VALUE > -1.2 And frmMain.myHrv.E_VALUE <= -1 Then
                FigX = frmMain.app_FilePath & "\raw\figd-1.raw"
            ElseIf frmMain.myHrv.E_VALUE > -1.5 And frmMain.myHrv.E_VALUE <= -1.2 Then
                FigX = frmMain.app_FilePath & "\raw\figd-12.raw"
            ElseIf frmMain.myHrv.E_VALUE >= -1.5 Then
                FigX = frmMain.app_FilePath & "\raw\figd-15.raw"
            ElseIf frmMain.myHrv.E_VALUE < -1.5 Then
                FigX = frmMain.app_FilePath & "\raw\figd-18.raw"
            End If
        End If
        

        

        'If frmMain.myHrv.E_VALUE >= 2.0 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd18.raw"
        'ElseIf 2.0 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 1.5 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd18.raw"
        'ElseIf 1.5 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 1.2 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd12.raw"
        'ElseIf 1.2 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 1 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd1.raw"
        'ElseIf 1 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 0.9 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd09.raw"
        '    'ElseIf 0.9 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 0.8 Then
        '    'FigX = frmMain.app_FilePath & "\raw\figd09.raw"
        'ElseIf 0.9 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE >= 0.8 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd08.raw"

        '    'ElseIf 0.8 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > 0 Then
        '    '   FigX = frmMain.app_FilePath & "\raw\figd0.raw"
        'ElseIf 0.8 > frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -0.8 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd0.raw"

        'ElseIf -0.8 >= frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -0.9 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-08.raw"
        'ElseIf -0.9 >= frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -1 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-09.raw"
        'ElseIf -1 >= frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -1.2 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-1.raw"
        'ElseIf -1.2 >= frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -1.5 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-12.raw"
        'ElseIf -1.5 >= frmMain.myHrv.E_VALUE And frmMain.myHrv.E_VALUE > -2.0 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-15.raw"
        'ElseIf frmMain.myHrv.E_VALUE <= -2.0 Then
        '    FigX = frmMain.app_FilePath & "\raw\figd-18.raw"
        'End If

        Response.Write("<img src=" & FigX & " alt=""自律神經偏向FigE.jpg"" style=""width:800px;height:450px;"">" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)


        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<div align=""right"">檢測報告第三頁/共四頁</div>" + Environment.NewLine)
        Response.Write("<div align=""right"">檢測報告第3頁</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()

    End Sub

    Private Sub Report_Page_4()


        Dim FILE_PATH As String = Application.StartupPath & "\data\report_page_4.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)


        'http://www.webtech.tw/info.php?tid=CSS_DIV_%E4%B8%A6%E6%8E%92%E8%AA%9E%E6%B3%95
        '#DIV1{
        'width:200px;　<!--DIV區塊寬度-->
        'line-height:50px;　//DIV區塊高度
        'padding:20px;　//DIV區塊內距，參閱：CSS padding 內距。
        'border:2px blue solid;　//DIV區塊邊框，參閱：CSS border 邊框設計。
        'margin-right:10px;　//靠右外距，參閱：CSS margin 邊界使用介紹範例教學。
        'float:left;
        '}

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write(".table-c table{padding:10px; border:0px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 100%; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)

        Response.Write("table.FFF{border: 1px solid black; border-collapse: collapse;padding: 1px;}" + Environment.NewLine)
        Response.Write("th.FFF{border: 1px solid black; border-collapse: collapse;padding: 1px;}" + Environment.NewLine)
        Response.Write("tr.FFF{border: 1px solid black; border-collapse: collapse;padding: 1px;}" + Environment.NewLine)
        Response.Write("td.FFF{border: 1px solid black; border-collapse: collapse;padding: 1px;}" + Environment.NewLine)
        Response.Write("th, td {padding: 1px;}" + Environment.NewLine)

        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;margin-top:0px; float:left;font-size: 15px;}" + Environment.NewLine)
        Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;margin-left:20px;margin-top:0px;}" + Environment.NewLine)
        'Response.Write("#DIV1{width:300px;line-height:10px;padding:30px;text-align: center;border:0px blue solid;margin-right:10px;float:left;font-size: 20px;}" + Environment.NewLine)
        'Response.Write("#DIV2{width:300px;line-height:30px;text-align: left;padding:10px;border02px blue solid;float:left;}" + Environment.NewLine)
        Response.Write("#HR{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)

        Response.Write("#wrap {width:auto;margin:0 auto;overflow:hidden;font-size:110%;} " + Environment.NewLine)
        Response.Write("#sideColumn {background-color:none;width:70%;float:left;}" + Environment.NewLine)
        Response.Write("#mainColumn {background-color:none;float:right;width:30%;}" + Environment.NewLine)
        Response.Write("#FigF{padding-left: 20px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Dim FigX As String = ""


        'Dim F_LEVEL As String = "SDNN " & frmMain.myHrv.F_SDNN & ", LF% " & frmMain.myHrv.SYM_SD & ", Ln(HF) " & frmMain.myHrv.ln_HF_SD & ": "
        Dim F_LEVEL As String = "(SDNN:" & frmMain.myHrv.SD & ")(SDNN區間上限:" & frmMain.myHrv.SDNN_UPPER_ & ")(SDNN區間下限:" & frmMain.myHrv.SDNN_LOWER_ & ")(交感神經標準差:" & frmMain.myHrv.SYM_SD & ")(副交感神經標準差:" & frmMain.myHrv.ln_HF_SD & ")"
        Dim F_MSG As String = ""

        If frmMain.myHrv.F_1 = 1 Then
            'F_LEVEL = F_LEVEL + "發病 "
            'F_MSG = "本受檢者自律神經功能處於發病期，自律神經功能已不正常，宜儘速就醫檢查。"
            'FigX = frmMain.app_FilePath & "\raw\f4.raw"
            F_LEVEL = F_LEVEL + "潛伏 "
            F_MSG = "本受測者自律神經功能處於潛伏期，報告檢測已經出現輕度身心壓力或是生理機能異常，建議加強自我健康管理，進行均衡飲食與調整正常作息，以維持身心健康平衡。"
            FigX = frmMain.app_FilePath & "\raw\f3.raw"
        ElseIf frmMain.myHrv.F_3 = 1 Then
            'F_LEVEL = F_LEVEL + "潛伏 "
            'F_MSG = "本受檢者自律神經功能處於潛伏期，生理年齡已提前老化，應注意保養與控制任何慢性疾病。"
            'F_MSG = "本受測者自律神經處於潛伏期，生理機能已提前老化，應注意身體保養與均衡作息飲食，做好身體健康管理，預防疾病發生。"
            F_MSG = "本受測者自律神經功能處於潛伏期，報告檢測已經出現輕度身心壓力或是生理機能異常，建議加強自我健康管理，進行均衡飲食與調整正常作息，以維持身心健康平衡。"
            FigX = frmMain.app_FilePath & "\raw\f3.raw"
        ElseIf frmMain.myHrv.F_2 = 1 Then
            'F_LEVEL = F_LEVEL + "警告 "
            F_MSG = "本受檢者自律神經功能處於警告期，應儘速控制病情，防止進入發病期。"
            FigX = frmMain.app_FilePath & "\raw\f2.raw"
        ElseIf frmMain.myHrv.F_4 = 1 Then
            'F_LEVEL = F_LEVEL + "正常 "
            F_MSG = "本受檢者自律神經功能處於正常期；若合併有其它的疾病或出現不正常的檢驗數值，仍需由醫師做整體評估。"
            FigX = frmMain.app_FilePath & "\raw\f4.raw"
        End If

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        'Response.Write("<td class=""table-c""><p class=""style-p""><strong>F: 整體評估 " & F_LEVEL & " </strong></p></td>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>F: 整體評估</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<div id=""FigF""align=""left"">" + Environment.NewLine)
        'Response.Write("<img src=" & FigX & " alt=""整體評估FigF.jpg"" style=""width:680px;height:425px;"">" + Environment.NewLine)
        Response.Write("<img src=" & FigX & " alt=""整體評估FigF.jpg"" style=""width:680px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div align=""center"" id=""HR"">" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<table style=""width:100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        Response.Write("<tr>" + Environment.NewLine)
        Response.Write("<td class=""table-c""><p class=""style-p""><strong>檢查說明</strong></p></td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div id=""FigF"" align=""left"">" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<strong>一、自律神經檢測提示:</strong>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<table class=""FFF"" style=""width:70%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" + Environment.NewLine)
        'Response.Write("<tbody>" + Environment.NewLine)
        'Response.Write("<tr class=""FFF"">" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""120"">" + Environment.NewLine)
        'Response.Write("<p><strong>項目</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""80"">" + Environment.NewLine)
        'Response.Write("<p><strong>結果說明&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""80"">" + Environment.NewLine)
        'Response.Write("<p><strong>檢測數值</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""FFF"">" + Environment.NewLine)
        'Response.Write("<td class=""FFF"">" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經總功能(ANS)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""161"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.ANS_LEVEL & "</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""75"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.ln_LF & "</p>" + Environment.NewLine)
        ''Response.Write("<p>" & frmMain.myHrv.ANS_SD & "</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""FFF"">" + Environment.NewLine)
        'Response.Write("<td class=""FFF"">" + Environment.NewLine)
        'Response.Write("<p><strong>交感神經(SYM)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""161"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.SYM_LEVEL & "</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""75"">" + Environment.NewLine)
        ''Response.Write("<p>" & frmMain.myHrv.SYM_SD & "</p>" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.SYM & "</p>" + Environment.NewLine) '20160312 客戶提示修改
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""FFF"">" + Environment.NewLine)
        'Response.Write("<td class=""FFF"">" + Environment.NewLine)
        'Response.Write("<p><strong>副交感神經(VAG)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""161"">" + Environment.NewLine)
        ''Response.Write("<p>" & frmMain.myHrv.C_VGA_LEVEL & "</p>" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.ln_HF_LEVEL & "</p>" + Environment.NewLine) '20160524 將平衡改為正常
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""75"">" + Environment.NewLine)
        ''Response.Write("<p>" & frmMain.myHrv.ln_HF_SD & "</p>" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.ln_HF & "</p>" + Environment.NewLine) '20160312 客戶提示修改
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""FFF"">" + Environment.NewLine)
        'Response.Write("<td class=""FFF"">" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經偏向</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""161"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.Blance_LEVEL & "</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td class=""FFF"" width=""75"">" + Environment.NewLine)
        'Response.Write("<p>" & frmMain.myHrv.Blance & "</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)
        'Response.Write("</tbody>" + Environment.NewLine)
        'Response.Write("</table>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div id=""FigF""  align=""left"">" + Environment.NewLine)
        'Response.Write("<strong>二、建議:</strong>" + Environment.NewLine)
        Response.Write("<strong>一、自律神經檢測提示:</strong>" + Environment.NewLine)
        Response.Write("</dib>" + Environment.NewLine)

        'Response.Write("<div style=""width:800px;"" >" + Environment.NewLine)
        Response.Write("<div>" + Environment.NewLine)
        'Response.Write("<p>自律神經檢測出你的自律神經生理年齡 " & frmMain.myHrv.ANS_LEVEL & " ，自律神經分析" & frmMain.myHrv.E_VALUE_LEVEL & "，" & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "<br><br>" & F_MSG & "<br><br>" & "這些診斷建議來自心律變異檢測，請關注自身身心狀況，並諮詢相關醫療機構。</p>" + Environment.NewLine)
        'Response.Write("<p>自律神經檢測出你的自律神經生理年齡 " & frmMain.myHrv.ANS_LEVEL & " ，自律神經分析" & frmMain.myHrv.E_VALUE_LEVEL & "，" & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "<br><br>" & F_MSG & "<br><br>" & "這些診斷建議來自心率變異檢測，請關注自身身心狀況。</p>" + Environment.NewLine)
        'Response.Write("<p>自律神經檢測出你的自律神經生理年齡 " & frmMain.myHrv.ANS_LEVEL & " ，身心壓力指標分析" & frmMain.myHrv.E_VALUE_LEVEL & "，" & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "<br><br>" & F_MSG & "<br><br>" & "這些診斷建議來自心率變異檢測，請關注自身身心狀況。</p>" + Environment.NewLine)
        Response.Write("<p>自律神經檢測出你的自律神經生理年齡 " & frmMain.myHrv.ANS_LEVEL & " ，身心壓力指標分析" & frmMain.myHrv.Blance_LEVEL & "，" & frmMain.myHrv.D_MSG(frmMain.myHrv.C_VGA_VALUE, frmMain.myHrv.C_SYM_VALUE) & "<br><br>" & F_MSG & "<br><br>" & "這些診斷建議來自心率變異檢測，請關注自身身心狀況。</p>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""left"">" + Environment.NewLine)
        'Response.Write("<br><strong>三、醫療診所診斷說明</strong>" + Environment.NewLine)
        Response.Write("<br><strong>二、醫療診所診斷說明</strong>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<div id=""wrap"">" + Environment.NewLine)
        Response.Write("<div id=""sideColumn"">" + frmMain.VERSION + "</div>" + Environment.NewLine)
        'Response.Write("<div id=""mainColumn"" align=""right"">檢測報告第四頁/共四頁</div> " + Environment.NewLine)
        Response.Write("<div id=""mainColumn"" align=""right"">檢測報告第4頁</div> " + Environment.NewLine)
        Response.Write("</div> " + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()

    End Sub

    Public Sub Compare_2_Report_NO_USE()

        Dim FILE_PATH As String = Application.StartupPath & "\data\compare.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 300px; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)
        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/ }" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("th, td {padding: 1px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)


        Response.Write("<div class=""table-a"">" + Environment.NewLine)

        'Response.Write("<p><strong>多次檢測紀錄：</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 受測者ID：" & frmMain.current_pid & "</p>" + Environment.NewLine)
        Response.Write("<h2 align=""center"">多次檢測紀錄  受測者ID：" & frmMain.current_pid & "</h2>" + Environment.NewLine)

        Response.Write("<table align=""center"" style=""width:100%"" width=""100"">" + Environment.NewLine)
        Response.Write("<tbody>" + Environment.NewLine)
        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>&nbsp;</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cTIME(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cTIME(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>心跳速率(Heart rate)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cHR(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cHR(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>NN</strong><strong>間距標準差(SDNN)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經年齡(ANS Age)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr  class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經總功能(ANS)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANS(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANS(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感神經功能(SYM)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYM(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYM(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>副交感神經功能(VAG)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cVGA(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cVGA(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經偏向(Balance)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感調控(SYM Modulation)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        'Response.Write("<td width=""215"">" + Environment.NewLine)
        'Response.Write("<p><strong>正常RR間距變化(RRIV)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p><strong>" & frmMain.myHrv. & "</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p><strong>" & frmMain.myHrv. & "</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        Response.Write("</tbody>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div>" + Environment.NewLine)
        Response.Write("<p align=""center""><strong>多次比較圖：</strong></p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)


        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""SYMMOD.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SYMMOD.jpg"" alt=""SYMMOD.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        'Response.Write("<img src=""xxx.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""blank.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()


    End Sub

    Public Sub Compare_3_Report_NO_USE()

        Dim FILE_PATH As String = Application.StartupPath & "\data\compare.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 300px; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)
        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/ }" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("th, td {padding: 1px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)
        Response.Write("<br>" + Environment.NewLine)


        Response.Write("<div class=""table-a"">" + Environment.NewLine)

        'Response.Write("<p><strong>多次檢測紀錄：</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 受測者ID：" & frmMain.current_pid & "</p>" + Environment.NewLine)
        Response.Write("<h2 align=""center"">多次檢測紀錄  受測者ID：" & frmMain.current_pid & "</h2>" + Environment.NewLine)

        Response.Write("<table align=""center"" style=""width:100%"" width=""100"">" + Environment.NewLine)
        Response.Write("<tbody>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>&nbsp;</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cTIME(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cTIME(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cTIME(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>心跳速率(Heart rate)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cHR(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cHR(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cHR(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>NN</strong><strong>間距標準差(SDNN)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經年齡(ANS Age)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr  class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經總功能(ANS)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANS(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANS(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cANS(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感神經功能(SYM)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYM(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYM(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYM(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>副交感神經功能(VAG)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cVGA(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cVGA(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cVGA(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經偏向(Balance)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感調控(SYM Modulation)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(0) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(1) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("<td width=""130"">" + Environment.NewLine)
        Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(2) & "</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        Response.Write("</tr>" + Environment.NewLine)

        'Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        'Response.Write("<td width=""215"">" + Environment.NewLine)
        'Response.Write("<p><strong>正常RR間距變化(RRIV)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        Response.Write("</tbody>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)

        Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div>" + Environment.NewLine)
        Response.Write("<p align=""center""><strong>多次比較圖：</strong></p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)


        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)
        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""SYMMOD.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SYMMOD.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        'Response.Write("<img src=""xxx.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("<img src=""blank.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()


    End Sub

    

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub Multi_Report_Page_1()

        Dim FILE_PATH As String = Application.StartupPath & "\data\compare_p1.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If


        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4;}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write("#cssTable td {text-align:center; vertical-align:middle;}" + Environment.NewLine)
        Response.Write(".style-p {width: 300px; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)
        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/ }" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("#go_left {margin-left:4%;}" + Environment.NewLine)

        Response.Write("th, td {padding: 1px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div class=""table-a"">" + Environment.NewLine)

        'Response.Write("<p><strong>多次檢測紀錄：</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 受測者ID：" & frmMain.current_pid & "</p>" + Environment.NewLine)
        Response.Write("<h2 align=""center"">多次檢測紀錄  受測者ID：" & frmMain.current_pid & "</h2>" + Environment.NewLine)

        Response.Write("<table align=""center"" style=""width:100%"" width=""100"" id=""cssTable"">" + Environment.NewLine)
        Response.Write("<tbody>" + Environment.NewLine)

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>&nbsp;</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            Response.Write("<p><strong>" & (i + 1) & "</strong></p>" + Environment.NewLine)
            Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================
        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>&nbsp;</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(i) & "</strong></p>" + Environment.NewLine)
            Response.Write("<p><strong>" & frmMain.myHrv.cTIME(i) & "</strong></p>" + Environment.NewLine)
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cTIME(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cDATEE(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cTIME(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>心跳速率(Heart Rate)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cHR_LEVEL(i) = "正常" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cHR(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cHR(i) & "</font></strong></p>" + Environment.NewLine)
            End If

            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cHR(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cHR(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next

        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>NN</strong><strong>間距標準差(SDNN)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cSDNN_LEVEL(i) = "正常" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cSDNN(i) & "</font></strong></p>" + Environment.NewLine)
            End If
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSDNN(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next

        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經年齡(ANS Age)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)

        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cANSAGE_LEVEL(i) = "自律神經年齡相當同年齡者" Then
                'Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(i) & "</strong></p>" + Environment.NewLine)
                Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE_LF_X_UP(i) & " - " & frmMain.myHrv.cANSAGE_LF_X(i) & "</strong></p>" + Environment.NewLine)
            Else
                'Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cANSAGE(i) & "</font></strong></p>" + Environment.NewLine)
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cANSAGE_LF_X_UP(i) & " - " & frmMain.myHrv.cANSAGE_LF_X(i) & "</font></strong></p>" + Environment.NewLine)
            End If
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cANSAGE(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next

        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr  class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>自律神經總功能(ANS)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cANS_LEVEL(i) = "正常" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cANS(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cANS(i) & "</font></strong></p>" + Environment.NewLine)
            End If

            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cANS(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cANS(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感神經功能(SYM)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cSYM_LEVEL(i) = "正常" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cSYM(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cSYM(i) & "</font></strong></p>" + Environment.NewLine)
            End If

            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSYM(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSYM(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>副交感神經功能(VAG)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cVGA_LEVEL(i) = "正常" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cVGA(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cVGA(i) & "</font></strong></p>" + Environment.NewLine)
            End If
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cVGA(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cVGA(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        'Response.Write("<p><strong>自律神經偏向(Balance)</strong></p>" + Environment.NewLine)
        Response.Write("<p><strong>壓力指標分析</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)

        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cBALANCE_LEVEL(i) = "自律神經平衡" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cBALANCE(i) & "</font></strong></p>" + Environment.NewLine)
            End If
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cBALANCE(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next

        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================

        Response.Write("<tr class=""title"">" + Environment.NewLine)
        Response.Write("<td width=""215"">" + Environment.NewLine)
        Response.Write("<p><strong>交感調控(SYM Modulation)</strong></p>" + Environment.NewLine)
        Response.Write("</td>" + Environment.NewLine)
        For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
            Response.Write("<td width=""130"">" + Environment.NewLine)
            If frmMain.myHrv.cSYMMOD_LEVEL(i) = "平衡" Then
                Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(i) & "</strong></p>" + Environment.NewLine)
            Else
                Response.Write("<p><strong><font color=""red"">" & frmMain.myHrv.cSYMMOD(i) & "</font></strong></p>" + Environment.NewLine)
            End If
            Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(1) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
            'Response.Write("<td width=""130"">" + Environment.NewLine)
            'Response.Write("<p><strong>" & frmMain.myHrv.cSYMMOD(2) & "</strong></p>" + Environment.NewLine)
            'Response.Write("</td>" + Environment.NewLine)
        Next
        Response.Write("</tr>" + Environment.NewLine)
        '==================================================================================================
        'Response.Write("<tr class=""title_1"">" + Environment.NewLine)
        'Response.Write("<td width=""215"">" + Environment.NewLine)
        'Response.Write("<p><strong>正常RR間距變化(RRIV)</strong></p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("<td width=""130"">" + Environment.NewLine)
        'Response.Write("<p>&nbsp;</p>" + Environment.NewLine)
        'Response.Write("</td>" + Environment.NewLine)
        'Response.Write("</tr>" + Environment.NewLine)

        Response.Write("</tbody>" + Environment.NewLine)
        Response.Write("</table>" + Environment.NewLine)

        Response.Write("<div>" + Environment.NewLine)
        Response.Write("<p align=""center""><strong>多次比較圖：</strong></p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""Balance.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<div id=""go_left"">" + Environment.NewLine)
        'Response.Write("<img src=""Balance.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""SYMM.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)

        'Response.Write("<div>" + Environment.NewLine)
        'Response.Write("<p align=""center""><strong>多次比較圖：</strong></p>" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""Balance.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        ''Response.Write("<img src=""SYMMOD.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:400px;height:350px;"">" + Environment.NewLine)
        ''Response.Write("<img src=""xxx.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        ''Response.Write("<img src=""SYMM.jpg"" alt="""" style=""width:400px;height:350px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SYMM.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("</div>" + Environment.NewLine)

        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()


    End Sub

    Public Sub Multi_Report_Page_2()

        Dim FILE_PATH As String = Application.StartupPath & "\data\compare_p2.html"

        If System.IO.File.Exists(FILE_PATH) Then
            System.IO.File.Delete(FILE_PATH)
        End If

        Dim Response As New System.IO.StreamWriter(FILE_PATH, False)

        Response.Write("<!DOCTYPE html>" + Environment.NewLine)

        Response.Write("<html>" + Environment.NewLine)
        Response.Write("<head>" + Environment.NewLine)

        Response.Write("<title>" + Environment.NewLine)
        Response.Write("HRV Report" + Environment.NewLine)
        Response.Write("</title>" + Environment.NewLine)

        Response.Write("<style>" + Environment.NewLine)

        Response.Write(".table-b table{padding:10px; border:3px solid #78C0D4}" + Environment.NewLine)
        Response.Write("td.table-c {width: 10%;	background-color:#92d050; border-radius:10px; box-shadow: 5px 5px 5px #888888;}" + Environment.NewLine)
        Response.Write(".style-p {width: 300px; margin: 10px 200px 10px 10px;}" + Environment.NewLine)
        Response.Write("tr.title{width:10%; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("tr.title_1{width:10%; background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title{width:180px; background-color:#78C0D4;}" + Environment.NewLine)
        Response.Write("td.title_1{background-color:#d2eaf1;}" + Environment.NewLine)
        Response.Write("td.title_2{background-color:#A5D5E2;}	" + Environment.NewLine)
        Response.Write("#outer2   {width:90%; height: 100%; margin:20px 0px 400px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar2 {width:20%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content2 {width:50%; float:right; /*height:200px; background:#338;*/ }" + Environment.NewLine)
        Response.Write("#outer3   {width:50%; height: 100%; margin:20px 0px 250px 0px; background:white;}" + Environment.NewLine)
        Response.Write("#sidebar3 {width:10%; float:left; /*height:200px; background:#383;*/}" + Environment.NewLine)
        Response.Write("#content3 {width:25%; float:right; /*height:200px; background:#338;*/}" + Environment.NewLine)
        Response.Write("th, td {padding: 1px;}" + Environment.NewLine)

        Response.Write("</style>" + Environment.NewLine)

        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" + Environment.NewLine)

        Response.Write("</head>" + Environment.NewLine)

        Response.Write("<body>" + Environment.NewLine)

        'Response.Write("<br>" + Environment.NewLine)
        'Response.Write("<br>" + Environment.NewLine)

        Response.Write("<div>" + Environment.NewLine)
        Response.Write("<p align=""center""><strong>多次比較圖：</strong></p>" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        'Response.Write("<div align=""center"">" + Environment.NewLine)
        'Response.Write("<img src=""Balance.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("<img src=""SYMM.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:608px;height:402px;"">" + Environment.NewLine)
        'Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""HR.jpg"" border=""1"" alt=""HR.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""SDNN.jpg"" border=""1"" alt=""SDNN.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""ANSAGE.jpg"" border=""1"" alt=""ANSAGE.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""ANS.jpg"" border=""1"" alt=""ANS.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""SYM.jpg"" border=""1"" alt=""SYM.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""VGA.jpg"" border=""1"" alt=""VGA.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)

        Response.Write("<div align=""center"">" + Environment.NewLine)
        Response.Write("<img src=""Balance.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("<img src=""SYMM.jpg"" border=""1"" alt=""SYMMOD.jpg"" style=""width:401px;height:402px;"">" + Environment.NewLine)
        Response.Write("</div>" + Environment.NewLine)


        Response.Write("</body>" + Environment.NewLine)
        Response.Write("</html>" + Environment.NewLine)

        Response.Close()


    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click

        'Dim select_count As Integer = 0

        'For i As Integer = 0 To REPORT_ARRAY_id.Count
        REPORT_ARRAY_id.Clear()
        'Next

        For i As Integer = 0 To (dgvData.Rows.Count - 1) Step 1
            If dgvData.Rows(i).Cells(0).Value Then
                'Console.WriteLine(i & " Checked")
                Console.WriteLine("記錄編號 " & dgvData.Item(1, i).Value())
                REPORT_ARRAY_id.Add(dgvData.Item(1, i).Value())
                'select_count = select_count + 1
            Else
                'Console.WriteLine(i & " Un-Check")
            End If
        Next

        If REPORT_ARRAY_id.Count = 0 Then
            MsgBox("提示 : 請先選擇資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        If REPORT_ARRAY_id.Count > 1 Then
            MsgBox("提示 : 只能查看單筆檢測報告", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        frmProgress.Show()
        frmProgress.TopMost = True
        frmProgress.ProgressBar.Value = 0

        btnReport.Enabled = False
        btnMultiReport.Enabled = False
        btnDel.Enabled = False
        btnSearch.Enabled = False
        btnExport.Enabled = False
        'btnExit.Enabled = False

        Try

            frmMain.myHrv.HRV_PROCESS_DATA_BY_ID(REPORT_ARRAY_id.Item(0), frmMain.current_birth, frmMain.current_sex)

            Make_Report()

            Dim work_file_name As String = Now().ToString("yyyyMMddHHmmss")
            Debug.WriteLine(work_file_name)
            'Dim path As String = Directory.GetCurrentDirectory() & "\data\"
            Dim path As String = Application.StartupPath() & "\data\"
            Console.WriteLine(path)

            'If File.Exists(path & "index.pdf") Then
            '    File.Delete(path & "index.pdf")
            'End If

            Dim pHelp As New ProcessStartInfo
            pHelp.FileName = path & HTML2PDF
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

            frmProgress.ProgressBar.Value = 100
            frmProgress.Dispose()

            MsgBox("檢測報告轉檔成功." & _
                   path & work_file_name & ".pdf", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, frmMain.HRVapp)

            frmMain.myHrv.HRV_CLEAN_FILES()

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

            'System.Diagnostics.Process.Start("C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe D:\index.pdf")
            ' New ProcessStartInfo created
            'Dim p As New ProcessStartInfo

            '' Specify the location of the binary
            'p.FileName = "C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe"

            '' Use these arguments for the process
            'p.Arguments = path & "index.pdf"

            '' Use a hidden window
            'p.WindowStyle = ProcessWindowStyle.Normal

            '' Start the process
            'Process.Start(p)
            frmMain.Enabled = True
            Me.Dispose()
            Me.Close()

        Catch noFile As FileNotFoundException
            MsgBox("# frmDataList FILE Exception " & noFile.ToString, MsgBoxStyle.Critical)
        Catch Ex As Exception
            MsgBox("* frmDataList FILE Exception " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

        btnReport.Enabled = True
        btnMultiReport.Enabled = True
        btnDel.Enabled = True
        btnSearch.Enabled = True
        btnExport.Enabled = True
        frmMain.gbTop.Enabled = True

        frmProgress.Dispose()
        frmMain.Enabled = True
        Me.Dispose()

    End Sub


    Private Sub btnMultiReport_Click(sender As Object, e As EventArgs) Handles btnMultiReport.Click

        'Dim select_count As Integer = 0
      
        'For i As Integer = 0 To REPORT_ARRAY_id.Count
        REPORT_ARRAY_id.Clear()
        'Next

        Console.WriteLine(REPORT_ARRAY_id.Count)

        'For i As Integer = 0 To (dgvData.Rows.Count - 1) Step 1
        For i As Integer = (dgvData.Rows.Count - 1) To 0 Step -1
            If dgvData.Rows(i).Cells(0).Value Then
                'Console.WriteLine(i & " Checked")
                Console.WriteLine("記錄編號 " & dgvData.Item(1, i).Value())
                REPORT_ARRAY_id.Add(dgvData.Item(1, i).Value())
                'select_count = select_count + 1
            Else
                'Console.WriteLine(i & " Un-Check")
            End If
        Next

        If REPORT_ARRAY_id.Count = 0 Then
            MsgBox("提示 : 請先選擇資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If
        If REPORT_ARRAY_id.Count = 1 Then
            MsgBox("提示 : 請選擇兩筆以上的資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If
        If REPORT_ARRAY_id.Count > frmMain.myHrv.MAX_REPORT_COUNT Then
            'MsgBox("提示 : 只能同時查閱三筆資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            MsgBox("提示 : 只能同時查閱十筆資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        frmProgress.Show()
        frmProgress.TopMost = True
        frmProgress.ProgressBar.Value = 0

        btnReport.Enabled = False
        btnMultiReport.Enabled = False
        btnDel.Enabled = False
        btnSearch.Enabled = False
        btnExport.Enabled = False
        'btnExit.Enabled = False

        Try

            For i As Integer = 0 To REPORT_ARRAY_id.Count - 1
                Console.WriteLine("*記錄編號 " & REPORT_ARRAY_id.Item(i))
                frmMain.myHrv.HRV_PROCESS_DATA_BY_ID(REPORT_ARRAY_id.Item(i), frmMain.current_birth, frmMain.current_sex)
                frmMain.myHrv.cDATEE(i) = frmMain.myHrv.test_date
                frmMain.myHrv.cTIME(i) = frmMain.myHrv.test_time

                frmMain.myHrv.cHR(i) = frmMain.myHrv.HR
                frmMain.myHrv.cHR_LEVEL(i) = frmMain.myHrv.HR_LEVEL

                frmMain.myHrv.cSDNN(i) = frmMain.myHrv.SD
                frmMain.myHrv.cSDNN_LEVEL(i) = frmMain.myHrv.SDNN_LEVEL

                'frmMain.myHrv.cANSAGE(i) = frmMain.myHrv.age_MAX

                '20160626
                frmMain.myHrv.cANSAGE_LF_X_UP(i) = frmMain.myHrv.LF_X_UP
                frmMain.myHrv.cANSAGE_LF_X(i) = frmMain.myHrv.LF_X

                frmMain.myHrv.cANSAGE(i) = frmMain.myHrv.ANS_AGE
                frmMain.myHrv.cANSAGE_LEVEL(i) = frmMain.myHrv.ANS_LEVEL

                'frmMain.myHrv.cANS(i) = frmMain.myHrv.ANS_SD
                frmMain.myHrv.cANS(i) = frmMain.myHrv.ln_LF
                frmMain.myHrv.cANS_LEVEL(i) = frmMain.myHrv.ANS_SD_LEVEL
                'frmMain.myHrv.cSYM(i) = frmMain.myHrv.SYM_SD
                'frmMain.myHrv.cVGA(i) = frmMain.myHrv.ln_HF_SD
                frmMain.myHrv.cSYM(i) = frmMain.myHrv.SYM '20160312客戶提示修改
                frmMain.myHrv.cSYM_LEVEL(i) = frmMain.myHrv.SYM_LEVEL

                frmMain.myHrv.cVGA(i) = frmMain.myHrv.ln_HF
                frmMain.myHrv.cVGA_LEVEL(i) = frmMain.myHrv.ln_HF_LEVEL

                frmMain.myHrv.cBALANCE(i) = frmMain.myHrv.Blance
                frmMain.myHrv.cBALANCE_LEVEL(i) = frmMain.myHrv.Blance_LEVEL

                frmMain.myHrv.cSYMMOD(i) = frmMain.myHrv.SYM_Modulation
                frmMain.myHrv.cSYMMOD_LEVEL(i) = frmMain.myHrv.SYM_Modulation_LEVEL
            Next

            frmProgress.ProgressBar.Value = 10
            frmMain.myHrv.Compare_Report_Image("HR", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 20
            frmMain.myHrv.Compare_Report_Image("SDNN", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 30
            frmMain.myHrv.Compare_Report_Image("ANSage", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 40
            frmMain.myHrv.Compare_Report_Image("ANS", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 50
            frmMain.myHrv.Compare_Report_Image("SYM", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 60
            frmMain.myHrv.Compare_Report_Image("VGA", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 70
            frmMain.myHrv.Compare_Report_Image("BALANCE", REPORT_ARRAY_id.Count)
            frmProgress.ProgressBar.Value = 80
            frmMain.myHrv.Compare_Report_Image("SYMM", REPORT_ARRAY_id.Count)

            Multi_Report_Page_1()
            'Multi_Report_Page_2()

            'If REPORT_ARRAY_id.Count = 2 Then
            '    Console.WriteLine("REPORT_ARRAY_id.Count = 2 ")
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("HR", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "心跳速率", frmMain.myHrv.cHR(0), frmMain.myHrv.cHR(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SDNN", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "間距標準差(SDNN)", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("ANSAGE", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "自律神經年齡(ANS Age)", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("ANS", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "自律神經總功能(ANS)", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SYM", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "交感神經功能(SYM)", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("VGA", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "副交感神經功能(VGA)", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SYMMOD", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), "交感調控(SYM Modulation)", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1))

            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("HR", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "HR", frmMain.myHrv.cHR(0), frmMain.mHrv.cHR(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SDNN", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "SDNN", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("ANSAGE", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "ANS,Age", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("ANS", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "ANS", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SYM", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "SYM", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("VGA", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), "VGA", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_2("SYMMOD", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1), "SYM,Modulation", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1))

            '    frmMain.myHrv.HRV_COMPARE_FIG_2("HR", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "HR", frmMain.myHrv.cHR(0), frmMain.myHrv.cHR(1))
            '    frmProgress.ProgressBar.Value = 10
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("SDNN", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "SDNN", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1))
            '    frmProgress.ProgressBar.Value = 20
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("ANSAGE", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "ANS,Age", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1))
            '    frmProgress.ProgressBar.Value = 30
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("ANS", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "ANS", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1))
            '    frmProgress.ProgressBar.Value = 40
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("SYM", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "SYM", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1))
            '    frmProgress.ProgressBar.Value = 50
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("VGA", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "VGA", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1))
            '    frmProgress.ProgressBar.Value = 60
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_2("SYMMOD", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), "SYM,Modulation", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1))
            '    frmProgress.ProgressBar.Value = 70

            '    Compare_2_Report()

            'End If

            'If REPORT_ARRAY_id.Count = 3 Then
            '    Console.WriteLine("REPORT_ARRAY_id.Count = 3 ")
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("HR", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "心跳速率", frmMain.myHrv.cHR(0), frmMain.myHrv.cHR(1), frmMain.myHrv.cHR(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SDNN", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "間距標準差(SDNN)", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1), frmMain.myHrv.cSDNN(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("ANSAGE", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "自律神經年齡(ANS Age)", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1), frmMain.myHrv.cANSAGE(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("ANS", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "自律神經總功能(ANS)", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1), frmMain.myHrv.cANS(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SYM", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "交感神經功能(SYM)", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1), frmMain.myHrv.cSYM(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("VGA", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "副交感神經功能(VGA)", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1), frmMain.myHrv.cVGA(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SYMMOD", frmMain.myHrv.cDATEE(0) & " " & frmMain.myHrv.cTIME(0), frmMain.myHrv.cDATEE(1) & " " & frmMain.myHrv.cTIME(1), frmMain.myHrv.cDATEE(2) & " " & frmMain.myHrv.cTIME(2), "交感調控(SYM Modulation)", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1), frmMain.myHrv.cSYMMOD(2))

            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("HR", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "HR", frmMain.myHrv.cHR(0), frmMain.myHrv.cHR(1), frmMain.myHrv.cHR(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SDNN", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "SDNN", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1), frmMain.myHrv.cSDNN(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("ANSAGE", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "ANS,Age", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1), frmMain.myHrv.cANSAGE(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("ANS", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "ANS", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1), frmMain.myHrv.cANS(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SYM", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "SYM", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1), frmMain.myHrv.cSYM(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("VGA", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "VGA", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1), frmMain.myHrv.cVGA(2))
            '    'frmMain.myHrv.HRV_COMPARE_FIG_3("SYMMOD", frmMain.myHrv.cDATEE(0).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ","), frmMain.myHrv.cDATEE(1).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ","), frmMain.myHrv.cDATEE(2).Replace(" ", ",") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ","), "SYM,Modulation", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1), frmMain.myHrv.cSYMMOD(2))

            '    frmMain.myHrv.HRV_COMPARE_FIG_3("HR", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "HR", frmMain.myHrv.cHR(0), frmMain.myHrv.cHR(1), frmMain.myHrv.cHR(2))
            '    frmProgress.ProgressBar.Value = 10
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("SDNN", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "SDNN", frmMain.myHrv.cSDNN(0), frmMain.myHrv.cSDNN(1), frmMain.myHrv.cSDNN(2))
            '    frmProgress.ProgressBar.Value = 20
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("ANSAGE", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "ANS,Age", frmMain.myHrv.cANSAGE(0), frmMain.myHrv.cANSAGE(1), frmMain.myHrv.cANSAGE(2))
            '    frmProgress.ProgressBar.Value = 30
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("ANS", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "ANS", frmMain.myHrv.cANS(0), frmMain.myHrv.cANS(1), frmMain.myHrv.cANS(2))
            '    frmProgress.ProgressBar.Value = 40
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("SYM", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "SYM", frmMain.myHrv.cSYM(0), frmMain.myHrv.cSYM(1), frmMain.myHrv.cSYM(2))
            '    frmProgress.ProgressBar.Value = 50
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("VGA", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "VGA", frmMain.myHrv.cVGA(0), frmMain.myHrv.cVGA(1), frmMain.myHrv.cVGA(2))
            '    frmProgress.ProgressBar.Value = 60
            '    Thread.Sleep(200)
            '    frmMain.myHrv.HRV_COMPARE_FIG_3("SYMMOD", frmMain.myHrv.cDATEE(0).Replace(" ", "") & "," & frmMain.myHrv.cTIME(0).Replace(" ", ""), frmMain.myHrv.cDATEE(1).Replace(" ", "") & "," & frmMain.myHrv.cTIME(1).Replace(" ", ""), frmMain.myHrv.cDATEE(2).Replace(" ", "") & "," & frmMain.myHrv.cTIME(2).Replace(" ", ""), "SYM,Modulation", frmMain.myHrv.cSYMMOD(0), frmMain.myHrv.cSYMMOD(1), frmMain.myHrv.cSYMMOD(2))
            '    frmProgress.ProgressBar.Value = 70


            '    Compare_3_Report()

            'End If

            frmProgress.ProgressBar.Value = 90

            Dim work_file_name As String = Now().ToString("yyyyMMddHHmmss")
            Debug.WriteLine(work_file_name)
            'Dim path As String = Directory.GetCurrentDirectory() & "\data\"
            Dim path As String = Application.StartupPath() & "\data\"
            Console.WriteLine(path)

            'If File.Exists(path & "index.pdf") Then
            'File.Delete(path & "index.pdf")
            'End If

            Dim pHelp As New ProcessStartInfo
            pHelp.FileName = path & HTML2PDF
            'pHelp.Arguments = path & "compare_p1.html " & path & "compare_p2.html " & path & work_file_name & ".pdf"
            pHelp.Arguments = path & "compare_p1.html " & path & work_file_name & ".pdf"
            'pHelp.FileName = path & "Html2PDF.bat"
            pHelp.UseShellExecute = True
            pHelp.WindowStyle = ProcessWindowStyle.Hidden
            Dim proc As Process = Process.Start(pHelp)
            proc.Dispose()

            frmProgress.ProgressBar.Value = 90

            While (1)
                If File.Exists(path & work_file_name & ".pdf") Then
                    Exit While
                End If
                Thread.Sleep(200)
                Application.DoEvents()
            End While

            frmProgress.ProgressBar.Value = 100
            frmProgress.Dispose()

            MsgBox("檢測報告轉檔成功." & _
                   path & work_file_name & ".pdf", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, frmMain.HRVapp)

            frmMain.myHrv.HRV_CLEAN_FILES()

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

            'System.Diagnostics.Process.Start("C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe D:\index.pdf")
            ' New ProcessStartInfo created
            'Dim p As New ProcessStartInfo

            '' Specify the location of the binary
            'p.FileName = "C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe"

            '' Use these arguments for the process
            'p.Arguments = path & "index.pdf"

            '' Use a hidden window
            'p.WindowStyle = ProcessWindowStyle.Normal

            '' Start the process
            'Process.Start(p)
            frmMain.Enabled = True
            Me.Dispose()
            Me.Close()

        Catch noFile As FileNotFoundException
            MsgBox("# frmDataList FILE Exception " & noFile.ToString, MsgBoxStyle.Critical)
        Catch Ex As Exception
            MsgBox("* frmDataList FILE Exception " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

        btnReport.Enabled = True
        btnMultiReport.Enabled = True
        btnDel.Enabled = True
        btnSearch.Enabled = True
        btnExport.Enabled = True
        frmMain.gbTop.Enabled = True

        frmProgress.Dispose()
        frmMain.Enabled = True
        Me.Dispose()

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click

        REPORT_ARRAY_id.Clear()
        Dim current_date = ""

        For i As Integer = 0 To (dgvData.Rows.Count - 1) Step 1
            If dgvData.Rows(i).Cells(0).Value Then
                'Console.WriteLine(i & " Checked")
                Console.WriteLine("記錄編號 " & dgvData.Item(1, i).Value())
                'Console.WriteLine("記錄編號 " & dgvData.Item(2, i).Value())
                REPORT_ARRAY_id.Add(dgvData.Item(1, i).Value())
                current_date = dgvData.Item(2, i).Value()
                'select_count = select_count + 1
            Else
                'Console.WriteLine(i & " Un-Check")
            End If
        Next

        If REPORT_ARRAY_id.Count = 0 Then
            MsgBox("提示 : 請先選擇資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        If REPORT_ARRAY_id.Count > 1 Then
            MsgBox("提示 : 只能刪除單筆檢測報告", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        btnReport.Enabled = False
        btnMultiReport.Enabled = False
        btnExport.Enabled = False

        Try

            If MessageBox.Show("確定刪除 " & frmMain.current_name & " - " & current_date & " 的檢測資料?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If frmMain.myHrv.HRV_DEL_DATA_BY_ID(REPORT_ARRAY_id.Item(0)) = True Then
                    MsgBox("資料已刪除成功", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
                    LoadDataRecord()
                Else
                    MsgBox("資料已刪除失敗", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
                End If

                frmMain.Enabled = True
                'Me.Dispose()
                'Me.Close()
            Else
                'Return
            End If

        Catch Ex As Exception
            MsgBox("資料刪除錯誤: " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

        btnReport.Enabled = True
        btnMultiReport.Enabled = True
        'btnExit.Enabled = True
        btnExport.Enabled = True
        frmMain.gbTop.Enabled = True

        frmMain.Enabled = True
        'Me.Dispose()

    End Sub

   
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        'https://msdn.microsoft.com/en-us/library/b5xbyt6f(v=vs.90).aspx

        Console.WriteLine(DTPFrom.Value.ToString("yyyy/MM/dd"))
        Console.WriteLine(DTPTo.Value.ToString("yyyy/MM/dd"))

        Dim diff As Long = DateDiff("d", DTPFrom.Value, DTPTo.Value)
        If diff < 0 Then
            MsgBox("日期格式錯誤, 請重新選擇.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        Else
            LoadDataRecord_By_Date(DTPFrom.Value.ToString("yyyy/MM/dd"), DTPTo.Value.ToString("yyyy/MM/dd"))
        End If

        'If diff > 0 Then
        '    Console.WriteLine("DTPFrom is greater than DTPTo " & diff)
        'ElseIf diff < 0 Then
        '    Console.WriteLine("DTPFrom is lesser than DTPTo " & diff)
        'Else
        '    Console.WriteLine("DTPFrom is equal to DTPTo " & diff)
        'End If

        'If Not DateTime.Compare(DTPFrom.Value, DTPTo.Value) = 0 Then
        '    'they are same
        '    Console.WriteLine("DTPFrom = DTPTo")
        'ElseIf DateTime.Compare(DTPFrom.Value, DTPTo.Value) > 0 Then
        '    'dd1 is later than dd2
        '    Console.WriteLine("DTPFrom > DTPTo")
        'Else
        '    'dd1 is prior to dd2
        '    Console.WriteLine("DTPFrom < DTPTo")
        'End If

    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        REPORT_ARRAY_id.Clear()

        Console.WriteLine(REPORT_ARRAY_id.Count)

        'For i As Integer = 0 To (dgvData.Rows.Count - 1) Step 1
        For i As Integer = (dgvData.Rows.Count - 1) To 0 Step -1
            If dgvData.Rows(i).Cells(0).Value Then
                'Console.WriteLine(i & " Checked")
                Console.WriteLine("記錄編號 " & dgvData.Item(1, i).Value())
                REPORT_ARRAY_id.Add(dgvData.Item(1, i).Value())
                'select_count = select_count + 1
            Else
                'Console.WriteLine(i & " Un-Check")
            End If
        Next

        If REPORT_ARRAY_id.Count = 0 Then
            MsgBox("提示 : 請先選擇資料", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return
        End If

        btnReport.Enabled = False
        btnMultiReport.Enabled = False
        btnDel.Enabled = False
        btnSearch.Enabled = False
        btnExport.Enabled = False

        frmProgress.Show()
        frmProgress.TopMost = True
        frmProgress.ProgressBar.Value = 10
        frmProgress.ProgressBar.Step = 20

        Try

            If xlApp Is Nothing Then
                MessageBox.Show("未安裝 Excel 元件")
                Return
            End If

            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            xlWorkBook = xlApp.Workbooks.Add(misValue)
            'xlWorkSheet = xlWorkBook.Sheets("工作表1")
            xlWorkSheet = xlWorkBook.ActiveSheet

            xlWorkSheet.Cells(1, 1) = "項目 / 時間"
            xlWorkSheet.Cells(2, 1) = "心跳速率 (HR)"
            xlWorkSheet.Cells(3, 1) = "NN間距標準差 (SDNN)"
            xlWorkSheet.Cells(4, 1) = "自律神經年齡 (ANS Age)"
            xlWorkSheet.Cells(5, 1) = "自律神經總功能(ANS)"
            xlWorkSheet.Cells(6, 1) = "交感神經功能 (SYM)"
            xlWorkSheet.Cells(7, 1) = "副交感神經功能 (VAG)"
            xlWorkSheet.Cells(8, 1) = "自律神經偏向 (Balence)"
            xlWorkSheet.Cells(9, 1) = "交感調控 (SYM Modulation)"
            xlWorkSheet.Cells(10, 1) = "高頻功率 (TP)"
            xlWorkSheet.Cells(11, 1) = "極低頻功率 (VL)"
            xlWorkSheet.Cells(12, 1) = "低頻功率 (LF)"
            xlWorkSheet.Cells(13, 1) = "高頻功率 (HF)"
            xlWorkSheet.Cells(14, 1) = "低頻百分率 (LF%)"
            xlWorkSheet.Cells(15, 1) = "高頻百分率 (HF%)"
            xlWorkSheet.Columns.AutoFit()

            For i As Integer = 0 To REPORT_ARRAY_id.Count - 1

                Console.WriteLine("*記錄編號 " & REPORT_ARRAY_id.Item(i))
                frmMain.myHrv.HRV_PROCESS_DATA_BY_ID(REPORT_ARRAY_id.Item(i), frmMain.current_birth, frmMain.current_sex)

                Console.WriteLine("=======================================================================")
                Console.WriteLine("Date " & frmMain.myHrv.test_date)
                Console.WriteLine("Time " & frmMain.myHrv.test_time)
                Console.WriteLine("HR " & frmMain.myHrv.HR) '心跳速率 (HR)
                Console.WriteLine("SDNN " & frmMain.myHrv.SD) 'NN間距標準差 (SDNN)
                Console.WriteLine("ANS AGE " & frmMain.myHrv.ANS_AGE) '自律神經年齡 (ANS Age)
                Console.WriteLine("ANS " & frmMain.myHrv.ln_LF) '自律神經總功能(ANS)
                Console.WriteLine("SYM " & frmMain.myHrv.SYM) '交感神經功能 (SYM)
                Console.WriteLine("VGA " & frmMain.myHrv.ln_HF) '副交感神經功能 (VAG)
                Console.WriteLine("Balance " & frmMain.myHrv.Blance) '自律神經偏向 (Balence)
                Console.WriteLine("SYM Modulation " & frmMain.myHrv.SYM_Modulation) '交感調控 (SYM Modulation)
                Console.WriteLine("TP " & frmMain.myHrv.TP) '高頻功率 (TP)
                Console.WriteLine("VL " & frmMain.myHrv.VL) '極低頻功率 (VL)
                Console.WriteLine("HF " & frmMain.myHrv.LF) '低頻功率 (LF)
                Console.WriteLine("HF " & frmMain.myHrv.HF) '高頻功率 (HF)
                Console.WriteLine("LF% " & frmMain.myHrv.sym_percent) '交感
                Console.WriteLine("HF% " & frmMain.myHrv.vga_percent) '副交感
                Console.WriteLine("=======================================================================")

                xlWorkSheet.Cells(1, i + 2) = frmMain.myHrv.test_date.ToString + vbCrLf + frmMain.myHrv.test_time.ToString
                xlWorkSheet.Cells(2, i + 2) = frmMain.myHrv.HR
                xlWorkSheet.Cells(3, i + 2) = frmMain.myHrv.SD
                'xlWorkSheet.Cells(4, i + 2) = frmMain.myHrv.ANS_AGE
                xlWorkSheet.Cells(4, i + 2) = frmMain.myHrv.LF_X_UP & "-" & frmMain.myHrv.LF_X
                xlWorkSheet.Cells(5, i + 2) = frmMain.myHrv.ln_LF
                xlWorkSheet.Cells(6, i + 2) = frmMain.myHrv.SYM
                xlWorkSheet.Cells(7, i + 2) = frmMain.myHrv.ln_HF
                xlWorkSheet.Cells(8, i + 2) = frmMain.myHrv.Blance
                xlWorkSheet.Cells(9, i + 2) = frmMain.myHrv.SYM_Modulation
                xlWorkSheet.Cells(10, i + 2) = frmMain.myHrv.TP
                xlWorkSheet.Cells(11, i + 2) = frmMain.myHrv.VL
                xlWorkSheet.Cells(12, i + 2) = frmMain.myHrv.LF
                xlWorkSheet.Cells(13, i + 2) = frmMain.myHrv.HF
                xlWorkSheet.Cells(14, i + 2) = frmMain.myHrv.sym_percent
                xlWorkSheet.Cells(15, i + 2) = frmMain.myHrv.vga_percent
                'xlWorkSheet.Columns.HorizontalAlignment = xlRight
                'xlWorkSheet.Range("B4").HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                xlWorkSheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                xlWorkSheet.Columns.AutoFit()

                'frmMain.myHrv.cDATEE(i) = frmMain.myHrv.test_date
                'frmMain.myHrv.cTIME(i) = frmMain.myHrv.test_time
                'frmMain.myHrv.cHR(i) = frmMain.myHrv.HR
                'frmMain.myHrv.cSDNN(i) = frmMain.myHrv.SD
                ''frmMain.myHrv.cANSAGE(i) = frmMain.myHrv.age_MAX
                'frmMain.myHrv.cANSAGE(i) = frmMain.myHrv.ANS_AGE
                'frmMain.myHrv.cANS(i) = frmMain.myHrv.ANS_SD
                ''frmMain.myHrv.cSYM(i) = frmMain.myHrv.SYM_SD
                ''frmMain.myHrv.cVGA(i) = frmMain.myHrv.ln_HF_SD
                'frmMain.myHrv.cSYM(i) = frmMain.myHrv.SYM '20160312客戶提示修改
                'frmMain.myHrv.cVGA(i) = frmMain.myHrv.ln_HF
                'frmMain.myHrv.cBALANCE(i) = frmMain.myHrv.Blance
                'frmMain.myHrv.cSYMMOD(i) = frmMain.myHrv.SYM_Modulation           

                frmProgress.ProgressBar.PerformStep()

            Next

            'btnReport.Enabled = True
            'btnMultiReport.Enabled = True
            'btnDel.Enabled = True
            'btnSearch.Enabled = True

            'Dim work_file_name As String = Now().ToString("yyyyMMddHHmmss")
            'Debug.WriteLine(work_file_name)
            ''Dim path As String = Directory.GetCurrentDirectory() & "\data\"
            'Dim path As String = Application.StartupPath() & "\data\"
            'Console.WriteLine(path)

            ''If File.Exists(path & "index.pdf") Then
            ''File.Delete(path & "index.pdf")
            ''End If

            'Dim pHelp As New ProcessStartInfo
            'pHelp.FileName = path & HTML2PDF
            'pHelp.Arguments = path & "compare.html " & path & work_file_name & ".pdf"
            ''pHelp.FileName = path & "Html2PDF.bat"
            'pHelp.UseShellExecute = True
            'pHelp.WindowStyle = ProcessWindowStyle.Hidden
            'Dim proc As Process = Process.Start(pHelp)
            'proc.Dispose()

            'frmProgress.ProgressBar.Value = 90

            'While (1)
            '    If File.Exists(path & work_file_name & ".pdf") Then
            '        Exit While
            '    End If
            '    Thread.Sleep(200)
            '    Application.DoEvents()
            'End While

            'frmProgress.ProgressBar.Value = 100
            frmProgress.ProgressBar.Value = 100
            frmProgress.Dispose()

            Dim work_file_name As String = Now().ToString("yyyyMMddHHmmss")
            Dim FILE_NAME As String = frmMain.app_FilePath & "\data\" & frmMain.current_pid & "_" & work_file_name & ".xls"
            xlWorkBook.SaveAs(FILE_NAME, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            Dim result As String = MsgBox("匯出完成,是否開啟檔案 " & FILE_NAME & " ?", MsgBoxStyle.YesNo, frmMain.HRVapp)

            If result = vbYes Then

                Dim excel_proc As New Process
                'Set the process details
                With excel_proc.StartInfo
                    'Set the information for the file to launch
                    '.Verb = "print"
                    .FileName = FILE_NAME
                    .UseShellExecute = True
                End With
                'Open the file
                excel_proc.Start()
                excel_proc.Dispose()


            Else

                'Me.Close()
            End If

            'MessageBox.Show("Excel file created , you can find the file d:\csharp-Excel.xls")

            'For i As Integer = 1 To REPORT_ARRAY_id.Count
            '    frmProgress.ProgressBar.Value = ((REPORT_ARRAY_id.Count / 100) * 1000) * i
            '    Console.WriteLine("--- " & ((REPORT_ARRAY_id.Count / 100) * 1000) * i)
            '    Thread.Sleep(100)
            'Next
            'frmProgress.Dispose()

            'MsgBox("檢測報告轉檔成功." & _
            '       path & work_file_name & ".pdf", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, frmMain.HRVapp)

            'frmMain.myHrv.HRV_CLEAN_FILES()

            'Dim pdf_proc As New Process
            ''Set the process details
            'With pdf_proc.StartInfo
            '    'Set the information for the file to launch
            '    '.Verb = "print"
            '    .FileName = path & work_file_name & ".pdf"
            '    .UseShellExecute = True
            'End With
            ''Open the file
            'pdf_proc.Start()
            'pdf_proc.Dispose()

            'System.Diagnostics.Process.Start("C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe D:\index.pdf")
            ' New ProcessStartInfo created
            'Dim p As New ProcessStartInfo

            '' Specify the location of the binary
            'p.FileName = "C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe"

            '' Use these arguments for the process
            'p.Arguments = path & "index.pdf"

            '' Use a hidden window
            'p.WindowStyle = ProcessWindowStyle.Normal

            '' Start the process
            'Process.Start(p)
            'frmMain.Enabled = True
            'Me.Dispose()
            'Me.Close()
            frmMain.Enabled = True
            Me.Dispose()
            Me.Close()

        Catch noFile As FileNotFoundException
            MsgBox("# btnExport_Click FILE Exception " & noFile.ToString, MsgBoxStyle.Critical)
        Catch Ex As Exception
            MsgBox("* btnExport_Click FILE Exception " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

        'btnReport.Enabled = True
        'btnMultiReport.Enabled = True
        ''btnExit.Enabled = True
        'frmMain.gbTop.Enabled = True

        'frmMain.Enabled = True
        'Me.Dispose()

        btnReport.Enabled = True
        btnMultiReport.Enabled = True
        btnDel.Enabled = True
        btnSearch.Enabled = True
        btnExport.Enabled = True
        frmMain.gbTop.Enabled = True

        frmMain.Enabled = True
        Me.Dispose()
        Me.Close()

    End Sub

    

End Class
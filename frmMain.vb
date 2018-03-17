Imports libhrv
Imports System.IO.MemoryMappedFiles
Imports System.IO
Imports System.Data.OleDb
Imports System.Threading

Public Class frmMain

    Dim APP_NAME As String = "HRVapp"
    'Dim VERSION As String = "v1.0" '201512xx
    'Dim VERSION As String = "v2.0" '20151230
    'Dim VERSION As String = "v2.1" '20151231
    'Dim VERSION As String = "v2.2" '20160101
    'Dim VERSION As String = "v2.3" '20160104
    'Dim VERSION As String = "v2.4" '20160107
    'Dim VERSION As String = "v2.5" '20160109
    'Public VERSION As String = "v2.6" '20160123
    'Public VERSION As String = "v2.7" '20160131
    'Public VERSION As String = "v2.8" '20160312
    'Public VERSION As String = "v2.81" '20160315
    'Public VERSION As String = "v2.82" '20160320
    'Public VERSION As String = "v2.83" '20160329
    'Public VERSION As String = "v2.84" '20160330
    'Public VERSION As String = "v2.85" '20160413
    'Public VERSION As String = "v2.86" '20160426
    'Public VERSION As String = "v2.87" '20160504
    'Public VERSION As String = "v2.87a" '20160509
    'Public VERSION As String = "v2.87b" '20160514
    'Public VERSION As String = "v2.87c" '20160516
    'Public VERSION As String = "v2.87d" '20160524
    'Public VERSION As String = "v2.87e" '20160527
    'Public VERSION As String = "v2.88" '20160609
    'Public VERSION As String = "v2.88a" '20160614
    'Public VERSION As String = "v2.88b" '20160623
    Public VERSION As String = "v2.88c" '20160626
    Public HRVapp As String = APP_NAME & " " & VERSION

    Public app_FilePath As String = Application.StartupPath

    Dim SHARE_NAME As String = "HRV"
    Dim SHARE_LEN As Integer = 129

    Public myHrv As New Hrvlib
    Public current_user_id As Integer = -1
    Public current_name As String
    Public current_birth As String
    Public current_sex As String
    Public current_pid As String

    'http://www.coderslexicon.com/quick-tutorial-building-menus-dynamically-with-the-menustrip-control-in-c/
    'http://www.coderslexicon.com/quick-tutorial-building-menus-dynamically-with-the-menustrip-control-in-c/
    'https://social.msdn.microsoft.com/Forums/en-US/c5fa7e88-a739-48d0-a23f-38401ca81345/event-handler-for-dynamically-created-menustip-vbnet-2008?forum=vbgeneral
    'http://www.codeproject.com/Articles/19223/Dynamic-Creation-Of-MenuStrip-VB-NET
    Dim HrvMenu As MenuStrip = New MenuStrip()

#If EngVersion Then
 Dim mSystem As ToolStripMenuItem = New ToolStripMenuItem("System")
    'Create our first sub item with a delete icon image
    'Dim firstSubitem As ToolStripMenuItem = New ToolStripMenuItem("First Sub Item", Image.FromFile("c:\\Delete.png"))
    Public mSystem_Login As ToolStripMenuItem = New ToolStripMenuItem("Login")
    Dim mSystem_End As ToolStripMenuItem = New ToolStripMenuItem("Exit")
    Public mConfig As ToolStripMenuItem = New ToolStripMenuItem("Setting")
    Public mConfig_Com As ToolStripMenuItem = New ToolStripMenuItem("COM Port")
    Public mConfig_Backup As ToolStripMenuItem = New ToolStripMenuItem("Backup")
    Public mConfig_PWD As ToolStripMenuItem = New ToolStripMenuItem("Modify Password")
    Public mConfig_TESTER As ToolStripMenuItem = New ToolStripMenuItem("Operator")
    Public mConfig_COUNT As ToolStripMenuItem = New ToolStripMenuItem("Detection times")


#else

    Dim mSystem As ToolStripMenuItem = New ToolStripMenuItem("系統")
    'Create our first sub item with a delete icon image
    'Dim firstSubitem As ToolStripMenuItem = New ToolStripMenuItem("First Sub Item", Image.FromFile("c:\\Delete.png"))
    Public mSystem_Login As ToolStripMenuItem = New ToolStripMenuItem("登入")
    Dim mSystem_End As ToolStripMenuItem = New ToolStripMenuItem("結束")

    Public mConfig As ToolStripMenuItem = New ToolStripMenuItem("設定")
    Public mConfig_Com As ToolStripMenuItem = New ToolStripMenuItem("COM Port")
    Public mConfig_Backup As ToolStripMenuItem = New ToolStripMenuItem("資料備份")
    Public mConfig_PWD As ToolStripMenuItem = New ToolStripMenuItem("變更密碼")
    Public mConfig_TESTER As ToolStripMenuItem = New ToolStripMenuItem("操作人員")
    Public mConfig_COUNT As ToolStripMenuItem = New ToolStripMenuItem("檢測次數")
#end if
    'Create two independent sub sub items
    'Dim m2_1_1 As ToolStripMenuItem() = {New ToolStripMenuItem("COM3"), New ToolStripMenuItem("COM4")}
    Dim mConfig_Com_Subcom As ToolStripMenuItem()
    'Dim m2_1_2 As ToolStripMenuItem = New ToolStripMenuItem("COM4")

    Dim COM_ARRAY As New ArrayList()
    Dim btn_Status As String
    Public logout_timer As Integer = 0
    Public logout_timer_max As Integer = 60 * 10 '10 Min
    'Public logout_timer_max As Integer = 10 * 1 '10 Min
    Public current_login_id As String
    Public current_login_name As String
    Public current_com_port As String
    Public current_login_unit As String
    Dim check_flag As Boolean = False
    Dim BK_PATH As String

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
#If EngVersion Then
        If MessageBox.Show("Exit this system?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
#else
        If MessageBox.Show("確定結束程式?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
#end if
            myHrv.HRV_Dispose()
        Else
            e.Cancel = True
        End If

    End Sub 'Form1_Closing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = HRVapp
        Me.lbTitle.Text = HRVapp

        myHrv.HRV_LICENSE_INIT()

        If myHrv.HRV_CHECK_LICENSE() = True Then
            MsgBox("請更新應用程式.", MsgBoxStyle.Critical)
            End
        End If

        If myHrv.HRV_INIT() <> True Then
            End
        End If

        'myHrv.HRV_LICENSE_INIT()

        'Dim E_VALUE As Double = -2.1
        'If E_VALUE > 1.5 Then
        '    Console.WriteLine("1")
        'ElseIf 1.5 >= E_VALUE And E_VALUE > 1.2 Then
        '    Console.WriteLine("2")
        'ElseIf 1.2 >= E_VALUE And E_VALUE > 1 Then
        '    Console.WriteLine("3")
        'ElseIf 1 >= E_VALUE And E_VALUE > 0.9 Then
        '    Console.WriteLine("4")
        'ElseIf 0.9 >= E_VALUE And E_VALUE > 0.8 Then
        '    Console.WriteLine("5")
        'ElseIf 0.8 >= E_VALUE And E_VALUE > 0 Then
        '    Console.WriteLine("6")
        'ElseIf 0 >= E_VALUE And E_VALUE > -0.8 Then
        '    Console.WriteLine("7")
        'ElseIf -0.8 >= E_VALUE And E_VALUE > -0.9 Then
        '    Console.WriteLine("9")
        'ElseIf -0.9 >= E_VALUE And E_VALUE > -1 Then
        '    Console.WriteLine("10")
        'ElseIf -1 >= E_VALUE And E_VALUE > -1.2 Then
        '    Console.WriteLine("11")
        'ElseIf -1.2 >= E_VALUE And E_VALUE > -1.5 Then
        '    Console.WriteLine("12")
        'ElseIf -1.5 >= E_VALUE And E_VALUE > -2.0 Then
        '    Console.WriteLine("14")
        'ElseIf E_VALUE <= -2.0 Then
        '    Console.WriteLine("15")
        'End If

        'Dim a As Integer = 1
        'Dim b As Integer = 2
        'Dim c As String = "3"
        'Dim d As String = "4"

        'Dim p As New ProcessStartInfo("C:\Users\Carrera\Qsync\myProject\TEST_CODE\MS_CHART\MS_CHART\bin\Debug\MS_CHART.exe")
        'p.Arguments = a & " " & b & " " & c & " " & d
        'p.WindowStyle = ProcessWindowStyle.Hidden
        'p.CreateNoWindow = True
        'Process.Start(p)

        'End

        'Dim current_dt As String = Now().ToString("yyyyMMddhhmmss")

        'USE FOR DEBUG ONLY
        'If myHrv.HRV_DATA_PARSER(2) <> True Then
        '    End
        'End If

        'If myHrv.HRV_CHECK_DATA() = True Then
        '    'MsgBox("檢測完成", MsgBoxStyle.Information)
        '    myHrv.HRV_PROCESS_OUT()
        '    If myHrv.HRV_DATA_PARSER(1) <> True Then
        '        'End
        '    End If
        'End If

        'MsgBox("注意: Access encrypt ", MsgBoxStyle.Exclamation)

        'If myHrv.HRV_NEW_TESTER("TEST", "admin", "admin") <> True Then
        '    End
        'End If

        '=======================================================================
        COM_ARRAY.Clear()
        myHrv.HRV_SCAN_COM(COM_ARRAY)

        'COM_ARRAY.Add("COM3")
        'COM_ARRAY.Add("COM4")
        'COM_ARRAY.Add("COM5")

        If COM_ARRAY.Count = 0 Then
            'COM_ARRAY.Add("None")
#if EngVersion
MsgBox("Note! Cannot find the COM port.", MsgBoxStyle.Exclamation)
#else
MsgBox("注意: 系統找不到 COM Port.", MsgBoxStyle.Exclamation)
#end if            
        Else
            Dim obj As [Object]
            For Each obj In COM_ARRAY
                Console.WriteLine("-->   {0}", obj)
            Next obj
        End If
        '=======================================================================

        LoadUserRecord()

        ''http://www.tutorialspoint.com/vb.net/vb.net_advanced_forms.htm
        ''defining the main menu bar
        'Dim mnuBar As New MainMenu()
        ''defining the menu items for the main menu bar
        'Dim myMenuItemFile As New MenuItem("&File")
        'Dim myMenuItemEdit As New MenuItem("&Edit")
        'Dim myMenuItemView As New MenuItem("&View")
        'Dim myMenuItemProject As New MenuItem("&Project")

        ''adding the menu items to the main menu bar
        'mnuBar.MenuItems.Add(myMenuItemFile)
        'mnuBar.MenuItems.Add(myMenuItemEdit)
        'mnuBar.MenuItems.Add(myMenuItemView)
        'mnuBar.MenuItems.Add(myMenuItemProject)

        '' defining some sub menus
        'Dim myMenuItemNew As New MenuItem("&New")
        'Dim myMenuItemOpen As New MenuItem("&Open")
        'Dim myMenuItemSave As New MenuItem("&Save")

        ''add sub menus to the File menu
        'myMenuItemFile.MenuItems.Add(myMenuItemNew)
        'myMenuItemFile.MenuItems.Add(myMenuItemOpen)
        'myMenuItemFile.MenuItems.Add(myMenuItemSave)

        ''add the main menu to the form
        'Me.Menu = mnuBar

        ''Create high level menu container
        'Dim strip As MenuStrip = New MenuStrip()
        ''Create a top level menu item called "File" with "F" being the access key (Alt + f)
        'Dim fileItem As ToolStripMenuItem = New ToolStripMenuItem("&File")
        ''Create one sub menu item on this menu
        'fileItem.DropDownItems.Add("First Menu Item")
        ''Add the high level menu item to the menu container
        'strip.Items.Add(fileItem)
        ''Add menu to form
        'Me.Controls.Add(strip)


        HrvMenu.GripStyle = ToolStripGripStyle.Visible
        HrvMenu.BackColor = Color.AliceBlue

        'Add the two items to the first sub item
        mSystem.DropDownItems.Add(mSystem_Login)
        mSystem.DropDownItems.Add(mSystem_End)
        HrvMenu.Items.Add(mSystem)
        'mSystem_End.Enabled = False

        mConfig.DropDownItems.Add(mConfig_Com)
        mConfig.DropDownItems.Add(mConfig_Backup)
        mConfig.DropDownItems.Add(mConfig_PWD)
        mConfig.DropDownItems.Add(mConfig_TESTER)
        mConfig.DropDownItems.Add(mConfig_COUNT)

        mConfig.Visible = False
        mConfig_Backup.Visible = False '備份資料
        mConfig_PWD.Enabled = True '變更密碼
        mConfig_TESTER.Visible = False '編修操作人員
        mConfig_COUNT.Visible = False '統計檢測次數

        ReDim mConfig_Com_Subcom(COM_ARRAY.Count)
        For i As Integer = 0 To COM_ARRAY.Count - 1
            mConfig_Com_Subcom(i) = New ToolStripMenuItem(COM_ARRAY(i).ToString)
            mConfig_Com_Subcom(i).Name = COM_ARRAY(i).ToString
            mConfig_Com.DropDownItems.Add(mConfig_Com_Subcom(i))
            AddHandler mConfig_Com_Subcom(i).Click, AddressOf MenuItemClicked
        Next

        If COM_ARRAY.Count > 0 Then

            current_com_port = myHrv.HRV_GET_COM()
            If current_com_port.Length = 0 Then
#if EngVersion
            MsgBox("Note! The COM port does not configure well.", MsgBoxStyle.Exclamation)    
#else
MsgBox("注意: 系統未設定 COM Port.", MsgBoxStyle.Exclamation)
#end if
            Else
                'Console.WriteLine("current_com_port " & current_com_port)

                For i As Integer = 0 To COM_ARRAY.Count - 1
                    If (COM_ARRAY(i).ToString = current_com_port) Then
                        mConfig_Com_Subcom(i).Checked = True
                        check_flag = True
                    End If
                Next
                If check_flag = False Then
#if EngVersion
 MsgBox("Note! Please reconfigure the COM port.", MsgBoxStyle.Exclamation)
#else
MsgBox("注意: 請重新設定 COM Port.", MsgBoxStyle.Exclamation)
#end if

                    
                End If
            End If

        End If

        'ReDim m2_1_1(2)
        'm2_1_1(0) = New ToolStripMenuItem("COM3")
        'm2_1_1(1) = New ToolStripMenuItem("COM4")
        'm2_1_1(0).Name = "111"
        'm2_1_1(1).Name = "222"
        'm2_1.DropDownItems.Add(m2_1_1(0))
        'm2_1.DropDownItems.Add(m2_1_1(1))

        HrvMenu.Items.Add(mConfig)

        'Add a Click handler to the DropDown Item
        AddHandler HrvMenu.Click, AddressOf MenuItemClicked
        AddHandler mSystem.Click, AddressOf MenuItemClicked
        AddHandler mSystem_Login.Click, AddressOf MenuItemClicked
        AddHandler mSystem_End.Click, AddressOf MenuItemClicked

        AddHandler mConfig.Click, AddressOf MenuItemClicked
        AddHandler mConfig_Com.Click, AddressOf MenuItemClicked
        AddHandler mConfig_Backup.Click, AddressOf MenuItemClicked
        AddHandler mConfig_PWD.Click, AddressOf MenuItemClicked
        AddHandler mConfig_TESTER.Click, AddressOf MenuItemClicked
        AddHandler mConfig_COUNT.Click, AddressOf MenuItemClicked
        'AddHandler m2_1_1(0).Click, AddressOf MenuItemClicked
        'AddHandler m2_1_1(1).Click, AddressOf MenuItemClicked

        Me.Controls.Add(HrvMenu)

        'mConfig.DropDownItems.Remove(mConfig_Com)
        'mConfig.DropDownItems.Remove(mConfig_Com)
        'mConfig.DropDownItems.Remove(mConfig_Com)

        'gbTop.Enabled = False
        gbTop.Visible = False
        'txtPid.Select()
        cbCondition.SelectedIndex = 0

        Set_Default()

    End Sub

    Private Sub Set_Default()

       
#if EngVersion
        btnNew.Text = "New"      
#else
 btnNew.Text = "新增"
#end if
        btnNew.Visible = True
        btnDel.Visible = False
        btnEdit.Visible = False
        btnSave.Visible = False

        txtPid.Text = ""
        txtPid.Enabled = False
        txtName.Text = ""
        txtName.Enabled = False
        dtpBrith.Enabled = False
        dtpBrith.Value = Now().ToString("yyyy/MM/dd")
        rbM.Enabled = False
        rbF.Enabled = False
        rbM.Checked = False
        rbF.Checked = False

        btnStartTest.Enabled = False
        gb3.Visible = True

        btnGetReport.Enabled = False

    End Sub

    Private Sub MenuItemClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'ArrayList
        'http://stackoverflow.com/questions/10578993/add-new-value-to-integer-array-visual-basic-2010

        'MessageBox.Show(sender.ToString)
        'Console.WriteLine(sender.ToString)

        Try
            If (sender.ToString.Length > 3) Then
                'Console.WriteLine(sender.ToString.Substring(0, 3))
                If (sender.ToString.Substring(0, 3) = "COM" And sender.ToString <> "COM Port") Then
                    Dim match_port As Integer = 0
                    For i As Integer = 0 To (COM_ARRAY.Count - 1)
                        If (COM_ARRAY(i).ToString = sender.ToString) Then
                            match_port = i
                        End If
                        If mConfig_Com_Subcom(i).Checked Then
                            mConfig_Com_Subcom(i).Checked = False
                        End If
                    Next
                    If COM_ARRAY.Count > 0 Then
                        Console.WriteLine("DO DB -> " & sender.ToString)
                        current_com_port = sender.ToString
                        myHrv.HRV_SET_COM(current_com_port)
                        mConfig_Com_Subcom(match_port).Checked = True
                        check_flag = True
                    End If
                End If
            End If
#if EngVersion
           If sender.ToString = "Login" Then  
#else
If sender.ToString = "登入" Then
#end if
            
                'gbTop.Enabled = True
                'gbTop.Visible = True
                Me.Enabled = False
                frmLogin.Show()
                frmLogin.TopMost = True
            End If

            If sender.ToString = "登出" Then

                If MessageBox.Show("確定登出?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    'gbTop.Enabled = True
                    gbTop.Visible = False
                    mConfig.Visible = False
#if EngVersion
             mSystem_Login.Text = "Login"
#else
 mSystem_Login.Text = "登入"
#end if
                   
                End If

            End If
#if EngVersion
If sender.ToString = "Exit" Then
                If MessageBox.Show("Exit ?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    myHrv.HRV_Dispose()
                    End
                End If
            End If
             
#else
If sender.ToString = "結束" Then
                If MessageBox.Show("確定結束程式?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    myHrv.HRV_Dispose()
                    End
                End If
            End If
#end if
            
            If sender.ToString = "操作人員" Then
                Me.Enabled = False
                frmTester.Show()
                frmTester.TopMost = True
            End If

            If sender.ToString = "變更密碼" Then
                Me.Enabled = False
                frmPwd.Show()
                frmPwd.TopMost = True
            End If

            If sender.ToString = "資料備份" Then

                Try
                    'If BK_PATH = "" Then
                    Dim dialog As New FolderBrowserDialog()
                    dialog.RootFolder = Environment.SpecialFolder.Desktop
                    dialog.SelectedPath = "C:\"
                    dialog.Description = "請選擇資料備份路徑:"
                    If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        BK_PATH = dialog.SelectedPath
                        Console.WriteLine("BK_PATH " & BK_PATH)

                        myHrv.HRV_BACKUP(BK_PATH)
                        'My.Computer.FileSystem.WriteAllText(BKpath & "apppath.txt", BKpath, False)
                        'End If
                    Else

                    End If
                Catch ex As Exception
                    MsgBox("備份失敗 " & ex.ToString, MsgBoxStyle.Critical)
                End Try

                

            End If

            If sender.ToString = "檢測次數" Then
                Me.Enabled = False
                frmCount.Show()
                frmCount.TopMost = True
            End If
            
            'subSubItem2.Checked = True
            'm2_1_1(0).Checked = True
        Catch ex As Exception

        End Try
        


    End Sub

    Public Sub LoadUserRecord()

        Dim OleDBC As New OleDbCommand
        Dim OleDBDR As OleDbDataReader
        Dim c As Integer
        c = 0

        Try

            OleDBC = myHrv.HRV_Get_tUser(cbCondition.Text, txtKeyWord.Text)

            OleDBDR = OleDBC.ExecuteReader
            dgvData.Rows.Clear()
            If OleDBDR.HasRows Then
                While OleDBDR.Read
                    dgvData.Rows.Add()

                    dgvData.Item(0, c).Value = OleDBDR.Item(0)
                    dgvData.Item(1, c).Value = OleDBDR.Item(1)
                    dgvData.Item(2, c).Value = OleDBDR.Item(2)
                    Dim val As Date = OleDBDR.Item(3).ToString
                    dgvData.Item(3, c).Value = val.ToString("yyyy/MM/dd")
                    If OleDBDR.Item(4) = 1 Then
                        dgvData.Item(4, c).Value = "男"
                    Else
                        dgvData.Item(4, c).Value = "女"
                    End If

                    'dgvData.Item(5, c).Value = OleDBDR.Item(5)
                    'dgvData.Item(6, c).Value = OleDBDR.Item(6)
                    'dgvData.Item(7, c).Value = OleDBDR.Item(7)
                    'dgvData.Item(8, c).Value = OleDBDR.Item(8)
                    'dgvData.Item(9, c).Value = OleDBDR.Item(9)
                    'dgvData.Item(10, c).Value = OleDBDR.Item(10)
                    'btnGetReport.Text = "查詢紀錄 ( " & (dgvData.RowCount - 1) & " 筆 )"
                    c = c + 1
                End While
            Else
                'btnGetReport.Text = "查詢紀錄 ( 0 筆 )"
            End If

        Catch ex As Exception
            'MsgBox("LoadUserRecord " & ex.ToString, MsgBoxStyle.Critical)
        End Try

       
    End Sub

    Private Sub ReadMemoryMappedFile()

        Dim LookUpTable As String = "0123456789ABCDEF"
        Dim RXArray(1024) As Char
        Try
            Using file = MemoryMappedFile.OpenExisting(SHARE_NAME)
                Using reader = file.CreateViewAccessor(0, SHARE_LEN)
                    Dim bytes = New Byte(SHARE_LEN - 1) {}
                    reader.ReadArray(Of Byte)(0, bytes, 0, bytes.Length)

                    'TextBox1.Text = ""
                    'For i As Integer = 0 To bytes.Length - 1
                    '    'TextBox1.AppendText(CStr(bytes(i)) + " ")
                    '    TextBox1.AppendText(bytes(i).ToString + " ")
                    '    RXArray(0) = LookUpTable(bytes(i) >> 4) ' Convert each byte to two hexadecimal characters
                    '    RXArray(1) = LookUpTable(bytes(i) And 15)
                    'Next

                    'TextBox1.AppendText(vbCrLf)
                    'For i As Integer = 0 To bytes.Length - 1
                    '    RXArray(0) = LookUpTable(bytes(i) >> 4) ' Convert each byte to two hexadecimal characters
                    '    RXArray(1) = LookUpTable(bytes(i) And 15)
                    '    TextBox1.AppendText(RXArray(0) & RXArray(1) & " ")
                    'Next

                End Using
            End Using
        Catch noFile As FileNotFoundException
            MsgBox("錯誤: MMF " & SHARE_NAME & " 不存在.", MsgBoxStyle.Critical)
        Catch Ex As Exception
            MsgBox("ReadMemoryMappedFile Exception " + Ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        ReadMemoryMappedFile()

    End Sub

    Private Sub 結束ToolStripMenuItem_Click(sender As Object, e As EventArgs)

        If MessageBox.Show("確定結束程式?", "HRVapp", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            myHrv.HRV_Dispose()
            End
        End If

    End Sub

    Private Sub btnStartTest_Click(sender As Object, e As EventArgs) Handles btnStartTest.Click

        myHrv.HRV_SET_LICENSE()

        If myHrv.HRV_CHECK_LICENSE = True Then
            MsgBox("請更新應用程式.", MsgBoxStyle.Critical)
            Return
        End If

        myHrv.HRV_PROCESS_OUT()
        myHrv.HRV_CLEAN_FILES()

        myHrv.HRV_COM_STARTUP()

        'Process.Start(app_FilePath & "\HRV_COM.exe")
        'If myHrv.HRV_COM() Then
        frmDoHRV.Show()
        'frmDoHRV.TopMost = True
        tmr_logout.Enabled = False
        'Me.Enabled = False
        gbTop.Enabled = False
        btnStartTest.Text = "                 檢測中..."
        'btnStartTest.Enabled = False

        'End If

    End Sub

    Private Sub dgvData_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvData.CellMouseClick

        If e.RowIndex < 0 Then
            Exit Sub
        End If

        current_user_id = dgvData.Item(0, e.RowIndex).Value

        txtPid.Text = dgvData.Item(1, e.RowIndex).Value
        current_pid = dgvData.Item(1, e.RowIndex).Value

        txtName.Text = dgvData.Item(2, e.RowIndex).Value
        current_name = dgvData.Item(2, e.RowIndex).Value

        current_birth = dgvData.Item(3, e.RowIndex).Value
        dtpBrith.Text = dgvData.Item(3, e.RowIndex).Value

        If dgvData.Item(4, e.RowIndex).Value = "男" Then
            current_sex = "男"
            rbM.Checked = True
        Else
            current_sex = "女"
            rbF.Checked = True
        End If

        If check_flag = True Then
            btnStartTest.Enabled = True
        End If

        btnEdit.Text = "修改"
        btnDel.Text = "刪除"
        btnEdit.Visible = True
        btnDel.Visible = True
        btnGetReport.Enabled = True

    End Sub

    Private Sub btnGetReport_Click(sender As Object, e As EventArgs) Handles btnGetReport.Click

        Try
            If current_user_id < 0 Then
                MsgBox("請選擇檢測人", MsgBoxStyle.Critical)
                Exit Sub

            Else
                'Me.Enabled = False
                frmDataList.TopMost = True
                frmDataList.Show()
                'gbTop.Enabled = False
            End If

        Catch noFile As FileNotFoundException
            'MsgBox("1.", MsgBoxStyle.Critical)
        Catch Ex As Exception
            'MsgBox("2" + Ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click

        If btnNew.Text = "新增" Then

            btn_Status = "新增"

            txtPid.Enabled = True
            txtPid.Text = ""
            txtPid.Select()

            txtName.Enabled = True
            txtName.Text = ""
            dtpBrith.Enabled = True
            rbM.Enabled = True
            rbF.Enabled = True
            btnNew.Text = "取消"
            'btnDel.Enabled = False
            'btnEdit.Enabled = False
            btnSave.Visible = True
            btnStartTest.Enabled = False
            btnEdit.Visible = False
            btnDel.Visible = False
            'dgvData.Enabled = False
            gb3.Visible = False
        Else
            btnNew.Text = "新增"
            Set_Default()
        End If

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        If btnEdit.Text = "修改" Then

            btn_Status = "修改"

            txtPid.Enabled = True
            txtPid.Select()
            txtName.Enabled = True
            dtpBrith.Enabled = True
            rbM.Enabled = True
            rbF.Enabled = True
            btnEdit.Text = "取消"

            btnNew.Visible = False
            btnDel.Visible = False
            btnSave.Visible = True
            btnStartTest.Enabled = False

            gb3.Visible = False
        Else
            btnEdit.Text = "修改"
            Set_Default()
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Try

            Console.WriteLine(btn_Status)
            If txtPid.Text = "" Then
                MsgBox("錯誤 : [編號] 欄位請勿空白", MsgBoxStyle.Critical)
                txtPid.Select()
                Exit Sub
            End If
            If txtName.Text = "" Then
                MsgBox("錯誤 : [姓名] 欄位請勿空白", MsgBoxStyle.Critical)
                txtName.Select()
                Exit Sub
            End If
            If rbM.Checked = False And rbF.Checked = False Then
                MsgBox("錯誤 : 請選擇 [性別]", MsgBoxStyle.Critical)
                Exit Sub
            End If

            Dim tmp_sex As Integer
            If rbM.Checked = True Then
                tmp_sex = 1
            Else
                tmp_sex = 0
            End If

            If btn_Status = "新增" Then
                'Console.WriteLine(dtpBrith.Value.ToShortDateString)
                If myHrv.HRV_NEW_USER(txtPid.Text, txtName.Text, dtpBrith.Value.ToShortDateString, tmp_sex, "編號") = True Then
                    MsgBox("資料已新增", MsgBoxStyle.Information)
                    LoadUserRecord()
                    Set_Default()
                End If
                txtPid.Select()
            End If

            If btn_Status = "修改" Then
                'Console.WriteLine(dtpBrith.Value.ToShortDateString)
                Console.WriteLine("刪除 " & current_user_id)
                If myHrv.HRV_UPDATE_USER(current_user_id, txtPid.Text, txtName.Text, dtpBrith.Value.ToShortDateString, tmp_sex) = True Then
                    MsgBox("資料已修改", MsgBoxStyle.Information)
                    LoadUserRecord()
                    Set_Default()
                End If
            End If

        Catch Ex As Exception
            MsgBox(btn_Status & "錯誤 : " + Ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click

        Console.WriteLine("刪除 " & current_user_id)

        Dim result As Integer = MessageBox.Show("確定刪除 " & txtPid.Text & " ?", HRVapp, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            If myHrv.HRV_DEL_USER(current_user_id) = True Then
                MsgBox("資料已刪除", MsgBoxStyle.Information)
                LoadUserRecord()
                Set_Default()
            End If
        End If

    End Sub

    Private Sub tmr_logout_Tick(sender As Object, e As EventArgs) Handles tmr_logout.Tick

        logout_timer = logout_timer + 1
        'Console.WriteLine(logout_timer)
        If (logout_timer > logout_timer_max) Then
            gbTop.Visible = False
            mSystem_Login.Text = "登入"
            mConfig.Enabled = False
            tmr_logout.Enabled = False
        End If

    End Sub

    Private Sub gbTop_MouseMove(sender As Object, e As MouseEventArgs) Handles gbTop.MouseMove

        logout_timer = 0

    End Sub

    Private Sub GroupBox2_MouseMove(sender As Object, e As MouseEventArgs) Handles GroupBox2.MouseMove

        logout_timer = 0

    End Sub

    Private Sub gb3_MouseMove(sender As Object, e As MouseEventArgs) Handles gb3.MouseMove

        logout_timer = 0

    End Sub

    Private Sub txtKeyWord_TextChanged(sender As Object, e As EventArgs) Handles txtKeyWord.TextChanged

        LoadUserRecord()

    End Sub

  
    Private Sub cbCondition_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCondition.SelectedIndexChanged

        If cbCondition.Text = "生日" Then
            txtKeyWord.Text = ""
            txtKeyWord.MaxLength = 8 'yyyymmdd
        Else
            txtKeyWord.Text = ""
            txtKeyWord.MaxLength = 0
        End If

        txtKeyWord.Focus()

    End Sub

End Class


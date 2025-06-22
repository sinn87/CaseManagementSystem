' 登录窗体
Imports System.Data.OleDb

Public Class LoginForm
    ' 控件声明
    Private lblTitle As Label
    Private lblStaffNo As Label
    Private txtStaffNo As TextBox
    Private lblPassword As Label
    Private txtPassword As TextBox
    Private btnLogin As Button
    Private btnExit As Button
    Private linkChangePwd As LinkLabel
    Private lblError As Label

    Public Sub New()
        ' 初始化界面
        Me.Text = "系统登录"
        Me.Width = 400
        Me.Height = 320
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        InitControls()
    End Sub

    Private Sub InitControls()
        lblTitle = New Label With {.Text = "系统登录", .Font = New Font("微软雅黑", 16, FontStyle.Bold), .AutoSize = True, .Location = New Point(140, 20)}
        lblStaffNo = New Label With {.Text = "员工番号", .Location = New Point(50, 70), .AutoSize = True}
        txtStaffNo = New TextBox With {.Location = New Point(130, 65), .Width = 180}
        lblPassword = New Label With {.Text = "密码", .Location = New Point(50, 110), .AutoSize = True}
        txtPassword = New TextBox With {.Location = New Point(130, 105), .Width = 180, .PasswordChar = "*"c}
        btnLogin = New Button With {.Text = "登录", .Location = New Point(80, 160), .Width = 90}
        btnExit = New Button With {.Text = "结束", .Location = New Point(200, 160), .Width = 90}
        linkChangePwd = New LinkLabel With {.Text = "密码变更", .Location = New Point(280, 200), .AutoSize = True}
        lblError = New Label With {.Text = "", .ForeColor = Color.Red, .Location = New Point(80, 200), .AutoSize = True, .Visible = False}

        AddHandler btnLogin.Click, AddressOf btnLogin_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click
        AddHandler linkChangePwd.Click, AddressOf linkChangePwd_Click

        Me.Controls.Add(lblTitle)
        Me.Controls.Add(lblStaffNo)
        Me.Controls.Add(txtStaffNo)
        Me.Controls.Add(lblPassword)
        Me.Controls.Add(txtPassword)
        Me.Controls.Add(btnLogin)
        Me.Controls.Add(btnExit)
        Me.Controls.Add(linkChangePwd)
        Me.Controls.Add(lblError)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        lblError.Visible = False
        Dim staffNo As String = txtStaffNo.Text.Trim()
        Dim password As String = txtPassword.Text
        If staffNo = "" Or password = "" Then
            lblError.Text = "员工番号和密码均为必填项"
            lblError.Visible = True
            Return
        End If
        ' 调用数据层
        Dim repo As New DataAccess.UserRepository()
        Dim result As Integer = repo.CheckLogin(staffNo, password)
        Select Case result
            Case 1
                lblError.Visible = False
                MessageBox.Show("登录成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 跳转主页面
                Me.Hide()
                Dim mainForm As New MainForm(staffNo)
                mainForm.ShowDialog()
                Me.Close()
            Case 2
                lblError.Text = "密码不正确"
                lblError.Visible = True
            Case 3
                lblError.Text = "账号已被禁用"
                lblError.Visible = True
            Case 0
                lblError.Text = "员工番号不存在"
                lblError.Visible = True
            Case -1
                lblError.Text = "数据库连接失败"
                lblError.Visible = True
        End Select
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Application.Exit()
    End Sub

    Private Sub linkChangePwd_Click(sender As Object, e As EventArgs)
        MessageBox.Show("请实现密码变更功能。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ' TODO: 跳转到密码变更窗体
    End Sub
End Class 
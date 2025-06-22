' 主窗体
Public Class MainForm
    Private btnNewCase As Button
    Private btnCaseList As Button
    Private btnDailyWork As Button
    Private btnUserManage As Button
    Private btnLogout As Button
    Private _currentUser As String

    Public Sub New(currentUser As String)
        _currentUser = currentUser
        Me.Text = "主页面"
        Me.Width = 500
        Me.Height = 350
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        InitControls()
    End Sub

    Private Sub InitControls()
        btnNewCase = New Button With {.Text = "登录新案件", .Width = 150, .Height = 50, .Location = New Point(50, 60)}
        btnCaseList = New Button With {.Text = "案件一览", .Width = 150, .Height = 50, .Location = New Point(270, 60)}
        btnDailyWork = New Button With {.Text = "日常工作", .Width = 150, .Height = 50, .Location = New Point(50, 140)}
        btnUserManage = New Button With {.Text = "用户管理", .Width = 150, .Height = 50, .Location = New Point(270, 140)}
        btnLogout = New Button With {.Text = "退出", .Width = 80, .Height = 35, .Location = New Point(Me.ClientSize.Width - 100, Me.ClientSize.Height - 55), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right}

        AddHandler btnNewCase.Click, AddressOf btnNewCase_Click
        AddHandler btnCaseList.Click, AddressOf btnCaseList_Click
        AddHandler btnDailyWork.Click, AddressOf btnDailyWork_Click
        AddHandler btnUserManage.Click, AddressOf btnUserManage_Click
        AddHandler btnLogout.Click, AddressOf btnLogout_Click

        Me.Controls.Add(btnNewCase)
        Me.Controls.Add(btnCaseList)
        Me.Controls.Add(btnDailyWork)
        Me.Controls.Add(btnUserManage)
        Me.Controls.Add(btnLogout)
    End Sub

    Private Sub btnNewCase_Click(sender As Object, e As EventArgs)
        Try
            ' 跳转到案件类型选择页面
            Me.Hide()
            Dim typeSelectForm As New CaseTypeSelectForm(_currentUser)
            Dim result As DialogResult = typeSelectForm.ShowDialog()
            
            ' 如果用户完成了案件录入，重新显示主页面
            If result = DialogResult.OK Then
                Me.Show()
            Else
                ' 用户取消，关闭应用程序
                Me.Close()
            End If
            
        Catch ex As Exception
            MessageBox.Show($"跳转到案件类型选择页面时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("主页面跳转失败", ex)
        End Try
    End Sub
    
    Private Sub btnCaseList_Click(sender As Object, e As EventArgs)
        Try
            ' 跳转到案件一览页面
            Me.Hide()
            Dim caseListForm As New CaseListForm(_currentUser)
            caseListForm.ShowDialog()
            Me.Show()
            
        Catch ex As Exception
            MessageBox.Show($"跳转到案件一览页面时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("主页面跳转失败", ex)
        End Try
    End Sub
    
    Private Sub btnDailyWork_Click(sender As Object, e As EventArgs)
        MessageBox.Show("日常工作功能待实现。", "提示")
    End Sub
    
    Private Sub btnUserManage_Click(sender As Object, e As EventArgs)
        MessageBox.Show("用户管理功能待实现。", "提示")
    End Sub
    
    Private Sub btnLogout_Click(sender As Object, e As EventArgs)
        ' 退出登录，返回登录页面
        Me.Hide()
        Dim loginForm As New LoginForm()
        loginForm.ShowDialog()
        Me.Close()
    End Sub
End Class 
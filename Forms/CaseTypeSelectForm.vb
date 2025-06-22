' 案件类型选择窗体
Public Class CaseTypeSelectForm
    Private btnPO As Button
    Private btnIPO As Button
    Private btnRPO As Button
    Private btnRIPO As Button
    Private _currentUser As String

    Public Sub New(currentUser As String)
        _currentUser = currentUser
        Me.Text = "案件类型选择"
        Me.Width = 400
        Me.Height = 300
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        InitControls()
    End Sub

    Private Sub InitControls()
        btnPO = New Button With {.Text = "PO", .Width = 120, .Height = 50, .Location = New Point(40, 60)}
        btnIPO = New Button With {.Text = "IPO", .Width = 120, .Height = 50, .Location = New Point(220, 60)}
        btnRPO = New Button With {.Text = "RPO", .Width = 120, .Height = 50, .Location = New Point(40, 140)}
        btnRIPO = New Button With {.Text = "RIPO", .Width = 120, .Height = 50, .Location = New Point(220, 140)}

        AddHandler btnPO.Click, AddressOf btnType_Click
        AddHandler btnIPO.Click, AddressOf btnType_Click
        AddHandler btnRPO.Click, AddressOf btnType_Click
        AddHandler btnRIPO.Click, AddressOf btnType_Click

        Me.Controls.Add(btnPO)
        Me.Controls.Add(btnIPO)
        Me.Controls.Add(btnRPO)
        Me.Controls.Add(btnRIPO)
    End Sub

    Private Sub btnType_Click(sender As Object, e As EventArgs)
        Try
            Dim btn As Button = CType(sender, Button)
            Dim caseType As String = btn.Text
            
            ' 跳转到案件详细录入页面
            Me.Hide()
            Dim detailForm As New CaseDetailEntryForm(caseType, _currentUser)
            Dim result As DialogResult = detailForm.ShowDialog()
            
            ' 根据详细录入页面的结果决定是否关闭类型选择页面
            If result = DialogResult.OK Then
                ' 保存成功，返回主页面
                Me.DialogResult = DialogResult.OK
            Else
                ' 用户取消或返回，重新显示类型选择页面
                Me.Show()
            End If
            
        Catch ex As Exception
            MessageBox.Show($"跳转到详细录入页面时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("案件类型选择跳转失败", ex)
        End Try
    End Sub
End Class 
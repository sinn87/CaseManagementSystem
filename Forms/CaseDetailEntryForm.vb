' 案件详细录入窗体
Imports System.Windows.Forms

Public Class CaseDetailEntryForm
    Private picHome As PictureBox
    Private lblType As Label
    Private lblTime As Label
    Private btnBack As Button
    Private btnSubmit As Button
    Private tabControl As TabControl
    Private _caseType As String
    Private _currentUser As String

    Public Sub New(caseType As String, currentUser As String)
        _caseType = caseType
        _currentUser = currentUser
        Me.Text = "案件详细录入"
        Me.Width = 1100
        Me.Height = 700
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        InitControls(caseType)
    End Sub

    Private Sub InitControls(caseType As String)
        ' 顶部图片按钮
        picHome = New PictureBox With {.ImageLocation = "https://cdn-icons-png.flaticon.com/512/25/25694.png", .Size = New Size(32, 32), .Location = New Point(10, 10), .Cursor = Cursors.Hand}
        AddHandler picHome.Click, AddressOf picHome_Click
        Me.Controls.Add(picHome)

        ' 类型显示
        lblType = New Label With {.Text = $"类型：{caseType}", .Location = New Point(60, 15), .Font = New Font("微软雅黑", 12, FontStyle.Bold), .AutoSize = True}
        Me.Controls.Add(lblType)

        ' 时间显示
        lblTime = New Label With {.Text = $"时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}", .Location = New Point(300, 15), .Font = New Font("微软雅黑", 10), .AutoSize = True}
        Me.Controls.Add(lblTime)

        ' 右上角返回按钮
        btnBack = New Button With {.Text = "返回", .Location = New Point(Me.Width - 130, 15), .Width = 80, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnBack.Click, AddressOf btnBack_Click
        Me.Controls.Add(btnBack)

        ' 登录新案件按钮
        btnSubmit = New Button With {.Text = "登录新案件", .Location = New Point(Me.Width - 250, 15), .Width = 100, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnSubmit.Click, AddressOf btnSubmit_Click
        Me.Controls.Add(btnSubmit)

        ' TabControl
        tabControl = New TabControl With {.Location = New Point(20, 60), .Size = New Size(1040, 580)}
        CreateTabPages()
        Me.Controls.Add(tabControl)
    End Sub

    Private Sub CreateTabPages()
        For i = 1 To 9
            Dim tabPage As New TabPage($"信息页{i}")
            ' 注意：实际项目中建议使用设计器创建控件
            ' 这里仅作为示例，控件生成代码已移动到 Controls/CodeGenerated/ 文件夹
            CreateSampleControls(tabPage, i)
            tabControl.TabPages.Add(tabPage)
            
            ' 自动调整TabPage大小
            Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, Me)
        Next
    End Sub

    Private Sub CreateSampleControls(tabPage As TabPage, pageIndex As Integer)
        ' 创建示例控件，实际开发中建议使用设计器
        ' 控件生成代码参考：Controls/CodeGenerated/TabPageSizeAdjuster_Generated.vb
        Dim y As Integer = 20
        
        ' 创建一些示例控件并设置Tag属性对应数据库字段
        For j = 1 To 5
            Dim lbl As New Label With {
                .Text = $"字段{j}:",
                .Location = New Point(20, y),
                .AutoSize = True
            }
            tabPage.Controls.Add(lbl)
            
            Dim txt As New TextBox With {
                .Location = New Point(120, y),
                .Width = 200,
                .Tag = $"Field_{pageIndex}_{j}" ' Tag对应数据库字段
            }
            tabPage.Controls.Add(txt)
            
            y += 30
        Next
        
        ' 添加一些其他类型的控件示例
        Dim cbo As New ComboBox With {
            .Location = New Point(20, y),
            .Width = 200,
            .Tag = $"Combo_{pageIndex}",
            .Items = {"选项1", "选项2", "选项3"}
        }
        tabPage.Controls.Add(cbo)
        
        y += 30
        
        Dim chk As New CheckBox With {
            .Text = "复选框",
            .Location = New Point(20, y),
            .Tag = $"Check_{pageIndex}"
        }
        tabPage.Controls.Add(chk)
        
        y += 30
        
        Dim dtp As New DateTimePicker With {
            .Location = New Point(20, y),
            .Width = 200,
            .Tag = $"Date_{pageIndex}"
        }
        tabPage.Controls.Add(dtp)
        
        y += 30
        
        ' 添加一个RichTextBox示例
        Dim rtb As New RichTextBox With {
            .Location = New Point(20, y),
            .Width = 400,
            .Height = 100,
            .Tag = $"RichText_{pageIndex}"
        }
        tabPage.Controls.Add(rtb)
        
        y += 120
        
        ' 添加DataGridView示例
        Dim dgv As New DataGridView With {
            .Name = $"dgvItems_{pageIndex}",
            .Location = New Point(450, 20),
            .Size = New Size(550, 180),
            .AllowUserToAddRows = True,
            .AllowUserToDeleteRows = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .Tag = $"Grid_{pageIndex}" ' Tag标识这是一个DataGridView
        }
        
        ' 添加列
        dgv.Columns.Add("ItemName", "项目名称")
        dgv.Columns.Add("ItemValue", "项目值")
        dgv.Columns.Add("LastUpdate", "最后更新时间")
        dgv.Columns.Add("ReviewTime", "审查时间")
        dgv.Columns.Add("Status", "状态")
        dgv.Columns.Add("Reviewer", "审查人员")
        
        ' 设置默认值
        dgv.Columns("LastUpdate").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
        dgv.Columns("ReviewTime").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
        dgv.Columns("Status").DefaultCellStyle.NullValue = "新登录"
        dgv.Columns("Reviewer").DefaultCellStyle.NullValue = _currentUser
        
        tabPage.Controls.Add(dgv)
        
        ' 自动调整TabPage大小
        Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, Me)
    End Sub

    Private Sub picHome_Click(sender As Object, e As EventArgs)
        ' 返回主页面
        Me.Hide()
        Dim mainForm As New MainForm(_currentUser)
        mainForm.ShowDialog()
        Me.Close()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs)
        ' 返回类型选择
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs)
        Try
            ' 禁用按钮防止重复提交
            btnSubmit.Enabled = False
            btnSubmit.Text = "保存中..."
            
            ' 调用业务逻辑层提取数据
            Dim tabData As Dictionary(Of Integer, Dictionary(Of String, String)) = BusinessLogic.CaseManager.ExtractModifiedData(tabControl)
            Dim tabGridData As Dictionary(Of Integer, List(Of Dictionary(Of String, String))) = BusinessLogic.CaseManager.ExtractGridData(tabControl)
            
            ' 检查是否有数据需要保存
            If tabData.Count = 0 AndAlso tabGridData.Count = 0 Then
                MessageBox.Show("没有检测到需要保存的数据，请至少在一个标签页中输入信息。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            
            ' 调用业务逻辑层保存数据
            Dim success As Boolean = BusinessLogic.CaseManager.CreateNewCase(_caseType, tabData, tabGridData, _currentUser)
            
            If success Then
                MessageBox.Show("案件数据保存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 保存成功后返回主页面
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("保存失败，请检查数据后重试。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"保存过程中发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("案件详细录入保存失败", ex)
        Finally
            ' 恢复按钮状态
            btnSubmit.Enabled = True
            btnSubmit.Text = "登录新案件"
        End Try
    End Sub
End Class 
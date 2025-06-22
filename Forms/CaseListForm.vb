' 案件一览窗体
Imports System.Windows.Forms

Public Class CaseListForm
    Private searchBoxes(10) As TextBox
    Private btnSearch As Button
    Private dgvCases As DataGridView
    Private btnBack As Button
    Private picHome As PictureBox
    Private _currentUser As String
    Private _cases As List(Of CaseInfo)

    Public Sub New(currentUser As String)
        _currentUser = currentUser
        Me.Text = "案件一览"
        Me.Width = 1200
        Me.Height = 700
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadCases()
        InitControls()
    End Sub

    Private Sub LoadCases()
        Try
            _cases = CaseRepository.GetAllCases()
        Catch ex As Exception
            MessageBox.Show($"加载案件数据失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("加载案件数据失败", ex)
            _cases = New List(Of CaseInfo)()
        End Try
    End Sub

    Private Sub InitControls()
        ' 顶部图片按钮（返回主页面）
        picHome = New PictureBox With {.ImageLocation = "https://cdn-icons-png.flaticon.com/512/25/25694.png", .Size = New Size(32, 32), .Location = New Point(10, 10), .Cursor = Cursors.Hand}
        AddHandler picHome.Click, AddressOf picHome_Click
        Me.Controls.Add(picHome)

        ' 右上角返回按钮
        btnBack = New Button With {.Text = "返回", .Location = New Point(Me.Width - 110, 15), .Width = 80, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnBack.Click, AddressOf btnBack_Click
        Me.Controls.Add(btnBack)

        ' 搜索区域
        CreateSearchArea()

        ' DataGridView
        CreateDataGridView()
        
        ' 加载数据
        RefreshDataGridView()
    End Sub

    Private Sub CreateSearchArea()
        ' 搜索标签和输入框
        Dim labels() As String = {"案件类型", "案件名称", "产品代码", "产品名称", "状态", "创建用户", "开始日期", "结束日期", "公司角色", "发布状态", "上市状态"}
        
        For i = 0 To 10
            Dim lbl As New Label With {
                .Text = labels(i), 
                .Location = New Point(30 + i Mod 4 * 280, 60 + i \ 4 * 35), 
                .AutoSize = True,
                .Font = New Font("微软雅黑", 9)
            }
            
            Dim txt As New TextBox With {
                .Location = New Point(100 + i Mod 4 * 280, 55 + i \ 4 * 35), 
                .Width = 150,
                .Font = New Font("微软雅黑", 9)
            }
            
            searchBoxes(i) = txt
            Me.Controls.Add(lbl)
            Me.Controls.Add(txt)
        Next

        ' 搜索按钮
        btnSearch = New Button With {
            .Text = "搜索", 
            .Location = New Point(1000, 135), 
            .Width = 80,
            .Height = 30,
            .Font = New Font("微软雅黑", 9)
        }
        AddHandler btnSearch.Click, AddressOf btnSearch_Click
        Me.Controls.Add(btnSearch)
        
        ' 重置按钮
        Dim btnReset As New Button With {
            .Text = "重置", 
            .Location = New Point(1090, 135), 
            .Width = 80,
            .Height = 30,
            .Font = New Font("微软雅黑", 9)
        }
        AddHandler btnReset.Click, AddressOf btnReset_Click
        Me.Controls.Add(btnReset)
    End Sub

    Private Sub CreateDataGridView()
        dgvCases = New DataGridView With {
            .Location = New Point(30, 180), 
            .Width = 1140, 
            .Height = 450, 
            .ReadOnly = True, 
            .AllowUserToAddRows = False, 
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .Font = New Font("微软雅黑", 9),
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False,
            .AllowUserToResizeRows = False
        }
        
        ' 添加列
        dgvCases.Columns.Add("CaseID", "案件ID")
        dgvCases.Columns.Add("CaseType", "案件类型")
        dgvCases.Columns.Add("CaseName", "案件名称")
        dgvCases.Columns.Add("ProductCode", "产品代码")
        dgvCases.Columns.Add("ProductName", "产品名称")
        dgvCases.Columns.Add("Status", "状态")
        dgvCases.Columns.Add("CreateUser", "创建用户")
        dgvCases.Columns.Add("CreateTime", "创建时间")
        dgvCases.Columns.Add("LastUpdate", "最后更新")
        dgvCases.Columns.Add("IsTerminated", "是否中止")
        
        ' 设置列宽
        dgvCases.Columns("CaseID").Width = 80
        dgvCases.Columns("CaseType").Width = 100
        dgvCases.Columns("CaseName").Width = 150
        dgvCases.Columns("ProductCode").Width = 100
        dgvCases.Columns("ProductName").Width = 150
        dgvCases.Columns("Status").Width = 80
        dgvCases.Columns("CreateUser").Width = 100
        dgvCases.Columns("CreateTime").Width = 120
        dgvCases.Columns("LastUpdate").Width = 120
        dgvCases.Columns("IsTerminated").Width = 80
        
        AddHandler dgvCases.CellDoubleClick, AddressOf dgvCases_CellDoubleClick
        Me.Controls.Add(dgvCases)
    End Sub

    Private Sub RefreshDataGridView()
        dgvCases.Rows.Clear()
        
        For Each caseInfo In _cases
            Dim statusText As String = GetStatusText(caseInfo.Status)
            Dim terminatedText As String = If(caseInfo.IsTerminated = 1, "是", "否")
            
            Dim rowIndex As Integer = dgvCases.Rows.Add(
                caseInfo.CaseID,
                caseInfo.CaseType,
                If(caseInfo.CaseName, ""),
                If(caseInfo.ProductCode, ""),
                If(caseInfo.ProductName, ""),
                statusText,
                caseInfo.CreateUser,
                caseInfo.CreateTime.ToString("yyyy-MM-dd HH:mm"),
                caseInfo.LastUpdate.ToString("yyyy-MM-dd HH:mm"),
                terminatedText
            )
            
            ' 如果案件已中止，设置行背景色为红色
            If caseInfo.IsTerminated = 1 Then
                dgvCases.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightCoral
            End If
        Next
    End Sub

    Private Function GetStatusText(status As Integer) As String
        Select Case status
            Case 1
                Return "新登录"
            Case 2
                Return "待审查"
            Case 3
                Return "审查完成"
            Case 4
                Return "需要修改"
            Case 5
                Return "终了"
            Case Else
                Return "未知"
        End Select
    End Function

    Private Sub picHome_Click(sender As Object, e As EventArgs)
        ' 返回主页面
        Me.Hide()
        Dim mainForm As New MainForm(_currentUser)
        mainForm.ShowDialog()
        Me.Close()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs)
        ' 返回上一页
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs)
        Try
            ' 获取搜索条件
            Dim caseType As String = searchBoxes(0).Text.Trim()
            Dim caseName As String = searchBoxes(1).Text.Trim()
            Dim productCode As String = searchBoxes(2).Text.Trim()
            Dim productName As String = searchBoxes(3).Text.Trim()
            Dim statusText As String = searchBoxes(4).Text.Trim()
            Dim createUser As String = searchBoxes(5).Text.Trim()
            
            ' 转换状态文本为数字
            Dim status As Integer? = Nothing
            If Not String.IsNullOrEmpty(statusText) Then
                status = ConvertStatusTextToNumber(statusText)
            End If
            
            ' 执行搜索
            _cases = CaseRepository.SearchCases(caseType, caseName, productCode, status)
            
            ' 刷新显示
            RefreshDataGridView()
            
            MessageBox.Show($"搜索完成，共找到 {_cases.Count} 条记录。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"搜索失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("案件搜索失败", ex)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs)
        ' 清空搜索条件
        For i = 0 To 10
            searchBoxes(i).Text = ""
        Next
        
        ' 重新加载所有数据
        LoadCases()
        RefreshDataGridView()
    End Sub

    Private Sub dgvCases_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.RowIndex < dgvCases.Rows.Count Then
            Try
                ' 获取选中的案件ID
                Dim caseId As Integer = Convert.ToInt32(dgvCases.Rows(e.RowIndex).Cells("CaseID").Value)
                
                ' 打开案件详细页面
                Me.Hide()
                Dim detailForm As New CaseDetailForm(caseId, _currentUser)
                detailForm.ShowDialog()
                Me.Show()
                
                ' 重新加载数据（可能案件状态有变化）
                LoadCases()
                RefreshDataGridView()
                
            Catch ex As Exception
                MessageBox.Show($"打开案件详细页面失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Utils.LogUtil.LogError("打开案件详细页面失败", ex)
            End Try
        End If
    End Sub

    Private Function ConvertStatusTextToNumber(statusText As String) As Integer?
        Select Case statusText.ToLower()
            Case "新登录", "1"
                Return 1
            Case "待审查", "2"
                Return 2
            Case "审查完成", "3"
                Return 3
            Case "需要修改", "4"
                Return 4
            Case "终了", "5"
                Return 5
            Case Else
                Return Nothing
        End Select
    End Function
End Class 
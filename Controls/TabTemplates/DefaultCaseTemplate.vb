''' <summary>
''' 默认案件Tab模板 - 适用于通用案件类型
''' </summary>
Imports System.Data

Public Class DefaultCaseTemplate
    Inherits BaseTabTemplate
    
    Public Sub New(tabControl As TabControl)
        MyBase.New(tabControl)
        _tabNames = {"基本信息", "案件详情", "相关文件", "处理记录", "备注信息", "履历信息"}
    End Sub
    
    Public Overrides Sub CreateTabPages(tabControl As TabControl)
        ' 清空现有标签页
        tabControl.TabPages.Clear()
        
        ' 创建默认案件专用的标签页
        For i As Integer = 0 To _tabNames.Length - 1
            Dim tabPage As New TabPage(_tabNames(i))
            
            ' 添加Tab顶部状态标签
            Dim statusLabel As New Label With {
                .Text = "未提交",
                .Location = New Point(10, 5),
                .Font = New Font("微软雅黑", 9),
                .ForeColor = Color.Gray,
                .AutoSize = True,
                .Name = $"lblStatus_{i}"
            }
            tabPage.Controls.Add(statusLabel)
            
            ' 创建Tab内容
            CreateDefaultTabContent(tabPage, i)
            tabControl.TabPages.Add(tabPage)
        Next
    End Sub
    
    Private Sub CreateDefaultTabContent(tabPage As TabPage, tabIndex As Integer)
        Dim y As Integer = 40
        
        ' 根据Tab索引创建不同的控件
        Select Case tabIndex
            Case 0 ' 基本信息
                CreateBasicInfoControls(tabPage, y)
            Case 1 ' 案件详情
                CreateCaseDetailControls(tabPage, y)
            Case 2 ' 相关文件
                CreateFileControls(tabPage, y)
            Case 3 ' 处理记录
                CreateProcessControls(tabPage, y)
            Case 4 ' 备注信息
                CreateMemoControls(tabPage, y)
            Case 5 ' 履历信息
                CreateHistoryControls(tabPage, y)
        End Select
    End Sub
    
    Private Sub CreateBasicInfoControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("案件编号", "CaseNo", New TextBox),
            ("案件名称", "CaseName", New TextBox),
            ("案件类型", "CaseType", New ComboBox),
            ("申请日期", "ApplyDate", New DateTimePicker),
            ("申请人", "Applicant", New TextBox),
            ("联系电话", "Phone", New TextBox),
            ("电子邮箱", "Email", New TextBox),
            ("案件状态", "CaseStatus", New ComboBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateCaseDetailControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("案件描述", "CaseDescription", New RichTextBox),
            ("处理要求", "ProcessRequirement", New RichTextBox),
            ("相关法规", "RelatedRegulations", New TextBox),
            ("处理期限", "ProcessDeadline", New DateTimePicker)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateFileControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("文件清单", "FileList", New RichTextBox),
            ("文件路径", "FilePath", New TextBox),
            ("文件状态", "FileStatus", New ComboBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateProcessControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("处理步骤", "ProcessSteps", New RichTextBox),
            ("处理结果", "ProcessResult", New ComboBox),
            ("处理意见", "ProcessOpinion", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
        
        ' 添加DataGridView示例（如果需要）
        ' 注意：实际的DataGridView应该在您的自定义控件中创建
        ' 这里只是示例，展示如何确保DataGridView有正确的DataSource
        ' CreateSampleDataGridView(tabPage, y + 150)
    End Sub
    
    ''' <summary>
    ''' 创建示例DataGridView（仅供参考，实际使用时应删除）
    ''' </summary>
    ''' <param name="tabPage">标签页</param>
    ''' <param name="y">Y坐标</param>
    Private Sub CreateSampleDataGridView(tabPage As TabPage, y As Integer)
        ' 创建标签
        Dim lblDgv As New Label With {
            .Text = "处理记录列表:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lblDgv)
        
        ' 创建DataGridView
        Dim dgv As New DataGridView With {
            .Location = New Point(20, y + 25),
            .Size = New Size(600, 200),
            .Name = "dgvProcessRecords",
            .Tag = "ProcessRecords",
            .AllowUserToAddRows = True,
            .AllowUserToDeleteRows = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        }
        
        ' 创建DataTable并设置列
        Dim dt As New DataTable()
        dt.Columns.Add("记录编号", GetType(String))
        dt.Columns.Add("处理日期", GetType(String))
        dt.Columns.Add("处理内容", GetType(String))
        dt.Columns.Add("处理人员", GetType(String))
        dt.Columns.Add("处理结果", GetType(String))
        
        ' 设置DataSource
        dgv.DataSource = dt
        
        tabPage.Controls.Add(dgv)
    End Sub
    
    Private Sub CreateMemoControls(tabPage As TabPage, y As Integer)
        Dim memoBox As New RichTextBox With {
            .Location = New Point(20, y),
            .Size = New Size(800, 400),
            .Tag = "Memo",
            .Font = New Font("微软雅黑", 10)
        }
        tabPage.Controls.Add(memoBox)
    End Sub
    
    Private Sub CreateHistoryControls(tabPage As TabPage, y As Integer)
        Dim historyBox As New RichTextBox With {
            .Location = New Point(20, y),
            .Size = New Size(800, 400),
            .Tag = "History",
            .Font = New Font("微软雅黑", 10),
            .ReadOnly = True
        }
        tabPage.Controls.Add(historyBox)
    End Sub
    
    Public Overrides Function GetSupportedCaseTypes() As List(Of String)
        Return New List(Of String) From {"通用案件", "其他案件", "未分类案件"}
    End Function
    
    Public Overrides Function GetTemplateName() As String
        Return "默认案件模板"
    End Function
End Class 
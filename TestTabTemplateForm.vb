''' <summary>
''' Tab模板测试窗体 - 用于测试不同案件类型的模板切换功能
''' </summary>
Public Class TestTabTemplateForm
    Inherits Form
    
    Private tabControl As TabControl
    Private cmbCaseType As ComboBox
    Private btnCreateTemplate As Button
    Private btnLoadData As Button
    Private btnSaveData As Button
    Private btnSetReadOnly As Button
    Private btnSetStyle As Button
    Private lblInfo As Label
    
    Private _currentTemplate As ITabTemplate
    Private _testData As List(Of CaseDetail)
    
    Public Sub New()
        Me.Text = "Tab模板测试窗体"
        Me.Width = 1200
        Me.Height = 800
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        
        InitControls()
        InitTestData()
    End Sub
    
    Private Sub InitControls()
        ' 案件类型选择
        Dim lblCaseType As New Label With {
            .Text = "案件类型:",
            .Location = New Point(20, 20),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 10)
        }
        Me.Controls.Add(lblCaseType)
        
        cmbCaseType = New ComboBox With {
            .Location = New Point(100, 20),
            .Size = New Size(200, 25),
            .Font = New Font("微软雅黑", 10),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        
        ' 添加支持的案件类型
        Dim supportedTypes = TabTemplateFactory.GetSupportedCaseTypes()
        For Each caseType In supportedTypes
            cmbCaseType.Items.Add(caseType)
        Next
        
        If cmbCaseType.Items.Count > 0 Then
            cmbCaseType.SelectedIndex = 0
        End If
        
        Me.Controls.Add(cmbCaseType)
        
        ' 创建模板按钮
        btnCreateTemplate = New Button With {
            .Text = "创建模板",
            .Location = New Point(320, 20),
            .Size = New Size(100, 25),
            .Font = New Font("微软雅黑", 10)
        }
        AddHandler btnCreateTemplate.Click, AddressOf btnCreateTemplate_Click
        Me.Controls.Add(btnCreateTemplate)
        
        ' 加载数据按钮
        btnLoadData = New Button With {
            .Text = "加载测试数据",
            .Location = New Point(440, 20),
            .Size = New Size(120, 25),
            .Font = New Font("微软雅黑", 10),
            .Enabled = False
        }
        AddHandler btnLoadData.Click, AddressOf btnLoadData_Click
        Me.Controls.Add(btnLoadData)
        
        ' 保存数据按钮
        btnSaveData = New Button With {
            .Text = "保存数据",
            .Location = New Point(580, 20),
            .Size = New Size(100, 25),
            .Font = New Font("微软雅黑", 10),
            .Enabled = False
        }
        AddHandler btnSaveData.Click, AddressOf btnSaveData_Click
        Me.Controls.Add(btnSaveData)
        
        ' 设置只读按钮
        btnSetReadOnly = New Button With {
            .Text = "设置只读",
            .Location = New Point(700, 20),
            .Size = New Size(100, 25),
            .Font = New Font("微软雅黑", 10),
            .Enabled = False
        }
        AddHandler btnSetReadOnly.Click, AddressOf btnSetReadOnly_Click
        Me.Controls.Add(btnSetReadOnly)
        
        ' 设置样式按钮
        btnSetStyle = New Button With {
            .Text = "设置样式",
            .Location = New Point(820, 20),
            .Size = New Size(100, 25),
            .Font = New Font("微软雅黑", 10),
            .Enabled = False
        }
        AddHandler btnSetStyle.Click, AddressOf btnSetStyle_Click
        Me.Controls.Add(btnSetStyle)
        
        ' 信息标签
        lblInfo = New Label With {
            .Text = "请选择案件类型并创建模板",
            .Location = New Point(20, 60),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 10),
            .ForeColor = Color.Blue
        }
        Me.Controls.Add(lblInfo)
        
        ' TabControl
        tabControl = New TabControl With {
            .Location = New Point(20, 90),
            .Size = New Size(1140, 650)
        }
        Me.Controls.Add(tabControl)
    End Sub
    
    Private Sub InitTestData()
        ' 创建测试数据
        _testData = New List(Of CaseDetail)()
        
        ' 基本信息测试数据
        _testData.Add(New CaseDetail With {.TabIndex = 0, .FieldNo = "CaseNo", .FieldValue = "TEST001"})
        _testData.Add(New CaseDetail With {.TabIndex = 0, .FieldNo = "CaseName", .FieldValue = "测试案件"})
        _testData.Add(New CaseDetail With {.TabIndex = 0, .FieldNo = "Applicant", .FieldValue = "测试申请人"})
        _testData.Add(New CaseDetail With {.TabIndex = 0, .FieldNo = "Phone", .FieldValue = "13800138000"})
        _testData.Add(New CaseDetail With {.TabIndex = 0, .FieldNo = "Email", .FieldValue = "test@example.com"})
        
        ' 产品信息测试数据
        _testData.Add(New CaseDetail With {.TabIndex = 1, .FieldNo = "ProductCode", .FieldValue = "PROD001"})
        _testData.Add(New CaseDetail With {.TabIndex = 1, .FieldNo = "ProductName", .FieldValue = "测试产品"})
        _testData.Add(New CaseDetail With {.TabIndex = 1, .FieldNo = "ProductType", .FieldValue = "电子产品"})
        _testData.Add(New CaseDetail With {.TabIndex = 1, .FieldNo = "Specification", .FieldValue = "标准规格"})
        _testData.Add(New CaseDetail With {.TabIndex = 1, .FieldNo = "Description", .FieldValue = "这是一个测试产品的详细描述信息。"})
        
        ' 技术参数测试数据
        _testData.Add(New CaseDetail With {.TabIndex = 2, .FieldNo = "TechnicalStandard", .FieldValue = "GB/T 12345"})
        _testData.Add(New CaseDetail With {.TabIndex = 2, .FieldNo = "Param1", .FieldValue = "参数1值"})
        _testData.Add(New CaseDetail With {.TabIndex = 2, .FieldNo = "Param2", .FieldValue = "参数2值"})
        _testData.Add(New CaseDetail With {.TabIndex = 2, .FieldNo = "TechnicalRequirement", .FieldValue = "技术要求：符合相关标准，性能稳定可靠。"})
        
        ' 备注信息测试数据
        _testData.Add(New CaseDetail With {.TabIndex = 7, .FieldNo = "Memo", .FieldValue = "这是测试案件的备注信息，包含重要的处理说明和注意事项。"})
    End Sub
    
    Private Sub btnCreateTemplate_Click(sender As Object, e As EventArgs)
        Try
            If cmbCaseType.SelectedItem Is Nothing Then
                MessageBox.Show("请选择案件类型！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            Dim caseType As String = cmbCaseType.SelectedItem.ToString()
            
            ' 创建模板
            _currentTemplate = TabTemplateFactory.CreateTemplate(caseType, tabControl)
            
            ' 使用模板创建标签页
            _currentTemplate.CreateTabPages(tabControl)
            
            ' 更新界面状态
            btnLoadData.Enabled = True
            btnSaveData.Enabled = True
            btnSetReadOnly.Enabled = True
            btnSetStyle.Enabled = True
            
            ' 显示模板信息
            Dim tabInfo = TabTemplateFactory.GetTabInfo(caseType)
            lblInfo.Text = $"已创建模板：{_currentTemplate.GetTemplateName()}，标签页数量：{tabInfo.TabCount}"
            
            MessageBox.Show($"成功创建{_currentTemplate.GetTemplateName()}！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"创建模板失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnLoadData_Click(sender As Object, e As EventArgs)
        Try
            If _currentTemplate Is Nothing Then
                MessageBox.Show("请先创建模板！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' 加载测试数据
            _currentTemplate.LoadData(_testData)
            
            lblInfo.Text = $"已加载测试数据，共{_testData.Count}条记录"
            MessageBox.Show("测试数据加载成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"加载数据失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnSaveData_Click(sender As Object, e As EventArgs)
        Try
            If _currentTemplate Is Nothing Then
                MessageBox.Show("请先创建模板！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' 保存数据
            Dim savedData = _currentTemplate.SaveData()
            
            ' 显示保存结果
            Dim totalFields As Integer = 0
            For Each kvp In savedData
                totalFields += kvp.Value.Count
            Next
            
            lblInfo.Text = $"已保存数据，共{totalFields}个字段"
            MessageBox.Show($"数据保存成功！共保存{totalFields}个字段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"保存数据失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnSetReadOnly_Click(sender As Object, e As EventArgs)
        Try
            If _currentTemplate Is Nothing Then
                MessageBox.Show("请先创建模板！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' 切换只读状态
            Static isReadOnly As Boolean = False
            isReadOnly = Not isReadOnly
            
            _currentTemplate.SetReadOnly(isReadOnly)
            
            Dim statusText = If(isReadOnly, "只读", "可编辑")
            lblInfo.Text = $"控件状态：{statusText}"
            btnSetReadOnly.Text = If(isReadOnly, "取消只读", "设置只读")
            
            MessageBox.Show($"控件已设置为{statusText}状态！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"设置只读状态失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnSetStyle_Click(sender As Object, e As EventArgs)
        Try
            If _currentTemplate Is Nothing Then
                MessageBox.Show("请先创建模板！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' 切换样式
            Static styleIndex As Integer = 0
            Dim colors() As Color = {Color.White, Color.LightBlue, Color.LightGreen, Color.LightYellow, Color.LightGray}
            
            _currentTemplate.SetStyle(colors(styleIndex))
            
            lblInfo.Text = $"控件样式：{colors(styleIndex).Name}"
            
            styleIndex = (styleIndex + 1) Mod colors.Length
            
            MessageBox.Show($"控件样式已更新！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"设置样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class 
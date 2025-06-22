''' <summary>
''' 产品案件Tab模板 - 适用于产品相关的案件类型
''' </summary>
Public Class ProductCaseTemplate
    Inherits BaseTabTemplate
    
    Public Sub New(tabControl As TabControl)
        MyBase.New(tabControl)
        _tabNames = {"基本信息", "产品信息", "技术参数", "质量标准", "生产信息", "检验信息", "包装信息", "备注信息", "履历信息"}
    End Sub
    
    Public Overrides Sub CreateTabPages(tabControl As TabControl)
        ' 清空现有标签页
        tabControl.TabPages.Clear()
        
        ' 创建产品案件专用的标签页
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
            CreateProductTabContent(tabPage, i)
            tabControl.TabPages.Add(tabPage)
        Next
    End Sub
    
    Private Sub CreateProductTabContent(tabPage As TabPage, tabIndex As Integer)
        Dim y As Integer = 40
        
        ' 根据Tab索引创建不同的控件
        Select Case tabIndex
            Case 0 ' 基本信息
                CreateBasicInfoControls(tabPage, y)
            Case 1 ' 产品信息
                CreateProductInfoControls(tabPage, y)
            Case 2 ' 技术参数
                CreateTechnicalControls(tabPage, y)
            Case 3 ' 质量标准
                CreateQualityControls(tabPage, y)
            Case 4 ' 生产信息
                CreateProductionControls(tabPage, y)
            Case 5 ' 检验信息
                CreateInspectionControls(tabPage, y)
            Case 6 ' 包装信息
                CreatePackageControls(tabPage, y)
            Case 7 ' 备注信息
                CreateMemoControls(tabPage, y)
            Case 8 ' 履历信息
                CreateHistoryControls(tabPage, y)
        End Select
    End Sub
    
    Private Sub CreateBasicInfoControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("案件编号", "CaseNo", New TextBox),
            ("案件名称", "CaseName", New TextBox),
            ("申请日期", "ApplyDate", New DateTimePicker),
            ("申请人", "Applicant", New TextBox),
            ("联系电话", "Phone", New TextBox),
            ("电子邮箱", "Email", New TextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateProductInfoControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("产品代码", "ProductCode", New TextBox),
            ("产品名称", "ProductName", New TextBox),
            ("产品类型", "ProductType", New ComboBox),
            ("规格型号", "Specification", New TextBox),
            ("计量单位", "Unit", New ComboBox),
            ("产品描述", "Description", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateTechnicalControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("技术标准", "TechnicalStandard", New TextBox),
            ("主要参数1", "Param1", New TextBox),
            ("主要参数2", "Param2", New TextBox),
            ("主要参数3", "Param3", New TextBox),
            ("技术要求", "TechnicalRequirement", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateQualityControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("质量标准", "QualityStandard", New TextBox),
            ("检验方法", "InspectionMethod", New TextBox),
            ("合格标准", "QualifiedStandard", New TextBox),
            ("不合格处理", "UnqualifiedProcess", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateProductionControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("生产厂家", "Manufacturer", New TextBox),
            ("生产地址", "ProductionAddress", New TextBox),
            ("生产能力", "ProductionCapacity", New TextBox),
            ("生产许可证", "ProductionLicense", New TextBox),
            ("生产设备", "ProductionEquipment", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateInspectionControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("检验机构", "InspectionOrg", New TextBox),
            ("检验日期", "InspectionDate", New DateTimePicker),
            ("检验结果", "InspectionResult", New ComboBox),
            ("检验报告号", "ReportNo", New TextBox),
            ("检验备注", "InspectionMemo", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreatePackageControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("包装方式", "PackageMethod", New ComboBox),
            ("包装材料", "PackageMaterial", New TextBox),
            ("包装规格", "PackageSpec", New TextBox),
            ("存储条件", "StorageCondition", New TextBox),
            ("运输要求", "TransportRequirement", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
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
        Return New List(Of String) From {"产品案件", "新产品", "产品变更", "产品认证", "产品备案"}
    End Function
    
    Public Overrides Function GetTemplateName() As String
        Return "产品案件模板"
    End Function
End Class 
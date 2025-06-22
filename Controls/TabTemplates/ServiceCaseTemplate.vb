''' <summary>
''' 服务案件Tab模板 - 适用于服务相关的案件类型
''' </summary>
Public Class ServiceCaseTemplate
    Inherits BaseTabTemplate
    
    Public Sub New(tabControl As TabControl)
        MyBase.New(tabControl)
        _tabNames = {"基本信息", "服务信息", "服务标准", "服务流程", "人员信息", "设备信息", "质量保证", "备注信息", "履历信息"}
    End Sub
    
    Public Overrides Sub CreateTabPages(tabControl As TabControl)
        ' 清空现有标签页
        tabControl.TabPages.Clear()
        
        ' 创建服务案件专用的标签页
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
            CreateServiceTabContent(tabPage, i)
            tabControl.TabPages.Add(tabPage)
        Next
    End Sub
    
    Private Sub CreateServiceTabContent(tabPage As TabPage, tabIndex As Integer)
        Dim y As Integer = 40
        
        ' 根据Tab索引创建不同的控件
        Select Case tabIndex
            Case 0 ' 基本信息
                CreateBasicInfoControls(tabPage, y)
            Case 1 ' 服务信息
                CreateServiceInfoControls(tabPage, y)
            Case 2 ' 服务标准
                CreateServiceStandardControls(tabPage, y)
            Case 3 ' 服务流程
                CreateServiceProcessControls(tabPage, y)
            Case 4 ' 人员信息
                CreatePersonnelControls(tabPage, y)
            Case 5 ' 设备信息
                CreateEquipmentControls(tabPage, y)
            Case 6 ' 质量保证
                CreateQualityAssuranceControls(tabPage, y)
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
    
    Private Sub CreateServiceInfoControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("服务代码", "ServiceCode", New TextBox),
            ("服务名称", "ServiceName", New TextBox),
            ("服务类型", "ServiceType", New ComboBox),
            ("服务范围", "ServiceScope", New TextBox),
            ("服务周期", "ServiceCycle", New ComboBox),
            ("服务描述", "ServiceDescription", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateServiceStandardControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("服务标准", "ServiceStandard", New TextBox),
            ("质量标准", "QualityStandard", New TextBox),
            ("响应时间", "ResponseTime", New TextBox),
            ("服务承诺", "ServiceCommitment", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateServiceProcessControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("服务流程", "ServiceProcess", New TextBox),
            ("关键节点", "KeyNodes", New TextBox),
            ("时间要求", "TimeRequirement", New TextBox),
            ("流程说明", "ProcessDescription", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreatePersonnelControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("服务人员", "ServicePersonnel", New TextBox),
            ("人员资质", "PersonnelQualification", New TextBox),
            ("培训情况", "TrainingStatus", New TextBox),
            ("人员配置", "PersonnelAllocation", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateEquipmentControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("服务设备", "ServiceEquipment", New TextBox),
            ("设备型号", "EquipmentModel", New TextBox),
            ("设备状态", "EquipmentStatus", New ComboBox),
            ("设备维护", "EquipmentMaintenance", New RichTextBox)
        }
        
        CreateControlGroup(tabPage, controls, y)
    End Sub
    
    Private Sub CreateQualityAssuranceControls(tabPage As TabPage, y As Integer)
        Dim controls() As (String, String, Control) = {
            ("质量体系", "QualitySystem", New TextBox),
            ("监控方法", "MonitoringMethod", New TextBox),
            ("改进措施", "ImprovementMeasures", New TextBox),
            ("质量保证", "QualityAssurance", New RichTextBox)
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
        Return New List(Of String) From {"服务案件", "服务认证", "服务备案", "服务变更", "服务评估"}
    End Function
    
    Public Overrides Function GetTemplateName() As String
        Return "服务案件模板"
    End Function
End Class 
''' <summary>
''' Tab模板工厂类 - 根据案件类型创建对应的模板
''' </summary>
Public Class TabTemplateFactory
    
    Private Shared ReadOnly _templates As New Dictionary(Of String, Type)
    
    ''' <summary>
    ''' 静态构造函数，注册所有模板
    ''' </summary>
    Shared Sub New()
        RegisterTemplates()
    End Sub
    
    ''' <summary>
    ''' 注册所有模板类型
    ''' </summary>
    Private Shared Sub RegisterTemplates()
        ' 注册产品案件模板
        RegisterTemplate("产品案件", GetType(ProductCaseTemplate))
        RegisterTemplate("新产品", GetType(ProductCaseTemplate))
        RegisterTemplate("产品变更", GetType(ProductCaseTemplate))
        RegisterTemplate("产品认证", GetType(ProductCaseTemplate))
        RegisterTemplate("产品备案", GetType(ProductCaseTemplate))
        
        ' 注册服务案件模板
        RegisterTemplate("服务案件", GetType(ServiceCaseTemplate))
        RegisterTemplate("服务认证", GetType(ServiceCaseTemplate))
        RegisterTemplate("服务备案", GetType(ServiceCaseTemplate))
        RegisterTemplate("服务变更", GetType(ServiceCaseTemplate))
        RegisterTemplate("服务评估", GetType(ServiceCaseTemplate))
        
        ' 注册默认模板（用于未匹配的案件类型）
        RegisterTemplate("通用案件", GetType(DefaultCaseTemplate))
        RegisterTemplate("其他案件", GetType(DefaultCaseTemplate))
        RegisterTemplate("未分类案件", GetType(DefaultCaseTemplate))
    End Sub
    
    ''' <summary>
    ''' 注册模板类型
    ''' </summary>
    Private Shared Sub RegisterTemplate(caseType As String, templateType As Type)
        If Not _templates.ContainsKey(caseType) Then
            _templates.Add(caseType, templateType)
        End If
    End Sub
    
    ''' <summary>
    ''' 根据案件类型创建模板
    ''' </summary>
    ''' <param name="caseType">案件类型</param>
    ''' <param name="tabControl">TabControl控件</param>
    ''' <returns>对应的模板实例</returns>
    Public Shared Function CreateTemplate(caseType As String, tabControl As TabControl) As ITabTemplate
        ' 查找对应的模板类型
        Dim templateType As Type = Nothing
        
        If _templates.TryGetValue(caseType, templateType) Then
            ' 创建模板实例
            Return DirectCast(Activator.CreateInstance(templateType, tabControl), ITabTemplate)
        Else
            ' 如果没有找到对应模板，使用默认模板
            Return New DefaultCaseTemplate(tabControl)
        End If
    End Function
    
    ''' <summary>
    ''' 获取所有支持的案件类型
    ''' </summary>
    ''' <returns>支持的案件类型列表</returns>
    Public Shared Function GetSupportedCaseTypes() As List(Of String)
        Return _templates.Keys.ToList()
    End Function
    
    ''' <summary>
    ''' 获取模板信息
    ''' </summary>
    ''' <returns>模板信息字典</returns>
    Public Shared Function GetTemplateInfo() As Dictionary(Of String, String)
        Dim info As New Dictionary(Of String, String)
        
        For Each kvp In _templates
            Dim template As ITabTemplate = CreateTemplate(kvp.Key, Nothing)
            info.Add(kvp.Key, template.GetTemplateName())
        Next
        
        Return info
    End Function
    
    ''' <summary>
    ''' 检查案件类型是否支持
    ''' </summary>
    ''' <param name="caseType">案件类型</param>
    ''' <returns>是否支持</returns>
    Public Shared Function IsSupported(caseType As String) As Boolean
        Return _templates.ContainsKey(caseType)
    End Function
    
    ''' <summary>
    ''' 获取模板的标签页信息
    ''' </summary>
    ''' <param name="caseType">案件类型</param>
    ''' <returns>标签页信息</returns>
    Public Shared Function GetTabInfo(caseType As String) As (TabCount As Integer, TabNames As String())
        Try
            Dim template As ITabTemplate = CreateTemplate(caseType, Nothing)
            Return (template.GetTabCount(), template.GetTabNames())
        Catch ex As Exception
            ' 如果出错，返回默认信息
            Return (6, {"基本信息", "案件详情", "相关文件", "处理记录", "备注信息", "履历信息"})
        End Try
    End Function
End Class 
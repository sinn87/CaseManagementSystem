''' <summary>
''' Tab模板接口 - 定义所有案件类型Tab模板必须实现的方法
''' </summary>
Public Interface ITabTemplate
    ''' <summary>
    ''' 创建标签页内容
    ''' </summary>
    ''' <param name="tabControl">目标TabControl</param>
    Sub CreateTabPages(tabControl As TabControl)
    
    ''' <summary>
    ''' 加载数据到标签页
    ''' </summary>
    ''' <param name="caseDetails">案件详细信息</param>
    Sub LoadData(caseDetails As List(Of CaseDetail))
    
    ''' <summary>
    ''' 保存标签页数据
    ''' </summary>
    ''' <returns>保存的数据字典</returns>
    Function SaveData() As Dictionary(Of Integer, Dictionary(Of String, String))
    
    ''' <summary>
    ''' 设置控件只读状态
    ''' </summary>
    ''' <param name="readOnly">是否只读</param>
    Sub SetReadOnly(readOnly As Boolean)
    
    ''' <summary>
    ''' 设置控件样式
    ''' </summary>
    ''' <param name="backColor">背景色</param>
    Sub SetStyle(backColor As Color)
    
    ''' <summary>
    ''' 获取支持的案件类型
    ''' </summary>
    ''' <returns>支持的案件类型列表</returns>
    Function GetSupportedCaseTypes() As List(Of String)
    
    ''' <summary>
    ''' 获取模板名称
    ''' </summary>
    ''' <returns>模板名称</returns>
    Function GetTemplateName() As String
    
    ''' <summary>
    ''' 获取标签页数量
    ''' </summary>
    ''' <returns>标签页数量</returns>
    Function GetTabCount() As Integer
    
    ''' <summary>
    ''' 获取标签页名称列表
    ''' </summary>
    ''' <returns>标签页名称数组</returns>
    Function GetTabNames() As String()
End Interface 
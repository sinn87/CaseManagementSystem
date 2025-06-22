# CaseManagementSystem - Tab模板功能

## 功能概述

CaseManagementSystem系统实现了基于案件类型的Tab模板切换功能，通过工厂模式动态创建不同类型的Tab模板，提供灵活的案件录入界面。

## 功能简介

案件类型Tab模板切换功能允许系统根据不同案件类型自动显示不同的标签页布局，提供更专业和针对性的界面。

## 快速开始

### 1. 启动测试程序

运行程序时，会弹出选择对话框：
- 选择"是"：启动主程序
- 选择"否"：启动Tab模板测试窗体

### 2. 测试功能

在测试窗体中：

1. **选择案件类型**：从下拉列表中选择要测试的案件类型
2. **创建模板**：点击"创建模板"按钮，系统会根据案件类型创建对应的标签页
3. **加载测试数据**：点击"加载测试数据"按钮，将测试数据填充到模板中
4. **测试保存功能**：点击"保存数据"按钮，测试数据保存功能
5. **测试只读状态**：点击"设置只读"按钮，切换控件的只读状态
6. **测试样式设置**：点击"设置样式"按钮，切换控件的背景颜色

## 支持的案件类型

### 产品案件模板
- 产品案件
- 新产品
- 产品变更
- 产品认证
- 产品备案

**标签页结构**：
1. 基本信息
2. 产品信息
3. 技术参数
4. 质量标准
5. 生产信息
6. 检验信息
7. 包装信息
8. 备注信息
9. 履历信息

### 服务案件模板
- 服务案件
- 服务认证
- 服务备案
- 服务变更
- 服务评估

**标签页结构**：
1. 基本信息
2. 服务信息
3. 服务标准
4. 服务流程
5. 人员信息
6. 设备信息
7. 质量保证
8. 备注信息
9. 履历信息

### 默认案件模板
- 通用案件
- 其他案件
- 未分类案件

**标签页结构**：
1. 基本信息
2. 案件详情
3. 相关文件
4. 处理记录
5. 备注信息
6. 履历信息

## 技术实现

### 核心文件

1. **ITabTemplate.vb** - 模板接口定义
2. **BaseTabTemplate.vb** - 基础模板类
3. **ProductCaseTemplate.vb** - 产品案件模板
4. **ServiceCaseTemplate.vb** - 服务案件模板
5. **DefaultCaseTemplate.vb** - 默认案件模板
6. **TabTemplateFactory.vb** - 模板工厂类
7. **TestTabTemplateForm.vb** - 测试窗体

### 关键方法

```vb
' 创建模板
Dim template = TabTemplateFactory.CreateTemplate(caseType, tabControl)

' 创建标签页
template.CreateTabPages(tabControl)

' 加载数据
template.LoadData(caseDetails)

' 保存数据
Dim savedData = template.SaveData()

' 设置只读状态
template.SetReadOnly(True)

' 设置样式
template.SetStyle(Color.LightBlue)
```

## 扩展新模板

### 1. 创建新的模板类

```vb
Public Class NewCaseTemplate
    Inherits BaseTabTemplate
    
    Public Sub New(tabControl As TabControl)
        MyBase.New(tabControl)
        _tabNames = {"基本信息", "专业信息", "备注信息", "履历信息"}
    End Sub
    
    Public Overrides Sub CreateTabPages(tabControl As TabControl)
        ' 实现具体的标签页创建逻辑
        For i As Integer = 0 To _tabNames.Length - 1
            Dim tabPage As New TabPage(_tabNames(i))
            CreateTabContent(tabPage, i)
            tabControl.TabPages.Add(tabPage)
        Next
    End Sub
    
    Private Sub CreateTabContent(tabPage As TabPage, tabIndex As Integer)
        ' 根据标签页索引创建不同的控件
        Select Case tabIndex
            Case 0
                CreateBasicInfoControls(tabPage, 40)
            Case 1
                CreateProfessionalControls(tabPage, 40)
            ' ... 其他标签页
        End Select
    End Sub
    
    Public Overrides Function GetSupportedCaseTypes() As List(Of String)
        Return New List(Of String) From {"新案件类型1", "新案件类型2"}
    End Function
    
    Public Overrides Function GetTemplateName() As String
        Return "新案件模板"
    End Function
End Class
```

### 2. 在工厂类中注册

```vb
Private Shared Sub RegisterTemplates()
    ' 注册新模板
    RegisterTemplate("新案件类型1", GetType(NewCaseTemplate))
    RegisterTemplate("新案件类型2", GetType(NewCaseTemplate))
End Sub
```

## 数据库配置

确保数据库中有以下表结构：

### Cases表
```sql
CREATE TABLE Cases (
    CaseID INT PRIMARY KEY IDENTITY(1,1),
    CaseType NVARCHAR(50) NOT NULL,
    Status INT DEFAULT 1,
    CreateTime DATETIME DEFAULT GETDATE(),
    CreateUser NVARCHAR(50),
    LastUpdate DATETIME,
    IsTerminated BIT DEFAULT 0
)
```

### CaseDetails表
```sql
CREATE TABLE CaseDetails (
    DetailID INT PRIMARY KEY IDENTITY(1,1),
    CaseID INT NOT NULL,
    TabIndex INT NOT NULL,
    FieldNo NVARCHAR(50) NOT NULL,
    FieldValue NVARCHAR(MAX),
    FieldStatus NVARCHAR(20),
    CreateTime DATETIME DEFAULT GETDATE(),
    FOREIGN KEY (CaseID) REFERENCES Cases(CaseID)
)
```

### TabFieldMapping表
```sql
CREATE TABLE TabFieldMapping (
    MappingID INT PRIMARY KEY IDENTITY(1,1),
    TabIndex INT NOT NULL,
    FieldNo NVARCHAR(50) NOT NULL,
    FieldName NVARCHAR(100) NOT NULL,
    FieldType NVARCHAR(20) DEFAULT 'TextBox',
    IsRequired BIT DEFAULT 0,
    DefaultValue NVARCHAR(MAX)
)
```

## 注意事项

1. **模板一致性**：确保所有模板都实现了ITabTemplate接口的所有方法
2. **数据验证**：在模板中实现必要的数据验证逻辑
3. **性能优化**：大量控件创建时注意性能优化
4. **错误处理**：模板操作需要完善的异常处理
5. **用户体验**：模板切换时保持界面流畅

## 常见问题

### Q: 如何添加新的案件类型？
A: 创建新的模板类并在TabTemplateFactory中注册即可。

### Q: 如何修改现有模板的标签页？
A: 修改对应模板类的CreateTabPages方法和相关控件创建方法。

### Q: 如何自定义控件样式？
A: 在模板类中重写SetStyle方法或修改CreateControlGroup方法。

### Q: 如何添加数据验证？
A: 在模板的SaveData方法中添加验证逻辑，或在控件事件中添加验证。

## 联系支持

如有问题或建议，请联系开发团队。 

---
description: 面向 VB.NET 项目的通用编码规范与最佳实践，涵盖代码结构、命名、注释、异常处理、性能、安全等方面，帮助团队编写可维护、规范、安全的业务逻辑。
globs: *.vb
---

# VB.NET 编码规范与最佳实践手册

本文档定义了在 VB.NET 项目中应遵循的编程规范与最佳实践，以提升代码的可读性、可维护性、安全性和一致性。

## 1. 代码结构与项目组织

### 1.1 项目结构
    ProjectRoot/
    │
    ├─ 📁 Forms/                  '界面层（UI Layer）'
    │   ├─ MainForm.vb
    │   ├─ LoginForm.vb
    │   └─ SettingsForm.vb
    │
    ├─ 📁 Controls/               '自定义控件（UserControls）'
    │   ├─ UC_SearchBox.vb
    │   ├─ UC_Pagination.vb
    │   └─ UC_StatusBar.vb
    │
    ├─ 📁 BusinessLogic/          '业务逻辑层（BLL）'
    │   ├─ CaseManager.vb
    │   ├─ ReportService.vb
    │   └─ Validator.vb
    │
    ├─ 📁 DataAccess/             '数据访问层（DAL）'
    │   ├─ DbHelper.vb
    │   ├─ CaseRepository.vb
    │   ├─ UserRepository.vb
    │   └─ SqlBuilder.vb
    │
    ├─ 📁 Models/                 '实体类（数据模型）'
    │   ├─ CaseInfo.vb
    │   ├─ UserInfo.vb
    │   └─ LogEntry.vb
    │
    ├─ 📁 Interfaces/             '接口定义（Interface Layer）'
    │   ├─ ICaseService.vb
    │   ├─ IRepository.vb
    │   └─ IExportable.vb
    │
    ├─ 📁 Enums/                  '枚举定义'
    │   ├─ CaseStatus.vb
    │   └─ UserRoles.vb
    │
    ├─ 📁 Utils/                  '通用工具类'
    │   ├─ LogUtil.vb
    │   ├─ ConfigReader.vb
    │   └─ ExcelHelper.vb
    │
    ├─ 📁 Constants/              '常量定义'
    │   ├─ GlobalStrings.vb
    │   └─ AppSettings.vb
    │
    ├─ 📁 Resources/              '图片/多语言资源等'
    │   ├─ Icons/
    │   └─ Strings/
    │
    ├─ 📁 Tests/                  '单元测试（如使用 NUnit、xUnit）'
    │   └─ BusinessLogic.Tests/
    │
    └─ Program.vb / App.vb        '应用程序入口'


### 1.2 类文件组织

- 每个类使用单独文件，文件名与类名一致
- 类中成员顺序推荐：常量 → 变量 → 属性 → 构造函数 → 公共方法 → 私有方法 → 事件

## 2. 命名规范

### 2.1 命名风格

- **PascalCase** 用于类、属性、方法、事件（例：`CustomerManager`、`LoadData()`）
- **camelCase** 用于局部变量和参数（例：`dataList`, `userId`）
- 控件命名使用前缀 + 功能名，如 `txtName`, `btnSave`, `lblMessage`

### 2.2 常见前缀推荐

| 控件类型 | 前缀  |
|----------|-------|
| TextBox  | txt   |
| Button   | btn   |
| Label    | lbl   |
| DataGridView | dgv |
| ComboBox | cbo   |
| CheckBox | chk   |
| Form     | frm   |

## 3. 注释与文档

### 3.1 注释要求

- 每个公共类和方法应写明用途
- 逻辑复杂的代码块前需注释说明目的或处理流程
- 不使用无意义的注释（如 `i += 1 'i加一`）

### 3.2 示例格式

```vb
''' <summary>
''' 加载指定客户编号的数据
''' </summary>
''' <param name="customerId">客户编号</param>
''' <returns>客户数据对象</returns>
Public Function LoadCustomerData(customerId As String) As CustomerInfo
````

## 4. 错误处理

* 所有可能抛出异常的代码块应使用 `Try...Catch...Finally` 包裹
* 禁止捕获泛型 `Exception` 而不记录日志
* 使用 `Throw` 而不是 `Throw ex` 保留堆栈信息

```vb
Try
    ' 调用可能出错的函数
Catch ex As SqlException
    LogError("数据库异常", ex)
    MessageBox.Show("无法连接数据库，请稍后再试。")
Finally
    conn.Close()
End Try
```

## 5. 性能与可维护性

* 避免在循环中重复创建对象或频繁访问数据库
* 使用 `With` 块简化对象赋值逻辑
* 拆分过长函数，每个函数控制在 50 行以内

## 6. 安全性与兼容性

* 禁止使用硬编码密码、路径或 SQL 语句
* 所有用户输入需进行验证与转义，防止 SQL 注入
* 数据访问应使用参数化查询或 ORM 框架（如 Dapper）

## 7. 测试与调试建议

* 所有关键函数应覆盖基本输入、边界输入、异常情况
* UI 层与业务逻辑应分离，方便单元测试
* 对于复杂流程，建议先写伪代码再实现，便于同事 Review

---



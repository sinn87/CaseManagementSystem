# DataGridView数据保存功能说明

## 功能概述

本功能支持在案件详细录入窗体中，多个不同类型的DataGridView（如人员表、材料表等）数据的提取、保存和管理。每个DGV的数据将保存到数据库中对应的独立表，字段与DGV列一一对应。

## 主要特性

- 支持任意数量和类型的DGV，每个DGV对应数据库中的一张表
- 通用保存方法：根据表名和字段动态生成SQL，无需为每种DGV写专用保存方法
- 自动补充caseID等关联字段
- 支持批量插入，保证事务一致性
- "编号"字段为各自表内自增或顺序号，不作为外键

## 数据提取与保存

- 提取数据时，按DGV类型分组，组装为Dictionary(Of String, List(Of Dictionary(Of String, String)))，Key为表名，Value为行数据
- 只有有数据的DGV才会执行插入操作，无数据的DGV不会写入数据库，避免产生空数据
- 保存时，遍历所有表，调用通用保存方法SaveGridDataWithTransaction，动态插入数据

## 代码示例

```vb
' 提取数据
Dim gridData As Dictionary(Of String, List(Of Dictionary(Of String, String))) = ExtractAllGridData(tabControl)
' 保存数据
Dim success = BusinessLogic.CaseManager.CreateNewCase(caseType, tabData, gridData, currentUser)
```

## 数据库设计

- 每个DGV类型有独立的数据表，字段与DGV列一致
- caseID字段用于案件关联，编号字段为本表自增或顺序号

## 扩展性

- 新增DGV类型时，只需新建对应表和配置列，无需修改保存逻辑

## 审查流程

### 1. 数据提取
- 自动识别每个TabPage中的DataGridView控件
- 提取所有非空行的数据
- 支持动态列名和数据值提取

### 2. 事务保存
- DataGridView数据与单项数据在同一事务中保存
- 保证数据一致性，任何一步失败都会回滚所有操作
- 支持批量插入，提高性能

### 3. 审查流程
- 每条表格数据都有独立的审查状态
- 包含更新时间、审查时间、状态、审查人员等字段
- 与单项数据保持相同的审查流程

## 数据库设计

### CaseTableItems表结构
```sql
CREATE TABLE CaseTableItems (
    ItemID AUTOINCREMENT PRIMARY KEY,
    CaseID INTEGER NOT NULL,           -- 关联案件ID
    TabIndex INTEGER NOT NULL,         -- 标签页索引
    ItemName VARCHAR(255) NOT NULL,    -- 项目名称
    ItemValue TEXT,                    -- 项目值
    LastUpdate DATETIME NOT NULL,      -- 最后更新时间
    ReviewTime DATETIME NOT NULL,      -- 审查时间
    Status VARCHAR(50) NOT NULL,       -- 状态
    Reviewer VARCHAR(100) NOT NULL,    -- 审查人员
    FOREIGN KEY (CaseID) REFERENCES Cases(CaseID) ON DELETE CASCADE
);
```

## 代码结构

### 1. UI层 (CaseDetailEntryForm.vb)
- `CreateSampleControls()`: 在每个TabPage中创建DataGridView
- `ExtractGridData()`: 提取DataGridView数据
- `GetAllControls()`: 递归获取所有子控件

### 2. 业务逻辑层 (CaseManager.vb)
- `CreateNewCase()`: 修改后支持DataGridView数据保存
- 数据验证和业务规则处理

### 3. 数据访问层 (CaseRepository.vb)
- `SaveCaseTableItemsWithTransaction()`: 事务中批量保存表格数据
- `GetCaseTableItemsByCaseId()`: 根据案件ID获取表格数据

### 4. 数据模型 (CaseTableItem.vb)
- 定义表格数据实体类

## 使用方法

### 1. 在TabPage中添加DataGridView
```vb
' 创建DataGridView
Dim dgv As New DataGridView With {
    .Name = $"dgvItems_{pageIndex}",
    .AllowUserToAddRows = True,
    .AllowUserToDeleteRows = True
}

' 添加列
dgv.Columns.Add("ItemName", "项目名称")
dgv.Columns.Add("ItemValue", "项目值")
dgv.Columns.Add("LastUpdate", "最后更新时间")
dgv.Columns.Add("ReviewTime", "审查时间")
dgv.Columns.Add("Status", "状态")
dgv.Columns.Add("Reviewer", "审查人员")
```

### 2. 提取和保存数据
```vb
' 提取数据
Dim tabData = BusinessLogic.CaseManager.ExtractModifiedData(tabControl)
Dim tabGridData = ExtractGridData(tabControl)

' 保存数据
Dim success = BusinessLogic.CaseManager.CreateNewCase(caseType, tabData, tabGridData, currentUser)
```

## 测试功能

### TestDataGridViewForm.vb
提供了完整的测试窗体，可以：
- 在多个标签页中添加单项控件和DataGridView
- 输入测试数据
- 验证数据提取功能
- 清空数据重新测试

## 注意事项

1. **数据验证**: 只有项目名称不为空的行才会被保存
2. **默认值**: 状态默认为"新登录"，审查人员默认为当前用户
3. **时间格式**: 时间字段使用"yyyy-MM-dd HH:mm:ss"格式
4. **事务处理**: 所有数据在同一事务中保存，确保一致性
5. **性能优化**: 使用批量插入提高大量数据的保存性能

## 扩展功能

### 1. 自定义列
可以根据业务需求添加更多列：
```vb
dgv.Columns.Add("CustomField1", "自定义字段1")
dgv.Columns.Add("CustomField2", "自定义字段2")
```

### 2. 数据验证
可以在保存前添加数据验证逻辑：
```vb
' 验证必填字段
If String.IsNullOrEmpty(rowData("ItemName")) Then
    ' 显示错误信息
End If
```

### 3. 数据加载
可以从数据库加载已保存的表格数据：
```vb
Dim tableItems = CaseRepository.GetCaseTableItemsByCaseId(caseId)
' 将数据绑定到DataGridView
```

## 系统字段自动补全
- 数据库表中的系统字段（如更新时间、更新人员、审查日期、审查人员、状态等）由数据访问层在插入时自动补全，无需在UI或业务层赋值。
- 只需保证DGV数据中包含业务字段，系统字段会自动添加。

## 相关文件

- `Forms/CaseDetailEntryForm.vb`: 主窗体
- `BusinessLogic/CaseManager.vb`: 业务逻辑
- `DataAccess/CaseRepository.vb`: 数据访问
- `Models/CaseTableItem.vb`: 数据模型
- `Database/CaseTableItems_CreateTable.sql`: 数据库表结构
- `TestDataGridViewForm.vb`: 测试窗体 
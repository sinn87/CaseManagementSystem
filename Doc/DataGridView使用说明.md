# DataGridView使用说明

## 1. 概述

本文档说明如何在案件管理系统中正确使用DataGridView，确保与DataTable行状态管理功能兼容。

## 2. DataGridView创建原则

### 2.1 不在模板中直接创建DataGridView
- **原因**：DataGridView是自制控件，应该单独存放
- **做法**：在模板中预留位置，DataGridView通过其他方式添加

### 2.2 确保DataSource正确设置
- **必须**：DataGridView的DataSource必须是DataTable类型
- **原因**：支持行状态管理和批量保存功能
- **检查**：在模板加载时确保DataSource不为空

## 3. 模板中的DataGridView处理

### 3.1 BaseTabTemplate中的支持
```vb
' 在LoadTabData方法中自动检查DataGridView
For Each control As Control In GetAllControls(tabPage)
    ' 确保DataGridView有正确的DataSource
    If TypeOf control Is DataGridView Then
        EnsureDataGridViewDataSource(DirectCast(control, DataGridView))
    End If
Next
```

### 3.2 自动DataSource设置
```vb
Protected Sub EnsureDataGridViewDataSource(dgv As DataGridView)
    If dgv.DataSource Is Nothing Then
        ' 创建空的DataTable
        Dim dt As New DataTable()
        
        ' 根据DataGridView名称设置列结构
        Select Case dgv.Name.ToLower()
            Case "dgvperson", "dgv_person"
                dt.Columns.Add("编号", GetType(String))
                dt.Columns.Add("姓名", GetType(String))
                dt.Columns.Add("性别", GetType(String))
            Case "dgvmaterial", "dgv_material"
                dt.Columns.Add("编号", GetType(String))
                dt.Columns.Add("材料名", GetType(String))
                dt.Columns.Add("数量", GetType(String))
        End Select
        
        dgv.DataSource = dt
    End If
End Sub
```

## 4. 自定义DataGridView控件

### 4.1 控件创建位置
- **建议位置**：`Controls/`目录下创建自定义控件
- **命名规范**：`UC_DataGridView.vb`或类似命名
- **继承关系**：继承自DataGridView或UserControl

### 4.2 基本结构示例
```vb
Public Class UC_CustomDataGridView
    Inherits DataGridView
    
    Public Sub New()
        ' 初始化DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("编号", GetType(String))
        dt.Columns.Add("名称", GetType(String))
        dt.Columns.Add("值", GetType(String))
        
        ' 设置DataSource
        Me.DataSource = dt
        
        ' 设置其他属性
        Me.AllowUserToAddRows = True
        Me.AllowUserToDeleteRows = True
        Me.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub
    
    ''' <summary>
    ''' 获取DataTable
    ''' </summary>
    Public Function GetDataTable() As DataTable
        Return DirectCast(Me.DataSource, DataTable)
    End Function
    
    ''' <summary>
    ''' 设置DataTable
    ''' </summary>
    Public Sub SetDataTable(dt As DataTable)
        Me.DataSource = dt
    End Sub
End Class
```

## 5. 在模板中使用自定义DataGridView

### 5.1 模板中的引用
```vb
Private Sub CreateTabContent(tabPage As TabPage, y As Integer)
    ' 创建自定义DataGridView
    Dim customDgv As New UC_CustomDataGridView With {
        .Location = New Point(20, y),
        .Size = New Size(600, 200),
        .Name = "dgvCustom",
        .Tag = "CustomData"
    }
    
    tabPage.Controls.Add(customDgv)
End Sub
```

### 5.2 数据加载
```vb
' 在模板的LoadData方法中加载DataGridView数据
Public Overrides Sub LoadData(caseDetails As List(Of CaseDetail))
    MyBase.LoadData(caseDetails)
    
    ' 加载DataGridView数据
    LoadDataGridViewData()
End Sub

Private Sub LoadDataGridViewData()
    ' 从数据库加载DataGridView数据
    ' 这里需要根据具体的表名来加载数据
    For Each tabPage As TabPage In _tabControl.TabPages
        For Each control As Control In GetAllControls(tabPage)
            If TypeOf control Is DataGridView Then
                Dim dgv As DataGridView = DirectCast(control, DataGridView)
                LoadDataGridViewFromDatabase(dgv)
            End If
        Next
    Next
End Sub
```

## 6. 数据保存流程

### 6.1 保存时的数据提取
```vb
' 在CaseManager.ExtractGridData方法中
Public Shared Function ExtractGridData(tabControl As TabControl) As Dictionary(Of String, (TabIndex As Integer, Table As DataTable))
    Dim result As New Dictionary(Of String, (TabIndex As Integer, Table As DataTable))
    
    For i As Integer = 0 To tabControl.TabPages.Count - 1
        Dim tabPage As TabPage = tabControl.TabPages(i)
        
        For Each ctrl As Control In GetAllControls(tabPage)
            If TypeOf ctrl Is DataGridView Then
                Dim dgv As DataGridView = CType(ctrl, DataGridView)
                
                ' 检查DataGridView是否有数据源
                If dgv.DataSource IsNot Nothing AndAlso TypeOf dgv.DataSource Is DataTable Then
                    Dim dataTable As DataTable = DirectCast(dgv.DataSource, DataTable)
                    
                    ' 检查DataTable是否有数据
                    If dataTable.Rows.Count > 0 Then
                        result(dgv.Name) = (i, dataTable)
                    End If
                End If
            End If
        Next
    Next
    
    Return result
End Function
```

### 6.2 保存时的状态处理
```vb
' 在CaseRepository.SaveDataTableWithTransaction方法中
Public Shared Function SaveDataTableWithTransaction(transaction As OleDbTransaction, tableName As String, dataTable As DataTable, currentUser As String, Optional caseId As Integer = 0) As Boolean
    ' 分别处理不同状态的行
    For Each dataRow As DataRow In dataTable.Rows
        Select Case dataRow.RowState
            Case DataRowState.Added
                ' 处理新增行
            Case DataRowState.Modified
                ' 处理修改行
            Case DataRowState.Deleted
                ' 处理删除行
        End Select
    Next
End Function
```

## 7. 最佳实践

### 7.1 DataGridView命名规范
- **命名格式**：`dgv[功能名]`，如`dgvPerson`、`dgvMaterial`
- **Tag属性**：设置对应的数据库表名
- **Name属性**：与Tag属性保持一致

### 7.2 数据验证
```vb
' 在保存前验证DataGridView数据
Private Function ValidateDataGridView(dgv As DataGridView) As Boolean
    Dim dt As DataTable = DirectCast(dgv.DataSource, DataTable)
    
    For Each row As DataRow In dt.Rows
        If row.RowState <> DataRowState.Deleted Then
            ' 验证必填字段
            If String.IsNullOrEmpty(row("名称").ToString()) Then
                Return False
            End If
        End If
    Next
    
    Return True
End Function
```

### 7.3 错误处理
```vb
Try
    ' 保存DataGridView数据
    If Not CaseRepository.SaveDataTableWithTransaction(transaction, tableName, dataTable, currentUser, caseId) Then
        Throw New Exception("保存DataGridView数据失败")
    End If
    
    ' 保存成功后清除状态
    dataTable.AcceptChanges()
    
Catch ex As Exception
    ' 保存失败时回滚状态
    dataTable.RejectChanges()
    Throw
End Try
```

## 8. 注意事项

### 8.1 数据源管理
- 确保DataGridView的DataSource始终是DataTable类型
- 避免直接操作DataGridView.Rows集合
- 使用DataTable的方法来管理数据

### 8.2 状态同步
- 保存成功后调用`AcceptChanges()`
- 保存失败时调用`RejectChanges()`
- 避免状态不一致导致的数据错误

### 8.3 性能优化
- 批量处理DataGridView数据
- 避免频繁的数据源切换
- 合理设置DataGridView的显示属性

## 9. 总结

正确使用DataGridView的关键点：

1. **DataSource设置**：确保使用DataTable作为数据源
2. **状态管理**：利用DataTable的行状态管理功能
3. **模板分离**：DataGridView不在模板中直接创建
4. **数据验证**：保存前进行必要的数据验证
5. **错误处理**：完善的异常处理和状态回滚机制

通过遵循这些原则，可以确保DataGridView与系统的DataTable行状态管理功能完美配合，提供高效、可靠的数据操作体验。 
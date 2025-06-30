# UC_DataGridView使用示例

## 1. 概述

本文档提供了UC_DataGridView控件的详细使用示例，包括创建、配置、数据操作和事件处理。

## 2. 基本使用

### 2.1 创建UC_DataGridView实例

```vb
' 使用默认构造函数
Dim dgv As New UC_DataGridView()

' 使用自定义列定义
Dim columns As New Dictionary(Of String, Type) From {
    {"编号", GetType(String)},
    {"姓名", GetType(String)},
    {"性别", GetType(String)},
    {"部门", GetType(String)},
    {"入职日期", GetType(DateTime)}
}
Dim dgvPerson As New UC_DataGridView("PersonTable", columns)
```

### 2.2 基本属性设置

```vb
' 设置基本属性
dgv.Location = New Point(20, 100)
dgv.Size = New Size(600, 300)
dgv.Name = "dgvPerson"
dgv.Tag = "CasePersonItems"

' 设置列标题
Dim columnHeaders As New Dictionary(Of String, String) From {
    {"编号", "员工编号"},
    {"姓名", "员工姓名"},
    {"性别", "性别"},
    {"部门", "所属部门"},
    {"入职日期", "入职时间"}
}
dgv.SetColumnHeaders(columnHeaders)

' 设置列宽度
Dim columnWidths As New Dictionary(Of String, Integer) From {
    {"编号", 80},
    {"姓名", 120},
    {"性别", 60},
    {"部门", 150},
    {"入职日期", 120}
}
dgv.SetColumnWidths(columnWidths)
```

## 3. 在模板中使用

### 3.1 在BaseTabTemplate中创建DataGridView

```vb
Private Sub CreateDataGridViewTab(tabPage As TabPage, y As Integer)
    ' 创建标签
    Dim lblDgv As New Label With {
        .Text = "人员信息列表:",
        .Location = New Point(20, y),
        .AutoSize = True,
        .Font = New Font("微软雅黑", 9)
    }
    tabPage.Controls.Add(lblDgv)
    
    ' 创建UC_DataGridView
    Dim dgv As New UC_DataGridView With {
        .Location = New Point(20, y + 25),
        .Size = New Size(600, 200),
        .Name = "dgvPerson",
        .Tag = "CasePersonItems"
    }
    
    ' 设置列标题
    Dim columnHeaders As New Dictionary(Of String, String) From {
        {"编号", "员工编号"},
        {"姓名", "员工姓名"},
        {"性别", "性别"},
        {"部门", "所属部门"}
    }
    dgv.SetColumnHeaders(columnHeaders)
    
    ' 添加事件处理
    AddHandler dgv.DataChanged, AddressOf OnDataGridViewChanged
    AddHandler dgv.RowAdded, AddressOf OnDataGridViewRowAdded
    AddHandler dgv.RowDeleted, AddressOf OnDataGridViewRowDeleted
    
    tabPage.Controls.Add(dgv)
End Sub
```

### 3.2 事件处理方法

```vb
Private Sub OnDataGridViewChanged(sender As Object, e As EventArgs)
    ' 数据改变时的处理
    Dim dgv As UC_DataGridView = DirectCast(sender, UC_DataGridView)
    Console.WriteLine($"DataGridView {dgv.Name} 数据已改变")
End Sub

Private Sub OnDataGridViewRowAdded(sender As Object, e As DataGridViewRowEventArgs)
    ' 行添加时的处理
    Dim dgv As UC_DataGridView = DirectCast(sender, UC_DataGridView)
    Console.WriteLine($"DataGridView {dgv.Name} 添加了新行")
End Sub

Private Sub OnDataGridViewRowDeleted(sender As Object, e As DataGridViewRowEventArgs)
    ' 行删除时的处理
    Dim dgv As UC_DataGridView = DirectCast(sender, UC_DataGridView)
    Console.WriteLine($"DataGridView {dgv.Name} 删除了行")
End Sub
```

## 4. 数据操作

### 4.1 加载数据

```vb
' 从数据库加载数据
Private Sub LoadDataGridViewData(dgv As UC_DataGridView, caseId As Integer)
    Try
        ' 从数据库获取数据
        Dim data As List(Of Dictionary(Of String, String)) = GetDataFromDatabase(dgv.Tag.ToString(), caseId)
        
        ' 加载到DataGridView
        dgv.LoadData(data)
        
        ' 接受更改（标记为未修改状态）
        dgv.AcceptChanges()
        
    Catch ex As Exception
        Utils.LogUtil.LogError($"加载DataGridView数据失败：{ex.Message}")
    End Try
End Sub

' 从数据库获取数据的示例方法
Private Function GetDataFromDatabase(tableName As String, caseId As Integer) As List(Of Dictionary(Of String, String))
    ' 这里应该调用实际的数据访问层方法
    ' 示例返回空数据
    Return New List(Of Dictionary(Of String, String))
End Function
```

### 4.2 添加数据

```vb
' 添加新行
Private Sub AddNewRow(dgv As UC_DataGridView)
    Dim newRowData As New Dictionary(Of String, String) From {
        {"编号", "001"},
        {"姓名", "张三"},
        {"性别", "男"},
        {"部门", "技术部"}
    }
    
    dgv.AddRow(newRowData)
End Sub

' 批量添加数据
Private Sub AddMultipleRows(dgv As UC_DataGridView)
    Dim dataList As New List(Of Dictionary(Of String, String)) From {
        New Dictionary(Of String, String) From {
            {"编号", "001"},
            {"姓名", "张三"},
            {"性别", "男"},
            {"部门", "技术部"}
        },
        New Dictionary(Of String, String) From {
            {"编号", "002"},
            {"姓名", "李四"},
            {"性别", "女"},
            {"部门", "人事部"}
        }
    }
    
    dgv.LoadData(dataList)
End Sub
```

### 4.3 数据验证

```vb
' 验证数据
Private Function ValidateDataGridView(dgv As UC_DataGridView) As Boolean
    ' 定义必填列
    Dim requiredColumns As New List(Of String) From {"编号", "姓名"}
    
    ' 执行验证
    Dim validationResult = dgv.ValidateData(requiredColumns)
    
    If Not validationResult.IsValid Then
        MessageBox.Show($"数据验证失败：{validationResult.ErrorMessage}", "验证错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return False
    End If
    
    Return True
End Function
```

### 4.4 保存数据

```vb
' 保存DataGridView数据
Private Function SaveDataGridViewData(dgv As UC_DataGridView, caseId As Integer) As Boolean
    Try
        ' 验证数据
        If Not ValidateDataGridView(dgv) Then
            Return False
        End If
        
        ' 获取修改的数据
        Dim modifiedData = dgv.GetModifiedData()
        
        If modifiedData.Count > 0 Then
            ' 保存到数据库
            Dim success = SaveDataToDatabase(dgv.Tag.ToString(), modifiedData, caseId)
            
            If success Then
                ' 保存成功后接受更改
                dgv.AcceptChanges()
                Return True
            Else
                ' 保存失败时拒绝更改
                dgv.RejectChanges()
                Return False
            End If
        End If
        
        Return True
        
    Catch ex As Exception
        Utils.LogUtil.LogError($"保存DataGridView数据失败：{ex.Message}")
        dgv.RejectChanges()
        Return False
    End Try
End Function
```

## 5. 高级功能

### 5.1 状态管理

```vb
' 检查是否有未保存的更改
Private Sub CheckUnsavedChanges(dgv As UC_DataGridView)
    If dgv.HasChanges() Then
        Dim result = MessageBox.Show("有未保存的更改，是否保存？", "确认", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        
        Select Case result
            Case DialogResult.Yes
                SaveDataGridViewData(dgv, _caseId)
            Case DialogResult.No
                dgv.RejectChanges()
            Case DialogResult.Cancel
                ' 取消操作
        End Select
    End If
End Sub

' 设置只读状态
Private Sub SetDataGridViewReadOnly(dgv As UC_DataGridView, readOnly As Boolean)
    dgv.SetReadOnly(readOnly)
    
    If readOnly Then
        dgv.BackgroundColor = Color.LightGray
    Else
        dgv.BackgroundColor = Color.White
    End If
End Sub
```

### 5.2 数据导出

```vb
' 导出到CSV
Private Sub ExportDataGridViewToCSV(dgv As UC_DataGridView)
    Try
        Dim saveFileDialog As New SaveFileDialog With {
            .Filter = "CSV文件 (*.csv)|*.csv",
            .FileName = $"{dgv.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        }
        
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            dgv.ExportToCSV(saveFileDialog.FileName)
            MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        
    Catch ex As Exception
        MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub
```

### 5.3 自定义样式

```vb
' 设置自定义样式
Private Sub SetCustomStyle(dgv As UC_DataGridView)
    ' 设置交替行颜色
    dgv.AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle With {
        .BackColor = Color.LightBlue,
        .Font = New Font("微软雅黑", 9)
    }
    
    ' 设置选中行样式
    dgv.DefaultCellStyle.SelectionBackColor = Color.DarkBlue
    dgv.DefaultCellStyle.SelectionForeColor = Color.White
    
    ' 设置网格线样式
    dgv.GridColor = Color.DarkGray
    dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical
End Sub
```

## 6. 完整示例

### 6.1 在模板中的完整实现

```vb
Public Class CustomTabTemplate
    Inherits BaseTabTemplate
    
    Private _dgvPerson As UC_DataGridView
    Private _dgvMaterial As UC_DataGridView
    
    Public Overrides Sub CreateTabPages(tabControl As TabControl)
        ' 清空现有标签页
        tabControl.TabPages.Clear()
        
        ' 创建人员信息标签页
        Dim tabPerson As New TabPage("人员信息")
        CreatePersonTab(tabPerson)
        tabControl.TabPages.Add(tabPerson)
        
        ' 创建材料信息标签页
        Dim tabMaterial As New TabPage("材料信息")
        CreateMaterialTab(tabMaterial)
        tabControl.TabPages.Add(tabMaterial)
    End Sub
    
    Private Sub CreatePersonTab(tabPage As TabPage)
        ' 创建人员DataGridView
        Dim columns As New Dictionary(Of String, Type) From {
            {"编号", GetType(String)},
            {"姓名", GetType(String)},
            {"性别", GetType(String)},
            {"部门", GetType(String)}
        }
        
        _dgvPerson = New UC_DataGridView("PersonTable", columns) With {
            .Location = New Point(20, 50),
            .Size = New Size(600, 300),
            .Name = "dgvPerson",
            .Tag = "CasePersonItems"
        }
        
        ' 设置列标题
        Dim columnHeaders As New Dictionary(Of String, String) From {
            {"编号", "员工编号"},
            {"姓名", "员工姓名"},
            {"性别", "性别"},
            {"部门", "所属部门"}
        }
        _dgvPerson.SetColumnHeaders(columnHeaders)
        
        ' 添加事件处理
        AddHandler _dgvPerson.DataChanged, AddressOf OnPersonDataChanged
        
        tabPage.Controls.Add(_dgvPerson)
    End Sub
    
    Private Sub CreateMaterialTab(tabPage As TabPage)
        ' 创建材料DataGridView
        Dim columns As New Dictionary(Of String, Type) From {
            {"编号", GetType(String)},
            {"材料名", GetType(String)},
            {"数量", GetType(String)},
            {"单位", GetType(String)}
        }
        
        _dgvMaterial = New UC_DataGridView("MaterialTable", columns) With {
            .Location = New Point(20, 50),
            .Size = New Size(600, 300),
            .Name = "dgvMaterial",
            .Tag = "CaseMaterialItems"
        }
        
        ' 设置列标题
        Dim columnHeaders As New Dictionary(Of String, String) From {
            {"编号", "材料编号"},
            {"材料名", "材料名称"},
            {"数量", "数量"},
            {"单位", "单位"}
        }
        _dgvMaterial.SetColumnHeaders(columnHeaders)
        
        tabPage.Controls.Add(_dgvMaterial)
    End Sub
    
    Private Sub OnPersonDataChanged(sender As Object, e As EventArgs)
        ' 人员数据改变时的处理
        Console.WriteLine("人员数据已改变")
    End Sub
    
    Public Overrides Function GetSupportedCaseTypes() As List(Of String)
        Return New List(Of String) From {"自定义案件"}
    End Function
    
    Public Overrides Function GetTemplateName() As String
        Return "自定义模板"
    End Function
End Class
```

## 7. 注意事项

### 7.1 性能优化
- 避免频繁的数据源切换
- 使用批量操作而不是逐行操作
- 合理设置DataGridView的显示属性

### 7.2 内存管理
- 及时释放不需要的资源
- 使用Using语句确保资源正确释放
- 避免内存泄漏

### 7.3 错误处理
- 始终使用Try-Catch包装数据操作
- 提供用户友好的错误信息
- 记录详细的错误日志

### 7.4 数据一致性
- 保存前进行数据验证
- 使用事务确保数据一致性
- 正确处理保存成功和失败的情况

## 8. 总结

UC_DataGridView控件提供了完整的DataTable行状态管理功能，通过合理使用这些功能，可以：

1. **简化数据操作**：自动处理行的增删改状态
2. **提高性能**：批量处理数据，减少数据库往返
3. **增强用户体验**：提供直观的数据操作界面
4. **确保数据一致性**：完善的状态管理和错误处理

通过遵循本文档中的示例和最佳实践，可以充分发挥UC_DataGridView控件的优势，构建高效、可靠的数据管理界面。 
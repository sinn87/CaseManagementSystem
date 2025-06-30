''' <summary>
''' 自定义DataGridView控件 - 支持DataTable行状态管理
''' </summary>
Imports System.Data
Imports System.Windows.Forms
Imports System.Drawing

Public Class UC_DataGridView
    Inherits DataGridView
    
    ' 事件声明
    Public Event DataChanged(sender As Object, e As EventArgs)
    Public Event RowAdded(sender As Object, e As DataGridViewRowEventArgs)
    Public Event RowDeleted(sender As Object, e As DataGridViewRowEventArgs)
    Public Event RowModified(sender As Object, e As DataGridViewRowEventArgs)
    
    ' 私有字段
    Private _originalDataTable As DataTable
    Private _isLoading As Boolean = False
    
    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New()
        InitializeDataGridView()
        InitializeDataTable()
    End Sub
    
    ''' <summary>
    ''' 带参数的构造函数
    ''' </summary>
    ''' <param name="tableName">表名</param>
    ''' <param name="columns">列定义</param>
    Public Sub New(tableName As String, columns As Dictionary(Of String, Type))
        InitializeDataGridView()
        InitializeDataTable(tableName, columns)
    End Sub
    
    ''' <summary>
    ''' 初始化DataGridView属性
    ''' </summary>
    Private Sub InitializeDataGridView()
        ' 基本属性设置
        Me.AllowUserToAddRows = True
        Me.AllowUserToDeleteRows = True
        Me.AllowUserToResizeRows = True
        Me.AllowUserToResizeColumns = True
        Me.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Me.BackgroundColor = Color.White
        Me.BorderStyle = BorderStyle.Fixed3D
        Me.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        Me.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        Me.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
            .BackColor = Color.LightGray,
            .Font = New Font("微软雅黑", 9, FontStyle.Bold),
            .ForeColor = Color.Black
        }
        Me.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DefaultCellStyle = New DataGridViewCellStyle With {
            .Font = New Font("微软雅黑", 9),
            .BackColor = Color.White,
            .ForeColor = Color.Black
        }
        Me.EnableHeadersVisualStyles = False
        Me.GridColor = Color.LightGray
        Me.MultiSelect = False
        Me.ReadOnly = False
        Me.RowHeadersVisible = True
        Me.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        
        ' 添加事件处理
        AddEventHandlers()
    End Sub
    
    ''' <summary>
    ''' 初始化默认DataTable
    ''' </summary>
    Private Sub InitializeDataTable()
        Dim dt As New DataTable()
        
        ' 添加默认列
        dt.Columns.Add("编号", GetType(String))
        dt.Columns.Add("名称", GetType(String))
        dt.Columns.Add("值", GetType(String))
        
        ' 设置DataSource
        Me.DataSource = dt
        _originalDataTable = dt.Copy()
    End Sub
    
    ''' <summary>
    ''' 初始化自定义DataTable
    ''' </summary>
    ''' <param name="tableName">表名</param>
    ''' <param name="columns">列定义</param>
    Private Sub InitializeDataTable(tableName As String, columns As Dictionary(Of String, Type))
        Dim dt As New DataTable(tableName)
        
        ' 添加自定义列
        For Each column In columns
            dt.Columns.Add(column.Key, column.Value)
        Next
        
        ' 设置DataSource
        Me.DataSource = dt
        _originalDataTable = dt.Copy()
    End Sub
    
    ''' <summary>
    ''' 添加事件处理程序
    ''' </summary>
    Private Sub AddEventHandlers()
        AddHandler Me.CellValueChanged, AddressOf OnCellValueChanged
        AddHandler Me.RowsAdded, AddressOf OnRowsAdded
        AddHandler Me.RowsRemoved, AddressOf OnRowsRemoved
        AddHandler Me.UserDeletingRow, AddressOf OnUserDeletingRow
        AddHandler Me.DataError, AddressOf OnDataError
    End Sub
    
    ''' <summary>
    ''' 获取DataTable
    ''' </summary>
    ''' <returns>当前DataTable</returns>
    Public Function GetDataTable() As DataTable
        Return DirectCast(Me.DataSource, DataTable)
    End Function
    
    ''' <summary>
    ''' 设置DataTable
    ''' </summary>
    ''' <param name="dt">要设置的DataTable</param>
    Public Sub SetDataTable(dt As DataTable)
        _isLoading = True
        Try
            Me.DataSource = dt
            If dt IsNot Nothing Then
                _originalDataTable = dt.Copy()
            End If
        Finally
            _isLoading = False
        End Try
    End Sub
    
    ''' <summary>
    ''' 加载数据
    ''' </summary>
    ''' <param name="data">数据列表</param>
    Public Sub LoadData(data As List(Of Dictionary(Of String, String)))
        _isLoading = True
        Try
            Dim dt As DataTable = GetDataTable()
            dt.Clear()
            
            For Each rowData In data
                Dim newRow As DataRow = dt.NewRow()
                For Each kvp In rowData
                    If dt.Columns.Contains(kvp.Key) Then
                        newRow(kvp.Key) = kvp.Value
                    End If
                Next
                dt.Rows.Add(newRow)
            Next
            
            ' 保存原始数据
            _originalDataTable = dt.Copy()
        Finally
            _isLoading = False
        End Try
    End Sub
    
    ''' <summary>
    ''' 清空数据
    ''' </summary>
    Public Sub ClearData()
        Dim dt As DataTable = GetDataTable()
        dt.Clear()
        _originalDataTable = dt.Copy()
    End Sub
    
    ''' <summary>
    ''' 添加新行
    ''' </summary>
    ''' <param name="rowData">行数据</param>
    Public Sub AddRow(rowData As Dictionary(Of String, String))
        Dim dt As DataTable = GetDataTable()
        Dim newRow As DataRow = dt.NewRow()
        
        For Each kvp In rowData
            If dt.Columns.Contains(kvp.Key) Then
                newRow(kvp.Key) = kvp.Value
            End If
        Next
        
        dt.Rows.Add(newRow)
    End Sub
    
    ''' <summary>
    ''' 删除选中行
    ''' </summary>
    Public Sub DeleteSelectedRow()
        If Me.SelectedRows.Count > 0 Then
            Dim selectedRow = Me.SelectedRows(0)
            Me.Rows.Remove(selectedRow)
        End If
    End Sub
    
    ''' <summary>
    ''' 获取修改的数据
    ''' </summary>
    ''' <returns>修改的数据列表</returns>
    Public Function GetModifiedData() As List(Of Dictionary(Of String, String))
        Dim result As New List(Of Dictionary(Of String, String))
        Dim dt As DataTable = GetDataTable()
        
        For Each row As DataRow In dt.Rows
            If row.RowState <> DataRowState.Unchanged Then
                Dim rowData As New Dictionary(Of String, String)
                
                For Each column As DataColumn In dt.Columns
                    Dim value As Object = row(column)
                    rowData(column.ColumnName) = If(value?.ToString(), "")
                Next
                
                ' 添加行状态信息
                rowData("RowState") = row.RowState.ToString()
                
                result.Add(rowData)
            End If
        Next
        
        Return result
    End Function
    
    ''' <summary>
    ''' 检查是否有未保存的更改
    ''' </summary>
    ''' <returns>是否有更改</returns>
    Public Function HasChanges() As Boolean
        Dim dt As DataTable = GetDataTable()
        Return dt.GetChanges() IsNot Nothing
    End Function
    
    ''' <summary>
    ''' 接受更改
    ''' </summary>
    Public Sub AcceptChanges()
        Dim dt As DataTable = GetDataTable()
        dt.AcceptChanges()
        _originalDataTable = dt.Copy()
    End Sub
    
    ''' <summary>
    ''' 拒绝更改
    ''' </summary>
    Public Sub RejectChanges()
        Dim dt As DataTable = GetDataTable()
        dt.RejectChanges()
        If _originalDataTable IsNot Nothing Then
            dt.Clear()
            For Each row As DataRow In _originalDataTable.Rows
                dt.ImportRow(row)
            Next
        End If
    End Sub
    
    ''' <summary>
    ''' 验证数据
    ''' </summary>
    ''' <param name="requiredColumns">必填列名列表</param>
    ''' <returns>验证结果</returns>
    Public Function ValidateData(requiredColumns As List(Of String)) As (IsValid As Boolean, ErrorMessage As String)
        Dim dt As DataTable = GetDataTable()
        
        For Each row As DataRow In dt.Rows
            If row.RowState <> DataRowState.Deleted Then
                For Each columnName In requiredColumns
                    If dt.Columns.Contains(columnName) Then
                        Dim value As Object = row(columnName)
                        If value Is DBNull.Value OrElse String.IsNullOrEmpty(value?.ToString()) Then
                            Return (False, $"第{dt.Rows.IndexOf(row) + 1}行的{columnName}不能为空")
                        End If
                    End If
                Next
            End If
        Next
        
        Return (True, "")
    End Function
    
    ''' <summary>
    ''' 设置只读状态
    ''' </summary>
    ''' <param name="readOnly">是否只读</param>
    Public Sub SetReadOnly(readOnly As Boolean)
        Me.ReadOnly = readOnly
        Me.AllowUserToAddRows = Not readOnly
        Me.AllowUserToDeleteRows = Not readOnly
    End Sub
    
    ''' <summary>
    ''' 设置列标题
    ''' </summary>
    ''' <param name="columnHeaders">列标题字典</param>
    Public Sub SetColumnHeaders(columnHeaders As Dictionary(Of String, String))
        For Each kvp In columnHeaders
            If Me.Columns.Contains(kvp.Key) Then
                Me.Columns(kvp.Key).HeaderText = kvp.Value
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' 设置列宽度
    ''' </summary>
    ''' <param name="columnWidths">列宽度字典</param>
    Public Sub SetColumnWidths(columnWidths As Dictionary(Of String, Integer))
        For Each kvp In columnWidths
            If Me.Columns.Contains(kvp.Key) Then
                Me.Columns(kvp.Key).Width = kvp.Value
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' 导出到CSV
    ''' </summary>
    ''' <param name="filePath">文件路径</param>
    Public Sub ExportToCSV(filePath As String)
        Try
            Dim dt As DataTable = GetDataTable()
            Dim csvContent As New System.Text.StringBuilder()
            
            ' 添加列标题
            For i As Integer = 0 To dt.Columns.Count - 1
                csvContent.Append(dt.Columns(i).ColumnName)
                If i < dt.Columns.Count - 1 Then
                    csvContent.Append(",")
                End If
            Next
            csvContent.AppendLine()
            
            ' 添加数据行
            For Each row As DataRow In dt.Rows
                If row.RowState <> DataRowState.Deleted Then
                    For i As Integer = 0 To dt.Columns.Count - 1
                        Dim value As Object = row(i)
                        csvContent.Append(If(value?.ToString(), ""))
                        If i < dt.Columns.Count - 1 Then
                            csvContent.Append(",")
                        End If
                    Next
                    csvContent.AppendLine()
                End If
            Next
            
            ' 写入文件
            System.IO.File.WriteAllText(filePath, csvContent.ToString(), System.Text.Encoding.UTF8)
            
        Catch ex As Exception
            Throw New Exception($"导出CSV失败：{ex.Message}")
        End Try
    End Sub
    
    #Region "事件处理"
    
    ''' <summary>
    ''' 单元格值改变事件
    ''' </summary>
    Private Sub OnCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Not _isLoading Then
            RaiseEvent DataChanged(Me, EventArgs.Empty)
        End If
    End Sub
    
    ''' <summary>
    ''' 行添加事件
    ''' </summary>
    Private Sub OnRowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
        If Not _isLoading Then
            RaiseEvent RowAdded(Me, New DataGridViewRowEventArgs(Me.Rows(e.RowIndex)))
        End If
    End Sub
    
    ''' <summary>
    ''' 行删除事件
    ''' </summary>
    Private Sub OnRowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
        If Not _isLoading Then
            RaiseEvent RowDeleted(Me, New DataGridViewRowEventArgs(Nothing))
        End If
    End Sub
    
    ''' <summary>
    ''' 用户删除行事件
    ''' </summary>
    Private Sub OnUserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs)
        ' 可以在这里添加删除确认逻辑
        ' If MessageBox.Show("确定要删除这行吗？", "确认", MessageBoxButtons.YesNo) = DialogResult.No Then
        '     e.Cancel = True
        ' End If
    End Sub
    
    ''' <summary>
    ''' 数据错误事件
    ''' </summary>
    Private Sub OnDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        ' 记录数据错误
        Utils.LogUtil.LogError($"DataGridView数据错误：{e.Exception.Message}")
    End Sub
    
    #End Region
    
    #Region "IDisposable"
    
    ''' <summary>
    ''' 释放资源
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            ' 清理托管资源
            If _originalDataTable IsNot Nothing Then
                _originalDataTable.Dispose()
                _originalDataTable = Nothing
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    
    #End Region
End Class 
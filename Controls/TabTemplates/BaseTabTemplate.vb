''' <summary>
''' Tab模板基础类 - 提供通用的模板功能实现
''' </summary>
Imports System.Data

Public MustInherit Class BaseTabTemplate
    Implements ITabTemplate
    
    Protected _tabControl As TabControl
    Protected _caseDetails As List(Of CaseDetail)
    Protected _tabNames As String()
    
    Public Sub New(tabControl As TabControl)
        _tabControl = tabControl
        _caseDetails = New List(Of CaseDetail)
    End Sub
    
    ''' <summary>
    ''' 创建标签页内容 - 子类必须实现
    ''' </summary>
    Public MustOverride Sub CreateTabPages(tabControl As TabControl) Implements ITabTemplate.CreateTabPages
    
    ''' <summary>
    ''' 加载数据到标签页
    ''' </summary>
    Public Overridable Sub LoadData(caseDetails As List(Of CaseDetail)) Implements ITabTemplate.LoadData
        _caseDetails = caseDetails
        
        For Each tabPage As TabPage In _tabControl.TabPages
            Dim tabIndex As Integer = _tabControl.TabPages.IndexOf(tabPage)
            LoadTabData(tabPage, tabIndex)
        Next
    End Sub
    
    ''' <summary>
    ''' 保存标签页数据
    ''' </summary>
    Public Overridable Function SaveData() As Dictionary(Of Integer, Dictionary(Of String, String)) Implements ITabTemplate.SaveData
        Dim result As New Dictionary(Of Integer, Dictionary(Of String, String))
        
        For Each tabPage As TabPage In _tabControl.TabPages
            Dim tabIndex As Integer = _tabControl.TabPages.IndexOf(tabPage)
            Dim fieldData As New Dictionary(Of String, String)
            
            ' 遍历标签页中的所有控件
            For Each control As Control In GetAllControls(tabPage)
                If Not String.IsNullOrEmpty(control.Tag?.ToString()) Then
                    Dim fieldName As String = control.Tag.ToString()
                    Dim fieldValue As String = GetControlValue(control)
                    
                    ' 如果有值，则保存
                    If Not String.IsNullOrEmpty(fieldValue) Then
                        fieldData(fieldName) = fieldValue
                    End If
                End If
            Next
            
            ' 如果该标签页有数据，则添加到结果中
            If fieldData.Count > 0 Then
                result(tabIndex) = fieldData
            End If
        Next
        
        Return result
    End Function
    
    ''' <summary>
    ''' 设置控件只读状态
    ''' </summary>
    Public Overridable Sub SetReadOnly(readOnly As Boolean) Implements ITabTemplate.SetReadOnly
        For Each tabPage As TabPage In _tabControl.TabPages
            For Each control As Control In GetAllControls(tabPage)
                If TypeOf control Is TextBox Then
                    DirectCast(control, TextBox).ReadOnly = readOnly
                ElseIf TypeOf control Is ComboBox Then
                    DirectCast(control, ComboBox).Enabled = Not readOnly
                ElseIf TypeOf control Is DateTimePicker Then
                    DirectCast(control, DateTimePicker).Enabled = Not readOnly
                ElseIf TypeOf control Is RichTextBox Then
                    DirectCast(control, RichTextBox).ReadOnly = readOnly
                ElseIf TypeOf control Is DataGridView Then
                    DirectCast(control, DataGridView).ReadOnly = readOnly
                    DirectCast(control, DataGridView).AllowUserToAddRows = Not readOnly
                    DirectCast(control, DataGridView).AllowUserToDeleteRows = Not readOnly
                End If
            Next
        Next
    End Sub
    
    ''' <summary>
    ''' 设置控件样式
    ''' </summary>
    Public Overridable Sub SetStyle(backColor As Color) Implements ITabTemplate.SetStyle
        For Each tabPage As TabPage In _tabControl.TabPages
            For Each control As Control In GetAllControls(tabPage)
                If TypeOf control Is TextBox OrElse TypeOf control Is ComboBox OrElse TypeOf control Is DateTimePicker OrElse TypeOf control Is RichTextBox Then
                    control.BackColor = backColor
                ElseIf TypeOf control Is DataGridView Then
                    DirectCast(control, DataGridView).BackgroundColor = backColor
                End If
            Next
        Next
    End Sub
    
    ''' <summary>
    ''' 获取支持的案件类型 - 子类必须实现
    ''' </summary>
    Public MustOverride Function GetSupportedCaseTypes() As List(Of String) Implements ITabTemplate.GetSupportedCaseTypes
    
    ''' <summary>
    ''' 获取模板名称 - 子类必须实现
    ''' </summary>
    Public MustOverride Function GetTemplateName() As String Implements ITabTemplate.GetTemplateName
    
    ''' <summary>
    ''' 获取标签页数量
    ''' </summary>
    Public Overridable Function GetTabCount() As Integer Implements ITabTemplate.GetTabCount
        Return If(_tabNames, New String() {}).Length
    End Function
    
    ''' <summary>
    ''' 获取标签页名称列表
    ''' </summary>
    Public Overridable Function GetTabNames() As String() Implements ITabTemplate.GetTabNames
        Return _tabNames
    End Function
    
    ''' <summary>
    ''' 加载标签页数据
    ''' </summary>
    Protected Overridable Sub LoadTabData(tabPage As TabPage, tabIndex As Integer)
        ' 从_caseDetails中加载对应Tab的数据
        Dim tabDetails = _caseDetails.Where(Function(d) d.TabIndex = tabIndex).ToList()
        
        For Each control As Control In GetAllControls(tabPage)
            If Not String.IsNullOrEmpty(control.Tag?.ToString()) Then
                Dim fieldName As String = control.Tag.ToString()
                Dim detail = tabDetails.FirstOrDefault(Function(d) d.FieldNo = fieldName)
                
                If detail IsNot Nothing Then
                    SetControlValue(control, detail.FieldValue)
                End If
            End If
            
            ' 确保DataGridView有正确的DataSource
            If TypeOf control Is DataGridView Then
                EnsureDataGridViewDataSource(DirectCast(control, DataGridView))
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' 确保DataGridView有正确的DataSource
    ''' </summary>
    ''' <param name="dgv">DataGridView控件</param>
    Protected Sub EnsureDataGridViewDataSource(dgv As DataGridView)
        If dgv.DataSource Is Nothing Then
            ' 创建一个空的DataTable作为默认数据源
            Dim dt As New DataTable()
            
            ' 根据DataGridView的名称或Tag来确定表结构
            If Not String.IsNullOrEmpty(dgv.Name) Then
                ' 可以根据DataGridView的名称来设置不同的列结构
                Select Case dgv.Name.ToLower()
                    Case "dgvperson", "dgv_person"
                        dt.Columns.Add("编号", GetType(String))
                        dt.Columns.Add("姓名", GetType(String))
                        dt.Columns.Add("性别", GetType(String))
                        dt.Columns.Add("部门", GetType(String))
                    Case "dgvmaterial", "dgv_material"
                        dt.Columns.Add("编号", GetType(String))
                        dt.Columns.Add("材料名", GetType(String))
                        dt.Columns.Add("数量", GetType(String))
                        dt.Columns.Add("单位", GetType(String))
                    Case Else
                        ' 默认列结构
                        dt.Columns.Add("项目名称", GetType(String))
                        dt.Columns.Add("项目值", GetType(String))
                End Select
            End If
            
            dgv.DataSource = dt
        End If
    End Sub
    
    ''' <summary>
    ''' 获取所有控件（包括子控件）
    ''' </summary>
    Protected Function GetAllControls(container As Control) As List(Of Control)
        Dim controls As New List(Of Control)()
        
        For Each control As Control In container.Controls
            controls.Add(control)
            ' 递归获取子控件
            controls.AddRange(GetAllControls(control))
        Next
        
        Return controls
    End Function
    
    ''' <summary>
    ''' 设置控件值
    ''' </summary>
    Protected Sub SetControlValue(control As Control, value As String)
        Select Case control.GetType().Name
            Case "TextBox"
                DirectCast(control, TextBox).Text = value
            Case "ComboBox"
                DirectCast(control, ComboBox).Text = value
            Case "CheckBox"
                DirectCast(control, CheckBox).Checked = (value = "1")
            Case "RadioButton"
                DirectCast(control, RadioButton).Checked = (value = "1")
            Case "DateTimePicker"
                If DateTime.TryParse(value, Nothing) Then
                    DirectCast(control, DateTimePicker).Value = DateTime.Parse(value)
                End If
            Case "RichTextBox"
                DirectCast(control, RichTextBox).Text = value
            Case "DataGridView"
                ' DataGridView的数据通过DataSource设置，这里不处理字符串值
                ' 如果需要设置DataGridView的数据，应该在模板中单独处理
                ' 确保DataGridView有正确的DataSource
                Dim dgv As DataGridView = DirectCast(control, DataGridView)
                If dgv.DataSource Is Nothing Then
                    ' 如果没有数据源，创建一个空的DataTable
                    Dim dt As New DataTable()
                    dgv.DataSource = dt
                End If
        End Select
    End Sub
    
    ''' <summary>
    ''' 获取控件值
    ''' </summary>
    Protected Function GetControlValue(control As Control) As String
        Select Case control.GetType().Name
            Case "TextBox"
                Return DirectCast(control, TextBox).Text
            Case "ComboBox"
                Return DirectCast(control, ComboBox).Text
            Case "CheckBox"
                Return If(DirectCast(control, CheckBox).Checked, "1", "0")
            Case "RadioButton"
                Return If(DirectCast(control, RadioButton).Checked, "1", "0")
            Case "DateTimePicker"
                Return DirectCast(control, DateTimePicker).Value.ToString("yyyy-MM-dd HH:mm:ss")
            Case "RichTextBox"
                Return DirectCast(control, RichTextBox).Text
            Case "DataGridView"
                ' DataGridView的数据通过DataSource获取，返回空字符串
                ' 实际的数据提取应该在模板中单独处理
                Return ""
            Case Else
                Return ""
        End Select
    End Function
    
    ''' <summary>
    ''' 创建控件组 - 基础版本，仅用于现有控件的布局调整
    ''' </summary>
    Protected Sub CreateControlGroup(tabPage As TabPage, controls() As (String, String, Control), ByRef y As Integer)
        For Each controlInfo In controls
            Dim label As New Label With {
                .Text = controlInfo.Item1 & ":",
                .Location = New Point(20, y),
                .AutoSize = True,
                .Font = New Font("微软雅黑", 9)
            }
            tabPage.Controls.Add(label)
            
            Dim control As Control = controlInfo.Item3
            control.Location = New Point(150, y)
            control.Tag = controlInfo.Item2
            control.Font = New Font("微软雅黑", 9)
            
            ' 设置控件大小
            If TypeOf control Is TextBox Then
                control.Size = New Size(200, 23)
            ElseIf TypeOf control Is ComboBox Then
                control.Size = New Size(200, 23)
                DirectCast(control, ComboBox).DropDownStyle = ComboBoxStyle.DropDownList
            ElseIf TypeOf control Is DateTimePicker Then
                control.Size = New Size(200, 23)
            ElseIf TypeOf control Is RichTextBox Then
                control.Size = New Size(400, 100)
                y += 80
            End If
            
            tabPage.Controls.Add(control)
            y += 30
        Next
        
        ' 在创建完控件组后，检查是否需要调整TabPage大小
        AdjustTabPageSizeIfNeeded(tabPage)
    End Sub
    
    ''' <summary>
    ''' 检查并调整TabPage大小
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    Private Sub AdjustTabPageSizeIfNeeded(tabPage As TabPage)
        Try
            ' 检查TabPage是否有父容器
            If tabPage.Parent IsNot Nothing AndAlso TypeOf tabPage.Parent Is TabControl Then
                Dim tabControl As TabControl = DirectCast(tabPage.Parent, TabControl)
                Dim parentForm As Form = tabControl.FindForm()
                
                If parentForm IsNot Nothing Then
                    ' 使用工具类调整TabPage大小
                    Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, parentForm)
                End If
            End If
        Catch ex As Exception
            ' 记录错误但不影响主要功能
            Utils.LogUtil.LogError("调整TabPage大小失败", ex)
        End Try
    End Sub
End Class 
''' <summary>
''' TabPage大小自动调整功能 - 控件生成代码
''' 此文件包含用于生成控件的代码，供开发参考使用
''' 实际项目中建议使用设计器创建控件，然后调用TabPageSizeAdjuster进行大小调整
''' </summary>

#Region "控件生成代码 - 仅供参考"

''' <summary>
''' 生成基础控件组的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="y">起始Y坐标</param>
Public Sub GenerateBasicControls(tabPage As TabPage, ByRef y As Integer)
    ' 生成标签和文本框控件
    For i As Integer = 1 To 5
        ' 生成标签
        Dim lbl As New Label With {
            .Text = $"字段{i}:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lbl)
        
        ' 生成文本框
        Dim txt As New TextBox With {
            .Location = New Point(150, y),
            .Size = New Size(200, 23),
            .Tag = $"Field_{i}",
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(txt)
        
        y += 30
    Next
    
    ' 生成下拉框
    Dim cbo As New ComboBox With {
        .Location = New Point(20, y),
        .Size = New Size(200, 23),
        .Tag = "Combo_1",
        .Font = New Font("微软雅黑", 9),
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Items = {"选项1", "选项2", "选项3"}
    }
    tabPage.Controls.Add(cbo)
    
    y += 30
    
    ' 生成日期选择器
    Dim dtp As New DateTimePicker With {
        .Location = New Point(20, y),
        .Size = New Size(200, 23),
        .Tag = "Date_1",
        .Font = New Font("微软雅黑", 9)
    }
    tabPage.Controls.Add(dtp)
    
    y += 30
    
    ' 生成富文本框
    Dim rtb As New RichTextBox With {
        .Location = New Point(20, y),
        .Size = New Size(400, 100),
        .Tag = "RichText_1",
        .Font = New Font("微软雅黑", 9)
    }
    tabPage.Controls.Add(rtb)
    
    y += 120
End Sub

''' <summary>
''' 生成多列布局控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="y">起始Y坐标</param>
''' <param name="columns">列数</param>
Public Sub GenerateMultiColumnControls(tabPage As TabPage, ByRef y As Integer, columns As Integer)
    Dim columnWidth As Integer = 350
    Dim currentColumn As Integer = 0
    Dim startY As Integer = y
    
    ' 生成多列控件
    For i As Integer = 1 To 8
        Dim x As Integer = 20 + (currentColumn * columnWidth)
        
        ' 生成标签
        Dim lbl As New Label With {
            .Text = $"字段{i}:",
            .Location = New Point(x, y),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lbl)
        
        ' 生成文本框
        Dim txt As New TextBox With {
            .Location = New Point(x + 130, y),
            .Size = New Size(200, 23),
            .Tag = $"Field_{i}",
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(txt)
        
        ' 移动到下一列或下一行
        currentColumn += 1
        If currentColumn >= columns Then
            currentColumn = 0
            y += 30
        End If
    Next
    
    ' 如果最后一行没有填满，也要换行
    If currentColumn > 0 Then
        y += 30
    End If
End Sub

''' <summary>
''' 生成网格布局控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="y">起始Y坐标</param>
''' <param name="rows">行数</param>
''' <param name="cols">列数</param>
Public Sub GenerateGridControls(tabPage As TabPage, ByRef y As Integer, rows As Integer, cols As Integer)
    Dim cellWidth As Integer = 300
    Dim cellHeight As Integer = 30
    Dim startY As Integer = y
    
    For row As Integer = 0 To rows - 1
        For col As Integer = 0 To cols - 1
            Dim index As Integer = row * cols + col + 1
            Dim x As Integer = 20 + (col * cellWidth)
            Dim currentY As Integer = startY + (row * cellHeight)
            
            ' 生成标签
            Dim lbl As New Label With {
                .Text = $"字段{index}:",
                .Location = New Point(x, currentY),
                .AutoSize = True,
                .Font = New Font("微软雅黑", 9)
            }
            tabPage.Controls.Add(lbl)
            
            ' 生成文本框
            Dim txt As New TextBox With {
                .Location = New Point(x + 130, currentY),
                .Size = New Size(150, 23),
                .Tag = $"Field_{index}",
                .Font = New Font("微软雅黑", 9)
            }
            tabPage.Controls.Add(txt)
        Next
    Next
    
    ' 更新Y坐标
    y = startY + (rows * cellHeight) + 20
End Sub

''' <summary>
''' 生成案件基本信息控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="y">起始Y坐标</param>
Public Sub GenerateCaseBasicInfoControls(tabPage As TabPage, ByRef y As Integer)
    ' 案件基本信息控件
    Dim controls() As (String, String, Control) = {
        ("案件编号", "CaseNo", New TextBox),
        ("案件名称", "CaseName", New TextBox),
        ("申请日期", "ApplyDate", New DateTimePicker),
        ("申请人", "Applicant", New TextBox),
        ("联系电话", "Phone", New TextBox),
        ("电子邮箱", "Email", New TextBox)
    }
    
    ' 使用2列布局生成控件
    GenerateControlsFromArray(tabPage, controls, y, 2)
End Sub

''' <summary>
''' 生成产品信息控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="y">起始Y坐标</param>
Public Sub GenerateProductInfoControls(tabPage As TabPage, ByRef y As Integer)
    ' 产品信息控件
    Dim controls() As (String, String, Control) = {
        ("产品代码", "ProductCode", New TextBox),
        ("产品名称", "ProductName", New TextBox),
        ("产品类型", "ProductType", New ComboBox),
        ("规格型号", "Specification", New TextBox),
        ("计量单位", "Unit", New ComboBox),
        ("产品描述", "Description", New RichTextBox)
    }
    
    ' 使用2列布局生成控件
    GenerateControlsFromArray(tabPage, controls, y, 2)
End Sub

''' <summary>
''' 从控件数组生成控件的通用方法
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="controls">控件信息数组</param>
''' <param name="y">起始Y坐标</param>
''' <param name="columns">列数</param>
Private Sub GenerateControlsFromArray(tabPage As TabPage, controls() As (String, String, Control), ByRef y As Integer, columns As Integer)
    Dim columnWidth As Integer = 350
    Dim currentColumn As Integer = 0
    Dim startY As Integer = y
    
    For i As Integer = 0 To controls.Length - 1
        Dim controlInfo = controls(i)
        Dim x As Integer = 20 + (currentColumn * columnWidth)
        
        ' 生成标签
        Dim label As New Label With {
            .Text = controlInfo.Item1 & ":",
            .Location = New Point(x, y),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(label)
        
        ' 生成控件
        Dim control As Control = controlInfo.Item3
        control.Location = New Point(x + 130, y)
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
            control.Size = New Size(300, 100)
            y += 80
        End If
        
        tabPage.Controls.Add(control)
        
        ' 移动到下一列或下一行
        currentColumn += 1
        If currentColumn >= columns Then
            currentColumn = 0
            y += 30
        End If
    Next
    
    ' 如果最后一行没有填满，也要换行
    If currentColumn > 0 Then
        y += 30
    End If
End Sub

#End Region

#Region "使用说明"

''' <summary>
''' 使用说明：
''' 1. 此文件包含控件生成的代码示例，仅供开发参考
''' 2. 实际项目中建议使用设计器创建控件，然后调用TabPageSizeAdjuster进行大小调整
''' 3. 如果需要在代码中生成控件，可以参考此文件中的方法
''' 4. 生成控件后，记得调用 Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, parentForm) 进行大小调整
''' </summary>

''' <summary>
''' 示例：在设计器创建的TabPage中调用大小调整
''' </summary>
''' <param name="tabPage">设计器创建的TabPage</param>
Public Sub AdjustDesignerTabPage(tabPage As TabPage)
    Try
        ' 获取父窗体
        Dim parentForm As Form = tabPage.FindForm()
        If parentForm IsNot Nothing Then
            ' 调用大小调整功能
            Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, parentForm)
        End If
    Catch ex As Exception
        Utils.LogUtil.LogError("调整设计器TabPage大小失败", ex)
    End Try
End Sub

''' <summary>
''' 示例：批量调整所有TabPage大小
''' </summary>
''' <param name="tabControl">TabControl</param>
Public Sub AdjustAllDesignerTabPages(tabControl As TabControl)
    Try
        Dim parentForm As Form = tabControl.FindForm()
        If parentForm IsNot Nothing Then
            Utils.TabPageSizeAdjuster.AdjustAllTabPages(tabControl, parentForm)
        End If
    Catch ex As Exception
        Utils.LogUtil.LogError("批量调整设计器TabPage大小失败", ex)
    End Try
End Sub

#End Region 
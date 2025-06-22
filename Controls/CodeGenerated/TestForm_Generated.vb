''' <summary>
''' 测试窗体控件生成代码 - 仅供参考
''' 此文件包含用于生成测试窗体控件的代码，供开发参考使用
''' 实际项目中建议使用设计器创建控件
''' </summary>

#Region "测试窗体控件生成代码"

''' <summary>
''' 生成测试窗体基础控件的代码示例
''' </summary>
''' <param name="form">窗体</param>
Public Sub GenerateTestFormControls(form As Form)
    ' 生成信息标签
    Dim lblInfo As New Label With {
        .Text = "TabPage大小调整功能测试 - 点击按钮测试不同场景",
        .Location = New Point(10, 10),
        .Font = New Font("微软雅黑", 10, FontStyle.Bold),
        .AutoSize = True
    }
    form.Controls.Add(lblInfo)
    
    ' 生成测试按钮
    Dim btnAddControls As New Button With {
        .Text = "添加更多控件",
        .Location = New Point(10, 40),
        .Size = New Size(120, 30)
    }
    form.Controls.Add(btnAddControls)
    
    Dim btnReset As New Button With {
        .Text = "重置TabPage",
        .Location = New Point(140, 40),
        .Size = New Size(120, 30)
    }
    form.Controls.Add(btnReset)
    
    Dim btnTestScroll As New Button With {
        .Text = "测试滚动条",
        .Location = New Point(270, 40),
        .Size = New Size(120, 30)
    }
    form.Controls.Add(btnTestScroll)
    
    ' 生成TabControl
    Dim tabControl As New TabControl With {
        .Location = New Point(10, 80),
        .Size = New Size(1160, 670),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    }
    form.Controls.Add(tabControl)
End Sub

''' <summary>
''' 生成测试TabPage控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="pageIndex">页面索引</param>
Public Sub GenerateTestTabPageControls(tabPage As TabPage, pageIndex As Integer)
    Dim y As Integer = 20
    
    ' 生成基础控件
    For j As Integer = 1 To 8
        ' 生成标签
        Dim lbl As New Label With {
            .Text = $"字段{j}:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lbl)
        
        ' 生成文本框
        Dim txt As New TextBox With {
            .Location = New Point(120, y),
            .Size = New Size(200, 23),
            .Tag = $"Field_{pageIndex}_{j}"
        }
        tabPage.Controls.Add(txt)
        
        y += 30
    Next
    
    ' 生成下拉框
    Dim cbo As New ComboBox With {
        .Location = New Point(20, y),
        .Size = New Size(200, 23),
        .Tag = $"Combo_{pageIndex}",
        .Items = {"选项1", "选项2", "选项3"}
    }
    tabPage.Controls.Add(cbo)
    
    y += 30
    
    ' 生成日期选择器
    Dim dtp As New DateTimePicker With {
        .Location = New Point(20, y),
        .Size = New Size(200, 23),
        .Tag = $"Date_{pageIndex}"
    }
    tabPage.Controls.Add(dtp)
    
    y += 30
    
    ' 生成富文本框
    Dim rtb As New RichTextBox With {
        .Location = New Point(20, y),
        .Size = New Size(400, 100),
        .Tag = $"RichText_{pageIndex}"
    }
    tabPage.Controls.Add(rtb)
End Sub

''' <summary>
''' 生成动态添加控件的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="count">控件数量</param>
Public Sub GenerateDynamicControls(tabPage As TabPage, count As Integer)
    ' 获取当前最大的Y坐标
    Dim maxY As Integer = 0
    For Each control As Control In tabPage.Controls
        If control.Location.Y + control.Height > maxY Then
            maxY = control.Location.Y + control.Height
        End If
    Next
    
    ' 添加新控件
    For i As Integer = 1 To count
        ' 生成标签
        Dim lbl As New Label With {
            .Text = $"新增字段{i}:",
            .Location = New Point(20, maxY + (i - 1) * 30),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lbl)
        
        ' 生成文本框
        Dim txt As New TextBox With {
            .Location = New Point(120, maxY + (i - 1) * 30),
            .Size = New Size(200, 23),
            .Tag = $"NewField_{i}"
        }
        tabPage.Controls.Add(txt)
    Next
End Sub

''' <summary>
''' 生成大量控件用于测试滚动条的代码示例
''' </summary>
''' <param name="tabPage">TabPage</param>
''' <param name="count">控件数量</param>
Public Sub GenerateManyControlsForScrollTest(tabPage As TabPage, count As Integer)
    ' 获取当前最大的Y坐标
    Dim maxY As Integer = 0
    For Each control As Control In tabPage.Controls
        If control.Location.Y + control.Height > maxY Then
            maxY = control.Location.Y + control.Height
        End If
    Next
    
    ' 添加大量控件
    For i As Integer = 1 To count
        ' 生成标签
        Dim lbl As New Label With {
            .Text = $"测试字段{i}:",
            .Location = New Point(20, maxY + (i - 1) * 30),
            .AutoSize = True,
            .Font = New Font("微软雅黑", 9)
        }
        tabPage.Controls.Add(lbl)
        
        ' 生成文本框
        Dim txt As New TextBox With {
            .Location = New Point(120, maxY + (i - 1) * 30),
            .Size = New Size(200, 23),
            .Tag = $"TestField_{i}"
        }
        tabPage.Controls.Add(txt)
    Next
End Sub

#End Region

#Region "使用说明"

''' <summary>
''' 使用说明：
''' 1. 此文件包含测试窗体控件生成的代码示例，仅供开发参考
''' 2. 实际项目中建议使用设计器创建控件
''' 3. 如果需要在代码中生成测试控件，可以参考此文件中的方法
''' 4. 生成控件后，记得调用大小调整功能
''' </summary>

''' <summary>
''' 示例：创建完整的测试窗体
''' </summary>
Public Sub CreateCompleteTestForm()
    ' 创建窗体
    Dim form As New Form With {
        .Text = "TabPage大小调整功能测试",
        .Width = 1200,
        .Height = 800,
        .FormBorderStyle = FormBorderStyle.Sizable,
        .StartPosition = FormStartPosition.CenterScreen
    }
    
    ' 生成窗体控件
    GenerateTestFormControls(form)
    
    ' 获取TabControl
    Dim tabControl As TabControl = form.Controls.OfType(Of TabControl)().FirstOrDefault()
    If tabControl IsNot Nothing Then
        ' 创建测试TabPage
        For i As Integer = 1 To 3
            Dim tabPage As New TabPage($"测试页面{i}")
            GenerateTestTabPageControls(tabPage, i)
            tabControl.TabPages.Add(tabPage)
            
            ' 自动调整TabPage大小
            Utils.TabPageSizeAdjuster.AdjustTabPageSize(tabPage, form)
        Next
    End If
    
    ' 显示窗体
    form.ShowDialog()
End Sub

#End Region 
' DataGridView数据保存功能测试窗体
Imports System.Windows.Forms

Public Class TestDataGridViewForm
    Private tabControl As TabControl
    Private btnTest As Button
    Private btnClear As Button
    Private lblResult As Label

    Public Sub New()
        Me.Text = "DataGridView数据保存功能测试"
        Me.Width = 1000
        Me.Height = 700
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        InitControls()
    End Sub

    Private Sub InitControls()
        ' 创建TabControl
        tabControl = New TabControl With {.Location = New Point(20, 20), .Size = New Size(900, 500)}
        CreateTestTabPages()
        Me.Controls.Add(tabControl)

        ' 测试按钮
        btnTest = New Button With {.Text = "测试数据保存", .Location = New Point(20, 540), .Width = 120, .Height = 30}
        AddHandler btnTest.Click, AddressOf btnTest_Click
        Me.Controls.Add(btnTest)

        ' 清空按钮
        btnClear = New Button With {.Text = "清空数据", .Location = New Point(160, 540), .Width = 100, .Height = 30}
        AddHandler btnClear.Click, AddressOf btnClear_Click
        Me.Controls.Add(btnClear)

        ' 结果显示标签
        lblResult = New Label With {.Text = "请在DataGridView中输入数据，然后点击测试按钮", .Location = New Point(20, 580), .AutoSize = True}
        Me.Controls.Add(lblResult)
    End Sub

    Private Sub CreateTestTabPages()
        For i = 1 To 3
            Dim tabPage As New TabPage($"测试页{i}")
            CreateTestControls(tabPage, i)
            tabControl.TabPages.Add(tabPage)
        Next
    End Sub

    Private Sub CreateTestControls(tabPage As TabPage, pageIndex As Integer)
        ' 创建一些单项控件
        Dim y As Integer = 20
        
        For j = 1 To 3
            Dim lbl As New Label With {
                .Text = $"测试字段{j}:",
                .Location = New Point(20, y),
                .AutoSize = True
            }
            tabPage.Controls.Add(lbl)
            
            Dim txt As New TextBox With {
                .Location = New Point(120, y),
                .Width = 200,
                .Tag = $"TestField_{pageIndex}_{j}"
            }
            tabPage.Controls.Add(txt)
            
            y += 30
        Next

        ' 创建DataGridView
        Dim dgv As New DataGridView With {
            .Name = $"dgvTest_{pageIndex}",
            .Location = New Point(20, y + 20),
            .Size = New Size(800, 200),
            .AllowUserToAddRows = True,
            .AllowUserToDeleteRows = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        }
        
        ' 添加列
        dgv.Columns.Add("ItemName", "项目名称")
        dgv.Columns.Add("ItemValue", "项目值")
        dgv.Columns.Add("LastUpdate", "最后更新时间")
        dgv.Columns.Add("ReviewTime", "审查时间")
        dgv.Columns.Add("Status", "状态")
        dgv.Columns.Add("Reviewer", "审查人员")
        
        ' 设置默认值
        dgv.Columns("LastUpdate").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
        dgv.Columns("ReviewTime").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
        dgv.Columns("Status").DefaultCellStyle.NullValue = "新登录"
        dgv.Columns("Reviewer").DefaultCellStyle.NullValue = "测试用户"
        
        tabPage.Controls.Add(dgv)
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs)
        Try
            ' 提取数据
            Dim tabData As Dictionary(Of Integer, Dictionary(Of String, String)) = BusinessLogic.CaseManager.ExtractModifiedData(tabControl)
            Dim tabGridData As Dictionary(Of Integer, List(Of Dictionary(Of String, String))) = ExtractGridData(tabControl)
            
            ' 显示提取的数据
            Dim result As String = "数据提取结果：" & vbCrLf
            result += $"单项数据：{tabData.Count} 个标签页" & vbCrLf
            result += $"表格数据：{tabGridData.Count} 个标签页" & vbCrLf
            
            For Each kvp In tabData
                result += $"  标签页{kvp.Key}：{kvp.Value.Count} 个字段" & vbCrLf
            Next
            
            For Each kvp In tabGridData
                result += $"  标签页{kvp.Key}：{kvp.Value.Count} 行表格数据" & vbCrLf
            Next
            
            lblResult.Text = result
            
            ' 模拟保存（不实际保存到数据库）
            MessageBox.Show("数据提取成功！" & vbCrLf & result, "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            lblResult.Text = $"测试失败：{ex.Message}"
            MessageBox.Show($"测试过程中发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        ' 清空所有数据
        For Each tabPage As TabPage In tabControl.TabPages
            For Each ctrl As Control In GetAllControls(tabPage)
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Clear()
                ElseIf TypeOf ctrl Is DataGridView Then
                    CType(ctrl, DataGridView).Rows.Clear()
                End If
            Next
        Next
        
        lblResult.Text = "数据已清空"
    End Sub

    ''' <summary>
    ''' 提取所有TabPage中DataGridView的数据
    ''' </summary>
    Private Function ExtractGridData(tabControl As TabControl) As Dictionary(Of Integer, List(Of Dictionary(Of String, String)))
        Dim result As New Dictionary(Of Integer, List(Of Dictionary(Of String, String)))
        
        For i As Integer = 0 To tabControl.TabPages.Count - 1
            Dim tabPage As TabPage = tabControl.TabPages(i)
            Dim rows As New List(Of Dictionary(Of String, String))
            
            For Each ctrl As Control In GetAllControls(tabPage)
                If TypeOf ctrl Is DataGridView Then
                    Dim dgv As DataGridView = CType(ctrl, DataGridView)
                    
                    For Each row As DataGridViewRow In dgv.Rows
                        If Not row.IsNewRow Then
                            Dim rowData As New Dictionary(Of String, String)
                            
                            For Each col As DataGridViewColumn In dgv.Columns
                                Dim cellValue As Object = row.Cells(col.Index).Value
                                rowData(col.Name) = If(cellValue?.ToString(), "")
                            Next
                            
                            If Not String.IsNullOrEmpty(rowData("ItemName")) Then
                                rows.Add(rowData)
                            End If
                        End If
                    Next
                End If
            Next
            
            If rows.Count > 0 Then
                result(i) = rows
            End If
        Next
        
        Return result
    End Function

    ''' <summary>
    ''' 递归获取所有子控件
    ''' </summary>
    Private Function GetAllControls(container As Control) As List(Of Control)
        Dim controls As New List(Of Control)
        
        For Each ctrl As Control In container.Controls
            controls.Add(ctrl)
            controls.AddRange(GetAllControls(ctrl))
        Next
        
        Return controls
    End Function
End Class 
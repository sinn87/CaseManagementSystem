' 案件管理业务逻辑
Imports System.Collections.Generic
Imports System.Data.OleDb

Public Class CaseManager
    ''' <summary>
    ''' 创建新案件并保存详细信息
    ''' </summary>
    ''' <param name="caseType">案件类型</param>
    ''' <param name="tabData">标签页数据字典，Key为标签页索引，Value为字段数据字典</param>
    ''' <param name="gridData">DataGridView数据字典，Key为DataGridView名称，Value为行数据列表</param>
    ''' <param name="currentUser">当前用户ID</param>
    ''' <returns>是否保存成功</returns>
    Public Shared Function CreateNewCase(caseType As String, tabData As Dictionary(Of Integer, Dictionary(Of String, String)), gridData As Dictionary(Of String, List(Of Dictionary(Of String, String))), currentUser As String) As Boolean
        Dim connection As OleDbConnection = Nothing
        Dim transaction As OleDbTransaction = Nothing
        
        Try
            ' 1. 创建案件主记录
            Dim caseInfo As New CaseInfo With {
                .CaseType = caseType,
                .Status = 1, ' 新登录状态
                .LastUpdate = DateTime.Now,
                .IsTerminated = 0,
                .CreateTime = DateTime.Now,
                .CreateUser = currentUser
            }
            
            ' 2. 准备详细信息数据
            Dim caseDetails As New List(Of CaseDetail)()
            Dim reviewLogs As New List(Of ReviewLog)()
            
            For Each kvp In tabData
                Dim tabIndex As Integer = kvp.Key
                Dim fieldData As Dictionary(Of String, String) = kvp.Value
                
                ' 检查是否有数据被修改（字段数据或DGV数据）
                Dim hasFieldData As Boolean = fieldData.Count > 0
                Dim hasGridData As Boolean = False
                
                ' 检查该Tab页是否有对应的DGV数据
                For Each gridKvp In gridData
                    Dim dgvTabIndex As Integer = GetTabIndexFromDgvName(gridKvp.Key)
                    If dgvTabIndex = tabIndex AndAlso gridKvp.Value.Count > 0 Then
                        hasGridData = True
                        Exit For
                    End If
                Next
                
                ' 如果有字段数据或DGV数据，则准备审查记录
                If hasFieldData OrElse hasGridData Then
                    ' 准备字段数据
                    If hasFieldData Then
                        For Each fieldKvp In fieldData
                            Dim detail As New CaseDetail With {
                                .TabIndex = tabIndex,
                                .FieldNo = fieldKvp.Key,
                                .FieldValue = fieldKvp.Value,
                                .FieldStatus = "新登录",
                                .CreateTime = DateTime.Now
                            }
                            caseDetails.Add(detail)
                        Next
                    End If
                    
                    ' 准备审查记录
                    Dim reviewLog As New ReviewLog With {
                        .TabIndex = tabIndex,
                        .ReviewerID = currentUser,
                        .ReviewStatus = "新登录",
                        .ReviewTime = DateTime.Now
                    }
                    reviewLogs.Add(reviewLog)
                End If
            Next
            
            ' 3. 开始数据库事务
            connection = DbHelper.GetConnection()
            transaction = DbHelper.BeginTransaction(connection)
            
            ' 4. 在事务中创建案件主记录
            Dim caseId As Integer = CaseRepository.CreateCaseWithTransaction(transaction, caseInfo)
            
            ' 5. 在事务中批量保存详细信息
            If caseDetails.Count > 0 Then
                ' 设置案件ID
                For Each detail In caseDetails
                    detail.CaseID = caseId
                Next
                CaseRepository.SaveCaseDetailsWithTransaction(transaction, caseDetails)
            End If
            
            ' 6. 在事务中保存所有DGV数据
            For Each kvp In gridData
                Dim tableName = kvp.Key
                Dim rows = kvp.Value
                If rows Is Nothing OrElse rows.Count = 0 Then
                    Continue For ' 跳过无数据的DGV
                End If
                For Each row In rows
                    If Not row.ContainsKey("caseID") Then row("caseID") = caseId
                Next
                If Not CaseRepository.SaveGridDataWithTransaction(transaction, tableName, rows, currentUser) Then
                    Throw New Exception($"保存表[{tableName}]数据失败")
                End If
            Next
            
            ' 7. 在事务中批量保存审查记录
            For Each reviewLog In reviewLogs
                reviewLog.CaseID = caseId
                CaseRepository.CreateReviewLogWithTransaction(transaction, reviewLog)
            Next
            
            ' 8. 提交事务
            transaction.Commit()
            
            Return True
            
        Catch ex As Exception
            ' 9. 回滚事务
            If transaction IsNot Nothing Then
                Try
                    transaction.Rollback()
                Catch rollbackEx As Exception
                    Utils.LogUtil.LogError("事务回滚失败", rollbackEx)
                End Try
            End If
            
            ' 记录错误日志
            Utils.LogUtil.LogError("创建新案件失败", ex)
            Return False
            
        Finally
            ' 10. 清理资源
            If transaction IsNot Nothing Then
                transaction.Dispose()
            End If
            If connection IsNot Nothing Then
                connection.Dispose()
            End If
        End Try
    End Function
    
    ''' <summary>
    ''' 从DataGridView名称中提取标签页索引
    ''' </summary>
    ''' <param name="dgvName">DataGridView名称</param>
    ''' <returns>标签页索引</returns>
    Private Shared Function GetTabIndexFromDgvName(dgvName As String) As Integer
        ' 假设命名规则为：dgvItems_1, dgvProducts_2 等
        Dim parts As String() = dgvName.Split("_"c)
        If parts.Length >= 2 Then
            Dim indexStr As String = parts(parts.Length - 1)
            Dim index As Integer
            If Integer.TryParse(indexStr, index) Then
                Return index - 1 ' 转换为0基索引
            End If
        End If
        Return 0
    End Function
    
    ''' <summary>
    ''' 从TabControl中提取有修改的数据
    ''' </summary>
    ''' <param name="tabControl">TabControl控件</param>
    ''' <returns>标签页数据字典</returns>
    Public Shared Function ExtractModifiedData(tabControl As TabControl) As Dictionary(Of Integer, Dictionary(Of String, String))
        Dim result As New Dictionary(Of Integer, Dictionary(Of String, String))()
        
        For i As Integer = 0 To tabControl.TabPages.Count - 1
            Dim tabPage As TabPage = tabControl.TabPages(i)
            Dim fieldData As New Dictionary(Of String, String)()
            
            ' 遍历标签页中的所有控件
            For Each control As Control In GetAllControls(tabPage)
                ' 检查控件是否有Tag属性（对应数据库字段）
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
                result(i) = fieldData
            End If
        Next
        
        Return result
    End Function
    
    ''' <summary>
    ''' 提取所有TabPage中DataGridView的数据
    ''' </summary>
    ''' <param name="tabControl">TabControl控件</param>
    ''' <returns>DataGridView名称和数据的字典</returns>
    Public Shared Function ExtractGridData(tabControl As TabControl) As Dictionary(Of String, List(Of Dictionary(Of String, String)))
        Dim result As New Dictionary(Of String, List(Of Dictionary(Of String, String)))
        
        For i As Integer = 0 To tabControl.TabPages.Count - 1
            Dim tabPage As TabPage = tabControl.TabPages(i)
            
            ' 遍历标签页中的所有控件
            For Each ctrl As Control In GetAllControls(tabPage)
                If TypeOf ctrl Is DataGridView Then
                    Dim dgv As DataGridView = CType(ctrl, DataGridView)
                    Dim rows As New List(Of Dictionary(Of String, String))
                    
                    ' 遍历当前DataGridView的所有行
                    For Each row As DataGridViewRow In dgv.Rows
                        If Not row.IsNewRow Then
                            Dim rowData As New Dictionary(Of String, String)
                            
                            ' 提取每一列的数据
                            For Each col As DataGridViewColumn In dgv.Columns
                                Dim cellValue As Object = row.Cells(col.Index).Value
                                rowData(col.Name) = If(cellValue?.ToString(), "")
                            Next
                            
                            ' 检查是否有有效数据（至少项目名称不为空）
                            If Not String.IsNullOrEmpty(rowData("ItemName")) Then
                                rows.Add(rowData)
                            End If
                        End If
                    Next
                    
                    ' 如果该DataGridView有数据，则添加到结果中
                    If rows.Count > 0 Then
                        result(dgv.Name) = rows
                    End If
                End If
            Next
        Next
        
        Return result
    End Function
    
    ''' <summary>
    ''' 获取DataGridView在标签页中的索引（用于区分多个DataGridView）
    ''' </summary>
    ''' <param name="tabPage">标签页</param>
    ''' <param name="targetDgv">目标DataGridView</param>
    ''' <returns>DataGridView索引</returns>
    Private Shared Function GetDataGridViewIndex(tabPage As TabPage, targetDgv As DataGridView) As Integer
        Dim index As Integer = 0
        
        For Each ctrl As Control In GetAllControls(tabPage)
            If TypeOf ctrl Is DataGridView Then
                If ctrl Is targetDgv Then
                    Return index
                End If
                index += 1
            End If
        Next
        
        Return 0
    End Function
    
    ''' <summary>
    ''' 保存案件数据
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="tabControl">TabControl控件</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否保存成功</returns>
    Public Shared Function SaveCaseData(caseId As Integer, tabControl As TabControl, currentUser As String) As Boolean
        Try
            ' 提取修改的数据
            Dim tabData As Dictionary(Of Integer, Dictionary(Of String, String)) = ExtractModifiedData(tabControl)
            
            If tabData.Count = 0 Then
                Return True ' 没有数据需要保存
            End If
            
            ' 保存案件详细信息
            Dim caseDetails As New List(Of CaseDetail)()
            
            For Each kvp In tabData
                Dim tabIndex As Integer = kvp.Key
                Dim fieldData As Dictionary(Of String, String) = kvp.Value
                
                For Each fieldKvp In fieldData
                    Dim detail As New CaseDetail With {
                        .CaseID = caseId,
                        .TabIndex = tabIndex,
                        .FieldNo = fieldKvp.Key,
                        .FieldValue = fieldKvp.Value,
                        .FieldStatus = "已保存",
                        .CreateTime = DateTime.Now
                    }
                    caseDetails.Add(detail)
                Next
            Next
            
            ' 批量保存详细信息
            If caseDetails.Count > 0 Then
                CaseRepository.SaveCaseDetails(caseDetails)
            End If
            
            ' 更新案件最后修改时间
            CaseRepository.UpdateCaseLastUpdate(caseId, DateTime.Now)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("保存案件数据失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 使用模板保存案件数据
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="templateData">模板数据字典</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否保存成功</returns>
    Public Shared Function SaveCaseDataWithTemplate(caseId As Integer, templateData As Dictionary(Of Integer, Dictionary(Of String, String)), currentUser As String) As Boolean
        Try
            If templateData Is Nothing OrElse templateData.Count = 0 Then
                Return True ' 没有数据需要保存
            End If
            
            ' 保存案件详细信息
            Dim caseDetails As New List(Of CaseDetail)()
            
            For Each kvp In templateData
                Dim tabIndex As Integer = kvp.Key
                Dim fieldData As Dictionary(Of String, String) = kvp.Value
                
                For Each fieldKvp In fieldData
                    Dim detail As New CaseDetail With {
                        .CaseID = caseId,
                        .TabIndex = tabIndex,
                        .FieldNo = fieldKvp.Key,
                        .FieldValue = fieldKvp.Value,
                        .FieldStatus = "已保存",
                        .CreateTime = DateTime.Now
                    }
                    caseDetails.Add(detail)
                Next
            Next
            
            ' 批量保存详细信息
            If caseDetails.Count > 0 Then
                CaseRepository.SaveCaseDetails(caseDetails)
            End If
            
            ' 更新案件最后修改时间
            CaseRepository.UpdateCaseLastUpdate(caseId, DateTime.Now)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("使用模板保存案件数据失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 提交审查
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="tabControl">TabControl控件</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否提交成功</returns>
    Public Shared Function SubmitForReview(caseId As Integer, tabControl As TabControl, currentUser As String) As Boolean
        Try
            ' 提取有数据的Tab
            Dim tabData As Dictionary(Of Integer, Dictionary(Of String, String)) = ExtractModifiedData(tabControl)
            
            If tabData.Count = 0 Then
                Return False ' 没有数据需要提交
            End If
            
            ' 为每个有数据的Tab创建审查记录
            For Each kvp In tabData
                Dim tabIndex As Integer = kvp.Key
                
                Dim reviewLog As New ReviewLog With {
                    .CaseID = caseId,
                    .TabIndex = tabIndex,
                    .ReviewerID = currentUser,
                    .ReviewStatus = "待审查",
                    .ReviewTime = DateTime.Now
                }
                
                CaseRepository.CreateReviewLog(reviewLog)
            Next
            
            ' 更新案件状态
            CaseRepository.UpdateCaseStatus(caseId, 2) ' 待审查状态
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("提交审查失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 同意审查
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="tabIndex">标签页索引</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否同意成功</returns>
    Public Shared Function ApproveReview(caseId As Integer, tabIndex As Integer, currentUser As String) As Boolean
        Try
            Dim reviewLog As New ReviewLog With {
                .CaseID = caseId,
                .TabIndex = tabIndex,
                .ReviewerID = currentUser,
                .ReviewStatus = "同意",
                .ReviewTime = DateTime.Now
            }
            
            CaseRepository.CreateReviewLog(reviewLog)
            
            ' 检查是否所有Tab都已同意
            If AllTabsApproved(caseId) Then
                CaseRepository.UpdateCaseStatus(caseId, 3) ' 审查完成状态
            End If
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("同意审查失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 不同意审查
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="tabIndex">标签页索引</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否不同意成功</returns>
    Public Shared Function RejectReview(caseId As Integer, tabIndex As Integer, currentUser As String) As Boolean
        Try
            Dim reviewLog As New ReviewLog With {
                .CaseID = caseId,
                .TabIndex = tabIndex,
                .ReviewerID = currentUser,
                .ReviewStatus = "不同意",
                .ReviewTime = DateTime.Now
            }
            
            CaseRepository.CreateReviewLog(reviewLog)
            
            ' 更新案件状态为需要修改
            CaseRepository.UpdateCaseStatus(caseId, 4) ' 需要修改状态
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("不同意审查失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 中止案件
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否中止成功</returns>
    Public Shared Function TerminateCase(caseId As Integer, currentUser As String) As Boolean
        Try
            ' 更新案件为中止状态
            CaseRepository.UpdateCaseTerminated(caseId, 1, currentUser)
            
            ' 创建中止记录
            Dim reviewLog As New ReviewLog With {
                .CaseID = caseId,
                .TabIndex = -1, ' 表示整个案件
                .ReviewerID = currentUser,
                .ReviewStatus = "中止",
                .ReviewTime = DateTime.Now
            }
            
            CaseRepository.CreateReviewLog(reviewLog)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("中止案件失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 恢复案件
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否恢复成功</returns>
    Public Shared Function RestoreCase(caseId As Integer, currentUser As String) As Boolean
        Try
            ' 更新案件为非中止状态
            CaseRepository.UpdateCaseTerminated(caseId, 0, currentUser)
            
            ' 创建恢复记录
            Dim reviewLog As New ReviewLog With {
                .CaseID = caseId,
                .TabIndex = -1, ' 表示整个案件
                .ReviewerID = currentUser,
                .ReviewStatus = "恢复",
                .ReviewTime = DateTime.Now
            }
            
            CaseRepository.CreateReviewLog(reviewLog)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("恢复案件失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 终了案件
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否终了成功</returns>
    Public Shared Function FinishCase(caseId As Integer, currentUser As String) As Boolean
        Try
            ' 检查是否所有Tab都已同意
            If Not AllTabsApproved(caseId) Then
                Return False
            End If
            
            ' 更新案件状态为终了
            CaseRepository.UpdateCaseStatus(caseId, 5) ' 终了状态
            
            ' 创建终了记录
            Dim reviewLog As New ReviewLog With {
                .CaseID = caseId,
                .TabIndex = -1, ' 表示整个案件
                .ReviewerID = currentUser,
                .ReviewStatus = "终了",
                .ReviewTime = DateTime.Now
            }
            
            CaseRepository.CreateReviewLog(reviewLog)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("终了案件失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 导出到Excel
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="caseInfo">案件信息</param>
    ''' <param name="caseDetails">案件详细信息</param>
    ''' <returns>是否导出成功</returns>
    Public Shared Function ExportToExcel(caseId As Integer, caseInfo As CaseInfo, caseDetails As List(Of CaseDetail)) As Boolean
        Try
            ' 这里实现Excel导出逻辑
            ' 可以使用EPPlus、NPOI等库
            ' 暂时返回True表示成功
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("导出Excel失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 导出到PDF
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="caseInfo">案件信息</param>
    ''' <param name="caseDetails">案件详细信息</param>
    ''' <returns>是否导出成功</returns>
    Public Shared Function ExportToPDF(caseId As Integer, caseInfo As CaseInfo, caseDetails As List(Of CaseDetail)) As Boolean
        Try
            ' 这里实现PDF导出逻辑
            ' 可以使用iTextSharp、PdfSharp等库
            ' 暂时返回True表示成功
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("导出PDF失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 检查是否所有Tab都已同意
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <returns>是否所有Tab都已同意</returns>
    Private Shared Function AllTabsApproved(caseId As Integer) As Boolean
        Try
            Dim reviewLogs As List(Of ReviewLog) = CaseRepository.GetReviewLogsByCaseId(caseId)
            
            ' 检查每个Tab的审查状态
            For i As Integer = 0 To 8
                Dim tabReviewLogs = reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
                If tabReviewLogs.Count > 0 Then
                    Dim lastReview = tabReviewLogs.OrderByDescending(Function(r) r.ReviewTime).First()
                    If lastReview.ReviewStatus <> "同意" Then
                        Return False
                    End If
                Else
                    Return False
                End If
            Next
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("检查Tab审查状态失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 递归获取所有控件（包括嵌套控件）
    ''' </summary>
    ''' <param name="container">容器控件</param>
    ''' <returns>所有控件列表</returns>
    Private Shared Function GetAllControls(container As Control) As List(Of Control)
        Dim controls As New List(Of Control)()
        
        For Each control As Control In container.Controls
            controls.Add(control)
            ' 递归获取子控件
            controls.AddRange(GetAllControls(control))
        Next
        
        Return controls
    End Function
    
    ''' <summary>
    ''' 获取控件的值
    ''' </summary>
    ''' <param name="control">控件</param>
    ''' <returns>控件值</returns>
    Private Shared Function GetControlValue(control As Control) As String
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
                Return DirectCast(control, DateTimePicker).Value.ToString("yyyy-MM-dd")
            Case "RichTextBox"
                Return DirectCast(control, RichTextBox).Text
            Case Else
                Return ""
        End Select
    End Function
End Class 
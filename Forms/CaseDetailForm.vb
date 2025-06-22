' 案件详细页面窗体
Imports System.Windows.Forms
Imports System.Drawing

Public Class CaseDetailForm
    ' 控件声明
    Private picHome As PictureBox
    Private lblCaseInfo As Label
    Private lblTerminated As Label
    Private btnView As Button
    Private btnEdit As Button
    Private btnReview As Button
    Private btnExport As Button
    Private btnTerminate As Button
    Private btnFinish As Button
    Private tabControl As TabControl
    Private panelTerminated As Panel
    
    ' 数据字段
    Private _caseId As Integer
    Private _currentUser As String
    Private _currentStatus As String
    Private _caseInfo As CaseInfo
    Private _caseDetails As List(Of CaseDetail)
    Private _reviewLogs As List(Of ReviewLog)
    
    ' 模板相关字段
    Private _currentTemplate As ITabTemplate
    
    ' 状态常量
    Private Const STATUS_VIEW As String = "阅览"
    Private Const STATUS_EDIT As String = "编辑"
    Private Const STATUS_REVIEW As String = "审查"
    Private Const STATUS_EXPORT As String = "导出"
    Private Const STATUS_TERMINATED As String = "中止"
    Private Const STATUS_FINISHED As String = "终了"
    
    Public Sub New(caseId As Integer, currentUser As String)
        _caseId = caseId
        _currentUser = currentUser
        Me.Text = "案件详细页面"
        Me.Width = 1200
        Me.Height = 800
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadCaseData()
        InitControls()
        SetStatus(_currentStatus)
    End Sub
    
    Private Sub LoadCaseData()
        Try
            ' 加载案件基本信息
            _caseInfo = CaseRepository.GetCaseById(_caseId)
            If _caseInfo Is Nothing Then
                MessageBox.Show("案件不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End If
            
            ' 加载案件详细信息
            _caseDetails = CaseRepository.GetCaseDetailsByCaseId(_caseId)
            
            ' 加载审查记录
            _reviewLogs = CaseRepository.GetReviewLogsByCaseId(_caseId)
            
            ' 确定当前状态
            _currentStatus = DetermineCurrentStatus()
            
            ' 清除之前的遮罩
            ClearReviewingOverlays()
            
        Catch ex As Exception
            MessageBox.Show($"加载案件数据失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("加载案件数据失败", ex)
            Me.Close()
        End Try
    End Sub
    
    Private Function DetermineCurrentStatus() As String
        ' 根据案件状态和用户权限确定当前状态
        If _caseInfo.IsTerminated = 1 Then
            Return STATUS_TERMINATED
        End If
        
        ' 检查是否所有Tab都已同意
        Dim allApproved As Boolean = True
        For i As Integer = 0 To 8
            Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
            If tabReviewLogs.Count > 0 Then
                Dim lastReview = tabReviewLogs.OrderByDescending(Function(r) r.ReviewTime).First()
                If lastReview.ReviewStatus <> "同意" Then
                    allApproved = False
                    Exit For
                End If
            Else
                allApproved = False
                Exit For
            End If
        Next
        
        If allApproved Then
            Return STATUS_FINISHED
        End If
        
        ' 根据用户角色和案件状态确定默认状态
        Return STATUS_VIEW
    End Function
    
    Private Sub InitControls()
        ' 顶部图片按钮
        picHome = New PictureBox With {
            .ImageLocation = "https://cdn-icons-png.flaticon.com/512/25/25694.png",
            .Size = New Size(32, 32),
            .Location = New Point(10, 10),
            .Cursor = Cursors.Hand
        }
        AddHandler picHome.Click, AddressOf picHome_Click
        Me.Controls.Add(picHome)
        
        ' 案件信息显示
        Dim caseInfoText As String = $"{_caseInfo.CaseType} | {_caseInfo.ProductCode} | {_caseInfo.ProductName} | {_caseInfo.CaseName} | {_caseInfo.CreateTime:yyyy-MM-dd HH:mm:ss}"
        lblCaseInfo = New Label With {
            .Text = caseInfoText,
            .Location = New Point(60, 15),
            .Font = New Font("微软雅黑", 12, FontStyle.Bold),
            .AutoSize = True
        }
        Me.Controls.Add(lblCaseInfo)
        
        ' 中止标记
        lblTerminated = New Label With {
            .Text = "中止",
            .Location = New Point(600, 15),
            .Font = New Font("微软雅黑", 12, FontStyle.Bold),
            .ForeColor = Color.Red,
            .AutoSize = True,
            .Visible = False
        }
        Me.Controls.Add(lblTerminated)
        
        ' 状态按钮
        CreateStatusButtons()
        
        ' TabControl
        tabControl = New TabControl With {
            .Location = New Point(20, 60),
            .Size = New Size(1140, 680)
        }
        CreateTabPages()
        Me.Controls.Add(tabControl)
        
        ' 中止面板
        CreateTerminatedPanel()
    End Sub
    
    Private Sub CreateStatusButtons()
        Dim buttonWidth As Integer = 80
        Dim buttonHeight As Integer = 30
        Dim startX As Integer = Me.Width - 500
        Dim y As Integer = 15
        
        ' 阅览按钮
        btnView = New Button With {
            .Text = "阅览",
            .Location = New Point(startX, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_VIEW
        }
        AddHandler btnView.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnView)
        
        ' 编辑按钮
        btnEdit = New Button With {
            .Text = "编辑",
            .Location = New Point(startX + buttonWidth + 5, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_EDIT
        }
        AddHandler btnEdit.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnEdit)
        
        ' 审查按钮
        btnReview = New Button With {
            .Text = "审查",
            .Location = New Point(startX + (buttonWidth + 5) * 2, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_REVIEW
        }
        AddHandler btnReview.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnReview)
        
        ' 导出按钮
        btnExport = New Button With {
            .Text = "导出",
            .Location = New Point(startX + (buttonWidth + 5) * 3, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_EXPORT
        }
        AddHandler btnExport.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnExport)
        
        ' 中止按钮
        btnTerminate = New Button With {
            .Text = "中止",
            .Location = New Point(startX + (buttonWidth + 5) * 4, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_TERMINATED
        }
        AddHandler btnTerminate.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnTerminate)
        
        ' 终了按钮
        btnFinish = New Button With {
            .Text = "终了",
            .Location = New Point(startX + (buttonWidth + 5) * 5, y),
            .Size = New Size(buttonWidth, buttonHeight),
            .Tag = STATUS_FINISHED
        }
        AddHandler btnFinish.Click, AddressOf StatusButton_Click
        Me.Controls.Add(btnFinish)
    End Sub
    
    Private Sub CreateTabPages()
        ' 根据案件类型创建对应的模板
        _currentTemplate = TabTemplateFactory.CreateTemplate(_caseInfo.CaseType, tabControl)
        
        ' 使用模板创建标签页
        _currentTemplate.CreateTabPages(tabControl)
        
        ' 加载数据到模板
        _currentTemplate.LoadData(_caseDetails)
        
        ' 添加Tab切换事件
        AddHandler tabControl.SelectedIndexChanged, AddressOf TabControl_SelectedIndexChanged
        
        ' 更新Tab状态
        UpdateTabStatus()
    End Sub
    
    Private Sub TabControl_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' Tab切换事件处理
        If _currentStatus = STATUS_REVIEW Then
            ' 在审查模式下，切换标签页时更新控件颜色
            Dim selectedTabPage As TabPage = tabControl.SelectedTab
            If selectedTabPage IsNot Nothing Then
                Dim tabIndex As Integer = tabControl.SelectedIndex
                SetTabPageControlColors(selectedTabPage, tabIndex)
            End If
        End If
    End Sub
    
    Private Sub CreateTerminatedPanel()
        panelTerminated = New Panel With {
            .Location = New Point(20, 60),
            .Size = New Size(1140, 680),
            .Visible = False,
            .BackColor = Color.LightGray
        }
        
        Dim lblTerminatedInfo As New Label With {
            .Text = "案件已中止",
            .Location = New Point(50, 50),
            .Font = New Font("微软雅黑", 16, FontStyle.Bold),
            .AutoSize = True
        }
        panelTerminated.Controls.Add(lblTerminatedInfo)
        
        Dim btnRestore As New Button With {
            .Text = "恢复案件",
            .Location = New Point(50, 100),
            .Size = New Size(100, 30)
        }
        AddHandler btnRestore.Click, AddressOf btnRestore_Click
        panelTerminated.Controls.Add(btnRestore)
        
        Me.Controls.Add(panelTerminated)
    End Sub
    
    Private Sub UpdateTabStatus()
        ' 更新每个Tab的状态显示
        For i As Integer = 0 To 8
            Dim tabPage As TabPage = tabControl.TabPages(i)
            Dim statusLabel As Label = tabPage.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = $"lblStatus_{i}")
            
            If statusLabel IsNot Nothing Then
                Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
                If tabReviewLogs.Count > 0 Then
                    Dim lastReview = tabReviewLogs.OrderByDescending(Function(r) r.ReviewTime).First()
                    statusLabel.Text = lastReview.ReviewStatus
                    Select Case lastReview.ReviewStatus
                        Case "同意"
                            statusLabel.ForeColor = Color.Green
                        Case "不同意"
                            statusLabel.ForeColor = Color.Red
                        Case Else
                            statusLabel.ForeColor = Color.Blue
                    End Select
                Else
                    statusLabel.Text = "未提交"
                    statusLabel.ForeColor = Color.Gray
                End If
            End If
        Next
    End Sub
    
    Private Sub SetStatus(status As String)
        _currentStatus = status
        
        ' 更新按钮状态
        UpdateButtonStates()
        
        ' 更新界面状态
        Select Case status
            Case STATUS_VIEW
                SetViewMode()
            Case STATUS_EDIT
                SetEditMode()
            Case STATUS_REVIEW
                SetReviewMode()
            Case STATUS_EXPORT
                SetExportMode()
            Case STATUS_TERMINATED
                SetTerminatedMode()
            Case STATUS_FINISHED
                SetFinishedMode()
        End Select
    End Sub
    
    Private Sub UpdateButtonStates()
        ' 根据当前状态和用户权限更新按钮状态
        btnView.Enabled = True
        btnEdit.Enabled = CanEdit()
        btnReview.Enabled = CanReview()
        btnExport.Enabled = True
        btnTerminate.Enabled = CanTerminate()
        btnFinish.Enabled = CanFinish()
    End Sub
    
    Private Function CanEdit() As Boolean
        ' 检查是否可以编辑
        Return _caseInfo.IsTerminated = 0 AndAlso _currentStatus <> STATUS_TERMINATED
    End Function
    
    Private Function CanReview() As Boolean
        ' 检查是否可以审查
        Return _caseInfo.IsTerminated = 0 AndAlso _currentStatus <> STATUS_TERMINATED
    End Function
    
    Private Function CanTerminate() As Boolean
        ' 检查是否可以中止
        Return _caseInfo.IsTerminated = 0 AndAlso _currentStatus <> STATUS_TERMINATED
    End Function
    
    Private Function CanFinish() As Boolean
        ' 检查是否可以终了
        Return _caseInfo.IsTerminated = 0 AndAlso _currentStatus <> STATUS_TERMINATED AndAlso AllTabsApproved()
    End Function
    
    Private Function AllTabsApproved() As Boolean
        ' 检查是否所有Tab都已同意
        For i As Integer = 0 To 8
            Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
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
    End Function
    
    Private Sub SetViewMode()
        ' 阅览模式：只读，蓝色样式
        tabControl.Visible = True
        panelTerminated.Visible = False
        lblTerminated.Visible = _caseInfo.IsTerminated = 1
        
        ' 清除编辑模式的遮罩
        ClearReviewingOverlays()
        
        SetControlsReadOnly(True)
        SetControlsStyle(Color.LightBlue)
    End Sub
    
    Private Sub SetEditMode()
        ' 编辑模式：可编辑，已提交审查的Tab显示遮罩
        tabControl.Visible = True
        panelTerminated.Visible = False
        lblTerminated.Visible = False
        
        SetControlsReadOnly(False)
        SetControlsStyle(Color.White)
        
        ' 为已提交审查的Tab添加遮罩
        AddReviewingOverlays()
        
        ' 添加保存和审查按钮
        AddActionButtons()
    End Sub
    
    Private Sub SetReviewMode()
        ' 审查模式：淡紫色，仅显示已提交Tab
        tabControl.Visible = True
        panelTerminated.Visible = False
        lblTerminated.Visible = False
        
        ' 清除编辑模式的遮罩
        ClearReviewingOverlays()
        
        SetControlsReadOnly(True)
        SetControlsStyle(Color.Lavender)
        
        ' 隐藏未提交的Tab
        HideUnsubmittedTabs()
        
        ' 设置项目级别的审查状态显示
        SetReviewModeControlColors()
        
        ' 添加同意/不同意按钮
        AddReviewButtons()
    End Sub
    
    Private Sub SetExportMode()
        ' 导出模式：结构同阅览/编辑/审查
        tabControl.Visible = True
        panelTerminated.Visible = False
        lblTerminated.Visible = _caseInfo.IsTerminated = 1
        
        ' 清除编辑模式的遮罩
        ClearReviewingOverlays()
        
        SetControlsReadOnly(True)
        SetControlsStyle(Color.White)
        
        ' 添加导出按钮
        AddExportButtons()
    End Sub
    
    Private Sub SetTerminatedMode()
        ' 中止模式：无TabControl，仅显示中止信息
        tabControl.Visible = False
        panelTerminated.Visible = True
        lblTerminated.Visible = True
        
        ' 清除编辑模式的遮罩
        ClearReviewingOverlays()
    End Sub
    
    Private Sub SetFinishedMode()
        ' 终了模式：只读，特殊显示规则
        tabControl.Visible = True
        panelTerminated.Visible = False
        lblTerminated.Visible = False
        
        ' 清除编辑模式的遮罩
        ClearReviewingOverlays()
        
        SetControlsReadOnly(True)
        SetControlsStyle(Color.LightGreen)
    End Sub
    
    Private Sub SetControlsReadOnly(readOnly As Boolean)
        ' 使用模板设置控件只读状态
        If _currentTemplate IsNot Nothing Then
            _currentTemplate.SetReadOnly(readOnly)
        End If
    End Sub
    
    Private Sub SetControlsStyle(backColor As Color)
        ' 使用模板设置控件样式
        If _currentTemplate IsNot Nothing Then
            _currentTemplate.SetStyle(backColor)
        End If
    End Sub
    
    Private Sub HideUnsubmittedTabs()
        ' 隐藏未提交的Tab
        For i As Integer = 0 To 8
            Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
            If tabReviewLogs.Count = 0 Then
                tabControl.TabPages(i).Visible = False
            End If
        Next
    End Sub
    
    Private Sub AddActionButtons()
        ' 添加保存和审查按钮
        Dim btnSave As New Button With {
            .Text = "保存",
            .Location = New Point(20, tabControl.Bottom + 10),
            .Size = New Size(80, 30)
        }
        AddHandler btnSave.Click, AddressOf btnSave_Click
        Me.Controls.Add(btnSave)
        
        Dim btnSubmit As New Button With {
            .Text = "提交审查",
            .Location = New Point(110, tabControl.Bottom + 10),
            .Size = New Size(100, 30)
        }
        AddHandler btnSubmit.Click, AddressOf btnSubmit_Click
        Me.Controls.Add(btnSubmit)
    End Sub
    
    Private Sub AddReviewButtons()
        ' 添加同意/不同意按钮
        Dim btnApprove As New Button With {
            .Text = "同意",
            .Location = New Point(20, tabControl.Bottom + 10),
            .Size = New Size(80, 30),
            .BackColor = Color.LightGreen
        }
        AddHandler btnApprove.Click, AddressOf btnApprove_Click
        Me.Controls.Add(btnApprove)
        
        Dim btnReject As New Button With {
            .Text = "不同意",
            .Location = New Point(110, tabControl.Bottom + 10),
            .Size = New Size(80, 30),
            .BackColor = Color.LightCoral
        }
        AddHandler btnReject.Click, AddressOf btnReject_Click
        Me.Controls.Add(btnReject)
    End Sub
    
    Private Sub AddExportButtons()
        ' 添加导出按钮
        Dim btnExportExcel As New Button With {
            .Text = "导出Excel",
            .Location = New Point(20, tabControl.Bottom + 10),
            .Size = New Size(100, 30)
        }
        AddHandler btnExportExcel.Click, AddressOf btnExportExcel_Click
        Me.Controls.Add(btnExportExcel)
        
        Dim btnExportPDF As New Button With {
            .Text = "导出PDF",
            .Location = New Point(130, tabControl.Bottom + 10),
            .Size = New Size(100, 30)
        }
        AddHandler btnExportPDF.Click, AddressOf btnExportPDF_Click
        Me.Controls.Add(btnExportPDF)
    End Sub
    
    ' 事件处理方法
    Private Sub picHome_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Dim mainForm As New MainForm(_currentUser)
        mainForm.ShowDialog()
        Me.Close()
    End Sub
    
    Private Sub StatusButton_Click(sender As Object, e As EventArgs)
        Dim button As Button = DirectCast(sender, Button)
        Dim status As String = button.Tag.ToString()
        SetStatus(status)
    End Sub
    
    Private Sub btnSave_Click(sender As Object, e As EventArgs)
        Try
            ' 使用模板保存案件数据
            If _currentTemplate IsNot Nothing Then
                Dim savedData = _currentTemplate.SaveData()
                Dim success As Boolean = BusinessLogic.CaseManager.SaveCaseDataWithTemplate(_caseId, savedData, _currentUser)
                If success Then
                    MessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ' 重新加载数据
                    LoadCaseData()
                    UpdateTabStatus()
                Else
                    MessageBox.Show("保存失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("模板未初始化！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"保存失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("保存案件数据失败", ex)
        End Try
    End Sub
    
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs)
        Try
            ' 提交审查
            Dim success As Boolean = BusinessLogic.CaseManager.SubmitForReview(_caseId, tabControl, _currentUser)
            If success Then
                MessageBox.Show("提交审查成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 重新加载数据
                LoadCaseData()
                UpdateTabStatus()
                SetStatus(STATUS_REVIEW)
            Else
                MessageBox.Show("提交审查失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"提交审查失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("提交审查失败", ex)
        End Try
    End Sub
    
    Private Sub btnApprove_Click(sender As Object, e As EventArgs)
        Try
            ' 同意审查
            Dim success As Boolean = BusinessLogic.CaseManager.ApproveReview(_caseId, tabControl.SelectedIndex, _currentUser)
            If success Then
                MessageBox.Show("审查同意成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 重新加载数据
                LoadCaseData()
                UpdateTabStatus()
                ' 更新审查模式下的控件颜色
                If _currentStatus = STATUS_REVIEW Then
                    SetReviewModeControlColors()
                End If
            Else
                MessageBox.Show("审查同意失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"审查同意失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("审查同意失败", ex)
        End Try
    End Sub
    
    Private Sub btnReject_Click(sender As Object, e As EventArgs)
        Try
            ' 不同意审查
            Dim success As Boolean = BusinessLogic.CaseManager.RejectReview(_caseId, tabControl.SelectedIndex, _currentUser)
            If success Then
                MessageBox.Show("审查不同意成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 重新加载数据
                LoadCaseData()
                UpdateTabStatus()
                ' 更新审查模式下的控件颜色
                If _currentStatus = STATUS_REVIEW Then
                    SetReviewModeControlColors()
                End If
            Else
                MessageBox.Show("审查不同意失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"审查不同意失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("审查不同意失败", ex)
        End Try
    End Sub
    
    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs)
        Try
            ' 导出Excel
            Dim success As Boolean = BusinessLogic.CaseManager.ExportToExcel(_caseId, _caseInfo, _caseDetails)
            If success Then
                MessageBox.Show("导出Excel成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("导出Excel失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"导出Excel失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("导出Excel失败", ex)
        End Try
    End Sub
    
    Private Sub btnExportPDF_Click(sender As Object, e As EventArgs)
        Try
            ' 导出PDF
            Dim success As Boolean = BusinessLogic.CaseManager.ExportToPDF(_caseId, _caseInfo, _caseDetails)
            If success Then
                MessageBox.Show("导出PDF成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("导出PDF失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"导出PDF失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("导出PDF失败", ex)
        End Try
    End Sub
    
    Private Sub btnRestore_Click(sender As Object, e As EventArgs)
        Try
            ' 恢复案件
            Dim success As Boolean = BusinessLogic.CaseManager.RestoreCase(_caseId, _currentUser)
            If success Then
                MessageBox.Show("恢复案件成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' 重新加载数据
                LoadCaseData()
                SetStatus(STATUS_VIEW)
            Else
                MessageBox.Show("恢复案件失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"恢复案件失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("恢复案件失败", ex)
        End Try
    End Sub
    
    ' 辅助方法
    Private Function GetAllControls(container As Control) As List(Of Control)
        Dim controls As New List(Of Control)()
        
        For Each control As Control In container.Controls
            controls.Add(control)
            ' 递归获取子控件
            controls.AddRange(GetAllControls(control))
        Next
        
        Return controls
    End Function
    
    Private Sub SetControlValue(control As Control, value As String)
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
        End Select
    End Sub
    
    Private Sub SetReviewModeControlColors()
        ' 设置项目级别的审查状态显示
        For Each tabPage As TabPage In tabControl.TabPages
            If tabPage.Visible Then
                Dim tabIndex As Integer = tabControl.TabPages.IndexOf(tabPage)
                SetTabPageControlColors(tabPage, tabIndex)
            End If
        Next
    End Sub
    
    Private Sub SetTabPageControlColors(tabPage As TabPage, tabIndex As Integer)
        ' 获取该标签页的审查历史
        Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = tabIndex).OrderBy(Function(r) r.ReviewTime).ToList()
        
        ' 遍历标签页中的所有控件
        For Each control As Control In GetAllControls(tabPage)
            If Not String.IsNullOrEmpty(control.Tag?.ToString()) Then
                Dim fieldName As String = control.Tag.ToString()
                Dim controlColor As Color = GetControlReviewColor(tabIndex, fieldName, tabReviewLogs)
                SetControlBackColor(control, controlColor)
            End If
        Next
    End Sub
    
    Private Function GetControlReviewColor(tabIndex As Integer, fieldName As String, tabReviewLogs As List(Of ReviewLog)) As Color
        ' 获取该字段的审查状态
        Dim reviewStatus As String = GetFieldReviewStatus(tabIndex, fieldName, tabReviewLogs)
        
        Select Case reviewStatus
            Case "未审查"
                Return Color.White ' 白色 - 历史数据中没审查过的项目
            Case "已通过"
                Return Color.LightGray ' 灰色 - 审查通过的项目
            Case "重新编辑"
                Return Color.LightCoral ' 红色 - 审查通过后又被编辑提交
            Case Else
                Return Color.White
        End Select
    End Function
    
    Private Function GetFieldReviewStatus(tabIndex As Integer, fieldName As String, tabReviewLogs As List(Of ReviewLog)) As String
        ' 检查该字段是否有审查记录
        If tabReviewLogs.Count = 0 Then
            Return "未审查"
        End If
        
        ' 获取最新的审查记录
        Dim lastReview = tabReviewLogs.Last()
        
        ' 如果最新状态是"同意"，检查是否有后续的编辑
        If lastReview.ReviewStatus = "同意" Then
            ' 检查该字段在同意后是否被重新编辑过
            If IsFieldModifiedAfterApproval(tabIndex, fieldName, lastReview.ReviewTime) Then
                Return "重新编辑"
            Else
                Return "已通过"
            End If
        ElseIf lastReview.ReviewStatus = "不同意" Then
            Return "未审查" ' 不同意后需要重新审查
        Else
            Return "未审查"
        End If
    End Function
    
    Private Function IsFieldModifiedAfterApproval(tabIndex As Integer, fieldName As String, approvalTime As DateTime) As Boolean
        ' 检查字段在同意后是否被修改过
        ' 这里需要检查案件详细信息的修改时间
        Dim fieldDetails = _caseDetails.Where(Function(d) d.TabIndex = tabIndex AndAlso d.FieldNo = fieldName).ToList()
        
        If fieldDetails.Count > 0 Then
            ' 获取该字段的最新修改时间
            Dim latestModification = fieldDetails.OrderByDescending(Function(d) d.CreateTime).First()
            
            ' 如果最新修改时间晚于同意时间，说明被重新编辑过
            Return latestModification.CreateTime > approvalTime
        End If
        
        Return False
    End Function
    
    Private Sub SetControlBackColor(control As Control, backColor As Color)
        ' 设置控件的背景色
        If TypeOf control Is TextBox Then
            DirectCast(control, TextBox).BackColor = backColor
        ElseIf TypeOf control Is ComboBox Then
            DirectCast(control, ComboBox).BackColor = backColor
        ElseIf TypeOf control Is DateTimePicker Then
            DirectCast(control, DateTimePicker).BackColor = backColor
        ElseIf TypeOf control Is RichTextBox Then
            DirectCast(control, RichTextBox).BackColor = backColor
        ElseIf TypeOf control Is CheckBox Then
            DirectCast(control, CheckBox).BackColor = backColor
        End If
    End Sub
    
    Private Sub AddReviewingOverlays()
        ' 清除之前的遮罩
        ClearReviewingOverlays()
        
        ' 为已提交审查的Tab添加遮罩层
        For i As Integer = 0 To 8
            Dim tabPage As TabPage = tabControl.TabPages(i)
            Dim tabReviewLogs = _reviewLogs.Where(Function(r) r.TabIndex = i).ToList()
            
            If tabReviewLogs.Count > 0 Then
                ' 检查是否有待审查的记录
                Dim hasPendingReview = tabReviewLogs.Any(Function(r) r.ReviewStatus = "待审查")
                
                If hasPendingReview Then
                    ' 设置该标签页的所有控件为只读/不可用/TabStop=False
                    For Each control As Control In GetAllControls(tabPage)
                        If TypeOf control Is TextBox Then
                            DirectCast(control, TextBox).ReadOnly = True
                        ElseIf TypeOf control Is ComboBox Then
                            DirectCast(control, ComboBox).Enabled = False
                        ElseIf TypeOf control Is DateTimePicker Then
                            DirectCast(control, DateTimePicker).Enabled = False
                        ElseIf TypeOf control Is RichTextBox Then
                            DirectCast(control, RichTextBox).ReadOnly = True
                        ElseIf TypeOf control Is CheckBox OrElse TypeOf control Is RadioButton Then
                            control.Enabled = False
                        End If
                        control.TabStop = False
                    Next
                    
                    ' 创建遮罩面板
                    Dim overlayPanel As New Panel With {
                        .Dock = DockStyle.Fill,
                        .BackColor = Color.FromArgb(128, 200, 200, 200), ' 半透明灰色
                        .Visible = True,
                        .Name = $"overlay_{i}",
                        .TabStop = False
                    }
                    ' 拦截所有鼠标和键盘事件
                    AddHandler overlayPanel.MouseDown, Sub(s, e) CType(e, MouseEventArgs).Handled = True
                    AddHandler overlayPanel.MouseMove, Sub(s, e) CType(e, MouseEventArgs).Handled = True
                    AddHandler overlayPanel.Click, Sub(s, e) CType(e, EventArgs).Handled = True
                    AddHandler overlayPanel.DoubleClick, Sub(s, e) CType(e, EventArgs).Handled = True
                    AddHandler overlayPanel.KeyDown, Sub(s, e) CType(e, KeyEventArgs).Handled = True
                    AddHandler overlayPanel.KeyPress, Sub(s, e) CType(e, KeyPressEventArgs).Handled = True
                    
                    ' 添加"正在审查中"标签
                    Dim lblReviewing As New Label With {
                        .Text = "正在审查中",
                        .Font = New Font("微软雅黑", 14, FontStyle.Bold),
                        .ForeColor = Color.DarkBlue,
                        .AutoSize = True,
                        .Location = New Point(50, 50),
                        .TabStop = False
                    }
                    overlayPanel.Controls.Add(lblReviewing)
                    
                    tabPage.Controls.Add(overlayPanel)
                    overlayPanel.BringToFront()
                End If
            End If
        Next
    End Sub
    
    Private Sub ClearReviewingOverlays()
        ' 清除所有遮罩层
        For i As Integer = 0 To 8
            Dim tabPage As TabPage = tabControl.TabPages(i)
            Dim overlayPanel = tabPage.Controls.OfType(Of Panel)().FirstOrDefault(Function(p) p.Name = $"overlay_{i}")
            If overlayPanel IsNot Nothing Then
                tabPage.Controls.Remove(overlayPanel)
                overlayPanel.Dispose()
            End If
        Next
    End Sub
End Class 
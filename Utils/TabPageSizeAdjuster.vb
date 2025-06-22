''' <summary>
''' TabPage大小自动调整工具类
''' 根据控件数量和布局自动调整TabPage大小，确保最佳的用户体验
''' </summary>
Public Class TabPageSizeAdjuster
    
    #Region "常量定义"
    
    ' 屏幕和窗体尺寸常量
    Private Const SCREEN_WIDTH As Integer = 1920
    Private Const SCREEN_HEIGHT As Integer = 1280
    Private Const DEFAULT_TAB_WIDTH As Integer = 1000
    Private Const DEFAULT_TAB_HEIGHT As Integer = 700
    Private Const MIN_TAB_WIDTH As Integer = 1000
    Private Const MIN_TAB_HEIGHT As Integer = 700
    Private Const MAX_TAB_WIDTH As Integer = 1000  ' 宽度固定
    Private Const MAX_TAB_HEIGHT As Integer = 1200 ' 留出一些边距
    
    ' 控件间距和边距常量
    Private Const CONTROL_MARGIN As Integer = 20
    Private Const CONTROL_SPACING As Integer = 30
    Private Const LABEL_WIDTH As Integer = 130
    Private Const CONTROL_WIDTH As Integer = 200
    Private Const RICH_TEXT_HEIGHT As Integer = 100
    Private Const STATUS_LABEL_HEIGHT As Integer = 25
    
    #End Region
    
    #Region "公共方法"
    
    ''' <summary>
    ''' 调整TabPage大小
    ''' </summary>
    ''' <param name="tabPage">要调整的TabPage</param>
    ''' <param name="parentForm">父窗体</param>
    Public Shared Sub AdjustTabPageSize(tabPage As TabPage, parentForm As Form)
        Try
            ' 计算所需的高度
            Dim requiredHeight As Integer = CalculateRequiredHeight(tabPage)
            
            ' 调整TabPage大小
            Dim newSize As Size = CalculateOptimalSize(requiredHeight)
            
            ' 应用新大小
            ApplyTabPageSize(tabPage, newSize, parentForm)
            
        Catch ex As Exception
            Utils.LogUtil.LogError("调整TabPage大小失败", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' 批量调整所有TabPage大小
    ''' </summary>
    ''' <param name="tabControl">TabControl</param>
    ''' <param name="parentForm">父窗体</param>
    Public Shared Sub AdjustAllTabPages(tabControl As TabControl, parentForm As Form)
        Try
            For Each tabPage As TabPage In tabControl.TabPages
                AdjustTabPageSize(tabPage, parentForm)
            Next
            
            Utils.LogUtil.LogInfo($"已自动调整{tabControl.TabPages.Count}个TabPage大小")
            
        Catch ex As Exception
            Utils.LogUtil.LogError("批量调整TabPage大小失败", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' 获取TabPage的推荐大小
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    ''' <returns>推荐大小</returns>
    Public Shared Function GetRecommendedSize(tabPage As TabPage) As Size
        Try
            Dim requiredHeight As Integer = CalculateRequiredHeight(tabPage)
            Return CalculateOptimalSize(requiredHeight)
        Catch ex As Exception
            Utils.LogUtil.LogError("获取TabPage推荐大小失败", ex)
            Return New Size(DEFAULT_TAB_WIDTH, DEFAULT_TAB_HEIGHT)
        End Try
    End Function
    
    ''' <summary>
    ''' 检查是否需要滚动条
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    ''' <returns>是否需要滚动条</returns>
    Public Shared Function NeedsScrollBar(tabPage As TabPage) As Boolean
        Try
            Dim requiredHeight As Integer = CalculateRequiredHeight(tabPage)
            Return requiredHeight > MAX_TAB_HEIGHT
        Catch ex As Exception
            Utils.LogUtil.LogError("检查TabPage滚动条需求失败", ex)
            Return False
        End Try
    End Function
    
    #End Region
    
    #Region "私有方法"
    
    ''' <summary>
    ''' 计算TabPage所需的高度
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    ''' <returns>所需高度</returns>
    Private Shared Function CalculateRequiredHeight(tabPage As TabPage) As Integer
        Dim maxY As Integer = STATUS_LABEL_HEIGHT + CONTROL_MARGIN ' 从状态标签开始
        
        ' 遍历所有控件，找到最底部的控件
        For Each control As Control In GetAllControls(tabPage)
            If control.Visible Then
                Dim controlBottom As Integer = control.Location.Y + control.Height
                If controlBottom > maxY Then
                    maxY = controlBottom
                End If
            End If
        Next
        
        ' 添加底部边距
        maxY += CONTROL_MARGIN
        
        Return maxY
    End Function
    
    ''' <summary>
    ''' 计算最优的TabPage大小
    ''' </summary>
    ''' <param name="requiredHeight">所需高度</param>
    ''' <returns>最优大小</returns>
    Private Shared Function CalculateOptimalSize(requiredHeight As Integer) As Size
        Dim width As Integer = MAX_TAB_WIDTH
        Dim height As Integer = Math.Max(MIN_TAB_HEIGHT, requiredHeight)
        
        ' 如果高度超过最大值，限制在最大值内
        If height > MAX_TAB_HEIGHT Then
            height = MAX_TAB_HEIGHT
        End If
        
        Return New Size(width, height)
    End Function
    
    ''' <summary>
    ''' 应用TabPage大小
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    ''' <param name="newSize">新大小</param>
    ''' <param name="parentForm">父窗体</param>
    Private Shared Sub ApplyTabPageSize(tabPage As TabPage, newSize As Size, parentForm As Form)
        ' 获取TabControl
        Dim tabControl As TabControl = DirectCast(tabPage.Parent, TabControl)
        If tabControl Is Nothing Then Return
        
        ' 计算窗体需要的新大小
        Dim formPadding As Integer = 40 ' 窗体边距
        Dim tabControlPadding As Integer = 40 ' TabControl边距
        Dim newFormWidth As Integer = newSize.Width + formPadding + tabControlPadding
        Dim newFormHeight As Integer = newSize.Height + formPadding + tabControlPadding + 60 ' 顶部按钮区域
        
        ' 检查是否超出屏幕范围
        If newFormWidth > SCREEN_WIDTH OrElse newFormHeight > SCREEN_HEIGHT Then
            ' 超出屏幕范围，启用滚动条
            EnableTabPageScroll(tabPage, newSize)
        Else
            ' 在屏幕范围内，调整窗体大小
            AdjustFormSize(parentForm, newFormWidth, newFormHeight)
            ' 调整TabControl大小
            tabControl.Size = newSize
        End If
    End Sub
    
    ''' <summary>
    ''' 启用TabPage滚动
    ''' </summary>
    ''' <param name="tabPage">TabPage</param>
    ''' <param name="contentSize">内容大小</param>
    Private Shared Sub EnableTabPageScroll(tabPage As TabPage, contentSize As Size)
        ' 创建Panel作为滚动容器
        Dim scrollPanel As New Panel With {
            .AutoScroll = True,
            .Dock = DockStyle.Fill,
            .AutoScrollMinSize = contentSize
        }
        
        ' 将所有控件移动到Panel中
        Dim controlsToMove As New List(Of Control)
        For Each control As Control In tabPage.Controls
            controlsToMove.Add(control)
        Next
        
        ' 清空TabPage并添加Panel
        tabPage.Controls.Clear()
        tabPage.Controls.Add(scrollPanel)
        
        ' 将控件添加到Panel中
        For Each control As Control In controlsToMove
            scrollPanel.Controls.Add(control)
        Next
        
        ' 设置Panel的滚动区域
        scrollPanel.AutoScrollMinSize = contentSize
    End Sub
    
    ''' <summary>
    ''' 调整窗体大小
    ''' </summary>
    ''' <param name="form">窗体</param>
    ''' <param name="width">宽度</param>
    ''' <param name="height">高度</param>
    Private Shared Sub AdjustFormSize(form As Form, width As Integer, height As Integer)
        ' 确保窗体大小在屏幕范围内
        width = Math.Min(width, SCREEN_WIDTH)
        height = Math.Min(height, SCREEN_HEIGHT)
        
        ' 调整窗体大小
        form.Size = New Size(width, height)
        
        ' 重新居中窗体
        form.StartPosition = FormStartPosition.CenterScreen
    End Sub
    
    ''' <summary>
    ''' 获取所有控件（包括嵌套控件）
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
    
    #End Region
    
End Class 
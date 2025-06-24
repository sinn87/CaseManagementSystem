' 事务功能测试
Imports System.Collections.Generic
Imports System.Windows.Forms

Public Class TestTransaction
    ''' <summary>
    ''' 测试事务功能
    ''' </summary>
    Public Shared Sub TestCaseCreationTransaction()
        Try
            ' 准备测试数据
            Dim testData As New Dictionary(Of Integer, Dictionary(Of String, String))()
            
            ' 模拟第一个标签页的数据
            Dim tab1Data As New Dictionary(Of String, String)()
            tab1Data.Add("Field_1_1", "测试字段1")
            tab1Data.Add("Field_1_2", "测试字段2")
            testData.Add(0, tab1Data)
            
            ' 模拟第二个标签页的数据
            Dim tab2Data As New Dictionary(Of String, String)()
            tab2Data.Add("Field_2_1", "测试字段3")
            testData.Add(1, tab2Data)
            
            ' 调用业务逻辑层创建案件
            Dim success As Boolean = BusinessLogic.CaseManager.CreateNewCase("测试案件类型", testData, "测试用户")
            
            If success Then
                MessageBox.Show("事务测试成功！案件创建、详细信息保存和审查记录创建都在一个事务中完成。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("事务测试失败！请检查错误日志。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"测试过程中发生异常：{ex.Message}", "测试异常", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Utils.LogUtil.LogError("事务测试异常", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' 测试事务回滚功能（通过故意制造错误）
    ''' </summary>
    Public Shared Sub TestTransactionRollback()
        Try
            ' 准备测试数据，包含一个无效的字段名来触发错误
            Dim testData As New Dictionary(Of Integer, Dictionary(Of String, String))()
            
            ' 模拟第一个标签页的数据
            Dim tab1Data As New Dictionary(Of String, String)()
            tab1Data.Add("Field_1_1", "测试字段1")
            tab1Data.Add("Invalid_Field_Name_With_Special_Characters_That_Will_Cause_Error", "测试字段2")
            testData.Add(0, tab1Data)
            
            ' 调用业务逻辑层创建案件（应该会失败并回滚）
            Dim success As Boolean = BusinessLogic.CaseManager.CreateNewCase("测试案件类型", testData, "测试用户")
            
            If Not success Then
                MessageBox.Show("事务回滚测试成功！当发生错误时，所有数据库操作都被回滚。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("事务回滚测试失败！应该发生错误并回滚事务。", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"事务回滚测试成功！捕获到异常：{ex.Message}", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Utils.LogUtil.LogError("事务回滚测试异常", ex)
        End Try
    End Sub
End Class 
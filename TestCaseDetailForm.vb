' 案件详细页面测试程序
Imports System.Windows.Forms

Public Class TestCaseDetailForm
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        
        Try
            ' 测试案件详细页面
            Dim testCaseId As Integer = 1 ' 假设存在案件ID为1的案件
            Dim testUser As String = "testuser"
            
            ' 创建并显示案件详细页面
            Dim detailForm As New CaseDetailForm(testCaseId, testUser)
            Application.Run(detailForm)
            
        Catch ex As Exception
            MessageBox.Show($"测试程序运行失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class 
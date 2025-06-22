' 程序入口点
Imports System.Windows.Forms

Module Program
    ''' <summary>
    ''' 应用程序的主入口点
    ''' </summary>
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        
        Try
            ' 启动登录窗体
            Application.Run(New LoginForm())
        Catch ex As Exception
            MessageBox.Show($"程序启动失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module 
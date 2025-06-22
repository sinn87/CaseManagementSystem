' 日志工具类
Imports System.IO

Public Class LogUtil
    Private Shared ReadOnly LogPath As String = Path.Combine(Application.StartupPath, "Logs")
    
    ''' <summary>
    ''' 记录错误日志
    ''' </summary>
    ''' <param name="msg">错误消息</param>
    ''' <param name="ex">异常对象</param>
    Public Shared Sub LogError(msg As String, ex As Exception)
        Try
            ' 确保日志目录存在
            If Not Directory.Exists(LogPath) Then
                Directory.CreateDirectory(LogPath)
            End If
            
            Dim logFile As String = Path.Combine(LogPath, $"Error_{DateTime.Now:yyyy-MM-dd}.log")
            Dim logContent As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {msg}" & vbCrLf & $"Exception: {ex.Message}" & vbCrLf & $"StackTrace: {ex.StackTrace}" & vbCrLf & "----------------------------------------" & vbCrLf
            
            File.AppendAllText(logFile, logContent)
        Catch
            ' 如果日志记录失败，忽略异常避免影响主程序
        End Try
    End Sub
    
    ''' <summary>
    ''' 记录信息日志
    ''' </summary>
    ''' <param name="msg">信息消息</param>
    Public Shared Sub LogInfo(msg As String)
        Try
            ' 确保日志目录存在
            If Not Directory.Exists(LogPath) Then
                Directory.CreateDirectory(LogPath)
            End If
            
            Dim logFile As String = Path.Combine(LogPath, $"Info_{DateTime.Now:yyyy-MM-dd}.log")
            Dim logContent As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {msg}" & vbCrLf
            
            File.AppendAllText(logFile, logContent)
        Catch
            ' 如果日志记录失败，忽略异常避免影响主程序
        End Try
    End Sub
End Class 
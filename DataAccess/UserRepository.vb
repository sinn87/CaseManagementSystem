' 用户数据访问仓储
Imports System.Data.OleDb

Public Class UserRepository
    Private connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=yourdb.accdb;Persist Security Info=False;"

    ''' <summary>
    ''' 检查登录信息，返回：0=不存在，1=正常，2=密码错误，3=禁用
    ''' </summary>
    Public Function CheckLogin(staffNo As String, password As String) As Integer
        Try
            Using conn As New OleDbConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT * FROM Users WHERE StaffNo=?"
                Using cmd As New OleDbCommand(sql, conn)
                    cmd.Parameters.AddWithValue("?", staffNo)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            If reader("Status") = 0 Then
                                Return 3 '禁用
                            ElseIf reader("Password").ToString() = password Then
                                Return 1 '正常
                            Else
                                Return 2 '密码错误
                            End If
                        Else
                            Return 0 '不存在
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return -1 '数据库异常
        End Try
    End Function
End Class 
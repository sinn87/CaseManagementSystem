' 数据库操作辅助类
Imports System.Data.OleDb
Imports System.Configuration

Public Class DbHelper
    Private Shared ReadOnly ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=CaseManagement.accdb"
    
    ''' <summary>
    ''' 获取数据库连接
    ''' </summary>
    ''' <returns>数据库连接对象</returns>
    Public Shared Function GetConnection() As OleDbConnection
        Return New OleDbConnection(ConnectionString)
    End Function
    
    ''' <summary>
    ''' 开始数据库事务
    ''' </summary>
    ''' <param name="connection">数据库连接</param>
    ''' <returns>事务对象</returns>
    Public Shared Function BeginTransaction(connection As OleDbConnection) As OleDbTransaction
        If connection.State <> ConnectionState.Open Then
            connection.Open()
        End If
        Return connection.BeginTransaction()
    End Function
    
    ''' <summary>
    ''' 在事务中执行非查询SQL语句
    ''' </summary>
    ''' <param name="transaction">事务对象</param>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>影响的行数</returns>
    Public Shared Function ExecuteNonQueryWithTransaction(transaction As OleDbTransaction, sql As String, Optional parameters As OleDbParameter() = Nothing) As Integer
        Using cmd As New OleDbCommand(sql, transaction.Connection, transaction)
            If parameters IsNot Nothing Then
                cmd.Parameters.AddRange(parameters)
            End If
            Return cmd.ExecuteNonQuery()
        End Using
    End Function
    
    ''' <summary>
    ''' 在事务中执行查询并返回第一行第一列的值
    ''' </summary>
    ''' <param name="transaction">事务对象</param>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>查询结果</returns>
    Public Shared Function ExecuteScalarWithTransaction(transaction As OleDbTransaction, sql As String, Optional parameters As OleDbParameter() = Nothing) As Object
        Using cmd As New OleDbCommand(sql, transaction.Connection, transaction)
            If parameters IsNot Nothing Then
                cmd.Parameters.AddRange(parameters)
            End If
            Return cmd.ExecuteScalar()
        End Using
    End Function
    
    ''' <summary>
    ''' 执行非查询SQL语句
    ''' </summary>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>影响的行数</returns>
    Public Shared Function ExecuteNonQuery(sql As String, Optional parameters As OleDbParameter() = Nothing) As Integer
        Using conn As OleDbConnection = GetConnection()
            conn.Open()
            Using cmd As New OleDbCommand(sql, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function
    
    ''' <summary>
    ''' 执行查询SQL语句
    ''' </summary>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>数据读取器</returns>
    Public Shared Function ExecuteReader(sql As String, Optional parameters As OleDbParameter() = Nothing) As OleDbDataReader
        Dim conn As OleDbConnection = GetConnection()
        conn.Open()
        Dim cmd As New OleDbCommand(sql, conn)
        If parameters IsNot Nothing Then
            cmd.Parameters.AddRange(parameters)
        End If
        Return cmd.ExecuteReader(CommandBehavior.CloseConnection)
    End Function
    
    ''' <summary>
    ''' 执行查询并返回第一行第一列的值
    ''' </summary>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>查询结果</returns>
    Public Shared Function ExecuteScalar(sql As String, Optional parameters As OleDbParameter() = Nothing) As Object
        Using conn As OleDbConnection = GetConnection()
            conn.Open()
            Using cmd As New OleDbCommand(sql, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If
                Return cmd.ExecuteScalar()
            End Using
        End Using
    End Function
    
    ''' <summary>
    ''' 执行查询并返回DataTable
    ''' </summary>
    ''' <param name="sql">SQL语句</param>
    ''' <param name="parameters">参数数组</param>
    ''' <returns>数据表</returns>
    Public Shared Function ExecuteDataTable(sql As String, Optional parameters As OleDbParameter() = Nothing) As DataTable
        Using conn As OleDbConnection = GetConnection()
            conn.Open()
            Using cmd As New OleDbCommand(sql, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If
                Using adapter As New OleDbDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function
End Class 
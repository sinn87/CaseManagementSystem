' 案件数据访问层
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.Data

Public Class CaseRepository
    ''' <summary>
    ''' 创建新案件
    ''' </summary>
    ''' <param name="caseInfo">案件信息</param>
    ''' <returns>新创建的案件ID</returns>
    Public Shared Function CreateCase(caseInfo As CaseInfo) As Integer
        Dim sql As String = "INSERT INTO Cases (CaseType, CaseName, ProductCode, ProductName, Status, PublishDate, ListingDate, CompanyRole, LastUpdate, IsTerminated, CreateTime, CreateUser) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseType", caseInfo.CaseType),
            New OleDbParameter("@CaseName", If(caseInfo.CaseName, DBNull.Value)),
            New OleDbParameter("@ProductCode", If(caseInfo.ProductCode, DBNull.Value)),
            New OleDbParameter("@ProductName", If(caseInfo.ProductName, DBNull.Value)),
            New OleDbParameter("@Status", caseInfo.Status),
            New OleDbParameter("@PublishDate", If(caseInfo.PublishDate, DBNull.Value)),
            New OleDbParameter("@ListingDate", If(caseInfo.ListingDate, DBNull.Value)),
            New OleDbParameter("@CompanyRole", If(caseInfo.CompanyRole, DBNull.Value)),
            New OleDbParameter("@LastUpdate", caseInfo.LastUpdate),
            New OleDbParameter("@IsTerminated", caseInfo.IsTerminated),
            New OleDbParameter("@CreateTime", caseInfo.CreateTime),
            New OleDbParameter("@CreateUser", caseInfo.CreateUser)
        }
        
        DbHelper.ExecuteNonQuery(sql, parameters)
        
        ' 获取新创建的案件ID
        Return Convert.ToInt32(DbHelper.ExecuteScalar("SELECT @@IDENTITY"))
    End Function
    
    ''' <summary>
    ''' 根据ID获取案件信息
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <returns>案件信息</returns>
    Public Shared Function GetCaseById(caseId As Integer) As CaseInfo
        Dim sql As String = "SELECT CaseID, CaseType, CaseName, ProductCode, ProductName, Status, PublishDate, ListingDate, CompanyRole, LastUpdate, IsTerminated, CreateTime, CreateUser FROM Cases WHERE CaseID = ?"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseID", caseId)
        }
        
        Dim dt As DataTable = DbHelper.ExecuteDataTable(sql, parameters)
        
        If dt.Rows.Count > 0 Then
            Dim row As DataRow = dt.Rows(0)
            Return New CaseInfo With {
                .CaseID = Convert.ToInt32(row("CaseID")),
                .CaseType = row("CaseType").ToString(),
                .CaseName = If(row("CaseName") Is DBNull.Value, Nothing, row("CaseName").ToString()),
                .ProductCode = If(row("ProductCode") Is DBNull.Value, Nothing, row("ProductCode").ToString()),
                .ProductName = If(row("ProductName") Is DBNull.Value, Nothing, row("ProductName").ToString()),
                .Status = Convert.ToInt32(row("Status")),
                .PublishDate = If(row("PublishDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("PublishDate"))),
                .ListingDate = If(row("ListingDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("ListingDate"))),
                .CompanyRole = If(row("CompanyRole") Is DBNull.Value, Nothing, row("CompanyRole").ToString()),
                .LastUpdate = Convert.ToDateTime(row("LastUpdate")),
                .IsTerminated = Convert.ToInt32(row("IsTerminated")),
                .CreateTime = Convert.ToDateTime(row("CreateTime")),
                .CreateUser = row("CreateUser").ToString()
            }
        End If
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' 根据案件ID获取案件详细信息
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <returns>案件详细信息列表</returns>
    Public Shared Function GetCaseDetailsByCaseId(caseId As Integer) As List(Of CaseDetail)
        Dim sql As String = "SELECT DetailID, CaseID, TabIndex, FieldNo, FieldValue, FieldStatus, CreateTime FROM CaseDetails WHERE CaseID = ? ORDER BY TabIndex, FieldNo"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseID", caseId)
        }
        
        Dim dt As DataTable = DbHelper.ExecuteDataTable(sql, parameters)
        Dim caseDetails As New List(Of CaseDetail)()
        
        For Each row As DataRow In dt.Rows
            Dim detail As New CaseDetail With {
                .DetailID = Convert.ToInt32(row("DetailID")),
                .CaseID = Convert.ToInt32(row("CaseID")),
                .TabIndex = Convert.ToInt32(row("TabIndex")),
                .FieldNo = row("FieldNo").ToString(),
                .FieldValue = If(row("FieldValue") Is DBNull.Value, Nothing, row("FieldValue").ToString()),
                .FieldStatus = row("FieldStatus").ToString(),
                .CreateTime = Convert.ToDateTime(row("CreateTime"))
            }
            caseDetails.Add(detail)
        Next
        
        Return caseDetails
    End Function
    
    ''' <summary>
    ''' 根据案件ID获取审查记录
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <returns>审查记录列表</returns>
    Public Shared Function GetReviewLogsByCaseId(caseId As Integer) As List(Of ReviewLog)
        Dim sql As String = "SELECT ReviewID, CaseID, TabIndex, ReviewerID, ReviewStatus, ReviewComment, ReviewTime FROM ReviewLogs WHERE CaseID = ? ORDER BY ReviewTime"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseID", caseId)
        }
        
        Dim dt As DataTable = DbHelper.ExecuteDataTable(sql, parameters)
        Dim reviewLogs As New List(Of ReviewLog)()
        
        For Each row As DataRow In dt.Rows
            Dim reviewLog As New ReviewLog With {
                .ReviewID = Convert.ToInt32(row("ReviewID")),
                .CaseID = Convert.ToInt32(row("CaseID")),
                .TabIndex = Convert.ToInt32(row("TabIndex")),
                .ReviewerID = row("ReviewerID").ToString(),
                .ReviewStatus = row("ReviewStatus").ToString(),
                .ReviewComment = If(row("ReviewComment") Is DBNull.Value, Nothing, row("ReviewComment").ToString()),
                .ReviewTime = Convert.ToDateTime(row("ReviewTime"))
            }
            reviewLogs.Add(reviewLog)
        Next
        
        Return reviewLogs
    End Function
    
    ''' <summary>
    ''' 批量保存案件详细信息
    ''' </summary>
    ''' <param name="caseDetails">案件详细信息列表</param>
    Public Shared Sub SaveCaseDetails(caseDetails As List(Of CaseDetail))
        If caseDetails Is Nothing OrElse caseDetails.Count = 0 Then
            Return
        End If
        
        Dim sql As String = "INSERT INTO CaseDetails (CaseID, TabIndex, FieldNo, FieldValue, FieldStatus, CreateTime) VALUES (?, ?, ?, ?, ?, ?)"
        
        For Each detail In caseDetails
            Dim parameters As OleDbParameter() = {
                New OleDbParameter("@CaseID", detail.CaseID),
                New OleDbParameter("@TabIndex", detail.TabIndex),
                New OleDbParameter("@FieldNo", detail.FieldNo),
                New OleDbParameter("@FieldValue", If(detail.FieldValue, DBNull.Value)),
                New OleDbParameter("@FieldStatus", detail.FieldStatus),
                New OleDbParameter("@CreateTime", detail.CreateTime)
            }
            
            DbHelper.ExecuteNonQuery(sql, parameters)
        Next
    End Sub
    
    ''' <summary>
    ''' 创建审查记录
    ''' </summary>
    ''' <param name="reviewLog">审查记录</param>
    Public Shared Sub CreateReviewLog(reviewLog As ReviewLog)
        Dim sql As String = "INSERT INTO ReviewLogs (CaseID, TabIndex, ReviewerID, ReviewStatus, ReviewComment, ReviewTime) VALUES (?, ?, ?, ?, ?, ?)"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseID", reviewLog.CaseID),
            New OleDbParameter("@TabIndex", reviewLog.TabIndex),
            New OleDbParameter("@ReviewerID", reviewLog.ReviewerID),
            New OleDbParameter("@ReviewStatus", reviewLog.ReviewStatus),
            New OleDbParameter("@ReviewComment", If(reviewLog.ReviewComment, DBNull.Value)),
            New OleDbParameter("@ReviewTime", reviewLog.ReviewTime)
        }
        
        DbHelper.ExecuteNonQuery(sql, parameters)
    End Sub
    
    ''' <summary>
    ''' 更新案件状态
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="status">新状态</param>
    Public Shared Sub UpdateCaseStatus(caseId As Integer, status As Integer)
        Dim sql As String = "UPDATE Cases SET Status = ?, LastUpdate = ? WHERE CaseID = ?"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@Status", status),
            New OleDbParameter("@LastUpdate", DateTime.Now),
            New OleDbParameter("@CaseID", caseId)
        }
        
        DbHelper.ExecuteNonQuery(sql, parameters)
    End Sub
    
    ''' <summary>
    ''' 更新案件最后修改时间
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="lastUpdate">最后修改时间</param>
    Public Shared Sub UpdateCaseLastUpdate(caseId As Integer, lastUpdate As DateTime)
        Dim sql As String = "UPDATE Cases SET LastUpdate = ? WHERE CaseID = ?"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@LastUpdate", lastUpdate),
            New OleDbParameter("@CaseID", caseId)
        }
        
        DbHelper.ExecuteNonQuery(sql, parameters)
    End Sub
    
    ''' <summary>
    ''' 更新案件中止状态
    ''' </summary>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="isTerminated">是否中止</param>
    ''' <param name="currentUser">当前用户</param>
    Public Shared Sub UpdateCaseTerminated(caseId As Integer, isTerminated As Integer, currentUser As String)
        Dim sql As String = "UPDATE Cases SET IsTerminated = ?, LastUpdate = ? WHERE CaseID = ?"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@IsTerminated", isTerminated),
            New OleDbParameter("@LastUpdate", DateTime.Now),
            New OleDbParameter("@CaseID", caseId)
        }
        
        DbHelper.ExecuteNonQuery(sql, parameters)
    End Sub
    
    ''' <summary>
    ''' 获取所有案件列表
    ''' </summary>
    ''' <returns>案件列表</returns>
    Public Shared Function GetAllCases() As List(Of CaseInfo)
        Dim sql As String = "SELECT CaseID, CaseType, CaseName, ProductCode, ProductName, Status, PublishDate, ListingDate, CompanyRole, LastUpdate, IsTerminated, CreateTime, CreateUser FROM Cases ORDER BY CreateTime DESC"
        
        Dim dt As DataTable = DbHelper.ExecuteDataTable(sql)
        Dim cases As New List(Of CaseInfo)()
        
        For Each row As DataRow In dt.Rows
            Dim caseInfo As New CaseInfo With {
                .CaseID = Convert.ToInt32(row("CaseID")),
                .CaseType = row("CaseType").ToString(),
                .CaseName = If(row("CaseName") Is DBNull.Value, Nothing, row("CaseName").ToString()),
                .ProductCode = If(row("ProductCode") Is DBNull.Value, Nothing, row("ProductCode").ToString()),
                .ProductName = If(row("ProductName") Is DBNull.Value, Nothing, row("ProductName").ToString()),
                .Status = Convert.ToInt32(row("Status")),
                .PublishDate = If(row("PublishDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("PublishDate"))),
                .ListingDate = If(row("ListingDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("ListingDate"))),
                .CompanyRole = If(row("CompanyRole") Is DBNull.Value, Nothing, row("CompanyRole").ToString()),
                .LastUpdate = Convert.ToDateTime(row("LastUpdate")),
                .IsTerminated = Convert.ToInt32(row("IsTerminated")),
                .CreateTime = Convert.ToDateTime(row("CreateTime")),
                .CreateUser = row("CreateUser").ToString()
            }
            cases.Add(caseInfo)
        Next
        
        Return cases
    End Function
    
    ''' <summary>
    ''' 根据条件搜索案件
    ''' </summary>
    ''' <param name="caseType">案件类型</param>
    ''' <param name="caseName">案件名称</param>
    ''' <param name="productCode">产品代码</param>
    ''' <param name="status">状态</param>
    ''' <returns>案件列表</returns>
    Public Shared Function SearchCases(caseType As String, caseName As String, productCode As String, status As Integer?) As List(Of CaseInfo)
        Dim sql As String = "SELECT CaseID, CaseType, CaseName, ProductCode, ProductName, Status, PublishDate, ListingDate, CompanyRole, LastUpdate, IsTerminated, CreateTime, CreateUser FROM Cases WHERE 1=1"
        Dim parameters As New List(Of OleDbParameter)()
        
        If Not String.IsNullOrEmpty(caseType) Then
            sql += " AND CaseType LIKE ?"
            parameters.Add(New OleDbParameter("@CaseType", "%" & caseType & "%"))
        End If
        
        If Not String.IsNullOrEmpty(caseName) Then
            sql += " AND CaseName LIKE ?"
            parameters.Add(New OleDbParameter("@CaseName", "%" & caseName & "%"))
        End If
        
        If Not String.IsNullOrEmpty(productCode) Then
            sql += " AND ProductCode LIKE ?"
            parameters.Add(New OleDbParameter("@ProductCode", "%" & productCode & "%"))
        End If
        
        If status.HasValue Then
            sql += " AND Status = ?"
            parameters.Add(New OleDbParameter("@Status", status.Value))
        End If
        
        sql += " ORDER BY CreateTime DESC"
        
        Dim dt As DataTable = DbHelper.ExecuteDataTable(sql, parameters.ToArray())
        Dim cases As New List(Of CaseInfo)()
        
        For Each row As DataRow In dt.Rows
            Dim caseInfo As New CaseInfo With {
                .CaseID = Convert.ToInt32(row("CaseID")),
                .CaseType = row("CaseType").ToString(),
                .CaseName = If(row("CaseName") Is DBNull.Value, Nothing, row("CaseName").ToString()),
                .ProductCode = If(row("ProductCode") Is DBNull.Value, Nothing, row("ProductCode").ToString()),
                .ProductName = If(row("ProductName") Is DBNull.Value, Nothing, row("ProductName").ToString()),
                .Status = Convert.ToInt32(row("Status")),
                .PublishDate = If(row("PublishDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("PublishDate"))),
                .ListingDate = If(row("ListingDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("ListingDate"))),
                .CompanyRole = If(row("CompanyRole") Is DBNull.Value, Nothing, row("CompanyRole").ToString()),
                .LastUpdate = Convert.ToDateTime(row("LastUpdate")),
                .IsTerminated = Convert.ToInt32(row("IsTerminated")),
                .CreateTime = Convert.ToDateTime(row("CreateTime")),
                .CreateUser = row("CreateUser").ToString()
            }
            cases.Add(caseInfo)
        Next
        
        Return cases
    End Function
    
    ''' <summary>
    ''' 在事务中创建新案件
    ''' </summary>
    ''' <param name="transaction">事务对象</param>
    ''' <param name="caseInfo">案件信息</param>
    ''' <returns>新创建的案件ID</returns>
    Public Shared Function CreateCaseWithTransaction(transaction As OleDbTransaction, caseInfo As CaseInfo) As Integer
        Dim sql As String = "INSERT INTO Cases (CaseType, CaseName, ProductCode, ProductName, Status, PublishDate, ListingDate, CompanyRole, LastUpdate, IsTerminated, CreateTime, CreateUser) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        Dim parameters As OleDbParameter() = {
            New OleDbParameter("@CaseType", caseInfo.CaseType),
            New OleDbParameter("@CaseName", If(caseInfo.CaseName, DBNull.Value)),
            New OleDbParameter("@ProductCode", If(caseInfo.ProductCode, DBNull.Value)),
            New OleDbParameter("@ProductName", If(caseInfo.ProductName, DBNull.Value)),
            New OleDbParameter("@Status", caseInfo.Status),
            New OleDbParameter("@PublishDate", If(caseInfo.PublishDate, DBNull.Value)),
            New OleDbParameter("@ListingDate", If(caseInfo.ListingDate, DBNull.Value)),
            New OleDbParameter("@CompanyRole", If(caseInfo.CompanyRole, DBNull.Value)),
            New OleDbParameter("@LastUpdate", caseInfo.LastUpdate),
            New OleDbParameter("@IsTerminated", caseInfo.IsTerminated),
            New OleDbParameter("@CreateTime", caseInfo.CreateTime),
            New OleDbParameter("@CreateUser", caseInfo.CreateUser)
        }
        
        DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters)
        
        ' 获取新创建的案件ID
        Return Convert.ToInt32(DbHelper.ExecuteScalarWithTransaction(transaction, "SELECT @@IDENTITY"))
    End Function
    
    ''' <summary>
    ''' 在事务中批量保存案件详细信息
    ''' </summary>
    ''' <param name="transaction">事务对象</param>
    ''' <param name="caseDetails">案件详细信息列表</param>
    Public Shared Sub SaveCaseDetailsWithTransaction(transaction As OleDbTransaction, caseDetails As List(Of CaseDetail))
        If caseDetails Is Nothing OrElse caseDetails.Count = 0 Then
            Return
        End If
        
        Dim sql As String = "INSERT INTO CaseDetails (CaseID, TabIndex, FieldNo, FieldValue, FieldStatus, CreateTime) VALUES (?, ?, ?, ?, ?, ?)"
        
        For Each detail In caseDetails
            Dim parameters As OleDbParameter() = {
                New OleDbParameter("@CaseID", detail.CaseID),
                New OleDbParameter("@TabIndex", detail.TabIndex),
                New OleDbParameter("@FieldNo", detail.FieldNo),
                New OleDbParameter("@FieldValue", If(detail.FieldValue, DBNull.Value)),
                New OleDbParameter("@FieldStatus", detail.FieldStatus),
                New OleDbParameter("@CreateTime", detail.CreateTime)
            }
            
            DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters)
        Next
    End Sub
    
    ''' <summary>
    ''' 在事务中创建审查记录
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="caseId">案件ID</param>
    ''' <param name="tabIndex">标签页索引</param>
    ''' <param name="reviewerId">审查人ID</param>
    ''' <param name="reviewStatus">审查状态</param>
    ''' <param name="reviewTime">审查时间</param>
    ''' <returns>是否创建成功</returns>
    Public Shared Function CreateReviewLogWithTransaction(transaction As OleDbTransaction, caseId As Integer, tabIndex As Integer, reviewerId As String, reviewStatus As String, reviewTime As DateTime) As Boolean
        Try
            Dim sql As String = "INSERT INTO ReviewLogs (CaseID, TabIndex, ReviewerID, ReviewStatus, ReviewTime) VALUES (?, ?, ?, ?, ?)"
            
            Dim parameters As OleDbParameter() = {
                New OleDbParameter("@CaseID", caseId),
                New OleDbParameter("@TabIndex", tabIndex),
                New OleDbParameter("@ReviewerID", reviewerId),
                New OleDbParameter("@ReviewStatus", reviewStatus),
                New OleDbParameter("@ReviewTime", reviewTime)
            }
            
            DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters)
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError("创建审查记录失败", ex)
            Return False
        End Try
    End Function
    
    ' 通用保存DGV数据方法，插入时自动补全系统字段
    Public Shared Function SaveGridDataWithTransaction(transaction As OleDbTransaction, tableName As String, rows As List(Of Dictionary(Of String, String)), currentUser As String) As Boolean
        If rows Is Nothing OrElse rows.Count = 0 Then Return True
        Try
            ' 获取数据库表的字段顺序
            Dim dbFieldNames As List(Of String) = GetTableFieldNames(transaction, tableName)
            If dbFieldNames Is Nothing OrElse dbFieldNames.Count = 0 Then
                Utils.LogUtil.LogError($"无法获取表[{tableName}]的字段信息")
                Return False
            End If
            
            For Each row In rows
                ' 自动补全系统字段
                If Not row.ContainsKey("更新时间") Then row("更新时间") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If Not row.ContainsKey("更新人员") Then row("更新人员") = currentUser
                If Not row.ContainsKey("审查日期") Then row("审查日期") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                If Not row.ContainsKey("审查人员") Then row("审查人员") = currentUser
                If Not row.ContainsKey("状态") Then row("状态") = "新登录"
            Next
            
            ' 构建INSERT语句，使用数据库字段顺序
            Dim insertFields As New List(Of String)()
            Dim valuePlaceholders As New List(Of String)()
            
            For Each fieldName In dbFieldNames
                ' 检查数据中是否包含该字段
                If rows(0).ContainsKey(fieldName) Then
                    insertFields.Add($"[{fieldName}]")
                    valuePlaceholders.Add("?")
                End If
            Next
            
            Dim sql As String = $"INSERT INTO [{tableName}] ({String.Join(",", insertFields)}) VALUES ({String.Join(",", valuePlaceholders)})"
            
            ' 执行插入
            For Each row In rows
                Dim parameters As New List(Of OleDb.OleDbParameter)()
                
                ' 按照数据库字段顺序添加参数
                For Each fieldName In dbFieldNames
                    If row.ContainsKey(fieldName) Then
                        parameters.Add(New OleDb.OleDbParameter("?", If(row(fieldName), DBNull.Value)))
                    End If
                Next
                
                DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters.ToArray())
            Next
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"保存表[{tableName}]数据失败", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 获取数据库表的字段名称列表（按数据库中的顺序）
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="tableName">表名</param>
    ''' <returns>字段名称列表</returns>
    Private Shared Function GetTableFieldNames(transaction As OleDbTransaction, tableName As String) As List(Of String)
        Try
            ' 使用Schema查询获取字段信息
            Dim sql As String = $"SELECT * FROM [{tableName}] WHERE 1=0"
            
            Using cmd As New OleDbCommand(sql, transaction.Connection, transaction)
                Using reader As OleDbDataReader = cmd.ExecuteReader(CommandBehavior.SchemaOnly)
                    Dim schemaTable As DataTable = reader.GetSchemaTable()
                    Dim fieldNames As New List(Of String)()
                    
                    For Each row As DataRow In schemaTable.Rows
                        Dim columnName As String = row("ColumnName").ToString()
                        fieldNames.Add(columnName)
                    Next
                    
                    Return fieldNames
                End Using
            End Using
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"获取表[{tableName}]字段信息失败", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' 在事务中保存DataTable数据，支持行状态管理（新增、修改、删除）
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="tableName">表名</param>
    ''' <param name="dataTable">数据表</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <param name="caseId">案件ID（可选，用于关联案件）</param>
    ''' <returns>是否保存成功</returns>
    Public Shared Function SaveDataTableWithTransaction(transaction As OleDbTransaction, tableName As String, dataTable As DataTable, currentUser As String, Optional caseId As Integer = 0) As Boolean
        Try
            If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
                Return True ' 没有数据需要处理
            End If
            
            ' 获取数据库表的字段信息
            Dim dbFieldNames As List(Of String) = GetTableFieldNames(transaction, tableName)
            If dbFieldNames Is Nothing OrElse dbFieldNames.Count = 0 Then
                Utils.LogUtil.LogError($"无法获取表[{tableName}]的字段信息")
                Return False
            End If
            
            ' 分别处理不同状态的行
            Dim addedRows As New List(Of DataRow)()
            Dim modifiedRows As New List(Of DataRow)()
            Dim deletedRows As New List(Of DataRow)()
            
            ' 分类行状态
            For Each dataRow As DataRow In dataTable.Rows
                Select Case dataRow.RowState
                    Case DataRowState.Added
                        addedRows.Add(dataRow)
                    Case DataRowState.Modified
                        modifiedRows.Add(dataRow)
                    Case DataRowState.Deleted
                        deletedRows.Add(dataRow)
                End Select
            Next
            
            ' 处理新增行
            If addedRows.Count > 0 Then
                If Not InsertRows(transaction, tableName, addedRows, dbFieldNames, currentUser, caseId) Then
                    Return False
                End If
            End If
            
            ' 处理修改行
            If modifiedRows.Count > 0 Then
                If Not UpdateRows(transaction, tableName, modifiedRows, dbFieldNames, currentUser) Then
                    Return False
                End If
            End If
            
            ' 处理删除行
            If deletedRows.Count > 0 Then
                If Not DeleteRows(transaction, tableName, deletedRows) Then
                    Return False
                End If
            End If
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"保存DataTable数据失败，表名：{tableName}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 插入新增行
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="tableName">表名</param>
    ''' <param name="rows">新增行列表</param>
    ''' <param name="dbFieldNames">数据库字段名列表</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <param name="caseId">案件ID</param>
    ''' <returns>是否插入成功</returns>
    Private Shared Function InsertRows(transaction As OleDbTransaction, tableName As String, rows As List(Of DataRow), dbFieldNames As List(Of String), currentUser As String, caseId As Integer) As Boolean
        Try
            ' 构建INSERT语句
            Dim insertFields As New List(Of String)()
            Dim valuePlaceholders As New List(Of String)()
            
            For Each fieldName In dbFieldNames
                ' 检查数据中是否包含该字段
                If rows(0).Table.Columns.Contains(fieldName) Then
                    insertFields.Add($"[{fieldName}]")
                    valuePlaceholders.Add("?")
                End If
            Next
            
            Dim sql As String = $"INSERT INTO [{tableName}] ({String.Join(",", insertFields)}) VALUES ({String.Join(",", valuePlaceholders)})"
            
            ' 执行插入
            For Each dataRow In rows
                Dim parameters As New List(Of OleDbParameter)()
                
                ' 按照数据库字段顺序添加参数
                For Each fieldName In dbFieldNames
                    If dataRow.Table.Columns.Contains(fieldName) Then
                        Dim value As Object = dataRow(fieldName)
                        
                        ' 自动补全系统字段
                        If fieldName = "更新时间" Then
                            value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        ElseIf fieldName = "更新人员" Then
                            value = currentUser
                        ElseIf fieldName = "审查日期" Then
                            value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        ElseIf fieldName = "审查人员" Then
                            value = currentUser
                        ElseIf fieldName = "状态" Then
                            value = "新登录"
                        ElseIf fieldName = "caseID" AndAlso caseId > 0 Then
                            value = caseId
                        End If
                        
                        parameters.Add(New OleDbParameter("?", If(value, DBNull.Value)))
                    End If
                Next
                
                DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters.ToArray())
            Next
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"插入行数据失败，表名：{tableName}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 更新修改行
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="tableName">表名</param>
    ''' <param name="rows">修改行列表</param>
    ''' <param name="dbFieldNames">数据库字段名列表</param>
    ''' <param name="currentUser">当前用户</param>
    ''' <returns>是否更新成功</returns>
    Private Shared Function UpdateRows(transaction As OleDbTransaction, tableName As String, rows As List(Of DataRow), dbFieldNames As List(Of String), currentUser As String) As Boolean
        Try
            ' 构建UPDATE语句
            Dim updateFields As New List(Of String)()
            
            For Each fieldName In dbFieldNames
                ' 跳过主键字段（假设第一个字段是主键）
                If fieldName <> dbFieldNames(0) AndAlso rows(0).Table.Columns.Contains(fieldName) Then
                    updateFields.Add($"[{fieldName}] = ?")
                End If
            Next
            
            Dim sql As String = $"UPDATE [{tableName}] SET {String.Join(",", updateFields)} WHERE [{dbFieldNames(0)}] = ?"
            
            ' 执行更新
            For Each dataRow In rows
                Dim parameters As New List(Of OleDbParameter)()
                
                ' 添加SET子句的参数
                For Each fieldName In dbFieldNames
                    If fieldName <> dbFieldNames(0) AndAlso dataRow.Table.Columns.Contains(fieldName) Then
                        Dim value As Object = dataRow(fieldName)
                        
                        ' 自动更新系统字段
                        If fieldName = "更新时间" Then
                            value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        ElseIf fieldName = "更新人员" Then
                            value = currentUser
                        ElseIf fieldName = "状态" Then
                            value = "已修改"
                        End If
                        
                        parameters.Add(New OleDbParameter("?", If(value, DBNull.Value)))
                    End If
                Next
                
                ' 添加WHERE子句的参数（主键）
                parameters.Add(New OleDbParameter("?", dataRow(dbFieldNames(0))))
                
                DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters.ToArray())
            Next
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"更新行数据失败，表名：{tableName}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 删除行
    ''' </summary>
    ''' <param name="transaction">数据库事务</param>
    ''' <param name="tableName">表名</param>
    ''' <param name="rows">删除行列表</param>
    ''' <returns>是否删除成功</returns>
    Private Shared Function DeleteRows(transaction As OleDbTransaction, tableName As String, rows As List(Of DataRow)) As Boolean
        Try
            ' 获取数据库表的字段信息
            Dim dbFieldNames As List(Of String) = GetTableFieldNames(transaction, tableName)
            If dbFieldNames Is Nothing OrElse dbFieldNames.Count = 0 Then
                Return False
            End If
            
            ' 构建DELETE语句
            Dim sql As String = $"DELETE FROM [{tableName}] WHERE [{dbFieldNames(0)}] = ?"
            
            ' 执行删除
            For Each dataRow In rows
                Dim parameters As OleDbParameter() = {
                    New OleDbParameter("?", dataRow(dbFieldNames(0), DataRowVersion.Original))
                }
                
                DbHelper.ExecuteNonQueryWithTransaction(transaction, sql, parameters)
            Next
            
            Return True
            
        Catch ex As Exception
            Utils.LogUtil.LogError($"删除行数据失败，表名：{tableName}", ex)
            Return False
        End Try
    End Function
End Class 
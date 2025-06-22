' 案件数据访问层
Imports System.Data.OleDb
Imports System.Collections.Generic

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
End Class 
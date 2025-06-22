' 案件信息实体
Public Class CaseInfo
    Public Property CaseID As Integer
    Public Property CaseType As String
    Public Property CaseName As String
    Public Property ProductCode As String
    Public Property ProductName As String
    Public Property Status As Integer
    Public Property PublishDate As DateTime?
    Public Property ListingDate As DateTime?
    Public Property CompanyRole As String
    Public Property LastUpdate As DateTime
    Public Property IsTerminated As Integer
    Public Property CreateTime As DateTime
    Public Property CreateUser As String
End Class 
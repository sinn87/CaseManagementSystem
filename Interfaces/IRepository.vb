' 通用仓储接口
Public Interface IRepository(Of T)
    Function GetAll() As List(Of T)
    Function GetById(id As Object) As T
    Sub Add(entity As T)
    Sub Update(entity As T)
    Sub Delete(id As Object)
End Interface 
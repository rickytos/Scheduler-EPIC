Public Class ResponseStatus

    Public StatusResp As Boolean = True
    Public StatusMessage As String = ""
    Public StatusInt As Integer = 0
    Public Sub New(ByVal statusResp As Boolean, ByVal statusMsg As String, Optional ByVal statusInt As Integer = 0)
        statusResp = statusResp
        StatusMessage = statusMsg
        statusInt = statusInt
    End Sub

End Class

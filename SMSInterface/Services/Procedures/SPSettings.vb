Public Class SPSettings
    Inherits DataBaseServices

    Public Shared Function UpdateEmailSettings(ByVal Param As SendEmailSettingParams, ByVal dbManager As DataBaseHelper, ByVal ExcludeParams As String()) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of SendEmailSettingParams) With {.StoreProcedure = "UpdateSendEmailSetting", .Parameter = Param, .ExcludeParameter = ExcludeParams}

        Return CallProcedures(Of SendEmailSettingParams)(newParam, dbManager)

    End Function

    Shared Function InsertEmailSettings(ByVal paramInsert As SendEmailSettingParams, dbManager As DataBaseHelper) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of SendEmailSettingParams) With {.StoreProcedure = "InsertSendEmailSetting", .Parameter = paramInsert}

        Return CallProcedures(Of SendEmailSettingParams)(newParam, dbManager)
    End Function

End Class

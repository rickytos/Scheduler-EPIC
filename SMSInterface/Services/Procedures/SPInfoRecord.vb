Imports System.Data.SqlClient
Imports System.Reflection

Public Class SPInfoRecord
    Inherits DataBaseServices

    Public Shared Function getInfoRecordXML() As CallProcedureResult
        Return CallProcedures("SP_Get_InfoRecord")
    End Function

    Public Shared Function UpdateInfoRecordRequest(ByVal parameter As IDictionary(Of String, Object)) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of IDictionary(Of String, Object)) With {.StoreProcedure = "Upd_InfoRecord_Statusheader", .Parameter = parameter}
        Return CallProcedures(newParam)
    End Function

    Public Shared Function UpdateSAPInfoRecordResponse(ByVal parameterResponse As InfoRecordRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of InfoRecordRes) With {.StoreProcedure = "Upd_InfoRecord_SAPNumber", .Parameter = parameterResponse}

        Return CallProcedures(Of InfoRecordRes)(newParam)

    End Function

End Class

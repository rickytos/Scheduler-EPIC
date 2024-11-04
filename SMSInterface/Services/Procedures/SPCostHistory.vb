Imports System.Data.SqlClient
Imports System.Reflection

Public Class SPCostHistory
    Inherits DataBaseServices

    'CostHistoryRequest
    'CostHistoryRequest
    'asdas
    Public Shared Function SAPCostHistoryRequest() As CallProcedureResult

        Return CallProcedures("SP_IF")

    End Function

    Public Shared Function SAPCostHistoryResponse(ByVal Param As CostHistorySAPRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of CostHistorySAPRes) With {.StoreProcedure = "Upd_SAPCostHistoryResponse", .Parameter = param}

        Return CallProcedures(Of CostHistorySAPRes)(newParam)

    End Function

    Public Shared Function KDPLCostHistoryResponse() As CallProcedureResult

        Return CallProcedures("Upd_KDPLCostHistoryResponse")

    End Function



End Class

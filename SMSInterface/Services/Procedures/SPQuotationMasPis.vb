Imports System.Data.SqlClient
Imports System.Reflection

Public Class SPQuotationMasPis
    Inherits DataBaseServices

    'TODO : add Quotation Request
    Public Shared Function SAPQuotationRequest() As CallProcedureResult

        Return CallProcedures("SP_IF_SendRequestQuotation")

    End Function

    'TODO : add Quotation Response 
    Public Shared Function SAPQuotationResponse(ByVal Param As QuotationMasPisRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of QuotationMasPisRes) With {.StoreProcedure = "SP_IF_GetResponseQuotation", .Parameter = Param}

        Return CallProcedures(Of QuotationMasPisRes)(newParam)

    End Function

    'TODO : add Cost Target
    Public Shared Function SAPCostTarget(ByVal Param As CostHistorySAPRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of CostHistorySAPRes) With {.StoreProcedure = "SP_IF_GetResponseQuotation", .Parameter = Param}

        Return CallProcedures(Of CostHistorySAPRes)(newParam)

    End Function

    Public Shared Function SAPInfoRecord(ByVal Param As CostHistorySAPRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of CostHistorySAPRes) With {.StoreProcedure = "Upd_SAPCostHistoryResponse", .Parameter = Param}

        Return CallProcedures(Of CostHistorySAPRes)(newParam)

    End Function

End Class

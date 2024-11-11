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
    Public Shared Function SAPCostTarget(ByVal Param As QuotationMaspisCostTarget) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of QuotationMaspisCostTarget) With {.StoreProcedure = "SP_IF_GetCostTargetQuotation", .Parameter = Param}

        Return CallProcedures(Of QuotationMaspisCostTarget)(newParam)

    End Function

    Public Shared Function IsAlreadyGetResp(ByVal quotNumber As String) As Boolean

        Dim queryString = String.Format("SELECT response FROM PA_Quotation_Interface WHERE Quotation_No = '{0}'", quotNumber)

        Return ExecuteQuery(queryString, QueryType.NonQuery).BooleanResult

    End Function

    Public Shared Function IsMtrlResExists(ByVal quotNumber As String, ByVal PartNo As String) As Object

        Dim queryString = String.Format("SELECT * FROM PAQuotationTrack_History WHERE [Quotation_No] = '{0}' AND [Part_No] = '{1}'", quotNumber, PartNo)

        Return ExecuteQuery(queryString, QueryType.NonQuery).BooleanResult

    End Function

    'TODO: For Info Record
    'Public Shared Function SAPInfoRecord(ByVal Param As CostHistorySAPRes) As CallProcedureResult

    '    Dim newParam As New ParamCallFunction(Of CostHistorySAPRes) With {.StoreProcedure = "Upd_SAPCostHistoryResponse", .Parameter = Param}

    '    Return CallProcedures(Of CostHistorySAPRes)(newParam)

    'End Function
    Public Shared Function SAPInfoRecordRequest()

        Return CallProcedures("SP_IF_GetInforecordMaspisRequest")

    End Function

    'Public Shared Function SAPInfoReordResponse()
    '    Dim newParam As New ParamCallFunction(Of QuotationMaspisCostTarget) With {.StoreProcedure = "SP_IF_GetCostTargetQuotation", .Parameter = Param}

    '    Return CallProcedures(Of QuotationMaspisCostTarget)(newParam)
    'End Function

End Class

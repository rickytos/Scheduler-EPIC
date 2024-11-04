Public Class SPSettings
    Inherits DataBaseServices

    Public Shared Function SAPQuotationResponse(ByVal Param As QuotationMasPisRes) As CallProcedureResult

        Dim newParam As New ParamCallFunction(Of QuotationMasPisRes) With {.StoreProcedure = "SP_IF_GetResponseQuotation", .Parameter = Param}

        Return CallProcedures(Of QuotationMasPisRes)(newParam)

    End Function

End Class

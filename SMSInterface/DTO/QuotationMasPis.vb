Public Class QuotationMasPisReq
    Public Property QUOTATION As String
    Public Property QUOT_SEQ As String
    Public Property SUPPLIER As String
    Public Property QUOT_DATE As Date
    Public Property ITEMNO As Integer
    Public Property MATERIAL As String
    Public Property COST_NO As Integer
    Public Property COST_NAME As String
    Public Property CURR_PRICE As Long
    Public Property QUOT_PRICE As Long
    Public Property CURRENCY As String
    Public Property QUOT_DEAL As String
End Class


Public Class QuotationMasPisRes
    Public Property QUOTATION As String
    Public Property QUOT_SEQ As String
    Public Property MATERIAL As String
    Public Property STATUS As String
    Public Property MESSAGE As String
End Class


Public Class QuotationMaspisCostTarget
    Public Property QUOTATION As String
    Public Property QUOT_SEQ As String
    Public Property PROP_NO As String
    Public Property SUPPLIER As String
    Public Property QUOT_DATE As String
    Public Property ITEMNO As String
    Public Property MATERIAL As String
    Public Property COST_NO As String
    Public Property COST_NAME As String
    Public Property COST_TARGET As String
    Public Property CURRENCY As String
    Public Property FINAL_APP As String
End Class

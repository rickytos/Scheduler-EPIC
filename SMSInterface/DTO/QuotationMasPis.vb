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


Public Class InfoRecordMaspisRequest
    Public Property VENDOR As String
    Public Property MATNR As String
    Public Property PURORG As String
    Public Property WERKS As String
    Public Property STANDARD As String
    Public Property REM1 As String
    Public Property REM2 As String
    Public Property REM3 As String
    Public Property VEN_MAT As String
    Public Property COUNTRY As String
    Public Property REGION As String
    Public Property VEN_MAT_GRP As String
    Public Property MANUF As String
    Public Property SALES_PER As String
    Public Property TELEPH As String
    Public Property ORD_UNIT As String
    Public Property ORD_DIN As String
    Public Property ORD_NUM As String
    Public Property NUMBERZ As String
    Public Property PLAN_DEL_TIME As String
    Public Property PUR_GRP As String
    Public Property STND_QTY As String
    Public Property MIN_QTY As String
    Public Property ROUND_PROF As String
    Public Property CONF_KEY As String
    Public Property TAX_CODE As String
    Public Property VALID_ON As String
    Public Property VALID_TO As String
    Public Property PRICE As String
    Public Property CURR As String

End Class

Public Class InfoRecordMaspisResponse
    Public Property VENDOR As String
    Public Property MATNR As String
    Public Property INFNR As String
    Public Property CONFIRM As String
    Public Property MESSAGE As String
End Class
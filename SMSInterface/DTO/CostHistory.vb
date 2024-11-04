Public Class CostHistorySAPReq
    Public Property MaterialCode As String = ""
    Public Property MaterialName As String = ""
    Public Property SupplierName As String = ""
    Public Property SupplierCode As String = ""
    Public Property VariantName As String = ""
    Public Property Dates As String = ""

End Class

Public Class CostHistorySAPRes
    Public Property MaterialCode As String = ""
    Public Property MaterialName As String = ""
    Public Property SupplierName As String = ""
    Public Property SupplierCode As String = ""
    Public Property Qty As String = ""
    Public Property Price As String = ""
End Class

Public Class CostHistoryKDPLRes
    Public Property MaterialNo As String = ""
    Public Property VariantNo As String = ""
    Public Property VariantName As String = ""
    Public Property PartNo As String = ""
    Public Property PartName As String = ""
    Public Property SupplierID As String = ""
    Public Property SupplierName As String = ""
    Public Property ParentNumber As String = ""
    Public Property Qty As String = ""
    Public Property BOMUsage As String = ""
    Public Property DMRegisterNo As String = ""
    Public Property ReasonOfChange As String = ""
    Public Property EffectiveDate As String = ""
    Public Property ChangeInformation As String = ""
    Public Property ChangeClassification As String = ""
    Public Property OldMaterialNo As String = ""
End Class
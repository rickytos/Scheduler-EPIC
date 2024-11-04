Public Class InterfaceSettingClassDto

    ' Public properties
    Public Property ConnectionString As String
    Public Property SMTP_User As String
    Public Property SMTP_Password As String
    Public Property SMTP_Port As String
    Public Property SMTP_Address As String
    Public Property Email As String
    Public Property JSON_Path As String
    Public Property JSON_PathBackup As String
    Public Property XMLESSBudget_Path As String
    Public Property XMLESSBudget_PathBackup As String
    Public Property XMLSAP_Path As String
    Public Property XMLSAP_PathBackup As String
    Public Property XMLIAPrice_Path As String
    Public Property IAMIPressPartLocalComponent_Path As String
    Public Property IAMIPressPartLocalComponent_PathBackup As String
    Public Property IAMIPressPartLocalComponentXMLToSAP_Path As String
    Public Property CostHistoryXMLToSAP_Path As String
    Public Property PurchaseRequest_Path As String
    Public Property PurchaseRequest_PathBackup As String
    Public Property PurchaseRequest_PathFailed As String
    Public Property PurchaseRequest_PathOut As String
    Public Property GoodsReceiveXMLToDMMS_Path As String
    Public Property MaterialMasterXMLtoSAP_Path As String
    Public Property MaterialMasterXML_BackupPath As String
    Public Property MaterialMasterXML_Resp_Pth As String
    Public Property MaterialMasterXML_Resp_BackupPth As String
    Public Property IR_XML_ReqPath As String
    Public Property IR_XML_ReqPath_Backup As String
    Public Property IR_XML_RespPath As String
    Public Property IR_XML_RespPath_Backup As String

    ' Constructor to initialize properties
    Public Sub New(ByVal pConStr As String, ByVal smtpUser As String, ByVal smtpPassword As String, ByVal smtpPort As String,
                   ByVal smtpAddress As String, ByVal email As String, ByVal jsonPath As String, ByVal jsonPathBackup As String,
                   ByVal xmlESSBudgetPath As String, ByVal xmlESSBudgetPathBackup As String, ByVal xmlSAPPath As String,
                   ByVal xmlSAPPathBackup As String, ByVal xmlIAPricePath As String, ByVal iaMIPressPartLocalComponentPath As String,
                   ByVal iaMIPressPartLocalComponentPathBackup As String, ByVal iaMIPressPartLocalComponentXMLToSAPPath As String,
                   ByVal costHistoryXMLToSAPPath As String, ByVal purchaseRequestPath As String,
                   ByVal purchaseRequestPathBackup As String, ByVal purchaseRequestPathFailed As String,
                   ByVal purchaseRequestPathOut As String, ByVal goodsReceiveXMLToDMMSPath As String,
                   ByVal materialMasterXMLtoSAPPath As String, ByVal materialMasterXMLBackupPath As String,
                   ByVal materialMasterXMLRespPth As String, ByVal materialMasterXMLRespBackupPth As String,
                   ByVal irXMLReqPath As String, ByVal irXMLReqPathBackup As String, ByVal irXMLRespPath As String,
                   ByVal irXMLRespPathBackup As String)
        ' Initialize properties
        Me.ConnectionString = pConStr
        Me.SMTP_User = smtpUser
        Me.SMTP_Password = smtpPassword
        Me.SMTP_Port = smtpPort
        Me.SMTP_Address = smtpAddress
        Me.Email = email
        Me.JSON_Path = jsonPath
        Me.JSON_PathBackup = jsonPathBackup
        Me.XMLESSBudget_Path = xmlESSBudgetPath
        Me.XMLESSBudget_PathBackup = xmlESSBudgetPathBackup
        Me.XMLSAP_Path = xmlSAPPath
        Me.XMLSAP_PathBackup = xmlSAPPathBackup
        Me.XMLIAPrice_Path = xmlIAPricePath
        Me.IAMIPressPartLocalComponent_Path = iaMIPressPartLocalComponentPath
        Me.IAMIPressPartLocalComponent_PathBackup = iaMIPressPartLocalComponentPathBackup
        Me.IAMIPressPartLocalComponentXMLToSAP_Path = iaMIPressPartLocalComponentXMLToSAPPath
        Me.CostHistoryXMLToSAP_Path = costHistoryXMLToSAPPath
        Me.PurchaseRequest_Path = purchaseRequestPath
        Me.PurchaseRequest_PathBackup = purchaseRequestPathBackup
        Me.PurchaseRequest_PathFailed = purchaseRequestPathFailed
        Me.PurchaseRequest_PathOut = purchaseRequestPathOut
        Me.GoodsReceiveXMLToDMMS_Path = goodsReceiveXMLToDMMSPath
        Me.MaterialMasterXMLtoSAP_Path = materialMasterXMLtoSAPPath
        Me.MaterialMasterXML_BackupPath = materialMasterXMLBackupPath
        Me.MaterialMasterXML_Resp_Pth = materialMasterXMLRespPth
        Me.MaterialMasterXML_Resp_BackupPth = materialMasterXMLRespBackupPth
        Me.IR_XML_ReqPath = irXMLReqPath
        Me.IR_XML_ReqPath_Backup = irXMLReqPathBackup
        Me.IR_XML_RespPath = irXMLRespPath
        Me.IR_XML_RespPath_Backup = irXMLRespPathBackup
    End Sub
End Class


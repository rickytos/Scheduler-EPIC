Imports System.Data.SqlClient

Public Class frmSetting
    Protected ConStr As String
    Dim cfg As clsConfig

    Private Sub frmSetting_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        cfg = New clsConfig

        gs_Server = cfg.Server
        gs_Database = cfg.Database
        gs_User = cfg.User
        gs_Password = cfg.Password

        Builder = New SqlConnectionStringBuilder
        Builder.DataSource = gs_Server
        Builder.InitialCatalog = gs_Database
        Builder.UserID = gs_User
        Builder.Password = gs_Password
        Builder.IntegratedSecurity = cfg.WinMode = ""
        Builder.ApplicationName = "SMS"

        ConStr = Builder.ConnectionString

        Dim dt As DataTable
        'SMTP DMMS PR
        dt = ClsInterfaceSettingDB.getDataSMTP_DMMS_PR(ConStr)
        If dt.Rows.Count > 0 Then
            gs_SMTPAddress_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Address"))
            gs_SMTPPort_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Port"))
            gs_SMTPUser_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_User"))
            gs_SMTPPassword_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Password"))
            gs_EmailFrom_DMMS_PR = Trim(dt.Rows(0).Item("FromEmail"))
            gs_EmailTo_DMMS_PR = Trim(dt.Rows(0).Item("ReceiptEmail"))
            gs_Subject_DMMS_PR = Trim(dt.Rows(0).Item("Subject"))
        Else
            gs_SMTPAddress_DMMS_PR = ""
            gs_SMTPPort_DMMS_PR = ""
            gs_SMTPUser_DMMS_PR = ""
            gs_SMTPPassword_DMMS_PR = ""
            gs_EmailFrom_DMMS_PR = ""
            gs_EmailTo_DMMS_PR = ""
            gs_Subject_DMMS_PR = ""
        End If
        dt.Reset()

        dt = ClsInterfaceSettingDB.getData(ConStr)
        
        SetupTextPath(dt)

        ActiveControl = txtSMTPAddress_DMMS_PR
    End Sub

    Public Sub SetupTextPath(ByVal dt As DataTable)
        If dt.Rows.Count > 0 Then
            gs_JSONPath = Trim(dt.Rows(0).Item("JSON_Path"))
            gs_JSONPathBackup = Trim(dt.Rows(0).Item("JSON_PathBackup"))
            gs_XMLPath = Trim(dt.Rows(0).Item("XML_IABudget_Path"))
            gs_XMLPathBackup = Trim(dt.Rows(0).Item("XML_IABudget_PathBackup"))
            gs_SAPPath = Trim(dt.Rows(0).Item("XML_SAP_Path"))
            gs_SAPPathBackup = Trim(dt.Rows(0).Item("XML_SAP_PathBackup"))
            gs_IAPricePath = Trim(dt.Rows(0).Item("XML_IAPrice_Path"))
            gs_IAMI_PressPartPath = Trim(dt.Rows(0).Item("IAMI_PressPart_Path"))
            gs_IAMI_PressPartPathBackup = Trim(dt.Rows(0).Item("IAMI_PressPart_PathBackup"))
            gs_IAMI_PressPartXMLToSAPPath = Trim(dt.Rows(0).Item("IAMI_PressPart_XMLToSAP_Path"))
            gs_CostHistoryXMLToSAPPath = Trim(dt.Rows(0).Item("CostHistory_XMLToSAP_Path"))
            gs_PurchaseRequestPath = Trim(dt.Rows(0).Item("XML_PurchaseRequest_Path"))
            gs_PurchaseRequestPathBackup = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathBackup"))
            gs_PurchaseRequestPathFailed = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathFailed"))
            gs_PurchaseRequestPathOut = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathOut"))
            gs_GoodsReceiveXMLToDMMSPath = Trim(dt.Rows(0).Item("GoodsReceive_XMLToDMMS_Path"))

            gs_MaterialMasterXMLtoSAP_Path_Req = Trim(dt.Rows(0).Item("MaterialMasterXML_ToSAP_Path"))
            gs_MaterialMasterXMLtoSAP_BackupPath_Req = Trim(dt.Rows(0).Item("MaterialMasterXML_ToSAP_PathBackup"))
            gs_MaterialMasterXMLtoSAP_Path_Res = Trim(dt.Rows(0).Item("MaterialMasterXML_Resp_Pth"))
            gs_MaterialMasterXMLtoSAP_BackupPath_Res = Trim(dt.Rows(0).Item("MaterialMasterXML_Resp_BackupPth"))
            gs_IR_XMLtoSAP_Path_Req = Trim(dt.Rows(0).Item("IR_XML_Req_Path"))
            gs_IR_XMLtoSAP_BackupPath_Req = Trim(dt.Rows(0).Item("IR_XML_Req_BackupPath"))
            gs_IR_XMLtoSAP_Path_Res = Trim(dt.Rows(0).Item("IR_XML_Res_Path"))
            gs_IR_XMLtoSAP_BackupPath_Res = Trim(dt.Rows(0).Item("IR_XML_Res_BackupPath"))
            gs_PAQuotReqPath = Trim(dt.Rows(0).Item("PAQuotReqPath"))
            gs_PAQuotResPath = Trim(dt.Rows(0).Item("PAQuotResPath"))
            gs_PAQuotResBackupPath = Trim(dt.Rows(0).Item("PAQuotResBackupPath"))

        Else
            gs_JSONPath = ""
            gs_JSONPathBackup = ""
            gs_XMLPath = ""
            gs_XMLPathBackup = ""
            gs_SAPPath = ""
            gs_SAPPathBackup = ""
            gs_IAPricePath = ""
            gs_IAMI_PressPartPath = ""
            gs_IAMI_PressPartPathBackup = ""
            gs_IAMI_PressPartXMLToSAPPath = ""
            gs_CostHistoryXMLToSAPPath = ""
            gs_PurchaseRequestPath = ""
            gs_PurchaseRequestPathBackup = ""
            gs_PurchaseRequestPathFailed = ""
            gs_PurchaseRequestPathOut = ""
            gs_GoodsReceiveXMLToDMMSPath = ""

            gs_MaterialMasterXMLtoSAP_Path_Req = ""
            gs_MaterialMasterXMLtoSAP_BackupPath_Req = ""
            gs_MaterialMasterXMLtoSAP_Path_Res = ""
            gs_MaterialMasterXMLtoSAP_BackupPath_Res = ""
            gs_IR_XMLtoSAP_Path_Req = ""
            gs_IR_XMLtoSAP_BackupPath_Req = ""
            gs_IR_XMLtoSAP_Path_Res = ""
            gs_IR_XMLtoSAP_BackupPath_Res = ""

            gs_PAQuotReqPath = String.Empty
            gs_PAQuotResPath = String.Empty
            gs_PAQuotResBackupPath = String.Empty
        End If

        txtSMTPAddress_DMMS_PR.Text = gs_SMTPAddress_DMMS_PR
        txtSMTPPort_DMMS_PR.Text = gs_SMTPPort_DMMS_PR
        txtSMTPUser_DMMS_PR.Text = gs_SMTPUser_DMMS_PR
        txtSMTPPassword_DMMS_PR.Text = gs_SMTPPassword_DMMS_PR
        txtEmailFrom_DMMS_PR.Text = gs_EmailFrom_DMMS_PR
        txtEmailTo_DMMS_PR.Text = gs_EmailTo_DMMS_PR
        txtSubject_DMMS_PR.Text = gs_Subject_DMMS_PR

        txtJSONPath.Text = gs_JSONPath
        txtJSONPathBackup.Text = gs_JSONPathBackup
        txtESSBudgetPath.Text = gs_XMLPath
        txtESSBudgetPathBackup.Text = gs_XMLPathBackup
        txtSAPFilePath.Text = gs_SAPPath
        txtSAPFilePathBackup.Text = gs_SAPPathBackup
        txtIAPriceDestinationPath.Text = gs_IAPricePath
        txtIAMIPressPartLocalComponentPath.Text = gs_IAMI_PressPartPath
        txtIAMIPressPartLocalComponentPathBackup.Text = gs_IAMI_PressPartPathBackup
        txtIAMIPressPartLocalComponentXMLToSAPPath.Text = gs_IAMI_PressPartXMLToSAPPath
        txtCostHistoryXMLToSAPPath.Text = gs_CostHistoryXMLToSAPPath
        txtPurchaseRequestPath.Text = gs_PurchaseRequestPath
        txtPurchaseRequestPathBackup.Text = gs_PurchaseRequestPathBackup
        txtPurchaseRequestPathFailed.Text = gs_PurchaseRequestPathFailed
        txtPurchaseRequestPathOut.Text = gs_PurchaseRequestPathOut
        txtGoodsReceiveXMLToDMMSPath.Text = gs_GoodsReceiveXMLToDMMSPath

        txtMaterialMasterXMLtoSAPPath_Req.Text = gs_MaterialMasterXMLtoSAP_Path_Req
        txtMaterialMasterXMLtoSAPBackupPath_Req.Text = gs_MaterialMasterXMLtoSAP_BackupPath_Req
        txtMaterialMasterXMLtoSAPPath_Res.Text = gs_MaterialMasterXMLtoSAP_Path_Res
        txtMaterialMasterXMLtoSAPBackupPath_Res.Text = gs_MaterialMasterXMLtoSAP_BackupPath_Res
        txtIR_XML_RequestPath.Text = gs_IR_XMLtoSAP_Path_Req
        txtIR_XML_RequestBackupPath.Text = gs_IR_XMLtoSAP_BackupPath_Req
        txtIR_XML_ResponsePath.Text = gs_IR_XMLtoSAP_Path_Res
        txtIR_XML_ResponseBackupPath.Text = gs_IR_XMLtoSAP_BackupPath_Res

        txtQuotationRequestPath.Text = gs_PAQuotReqPath
        txtQuotationResponsePath.Text = gs_PAQuotResPath
        txtQuotationResponseBackupPath.Text = gs_PAQuotResBackupPath
    End Sub

    Public Shared Function UpdateSMTP_Setting_DMMS_PR(ByVal pConStr As String, ByVal pSMTPAddress_DMMS_PR As String, ByVal pSMTPPort_DMMS_PR As String, ByVal pSMTPUser_DMMS_PR As String, ByVal pSMTPPassword_DMMS_PR As String, _
                                            ByVal pEmailFrom_DMMS_PR As String, ByVal pEmailTo_DMMS_PR As String, ByVal pSubject_DMMS_PR As String) As Integer

        Dim sql As String
        Dim retValue As Integer = 0

        Dim sSystemFrom As String = "DMMS"
        Dim sModule As String = "PR"

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "sp_SMTP_Settings_Ins"

                Dim cmd As New SqlCommand()
                cmd.CommandText = sql
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@SystemFrom", sSystemFrom)
                cmd.Parameters.AddWithValue("@Module", sModule)
                cmd.Parameters.AddWithValue("@SMTP_Address", pSMTPAddress_DMMS_PR)
                cmd.Parameters.AddWithValue("@SMTP_Port", pSMTPPort_DMMS_PR)
                cmd.Parameters.AddWithValue("@SMTP_User", pSMTPUser_DMMS_PR)
                cmd.Parameters.AddWithValue("@SMTP_Password", pSMTPPassword_DMMS_PR)
                cmd.Parameters.AddWithValue("@FromEmail", pEmailFrom_DMMS_PR)
                cmd.Parameters.AddWithValue("@ReceiptEmail", pEmailTo_DMMS_PR)
                cmd.Parameters.AddWithValue("@Subject", pSubject_DMMS_PR)
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try
    End Function

    Public Shared Function UpdateInterface_Setting(ByVal paramUpdate As SendEmailSettingParams) As Boolean

        Try

            Using dbManager As New DataBaseHelper()
                Try

                    Dim excludedParams = {"SMTP_User", "SMTP_Password", "SMTP_Port", "SMTP_Address", "Email"}
                    Dim result As CallProcedureResult = SPSettings.UpdateEmailSettings(paramUpdate, dbManager, excludedParams)

                    If Not String.IsNullOrEmpty(result.ResultMessage) Then
                        'Throw New ArgumentException(result.ResultMessage)
                        Dim resultInsert = SPSettings.InsertEmailSettings(paramUpdate, dbManager)

                        If Not String.IsNullOrEmpty(resultInsert.ResultMessage) Then
                            Throw New ArgumentException(resultInsert.ResultMessage)
                        End If

                    End If

                    dbManager.Commit()

                    Return True
                Catch ex As Exception
                    dbManager.Rollback()

                    Return False
                End Try
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Function

    Private Function ValidateTextBox(textBox As TextBox, message As String) As Boolean
        If textBox.Text.Trim() = "" Then
            MsgBox(message)
            textBox.Focus()
            Return False
        End If
        Return True
    End Function

    Private Function CheckValidate() As Boolean
        ' Validate SMTP DMMS - PR
        If Not ValidateTextBox(txtSMTPAddress_DMMS_PR, "Please input SMTP Address for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtSMTPPort_DMMS_PR, "Please input SMTP Port for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtSMTPUser_DMMS_PR, "Please input User for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtSMTPPassword_DMMS_PR, "Please input Password for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtEmailFrom_DMMS_PR, "Please input Email From for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtEmailTo_DMMS_PR, "Please input Email To for DMMS PR!") Then Return False
        If Not ValidateTextBox(txtSubject_DMMS_PR, "Please input Email Subject for DMMS PR!") Then Return False

        ' Validate path folders
        If Not ValidateTextBox(txtJSONPath, "Please input EPIC 1 Supplier Management's Path Folder!") Then Return False
        If Not ValidateTextBox(txtJSONPathBackup, "Please input EPIC 1 Supplier Management's Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtESSBudgetPath, "Please input ESS Budget's Path Folder!") Then Return False
        If Not ValidateTextBox(txtESSBudgetPathBackup, "Please input ESS Budget's Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtSAPFilePath, "Please input SAP PO GR's Path Folder!") Then Return False
        If Not ValidateTextBox(txtSAPFilePathBackup, "Please input SAP PO GR's Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtIAPriceDestinationPath, "Please input IA Price Destination's Path Folder!") Then Return False
        If Not ValidateTextBox(txtIAMIPressPartLocalComponentXMLToSAPPath, "Please input IAMI Press Part & Local Component XML To SAP's Path Folder!") Then Return False
        If Not ValidateTextBox(txtIAMIPressPartLocalComponentPath, "Please input IAMI Press Part & Local Component's Path Folder!") Then Return False
        If Not ValidateTextBox(txtIAMIPressPartLocalComponentPathBackup, "Please input IAMI Press Part & Local Component's Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtCostHistoryXMLToSAPPath, "Please input Cost History XML To SAP's Path Folder!") Then Return False
        If Not ValidateTextBox(txtPurchaseRequestPath, "Please input Purchase Request's Path Folder!") Then Return False
        If Not ValidateTextBox(txtPurchaseRequestPathBackup, "Please input Purchase Request's Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtPurchaseRequestPathFailed, "Please input Purchase Request's Failed Path Folder!") Then Return False
        If Not ValidateTextBox(txtPurchaseRequestPathOut, "Please input Purchase Request Response's Path Folder!") Then Return False
        If Not ValidateTextBox(txtGoodsReceiveXMLToDMMSPath, "Please input Goods Receive XML To DMMS's Path Folder!") Then Return False

        ' Validate Material Master paths
        If Not ValidateTextBox(txtMaterialMasterXMLtoSAPPath_Req, "Please input Material Master's Request Path Folder!") Then Return False
        If Not ValidateTextBox(txtMaterialMasterXMLtoSAPBackupPath_Req, "Please input Material Master's Request Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtMaterialMasterXMLtoSAPBackupPath_Res, "Please input Material Master's Response Path Folder!") Then Return False
        If Not ValidateTextBox(txtMaterialMasterXMLtoSAPBackupPath_Res, "Please input Material Master's Response Backup Path Folder!") Then Return False

        ' Validate Info Record paths
        If Not ValidateTextBox(txtIR_XML_RequestPath, "Please input Info Record's Request Path Folder!") Then Return False
        If Not ValidateTextBox(txtIR_XML_RequestBackupPath, "Please input Info Record's Request Backup Path Folder!") Then Return False
        If Not ValidateTextBox(txtIR_XML_ResponsePath, "Please input Info Record's Response Path Folder!") Then Return False
        If Not ValidateTextBox(txtIR_XML_ResponseBackupPath, "Please input Info Record's Response Backup Path Folder!") Then Return False

        ' Validate Quotation Master paths
        If Not ValidateTextBox(txtQuotationRequestPath, "Please input Quotation Request's Path Folder!") Then Return False
        If Not ValidateTextBox(txtQuotationResponsePath, "Please input Quotation Response's Path Folder!") Then Return False
        If Not ValidateTextBox(txtQuotationResponseBackupPath, "Please input Quotation Response's Backup Path Folder!") Then Return False

        Return True
    End Function

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            Dim checkValidation = CheckValidate()
            If Not checkValidation Then
                Exit Sub
            End If

            Dim errorCount As Integer = 0

            'update SMTP DMMS PR
            If UpdateSMTP_Setting_DMMS_PR(ConStr, txtSMTPAddress_DMMS_PR.Text.Trim, txtSMTPPort_DMMS_PR.Text.Trim, txtSMTPUser_DMMS_PR.Text.Trim, txtSMTPPassword_DMMS_PR.Text.Trim, _
                                          txtEmailFrom_DMMS_PR.Text.Trim, txtEmailTo_DMMS_PR.Text.Trim, txtSubject_DMMS_PR.Text.Trim) Then

                cfg = New clsConfig

                gs_Server = cfg.Server
                gs_Database = cfg.Database
                gs_User = cfg.User
                gs_Password = cfg.Password

                Builder = New SqlConnectionStringBuilder
                Builder.DataSource = gs_Server
                Builder.InitialCatalog = gs_Database
                Builder.UserID = gs_User
                Builder.Password = gs_Password
                Builder.IntegratedSecurity = cfg.WinMode = ""
                Builder.ApplicationName = "SMS"

                ConStr = Builder.ConnectionString

                Dim dt As DataTable
                dt = ClsInterfaceSettingDB.getDataSMTP_DMMS_PR(ConStr)

                If dt.Rows.Count > 0 Then
                    gs_SMTPAddress_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Address"))
                    gs_SMTPPort_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Port"))
                    gs_SMTPUser_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_User"))
                    gs_SMTPPassword_DMMS_PR = Trim(dt.Rows(0).Item("SMTP_Password"))
                    gs_EmailFrom_DMMS_PR = Trim(dt.Rows(0).Item("FromEmail"))
                    gs_EmailTo_DMMS_PR = Trim(dt.Rows(0).Item("ReceiptEmail"))
                    gs_Subject_DMMS_PR = Trim(dt.Rows(0).Item("Subject"))
                Else
                    gs_SMTPAddress_DMMS_PR = ""
                    gs_SMTPPort_DMMS_PR = ""
                    gs_SMTPUser_DMMS_PR = ""
                    gs_SMTPPassword_DMMS_PR = ""
                    gs_EmailFrom_DMMS_PR = ""
                    gs_EmailTo_DMMS_PR = ""
                    gs_Subject_DMMS_PR = ""
                End If

                txtSMTPAddress_DMMS_PR.Text = gs_SMTPAddress_DMMS_PR
                txtSMTPPort_DMMS_PR.Text = gs_SMTPPort_DMMS_PR
                txtSMTPUser_DMMS_PR.Text = gs_SMTPUser_DMMS_PR
                txtSMTPPassword_DMMS_PR.Text = gs_SMTPPassword_DMMS_PR
                txtEmailFrom_DMMS_PR.Text = gs_EmailFrom_DMMS_PR
                txtEmailTo_DMMS_PR.Text = gs_EmailTo_DMMS_PR
                txtSubject_DMMS_PR.Text = gs_Subject_DMMS_PR
            Else
                errorCount = errorCount + 1
            End If

            'update Interface Setting
            Dim paramSettings = SetupParams()
            If UpdateInterface_Setting(paramSettings) Then

                cfg = New clsConfig

                gs_Server = cfg.Server
                gs_Database = cfg.Database
                gs_User = cfg.User
                gs_Password = cfg.Password

                Builder = New SqlConnectionStringBuilder
                Builder.DataSource = gs_Server
                Builder.InitialCatalog = gs_Database
                Builder.UserID = gs_User
                Builder.Password = gs_Password
                Builder.IntegratedSecurity = cfg.WinMode = ""
                Builder.ApplicationName = "SMS"

                ConStr = Builder.ConnectionString

                Dim dt As DataTable
                dt = ClsInterfaceSettingDB.getData(ConStr)

                SetupTextPath(dt)

            Else
                errorCount = errorCount + 1
            End If

            If errorCount = 0 Then
                MsgBox("Update data successfully!", MsgBoxStyle.Information, "Info")
            Else
                MsgBox("Update data is not successfully!", MsgBoxStyle.Critical, "Error")
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Public Function SetupParams() As SendEmailSettingParams
        ' Create an instance of SendEmailSettingParams
        Return New SendEmailSettingParams With {
            .SMTP_User = txtSMTPUser_DMMS_PR.Text.Trim(),
            .SMTP_Password = txtSMTPPassword_DMMS_PR.Text.Trim(),
            .SMTP_Port = txtSMTPPort_DMMS_PR.Text.Trim(),
            .SMTP_Address = txtSMTPAddress_DMMS_PR.Text.Trim(),
            .Email = txtSMTPUser_DMMS_PR.Text.Trim(),
            .JSON_Path = txtJSONPath.Text.Trim(),
            .JSON_PathBackup = txtJSONPathBackup.Text.Trim(),
            .XMLESSBudget_Path = txtESSBudgetPath.Text.Trim(),
            .XMLESSBudget_PathBackup = txtESSBudgetPathBackup.Text.Trim(),
            .XMLSAP_Path = txtSAPFilePath.Text.Trim(),
            .XMLSAP_PathBackup = txtSAPFilePathBackup.Text.Trim(),
            .XMLIAPrice_Path = txtIAPriceDestinationPath.Text.Trim(),
            .IAMIPressPartLocalComponent_Path = txtIAMIPressPartLocalComponentPath.Text.Trim(),
            .IAMIPressPartLocalComponent_PathBackup = txtIAMIPressPartLocalComponentPathBackup.Text.Trim(),
            .IAMIPressPartLocalComponentXMLToSAP_Path = txtIAMIPressPartLocalComponentXMLToSAPPath.Text.Trim(),
            .CostHistoryXMLToSAP_Path = txtCostHistoryXMLToSAPPath.Text.Trim(),
            .PurchaseRequest_Path = txtPurchaseRequestPath.Text.Trim(),
            .PurchaseRequest_PathBackup = txtPurchaseRequestPathBackup.Text.Trim(),
            .PurchaseRequest_PathFailed = txtPurchaseRequestPathFailed.Text.Trim(),
            .PurchaseRequest_PathOut = txtPurchaseRequestPathOut.Text.Trim(),
            .GoodsReceiveXMLToDMMS_Path = txtGoodsReceiveXMLToDMMSPath.Text.Trim(),
            .MaterialMasterXMLtoSAP_Path = txtMaterialMasterXMLtoSAPPath_Req.Text.Trim(),
            .MaterialMasterXMLtoSAP_BackupPath = txtMaterialMasterXMLtoSAPBackupPath_Req.Text.Trim(),
            .MaterialMasterXML_Resp_Pth = txtMaterialMasterXMLtoSAPPath_Res.Text.Trim(),
            .MaterialMasterXML_Resp_BackupPth = txtMaterialMasterXMLtoSAPBackupPath_Res.Text.Trim(),
            .IR_XML_ReqPath = txtIR_XML_RequestPath.Text.Trim(),
            .IR_XML_ReqPath_Backup = txtIR_XML_RequestBackupPath.Text.Trim(),
            .IR_XML_RespPath = txtIR_XML_ResponsePath.Text.Trim(),
            .IR_XML_RespPath_Backup = txtIR_XML_ResponseBackupPath.Text.Trim(),
            .PAQuotReqPath = txtQuotationRequestPath.Text.Trim(),
            .PAQuotResPath = txtQuotationResponsePath.Text.Trim(),
            .PAQuotResBackupPath = txtQuotationResponseBackupPath.Text.Trim()
        }

    End Function

End Class
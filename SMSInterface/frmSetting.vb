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

        ActiveControl = txtSMTPAddress_DMMS_PR
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

    Public Shared Function UpdateInterface_Setting(ByVal pConStr As String, ByVal SMTP_User As String, ByVal SMTP_Password As String, ByVal SMTP_Port As String, ByVal SMTP_Address As String, ByVal Email As String, _
                                                   ByVal JSON_Path As String, ByVal JSON_PathBackup As String, ByVal XMLESSBudget_Path As String, ByVal XMLESSBudget_PathBackup As String, _
                                                   ByVal XMLSAP_Path As String, ByVal XMLSAP_PathBackup As String, ByVal XMLIAPrice_Path As String, ByVal IAMIPressPartLocalComponent_Path As String, _
                                                   ByVal IAMIPressPartLocalComponent_PathBackup As String, ByVal IAMIPressPartLocalComponentXMLToSAP_Path As String, ByVal CostHistoryXMLToSAP_Path As String, _
                                                   ByVal PurchaseRequest_Path As String, ByVal PurchaseRequest_PathBackup As String, ByVal PurchaseRequest_PathFailed As String, _
                                                   ByVal PurchaseRequest_PathOut As String, ByVal GoodsReceiveXMLToDMMS_Path As String,
                                                   ByVal MaterialMasterXMLtoSAP_Path As String, ByVal MaterialMasterXMLtoSAP_BackupPath As String, ByVal MaterialMasterXML_Resp_Pth As String,
                                                   ByVal MaterialMasterXML_Resp_BackupPth As String, ByVal IR_XML_ReqPath As String, ByVal IR_XML_ReqPath_Backup As String,
                                                   ByVal IR_XML_RespPath As String, ByVal IR_XML_RespPath_Backup As String
                                                    ) As Integer

        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()
                sql = " update   SendEmail_Setting set JSON_Path = '" & JSON_Path & "', JSON_PathBackup = '" & JSON_PathBackup & "', XML_IABudget_Path = '" & XMLESSBudget_Path & "', XML_IABudget_PathBackup = '" & XMLESSBudget_PathBackup & "', " & vbCrLf & _
                    "           XML_SAP_Path = '" & XMLSAP_Path & "', XML_SAP_PathBackup = '" & XMLSAP_PathBackup & "', XML_IAPrice_Path = '" & XMLIAPrice_Path & "', IAMI_PressPart_Path = '" & IAMIPressPartLocalComponent_Path & "', " & vbCrLf & _
                    "           IAMI_PressPart_PathBackup = '" & IAMIPressPartLocalComponent_PathBackup & "', IAMI_PressPart_XMLToSAP_Path = '" & IAMIPressPartLocalComponentXMLToSAP_Path & "', CostHistory_XMLToSAP_Path = '" & CostHistoryXMLToSAP_Path & "', " & vbCrLf & _
                    "           XML_PurchaseRequest_Path = '" & PurchaseRequest_Path & "', XML_PurchaseRequest_PathBackup = '" & PurchaseRequest_PathBackup & "', XML_PurchaseRequest_PathFailed = '" & PurchaseRequest_PathFailed & "', " & vbCrLf & _
                    "           XML_PurchaseRequest_PathOut = '" & PurchaseRequest_PathOut & "', GoodsReceive_XMLToDMMS_Path = '" & GoodsReceiveXMLToDMMS_Path & "', " & vbCrLf & _
                    "           MaterialMasterXML_ToSAP_Path = '" & MaterialMasterXMLtoSAP_Path & "', MaterialMasterXML_ToSAP_PathBackup = '" & MaterialMasterXMLtoSAP_BackupPath & "'," & vbCrLf & _
                    "           MaterialMasterXML_Resp_Pth = '" & MaterialMasterXML_Resp_Pth & "', MaterialMasterXML_Resp_BackupPth = '" & MaterialMasterXML_Resp_BackupPth & "', " & vbCrLf & _
                    "           IR_XML_Req_Path = '" & IR_XML_ReqPath & "', IR_XML_Req_BackupPath = '" & IR_XML_ReqPath_Backup & "', " & vbCrLf & _
                    "           IR_XML_Res_Path = '" & IR_XML_RespPath & "', IR_XML_Res_BackupPath = '" & IR_XML_RespPath_Backup & "'"

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = CInt(cmd.ExecuteNonQuery())
                If retValue = 0 Then
                    sql = " insert into SendEmail_Setting " & vbCrLf & _
                            " ( " & vbCrLf & _
                            "   SMTP_User, SMTP_Password, SMTP_Port, SMTP_Address, Email, " & vbCrLf & _
                            "   JSON_Path, JSON_PathBackup, XML_IABudget_Path, XML_IABudget_PathBackup, " & vbCrLf & _
                            "   XML_SAP_Path, XML_SAP_PathBackup, XML_IAPrice_Path, " & vbCrLf & _
                            "   IAMI_PressPart_Path, IAMI_PressPart_PathBackup, IAMI_PressPart_XMLToSAP_Path, CostHistory_XMLToSAP_Path, " & vbCrLf & _
                            "   XML_PurchaseRequest_Path, XML_PurchaseRequest_PathBackup, XML_PurchaseRequest_PathFailed, " & vbCrLf & _
                            "   XML_PurchaseRequest_PathOut, GoodsReceive_XMLToDMMS_Path, " & vbCrLf & _
                            "   MaterialMasterXML_ToSAP_Path, MaterialMasterXML_ToSAP_PathBackup, MaterialMasterXML_Resp_Pth, MaterialMasterXML_Resp_BackupPth, " & vbCrLf &
                            "   IR_XML_Req_Path, IR_XML_Req_BackupPath, IR_XML_Res_Path, IR_XML_Res_BackupPath" & vbCrLf &
                            " ) " & vbCrLf & _
                            " values " & vbCrLf & _
                            " ( " & vbCrLf & _
                            "   '" & SMTP_User & "', '" & SMTP_Password & "', '" & SMTP_Port & "', '" & SMTP_Address & "', '" & Email & "', " & vbCrLf & _
                            "   '" & JSON_Path & "', '" & JSON_PathBackup & "', '" & XMLESSBudget_Path & "', '" & XMLESSBudget_PathBackup & "', " & vbCrLf & _
                            "   '" & XMLSAP_Path & "', '" & XMLSAP_PathBackup & "', '" & XMLIAPrice_Path & "', " & vbCrLf & _
                            "   '" & IAMIPressPartLocalComponent_Path & "', '" & IAMIPressPartLocalComponent_PathBackup & "', '" & IAMIPressPartLocalComponentXMLToSAP_Path & "', '" & CostHistoryXMLToSAP_Path & "', " & vbCrLf & _
                            "   '" & PurchaseRequest_Path & "', '" & PurchaseRequest_PathBackup & "', '" & PurchaseRequest_PathFailed & "', " & vbCrLf & _
                            "   '" & PurchaseRequest_PathOut & "', '" & GoodsReceiveXMLToDMMS_Path & "', " & vbCrLf & _
                            "   '" & MaterialMasterXMLtoSAP_Path & "','" & MaterialMasterXMLtoSAP_BackupPath & "','" & MaterialMasterXML_Resp_Pth & "','" & MaterialMasterXML_Resp_BackupPth & "', " & vbCrLf & _
                            "   '" & IR_XML_ReqPath & "','" & IR_XML_ReqPath_Backup & "','" & IR_XML_RespPath & "','" & IR_XML_RespPath_Backup & "' " & vbCrLf & _
                            " ) "
                End If

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try
    End Function

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        'validate SMTP DMMS - PR
        '################################################################
        If txtSMTPAddress_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input SMTP Address for DMMS PR!")
            txtSMTPAddress_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtSMTPPort_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input SMTP Port for DMMS PR!")
            txtSMTPPort_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtSMTPUser_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input User for DMMS PR!")
            txtSMTPUser_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtSMTPPassword_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input Password for DMMS PR!")
            txtSMTPPassword_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtEmailFrom_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input Email From for DMMS PR!")
            txtEmailFrom_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtEmailTo_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input Email To for DMMS PR!")
            txtEmailTo_DMMS_PR.Focus()
            Exit Sub
        End If

        If txtSubject_DMMS_PR.Text.Trim = "" Then
            MsgBox("Please input Email Subject for DMMS PR!")
            txtSubject_DMMS_PR.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder EPIC 1 Supplier Management
        '################################################################
        If txtJSONPath.Text.Trim = "" Then
            MsgBox("Please input EPIC 1 Supplier Management's Path Folder!")
            txtJSONPath.Focus()
            Exit Sub
        End If

        If txtJSONPathBackup.Text.Trim = "" Then
            MsgBox("Please input EPIC 1 Supplier Management's Backup Path Folder!")
            txtJSONPathBackup.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder ESS Budget
        '################################################################
        If txtESSBudgetPath.Text.Trim = "" Then
            MsgBox("Please input ESS Budget's Path Folder!")
            txtESSBudgetPath.Focus()
            Exit Sub
        End If

        If txtESSBudgetPathBackup.Text.Trim = "" Then
            MsgBox("Please input ESS Budget's Backup Path Folder!")
            txtESSBudgetPathBackup.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder SAP PO GR
        '################################################################
        If txtSAPFilePath.Text.Trim = "" Then
            MsgBox("Please input SAP PO GR's Path Folder!")
            txtSAPFilePath.Focus()
            Exit Sub
        End If

        If txtSAPFilePathBackup.Text.Trim = "" Then
            MsgBox("Please input SAP PO GR's Backup Path Folder!")
            txtSAPFilePathBackup.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder IA Price
        '################################################################
        If txtIAPriceDestinationPath.Text.Trim = "" Then
            MsgBox("Please input IA Price Destination's Path Folder!")
            txtIAPriceDestinationPath.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder SAP - Cost Monitoring
        '################################################################
        If txtIAMIPressPartLocalComponentXMLToSAPPath.Text.Trim = "" Then
            MsgBox("Please input IAMI Press Part & Local Component XML To SAP's Path Folder!")
            txtIAMIPressPartLocalComponentXMLToSAPPath.Focus()
            Exit Sub
        End If

        If txtIAMIPressPartLocalComponentPath.Text.Trim = "" Then
            MsgBox("Please input IAMI Press Part & Local Component's Path Folder!")
            txtIAMIPressPartLocalComponentPath.Focus()
            Exit Sub
        End If

        If txtIAMIPressPartLocalComponentPathBackup.Text.Trim = "" Then
            MsgBox("Please input IAMI Press Part & Local Component's Backup Path Folder!")
            txtIAMIPressPartLocalComponentPathBackup.Focus()
            Exit Sub
        End If

        If txtCostHistoryXMLToSAPPath.Text.Trim = "" Then
            MsgBox("Please input Cost History XML To SAP's Path Folder!")
            txtCostHistoryXMLToSAPPath.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder DMMS - PR & GR
        '################################################################
        If txtPurchaseRequestPath.Text.Trim = "" Then
            MsgBox("Please input Purchase Request's Path Folder!")
            txtPurchaseRequestPath.Focus()
            Exit Sub
        End If

        If txtPurchaseRequestPathBackup.Text.Trim = "" Then
            MsgBox("Please input Purchase Request's Backup Path Folder!")
            txtPurchaseRequestPathBackup.Focus()
            Exit Sub
        End If

        If txtPurchaseRequestPathFailed.Text.Trim = "" Then
            MsgBox("Please input Purchase Request's Failed Path Folder!")
            txtPurchaseRequestPathFailed.Focus()
            Exit Sub
        End If

        If txtPurchaseRequestPathOut.Text.Trim = "" Then
            MsgBox("Please input Purchase Request Response's Path Folder!")
            txtPurchaseRequestPathOut.Focus()
            Exit Sub
        End If

        If txtGoodsReceiveXMLToDMMSPath.Text.Trim = "" Then
            MsgBox("Please input Goods Receive XML To DMMS's Path Folder!")
            txtGoodsReceiveXMLToDMMSPath.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder Material Master
        '################################################################
        If txtMaterialMasterXMLtoSAPPath_Req.Text.Trim = "" Then
            MsgBox("Please input Material Master's Request Path Folder!")
            txtMaterialMasterXMLtoSAPPath_Req.Focus()
            Exit Sub
        End If

        If txtMaterialMasterXMLtoSAPBackupPath_Req.Text.Trim = "" Then
            MsgBox("Please input Material Master's Request Backup Path Folder!")
            txtMaterialMasterXMLtoSAPBackupPath_Req.Focus()
            Exit Sub
        End If

        If txtMaterialMasterXMLtoSAPBackupPath_Res.Text.Trim = "" Then
            MsgBox("Please input Material Master's Response Path Folder!")
            txtMaterialMasterXMLtoSAPBackupPath_Res.Focus()
            Exit Sub
        End If

        If txtMaterialMasterXMLtoSAPBackupPath_Res.Text.Trim = "" Then
            MsgBox("Please input Materal Master's Response Backup Path Folder!")
            txtMaterialMasterXMLtoSAPBackupPath_Res.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder Material Master
        '################################################################
        If txtIR_XML_RequestPath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Request Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If

        If txtIR_XML_RequestBackupPath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Request Backup Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If

        If txtIR_XML_ResponsePath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Response Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If

        If txtIR_XML_ResponseBackupPath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Response Backup Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If
        '################################################################

        'validate path folder Quotation Master
        '################################################################
        If txtQuotationRequestPath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Request Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If

        If txtQuotationResponsePath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Response Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If

        If txtQuotationResponseBackupPath.Text.Trim = "" Then
            MsgBox("Please input Info Record's Response Backup Path Folder!")
            txtIR_XML_RequestPath.Focus()
            Exit Sub
        End If
        '################################################################

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
        If UpdateInterface_Setting(ConStr, txtSMTPUser_DMMS_PR.Text.Trim, txtSMTPPassword_DMMS_PR.Text.Trim, txtSMTPPort_DMMS_PR.Text.Trim, txtSMTPAddress_DMMS_PR.Text.Trim, txtSMTPUser_DMMS_PR.Text.Trim, _
                                    txtJSONPath.Text.Trim, txtJSONPathBackup.Text.Trim, txtESSBudgetPath.Text.Trim, txtESSBudgetPathBackup.Text.Trim, _
                                    txtSAPFilePath.Text.Trim, txtSAPFilePathBackup.Text.Trim, txtIAPriceDestinationPath.Text.Trim, _
                                    txtIAMIPressPartLocalComponentPath.Text.Trim, txtIAMIPressPartLocalComponentPathBackup.Text.Trim, txtIAMIPressPartLocalComponentXMLToSAPPath.Text.Trim, _
                                    txtCostHistoryXMLToSAPPath.Text.Trim, txtPurchaseRequestPath.Text.Trim, txtPurchaseRequestPathBackup.Text.Trim, txtPurchaseRequestPathFailed.Text.Trim, _
                                    txtPurchaseRequestPathOut.Text.Trim, txtGoodsReceiveXMLToDMMSPath.Text.Trim,
                                    txtMaterialMasterXMLtoSAPPath_Req.Text.Trim, txtMaterialMasterXMLtoSAPBackupPath_Req.Text.Trim,
                                    txtMaterialMasterXMLtoSAPPath_Res.Text.Trim, txtMaterialMasterXMLtoSAPBackupPath_Res.Text.Trim,
                                    txtIR_XML_RequestPath.Text.Trim, txtIR_XML_RequestBackupPath.Text.Trim,
                                    txtIR_XML_ResponsePath.Text.Trim, txtIR_XML_ResponseBackupPath.Text.Trim
                                    ) > 0 Then

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
            End If

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

        Else
            errorCount = errorCount + 1
        End If

        If errorCount = 0 Then
            MsgBox("Update data successfully!", MsgBoxStyle.Information, "Info")
        Else
            MsgBox("Update data is not successfully!", MsgBoxStyle.Critical, "Error")
        End If

    End Sub

End Class
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient

Public Class ClsInterfaceSettingDB
    Public Shared Function getDataSMTP_DMMS_PR(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = " select  isnull(SMTP_Address, '') as SMTP_Address, isnull(SMTP_Port, '') as SMTP_Port, " & vbCrLf & _
                "           isnull(SMTP_User, '') as SMTP_User, isnull(SMTP_Password, '') as SMTP_Password, " & vbCrLf & _
                "           isnull(FromEmail, '') as FromEmail, isnull(ReceiptEmail, '') as ReceiptEmail, " & vbCrLf & _
                "           isnull(Subject, '') as Subject " & vbCrLf & _
                "   from    SMTP_Settings " & vbCrLf & _
                "   where   SystemFrom = 'DMMS' and " & vbCrLf & _
                "           Module = 'PR' "

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getData(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = " select  isnull(JSON_Path, '') as JSON_Path, isnull(JSON_PathBackup, '') as JSON_PathBackup, " & vbCrLf & _
                "           isnull(XML_IABudget_Path, '') as XML_IABudget_Path, isnull(XML_IABudget_PathBackup, '') as XML_IABudget_PathBackup, " & vbCrLf & _
                "           isnull(XML_SAP_Path, '') as XML_SAP_Path, isnull(XML_SAP_PathBackup, '') as XML_SAP_PathBackup, " & vbCrLf & _
                "           isnull(XML_IAPrice_Path, '') as XML_IAPrice_Path, isnull(IAMI_PressPart_Path, '') as IAMI_PressPart_Path, isnull(IAMI_PressPart_PathBackup, '') as IAMI_PressPart_PathBackup, " & vbCrLf & _
                "           isnull(IAMI_PressPart_XMLToSAP_Path, '') as IAMI_PressPart_XMLToSAP_Path, " & vbCrLf & _
                "           isnull(CostHistory_XMLToSAP_Path, '') as CostHistory_XMLToSAP_Path, " & vbCrLf & _
                "           isnull(XML_PurchaseRequest_Path, '') as XML_PurchaseRequest_Path, " & vbCrLf & _
                "           isnull(XML_PurchaseRequest_PathBackup, '') as XML_PurchaseRequest_PathBackup, " & vbCrLf & _
                "           isnull(XML_PurchaseRequest_PathFailed, '') as XML_PurchaseRequest_PathFailed, " & vbCrLf & _
                "           isnull(XML_PurchaseRequest_PathOut, '') as XML_PurchaseRequest_PathOut, " & vbCrLf & _
                "           isnull(GoodsReceive_XMLToDMMS_Path, '') as GoodsReceive_XMLToDMMS_Path, " & vbCrLf & _
                "           IsNull(MaterialMasterXML_ToSAP_Path,'') as MaterialMasterXML_ToSAP_Path, " & vbCrLf & _
                "           IsNull(MaterialMasterXML_ToSAP_PathBackup,'') as MaterialMasterXML_ToSAP_PathBackup, " & vbCrLf & _
                "           IsNull(MaterialMasterXML_Resp_Pth,'') as MaterialMasterXML_Resp_Pth, " & vbCrLf & _
                "           IsNull(MaterialMasterXML_Resp_BackupPth,'') as MaterialMasterXML_Resp_BackupPth, " & vbCrLf & _
                "           IsNull(IR_XML_Req_Path,'') as IR_XML_Req_Path, " & vbCrLf & _
                "           IsNull(IR_XML_Req_BackupPath,'') as IR_XML_Req_BackupPath, " & vbCrLf & _
                "           IsNull(IR_XML_Res_Path,'') as IR_XML_Res_Path, " & vbCrLf & _
                "           IsNull(IR_XML_Res_BackupPath,'') as IR_XML_Res_BackupPath " & vbCrLf & _
                "           IsNull(PAQuotReqPath,'') as PAQuotReqPath, " & vbCrLf & _
                "           IsNull(PAQuotResPath,'') as PAQuotResPath " & vbCrLf & _
                "           IsNull(PAQuotResBackupPath,'') as PAQuotResBackupPath, " & vbCrLf & _
                "   from    SendEmail_Setting "

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataApprovalSchedule(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_DataApproval"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataIAPrice(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_GET_IA_PRICE_FOR_INTERFACE"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataApprovalInterval(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_DataApproval_Interval"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getNextProcess(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = " SELECT TOP 1 Convert(char(8),[Time],108) NextProcess,* " & vbCrLf & _
                    " 	FROM sendEmail_Schedule " & vbCrLf & _
                    " 	ORDER BY [Time]  "

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function


End Class

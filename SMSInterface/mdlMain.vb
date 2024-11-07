Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Drawing

Module mdlMain
    Public gs_Server As String = "serverdb\sqlr2"
    Public gs_Database As String = "PIS3"
    Public gs_User As String = "sa"
    Public gs_Password As String = "y2k"
    Public gs_AppID As String = "P01"
    Public gs_UserLogin As String = ""
    Public gs_RegisterProgram As String = "SMS"

    Public gs_SAPServer As String = "serverdb\sqlr2"
    Public gs_SAPDatabase As String = "PIS3"
    Public gs_SAPDI_User As String = "sa"
    Public gs_SAPDI_Pass As String = "y2k"
    Public gs_SAPPass As String = "P01"
    Public gs_SAPUser As String = ""

    Public gs_SMTPAddress_DMMS_PR As String = "mail.fast.net.id"
    Public gs_SMTPPort_DMMS_PR As String = "25"
    Public gs_SMTPUser_DMMS_PR As String = "tos.yudi@gmail.com"
    Public gs_SMTPPassword_DMMS_PR As String = "bukanpks"
    Public gs_EmailFrom_DMMS_PR As String = ""
    Public gs_EmailTo_DMMS_PR As String = ""
    Public gs_Subject_DMMS_PR As String = ""

    Public gv_SoNo As String = ""

    Public gs_JSONPath As String = ""
    Public gs_JSONPathBackup As String = ""
    Public gs_XMLPath As String = ""
    Public gs_XMLPathBackup As String = ""

    Public gs_SAPPath As String = ""
    Public gs_SAPPathBackup As String = ""
    Public gs_IAPricePath As String = ""

    Public gs_IAMI_PressPartPath As String = ""
    Public gs_IAMI_PressPartPathBackup As String = ""
    Public gs_IAMI_PressPartXMLToSAPPath As String = ""
    Public gs_CostHistoryXMLToSAPPath As String = ""

    'DMMS
    Public gs_PurchaseRequestPath As String = ""
    Public gs_PurchaseRequestPathBackup As String = ""
    Public gs_PurchaseRequestPathFailed As String = ""
    Public gs_PurchaseRequestPathOut As String = ""
    Public gs_GoodsReceiveXMLToDMMSPath As String = ""

    'Material Master Path
    Public gs_MaterialMasterXMLtoSAP_Path_Req As String = ""
    Public gs_MaterialMasterXMLtoSAP_BackupPath_Req As String = ""
    Public gs_MaterialMasterXMLtoSAP_Path_Res As String = ""
    Public gs_MaterialMasterXMLtoSAP_BackupPath_Res As String = ""

    'InfoRecord
    Public gs_IR_XMLtoSAP_Path_Req As String = ""
    Public gs_IR_XMLtoSAP_BackupPath_Req As String = ""
    Public gs_IR_XMLtoSAP_Path_Res As String = ""
    Public gs_IR_XMLtoSAP_BackupPath_Res As String = ""

    'MASPIS
    Public gs_PAQuotReqPath As String = ""
    Public gs_PAQuotResPath As String = ""
    Public gs_PAQuotResBackupPath As String = ""


    Public Db As SqlConnection
    Public Builder As SqlConnectionStringBuilder
    Public nRetry As Integer
    Public gi_RetryInterval As Integer
    Public gi_MaxRetry As Integer
    Public ErrMsg As String = ""
    Public ErrNo As String = ""
    Public Assem As Reflection.Assembly = Assembly.GetCallingAssembly

    Public gs_FormatAmount As String = "###,###.#0"
    Public gs_FormatAmount2 As String = "#,##0.00"
    Public gs_FormatQty As String = "###,###"
    Public gs_FormatDate As String = "dd.MM.yyyy"
    Public gs_FormatWeight As String = "###,###"
    Public gs_FormatPercentage As String = "#,###.#0"

    Public gs_Cls_ContractType As Integer = 1
    Public gs_Cls_OrderClassI As Integer = 2
    Public gs_Cls_OrderClassII As Integer = 3
    Public gs_Cls_EndUserSegmentI As Integer = 4
    Public gs_Cls_EndUSerSegmentII As Integer = 5
    Public gs_Cls_DeliveryTerm As Integer = 6
    Public gs_Cls_PaymentType As Integer = 7
    Public gs_Cls_PaymentTerm As Integer = 8
    Public gs_Cls_DPPaymentTerm As Integer = 9
    Public gs_Cls_Unit As Integer = 10
    Public gs_Cls_Currency As Integer = 11
    Public gs_Cls_TaxGroup As Integer = 12
    Public gs_Cls_ProductType As Integer = 13
    Public gs_Cls_Standard As Integer = 14
    Public gs_Cls_Grade As Integer = 15
    Public gs_Cls_ProductSize As Integer = 16
    Public gs_Cls_Length As Integer = 17
    Public gs_Cls_LengthByType As Integer = 18
    Public gs_Cls_LengthBySize As Integer = 19
    Public gs_Cls_UsageClassI As Integer = 20
    Public gs_Cls_UsageClassII As Integer = 21
    Public gs_Cls_CountryCode As Integer = 26
    Public gs_Cls_DeliveryArea As Integer = 27
    Public gs_Cls_AreaCode As Integer = 28
    Public gs_Cls_PortCode As Integer = 29
    Public gs_Cls_OrderCategory As Integer = 1
    Public gs_Cls_SpecialLength As Integer = 0
    Public gs_Cls_ExtraPriceCode As Integer = 22

    Public gi_Cls_Contract_Project As Integer = 0
    Public gi_Cls_Contract_General As Integer = 0
    Public gi_Cls_Contract_Scrap As Integer = 0

    Public Function up_LoadDynamicForm(ByVal FormName As String) As String
        Dim ls_ErrMsg As String = ""
        If ls_ErrMsg.Trim <> "" Then Return ls_ErrMsg
        Dim Frm As Form = Nothing
        Try
            For Each Frm In My.Application.OpenForms
                If Frm.Name.ToLower = FormName.ToLower Then
                    Frm.Dispose()
                    Exit For
                End If
            Next

            Frm = Assem.CreateInstance(Assem.GetName.Name & "." & Trim(FormName), True)
            If Frm Is Nothing Then
                Throw New Exception("Invalid menu ID")
            Else
                Frm.WindowState = FormWindowState.Maximized
                Frm.StartPosition = FormStartPosition.CenterScreen
                Frm.Show()
            End If
        Catch ex As Exception
            Try
                Frm.Dispose()
            Catch ex1 As Exception
                Return "Error loading Form !"
            End Try
            Return "Error loading Form !"
        End Try
        Return ""
    End Function

    
End Module

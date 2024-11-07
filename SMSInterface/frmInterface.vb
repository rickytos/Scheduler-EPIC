Imports System.Threading
Imports System.Data.SqlClient
Imports System.Security.Cryptography.X509Certificates
Imports System.Net
Imports System.Net.Security
Imports System.Net.Mail
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Xml
Imports System.Text

'NEW Scheduler 
Public Class frmInterface

#Region "Initial"
    Protected Gb As clsGlobal
    Protected ConStr As String

    Dim cfg As clsConfig
    Dim jsonfilename As String = ""
    Dim jsonfilenameLocation As String = ""
    Dim xmlfilename As String = ""
    Dim xmlfilenameLocation As String = ""
    Dim ls_StepProcess As String
    Delegate Sub SetTextCallback(ByVal [text] As String, ByVal textControl As Windows.Forms.TextBox)

    Dim li_Counter As Integer = 0
#End Region

#Region "Events"
    Private Sub frmInterface_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        txtMsg.Text = ""

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
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        txtMsg.Text = ""
        If (MsgBox("Do you want to close?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirmation") = MsgBoxResult.Yes) Then
            Me.Close()
        End If
    End Sub

    Private Sub btnProcess_Click(sender As System.Object, e As System.EventArgs) Handles btnProcess.Click
        txtMsg.Text = ""

        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Manual Process ..." & vbCrLf
        up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Try

            up_Interface_EPIC1_SupplierManagement_Process()
        Catch ex As IOException When ex.Message.Contains("The network path was not found")
            ' Handle the specific IOException with network path not found
            ls_StepProcess = "[" & ex.Message.ToString() & "] The network path was not found. Interface EPIC1 - Supplier Management!" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch exIO As IOException
            ls_StepProcess = "[" & exIO.Message.ToString() & "] Error Interface EPIC1 - Supplier Management!" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch ex As Exception
            ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface EPIC1 - Supplier Management!" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

        End Try

        Try
            up_Interface_ESS_Budget_Process()
        Catch ex As IOException When ex.Message.Contains("The network path was not found")
            ' Handle the specific IOException with network path not found
            ls_StepProcess = "[" & ex.Message.ToString() & "] The network path was not found. ESS Budget" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch exIO As IOException
            ls_StepProcess = "[" & exIO.Message.ToString() & "] Error Interface ESS Budget" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch ex As Exception
            ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface ESS Budget" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

        End Try

        Try
            up_Interface_SAP_Process()
        Catch ex As IOException When ex.Message.Contains("The network path was not found")
            ls_StepProcess = "[" & ex.Message.ToString() & "] The network path was not found. SAP" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch exIO As IOException
            ls_StepProcess = "[" & exIO.Message.ToString() & "] Error Interface SAP" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        Catch ex As Exception
            ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface SAP" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

        End Try
        Try
            up_Interface_IAMI_PressPart_LocalComponent_Process()
        Catch ex As Exception
            ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface IAMI Press Part" & vbCrLf
            up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        End Try

        generateXMLIAMI_PressPartLocalComponent()
        generateXMLCMHistory()

        'DMMS
        generateXMLGRToDMMS()
        up_Interface_PurchaseRequest_Process()

        'Material Master and Info Record
        MaterialMasterServices.generateXMLMaterialMaster(ConStr, Rtb1)
        InfoRecordServices.generateXMLInfoRecord(ConStr, Rtb1)

        'Interace MasPIS
        MASPISServices.GenQuotationMasPis(ConStr, Rtb1)
        MASPISServices.MASPISProcess(ConStr, Rtb1)

        'Interface CostHistory

        'Try
        '    MaterialMasterServices.InterfaceMaterialMasterProcess(ConStr, Rtb1)
        'Catch ex As IOException When ex.Message.Contains("The network path was not found")
        '    ls_StepProcess = "[" & ex.Message.ToString() & "] END The network path was not found. Material Master" & vbCrLf
        '    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        'Catch ex As Exception
        '    ls_StepProcess = "[" & ex.Message.ToString() & "] END Error Interface Material Master Response" & vbCrLf
        '    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        'End Try

        'Try
        '    InfoRecordServices.InterfaceRecordProcess(ConStr, Rtb1)
        'Catch ex As IOException When ex.Message.Contains("The network path was not found")
        '    ls_StepProcess = "[" & ex.Message.ToString() & "] END The network path was not found. Info Record" & vbCrLf
        '    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        'Catch ex As Exception
        '    ls_StepProcess = "[" & ex.Message.ToString() & "] END Error Interface Info Record Response" & vbCrLf
        '    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
        'End Try

        ls_StepProcess = "" & vbCrLf
        ls_StepProcess = ls_StepProcess & "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End Manual Process ..." & vbCrLf
        up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
    End Sub

    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        txtMsg.Text = ""
        up_StartScheduler()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        dtpNow.Value = Format(Now(), "dd MMM yyyy HH:mm:ss")

        If dtpNext.Value = dtpNow.Value Then
            Timer1.Enabled = False

            Dim dt As DataTable

            dt = ClsInterfaceSettingDB.getNextProcess(ConStr)

            If dt.Rows.Count > 0 Then
                dtpNext.Value = dt.Rows(0).Item("Time")
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Refresh Next Process Time ... " & vbCrLf
                up_gridProcess(Rtb1, 3, 0, ls_StepProcess)
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please set Next Process Time " & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Timer1.Enabled = True
                Exit Sub
            End If

            Timer1.Enabled = True
        End If

        If li_Counter >= txtInterval.Text Then
            Timer1.Enabled = False
            Dim timerealtime As DateTime = Format(Now(), "dd MMM yyyy HH:mm:ss")
            Dim timesetting As DateTime = Format(Now(), "yyyy-MM-dd") & " 23:50:00.000"

            Dim timesettingStart As DateTime = Format(Now(), "yyyy-MM-dd") & " 09:00:00.000"
            Dim timesettingFinish As DateTime = Format(Now(), "yyyy-MM-dd") & " 21:00:00.000"
            Dim timesettingInterval As DateTime = Format(dtpNext.Value, "dd MMM yyyy HH:mm:ss")

            If timerealtime >= timesettingInterval And timesettingInterval <= timesettingFinish Then
                Try
                    up_Interface_EPIC1_SupplierManagement_Process()
                Catch ex As IOException When ex.Message.Contains("The network path was not found")
                    ' Handle the specific IOException with network path not found
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Interface EPIC1 - Supplier" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch exIO As IOException
                    ls_StepProcess = "[" & exIO.Message.ToString() & "] Interface EPIC1 - Supplier" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Interface EPIC1 - Supplier" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

                End Try

                Try
                    up_Interface_ESS_Budget_Process()
                Catch ex As IOException When ex.Message.Contains("The network path was not found")
                    ' Handle the specific IOException with network path not found
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface ESS Budget" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch exIO As IOException
                    ls_StepProcess = "[" & exIO.Message.ToString() & "] Error Interface ESS Budget" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface ESS Budget" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

                End Try

                Try
                    up_Interface_SAP_Process()
                Catch ex As IOException When ex.Message.Contains("The network path was not found")
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface SAP" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch exIO As IOException
                    ls_StepProcess = "[" & exIO.Message.ToString() & "] Error Interface SAP" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface SAP" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)




                End Try

                Try
                    up_Interface_IAMI_PressPart_LocalComponent_Process()
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface IAMI Press Part" & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                End Try

                generateXMLIAMI_PressPartLocalComponent()
                generateXMLCMHistory()

                'DMMS
                generateXMLGRToDMMS()
                up_Interface_PurchaseRequest_Process()

                'Material  & Info Record
                'generateXMLMaterialMaster()
                MaterialMasterServices.generateXMLMaterialMaster(ConStr, Rtb1)
                'generateXMLInfoRecord()
                InfoRecordServices.generateXMLInfoRecord(ConStr, Rtb1)

                Try
                    'up_Interface_MaterialMaster_Process()
                    MaterialMasterServices.InterfaceMaterialMasterProcess(ConStr, Rtb1)
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface Material Master Response " & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                End Try

                Try
                    'up_Interface_InfoRecord_Process()
                    InfoRecordServices.InterfaceRecordProcess(ConStr, Rtb1)
                Catch ex As Exception
                    ls_StepProcess = "[" & ex.Message.ToString() & "] Error Interface InfoRecord Response " & vbCrLf
                    up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                End Try

                'Interface MasPis
                MASPISServices.GenQuotationMasPis(ConStr, Rtb1)
                MASPISServices.MASPISProcess(ConStr, Rtb1)

            End If

            li_Counter = 0

            Timer1.Enabled = True
        End If
        li_Counter = li_Counter + 1

    End Sub
#End Region

#Region "Method"
    Public Shared Function UpdateTrans(pConStr As String, TaskGroup As String, TransNumber As String, TransRev As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_DataApprovalUpdate '" & TaskGroup & "', '" & TransNumber & "', '" & TransRev & "' "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function EPIC1SupplierTransaction(pConStr As String, Supplier_Code As String, Supplier_Name As String, Address As String, Phone As String, Fax As String, Country As String, Email As String, Period_Code As String, SupplierType As String, President_Director As String, President_Director_Email As String, Vice_President_Director As String, Vice_President_Director_Email As String, Marketing_Director As String, Marketing_Director_Email As String, Marketing_General_Manager As String, Marketing_General_Manager_Email As String, Marketing_Manager As String, Marketing_Manager_Email As String, Plant_Director As String, Plant_Director_Email As String, Plant_Manager As String, Plant_Manager_Email As String, Marketing_Sales As String, Marketing_Sales_Email As String, Certification As String, QualityAudit_Score As String, DeliveryAudit_Score As String, EHS_Value As String, PPM As String, Mother_Company As String, Product_Name As String, OEM_Customer As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_interface_mst_Supplier '" & Supplier_Code & "', '" & Supplier_Name & "', '" & Address & "', '" & Phone & "', '" & Fax & "', '" & Country & "', '" & President_Director & "', '" & President_Director_Email & "', '" & Vice_President_Director & "', '" & Vice_President_Director_Email & "', '" & Marketing_Director & "', '" & Marketing_Director_Email & "', '" & Marketing_General_Manager & "', '" & Marketing_General_Manager_Email & "', '" & Marketing_Manager & "', '" & Marketing_Manager_Email & "', '" & Plant_Director & "', '" & Plant_Director_Email & "', '" & Plant_Manager & "', '" & Plant_Manager_Email & "', '" & Marketing_Sales & "', '" & Marketing_Sales_Email & "', '" & Email & "', '" & Period_Code & "', '" & SupplierType & "', '" & Certification & "', '" & QualityAudit_Score & "', '" & DeliveryAudit_Score & "', '" & EHS_Value & "', '" & PPM & "', '" & Mother_Company & "', '" & Product_Name & "', '" & OEM_Customer & "' "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function ESSBudgetTransaction(pConStr As String, IA_Number As String, IA_date As String, Departemen_ID As String, Departemen_Name As String, Cost_Center As String, IO_Number As String, GL_Account As String, NPK As String, Employee_Name As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_Interface_ESSBudget '" & IA_Number & "', '" & IA_date & "', '" & Departemen_ID & "', '" & Departemen_Name & "', '" & Cost_Center & "', '" & IO_Number & "', '" & GL_Account & "', '" & NPK & "', '" & Employee_Name & "' "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function SAPPOTransaction(pConStr As String, PONumber As String, PODate As String, IABudgetNo As String, PRNumber As String, ItemCode As String, Qty As String, Currency As String, POPrice As String, POAmount As String, CreateDate As String, Createdby As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_Interface_SAPPO '" & PONumber & "', '" & PODate & "', '" & IABudgetNo & "', '" & PRNumber & "', '" & ItemCode & "', '" & Qty & "', '" & Currency & "', '" & POPrice & "', '" & POAmount & "', '" & CreateDate & "', '" & Createdby & "' "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function SAP_Item_Transaction(pConStr As String, Material_No As String, SAP_No As String, Descs As String, UOM As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_Interface_SAP_Item '" & Material_No & "', '" & SAP_No & "', '" & Descs & "', '" & UOM & "'"

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function SAP_GR_Transaction(pConStr As String, ReceivingNo As String, ReceivingDate As String, PONo As String, ItemCode As String, Qty As String, Status As String, CreateDate As String, Createdby As String)
        Dim sql As String
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = "exec sp_Interface_SAP_GR '" & ReceivingNo & "', '" & ReceivingDate & "', '" & PONo & "', '" & ItemCode & "', '" & Qty & "', '" & Status & "', '" & CreateDate & "', '" & Createdby & "' "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            Return retValue
        End Try

    End Function

    Public Shared Function getDataIAMI_PressPartLocalComponent_XMLToSAP(ByVal pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_get_IAMI_PressPart_LocalComponent_XMLToSAP"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function


    Public Shared Function UPD_InfoRecord_Request(ByVal infoRecord As InfoRecordReq, ByVal StatusHeader As Integer, ByVal pConStr As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "Upd_InfoRecord_Statusheader"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@MaterialNumb", infoRecord.MATNR)
            cmd.Parameters.AddWithValue("@Supplier", infoRecord.VENDOR)
            cmd.Parameters.AddWithValue("@StatusHeader", StatusHeader)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function UPD_MaterialMaster_Request(ByVal materialMaster As MaterialMaster, ByVal StatusHeader As Integer, ByVal pConStr As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "Upd_MaterialMaster_StatusHeader"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@MaterialNumb", materialMaster.MATERIAL)
            cmd.Parameters.AddWithValue("@StatusHeader", StatusHeader)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function update_CM_CostHistory_Header_StatusRequest(ByVal pConStr As String, ByVal pYear As String, ByVal pMonth As String, ByVal pVariantCode As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_CM_CostHistory_Header_StatusRequest_Upd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Year", pYear)
            cmd.Parameters.AddWithValue("@Month", pMonth)
            cmd.Parameters.AddWithValue("@VariantCode", pVariantCode)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function update_IF_IAMI_PressPartLocalComponent_XMLFromSAP(ByVal pConStr As String, ByVal pYear As String, ByVal pMonth As String, ByVal pMaterialNoFinishedUnit As String, _
                                                                             ByVal pSubstitutionOldVariant As String, ByVal pPressPart As String, ByVal pLocalComponent As String, ByVal pMessageError As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_IAMI_PressPartLocalComponent_XMLFromSAP_InsUpd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Year", pYear)
            cmd.Parameters.AddWithValue("@Month", pMonth)
            cmd.Parameters.AddWithValue("@MaterialNoFinishedUnit", pMaterialNoFinishedUnit)
            cmd.Parameters.AddWithValue("@SubstitutionOldVariant", pSubstitutionOldVariant)
            cmd.Parameters.AddWithValue("@PressPart", IIf(pPressPart = "", 0, pPressPart))
            cmd.Parameters.AddWithValue("@LocalComponent", IIf(pLocalComponent = "", 0, pLocalComponent))
            cmd.Parameters.AddWithValue("@MessageError", pMessageError)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function delete_IF_PR_XMLFromDMMS_ProcessDate(ByVal pConStr As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_ProcessDate_Del"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function delete_IF_PR_DMMS_RESP(ByVal pConStr As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_DMMS_RESP_Del"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function check_DMMS_PR_IF_History(ByVal pConStr As String, ByVal pFileName As String) As Boolean
        Dim status As Boolean = False

        Dim sSystemFrom As String = "DMMS"
        Dim sModule As String = "PR"

        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_History_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@SystemFrom", sSystemFrom)
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@Module", sModule)
            i = CInt(cmd.ExecuteScalar())
            If i > 0 Then
                status = True
            End If
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function update_DMMS_PR_IF_History(ByVal pConStr As String, ByVal pFileName As String, ByVal pStatus As String, ByVal pErrorMessage As String) As Boolean
        Dim status As Boolean = False

        Dim sSystemFrom As String = "DMMS"
        Dim sModule As String = "PR"

        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_History_InsUpd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@SystemFrom", sSystemFrom)
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@Module", sModule)
            cmd.Parameters.AddWithValue("@Status", pStatus)
            cmd.Parameters.AddWithValue("@ErrorMessage", pErrorMessage)
            i = CInt(cmd.ExecuteNonQuery())
            If i > 0 Then
                status = True
            End If
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function insert_DMMS_PR_IF_RESP(ByVal pConStr As String, ByVal pFileName As String, ByVal pPRNo As String, ByVal pStatus As String, ByVal pErrorMessage As String) As Boolean
        Dim status As Boolean = False

        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_DMMS_RESP_Ins"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@PR_Number", pPRNo)
            cmd.Parameters.AddWithValue("@Status", pStatus)
            cmd.Parameters.AddWithValue("@ErrorMessage", pErrorMessage)
            i = CInt(cmd.ExecuteNonQuery())
            If i > 0 Then
                status = True
            End If
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function insert_IF_PR_XMLFromDMMS(ByVal pConStr As String, ByVal pFileName As String, ByVal pPRNumber As String, ByVal pPRDate As String, _
                                                    ByVal pBudgetCode As String, ByVal pPRTypeCode As String, ByVal pDepartmentCode As String, _
                                                    ByVal pSectionCode As String, ByVal pCostCenter As String, ByVal pProject As String, _
                                                    ByVal pUrgentStatus As String, ByVal pUrgentNote As String, ByVal pReqPOIssueDate As String, _
                                                    ByVal pMaterialNo As String, ByVal pMaterialNoDMMS As String, ByVal pQty As String, ByVal pRemarks As String, _
                                                    ByVal pUpdateDate As String, ByVal pUpdateUser As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_Ins"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@PR_Number", pPRNumber)
            cmd.Parameters.AddWithValue("@PR_Date", pPRDate)
            cmd.Parameters.AddWithValue("@Budget_Code", pBudgetCode)
            cmd.Parameters.AddWithValue("@PRType_Code", pPRTypeCode)
            cmd.Parameters.AddWithValue("@Department_Code", pDepartmentCode)
            cmd.Parameters.AddWithValue("@Section_Code", pSectionCode)
            cmd.Parameters.AddWithValue("@CostCenter", pCostCenter)
            cmd.Parameters.AddWithValue("@Project", pProject)
            cmd.Parameters.AddWithValue("@Urgent_Status", pUrgentStatus)
            cmd.Parameters.AddWithValue("@Urgent_Note", pUrgentNote)
            cmd.Parameters.AddWithValue("@Req_POIssueDate", pReqPOIssueDate)
            cmd.Parameters.AddWithValue("@Material_No", pMaterialNo)
            cmd.Parameters.AddWithValue("@Material_No_DMMS", pMaterialNoDMMS)
            cmd.Parameters.AddWithValue("@Qty", pQty)
            cmd.Parameters.AddWithValue("@Remarks", pRemarks)
            cmd.Parameters.AddWithValue("@UpdateDate", pUpdateDate)
            cmd.Parameters.AddWithValue("@UpdateUser", pUpdateUser)
            i = CInt(cmd.ExecuteNonQuery())

            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function insert_PR_Header(ByVal pConStr As String, ByVal pPRNumber As String, ByVal pPRDate As String, _
                                                    ByVal pBudgetCode As String, ByVal pPRTypeCode As String, ByVal pDepartmentCode As String, _
                                                    ByVal pSectionCode As String, ByVal pCostCenter As String, ByVal pProject As String, _
                                                    ByVal pUrgentStatus As String, ByVal pUrgentNote As String, ByVal pReqPOIssueDate As String, _
                                                    ByVal pUpdateDate As String, ByVal pUpdateUser As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_PRHeader_DMMS_Ins"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PRNo", pPRNumber)
            cmd.Parameters.AddWithValue("@PRDate", pPRDate)
            cmd.Parameters.AddWithValue("@Budget", pBudgetCode)
            cmd.Parameters.AddWithValue("@PRType", pPRTypeCode)
            cmd.Parameters.AddWithValue("@Department", pDepartmentCode)
            cmd.Parameters.AddWithValue("@Section", pSectionCode)
            cmd.Parameters.AddWithValue("@CostCenter", pCostCenter)
            cmd.Parameters.AddWithValue("@Project", pProject)
            cmd.Parameters.AddWithValue("@UrgentStatus", pUrgentStatus)
            cmd.Parameters.AddWithValue("@UrgentNote", pUrgentNote)
            cmd.Parameters.AddWithValue("@PODate", pReqPOIssueDate)
            cmd.Parameters.AddWithValue("@UpdateDate", pUpdateDate)
            cmd.Parameters.AddWithValue("@UpdateUser", pUpdateUser)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function insert_PR_Approval(ByVal pConStr As String, ByVal pPRNumber As String, ByVal pDepartmentCode As String, ByVal pUrgentStatus As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_PRHeader_DMMS_Approval_Ins"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PRNo", pPRNumber)
            cmd.Parameters.AddWithValue("@Department", pDepartmentCode)
            cmd.Parameters.AddWithValue("@UrgentStatus", pUrgentStatus)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function insert_PR_Detail(ByVal pConStr As String, ByVal pPRNumber As String, ByVal pMaterialNo As String, ByVal pMaterialNoDMMS As String, _
                                            ByVal pQty As String, ByVal pRemarks As String, ByVal pUpdateDate As String, ByVal pUpdateUser As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_PRDetail_DMMS_Ins"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PRNo", pPRNumber)
            cmd.Parameters.AddWithValue("@MaterialNo", pMaterialNo)
            cmd.Parameters.AddWithValue("@MaterialNoDMMS", pMaterialNoDMMS)
            cmd.Parameters.AddWithValue("@Qty", pQty)
            cmd.Parameters.AddWithValue("@Remarks", pRemarks)
            cmd.Parameters.AddWithValue("@UpdateDate", pUpdateDate)
            cmd.Parameters.AddWithValue("@UpdateUser", pUpdateUser)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function update_IF_PR_XMLFromDMMS_Process(ByVal pConStr As String, ByVal pFileName As String, ByVal pPRNumber As String, ByVal pMaterialNo As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_ProcessDate_Upd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@PRNo", pPRNumber)
            cmd.Parameters.AddWithValue("@MaterialNo", pMaterialNo)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function delete_IF_PR_XMLFromDMMS(ByVal pConStr As String, ByVal pFileName As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_Del"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function delete_IF_PR_XMLFromDMMS_All(ByVal pConStr As String, ByVal pFileName As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_Del_All"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function check_PR_Header(ByVal pConStr As String, ByVal pPRNumber As String) As Boolean
        Dim status As Boolean = False

        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_PRHeader_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PR_Number", pPRNumber)
            i = CInt(cmd.ExecuteScalar())
            If i > 0 Then
                status = True
            End If
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function check_Mst_Item(ByVal pConStr As String, ByVal pMaterialNo As String) As Boolean
        Dim status As Boolean = False

        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_ItemList_GetDataItem"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@ItemCode", pMaterialNo)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                status = True
            End If
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function check_UserSetup(ByVal pConStr As String, ByVal pUserID As String) As Boolean
        Dim status As Boolean = False

        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_UserSetup_sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@UserID", pUserID)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                status = True
            End If
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return status
        End Using
    End Function

    Public Shared Function getData_IF_PR_DMMS_RESP(ByVal pConStr As String, ByVal pFileName As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_DMMS_RESP_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return dt
        End Using
    End Function

    Public Shared Function getData_IF_PR_XMLFromDMMS(ByVal pConStr As String, ByVal pFileName As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return dt
        End Using
    End Function

    Public Shared Function getData_IF_PR_XMLFromDMMS_Detail(ByVal pConStr As String, ByVal pFileName As String, ByVal pPRNumber As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_IF_PR_XMLFromDMMS_Detail_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FileName", pFileName)
            cmd.Parameters.AddWithValue("@PRNumber", pPRNumber)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return dt
        End Using
    End Function

    Public Shared Function getDataCM_Status_XMLToSAP(ByVal pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_CM_CostHistory_Header_Status_Sel"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataCM_Detail_XMLToSAP(ByVal pConStr As String, ByVal pYear As String, ByVal pMonth As String, ByVal pVariantCode As String, _
                                                     ByVal pMaterialFUDescription As String, ByVal pSubstitutionOldVariant As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_CM_CostHistory_Pivot_XMLToSAP_Sel"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Year", pYear)
            cmd.Parameters.AddWithValue("@Month", pMonth)
            cmd.Parameters.AddWithValue("@VariantCode", pVariantCode)
            cmd.Parameters.AddWithValue("@MaterialFUDescription", pMaterialFUDescription)
            cmd.Parameters.AddWithValue("@SubstitutionOldVariant", pSubstitutionOldVariant)
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            da.Dispose()
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return dt
        End Using
    End Function

    Public Shared Function getDataGR_XMLToDMMS(ByVal pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_Interface_GRToDMMS"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function update_CM_CostHistory_Header_Status(ByVal pConStr As String, ByVal pYear As String, ByVal pMonth As String, ByVal pVariantCode As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_CM_CostHistory_Header_Status_Upd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Year", pYear)
            cmd.Parameters.AddWithValue("@Month", pMonth)
            cmd.Parameters.AddWithValue("@VariantCode", pVariantCode)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Public Shared Function update_GR_TransferDate(ByVal pConStr As String, ByVal pReceivingNumber As String, ByVal pPONumber As String, ByVal pItemCode As String) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "sp_Receiving_Master_TransferDMMS_Upd"

            cmd = New SqlCommand()
            cmd.CommandText = sql
            cmd.Connection = con
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@ReceivingNumber", pReceivingNumber)
            cmd.Parameters.AddWithValue("@PONumber", pPONumber)
            cmd.Parameters.AddWithValue("@ItemCode", pItemCode)
            i = CInt(cmd.ExecuteNonQuery())
            cmd.Parameters.Clear()
            cmd.Dispose()

            Return i
        End Using
    End Function

    Private Sub createNode_InfoRecord(ByVal InfoRecord As InfoRecordReq, ByVal writer As XmlTextWriter)
        ''TOADVANCE: advancing this code into more small line by using reflections and array
        writer.WriteStartElement("Item")
        writer.WriteStartElement("VENDOR")
        writer.WriteString(InfoRecord.VENDOR)
        writer.WriteEndElement()
        writer.WriteStartElement("MATNR")
        writer.WriteString(InfoRecord.MATNR)
        writer.WriteEndElement()
        writer.WriteStartElement("PURORG")
        writer.WriteString(InfoRecord.PURORG)
        writer.WriteEndElement()
        writer.WriteStartElement("WERKS")
        writer.WriteString(InfoRecord.WERKS)
        writer.WriteEndElement()
        writer.WriteStartElement("STANDARD")
        writer.WriteString(InfoRecord.STANDARD)
        writer.WriteEndElement()
        writer.WriteStartElement("SUBCONT")
        writer.WriteString(InfoRecord.SUBCONT)
        writer.WriteEndElement()
        writer.WriteStartElement("PIPEZ")
        writer.WriteString(InfoRecord.PIPEZ)
        writer.WriteEndElement()
        writer.WriteStartElement("CON")
        writer.WriteString(InfoRecord.CON)
        writer.WriteEndElement()
        writer.WriteStartElement("REM1")
        writer.WriteString(InfoRecord.REM1)
        writer.WriteEndElement()
        writer.WriteStartElement("REM2")
        writer.WriteString(InfoRecord.REM2)
        writer.WriteEndElement()
        writer.WriteStartElement("REM3")
        writer.WriteString(InfoRecord.REM3)
        writer.WriteEndElement()
        writer.WriteStartElement("VEN_MAT")
        writer.WriteString(InfoRecord.VEN_MAT)
        writer.WriteEndElement()
        writer.WriteStartElement("COUNTRY")
        writer.WriteString(InfoRecord.COUNTRY)
        writer.WriteEndElement()
        writer.WriteStartElement("REGION")
        writer.WriteString(InfoRecord.REGION)
        writer.WriteEndElement()
        writer.WriteStartElement("VEN_MAT_GRP")
        writer.WriteString(InfoRecord.VEN_MAT_GRP)
        writer.WriteEndElement()
        writer.WriteStartElement("MANUF")
        writer.WriteString(InfoRecord.MANUF)
        writer.WriteEndElement()
        writer.WriteStartElement("SALES_PER")
        writer.WriteString(InfoRecord.SALES_PER)
        writer.WriteEndElement()
        writer.WriteStartElement("TELEPH")
        writer.WriteString(InfoRecord.TELEPH)
        writer.WriteEndElement()
        writer.WriteStartElement("ORD_UNIT")
        writer.WriteString(InfoRecord.ORD_UNIT)
        writer.WriteEndElement()
        writer.WriteStartElement("ORD_DIN")
        writer.WriteString(InfoRecord.ORD_DIN)
        writer.WriteEndElement()
        writer.WriteStartElement("ORD_NUM")
        writer.WriteString(InfoRecord.ORD_NUM)
        writer.WriteEndElement()
        writer.WriteStartElement("NUMBERZ")
        writer.WriteString(InfoRecord.NUMBERZ)
        writer.WriteEndElement()
        writer.WriteStartElement("PLAN_DEL_TIME")
        writer.WriteString(InfoRecord.PLAN_DEL_TIME)
        writer.WriteEndElement()
        writer.WriteStartElement("PUR_GRP")
        writer.WriteString(InfoRecord.PUR_GRP)
        writer.WriteEndElement()
        writer.WriteStartElement("STND_QTY")
        writer.WriteString(InfoRecord.STND_QTY)
        writer.WriteEndElement()
        writer.WriteStartElement("MIN_QTY")
        writer.WriteString(InfoRecord.MIN_QTY)
        writer.WriteEndElement()
        writer.WriteStartElement("ROUND_PROF")
        writer.WriteString(InfoRecord.ROUND_PROF)
        writer.WriteEndElement()
        writer.WriteStartElement("CONF_KEY")
        writer.WriteString(InfoRecord.CONF_KEY)
        writer.WriteEndElement()
        writer.WriteStartElement("TAX_CODE")
        writer.WriteString(InfoRecord.TAX_CODE)
        writer.WriteEndElement()
        writer.WriteStartElement("NET_PRICE")
        writer.WriteString(InfoRecord.NET_PRICE)
        writer.WriteEndElement()
        writer.WriteStartElement("WAERS")
        writer.WriteString(InfoRecord.WAERS)
        writer.WriteEndElement()
        writer.WriteStartElement("VALID_ON")
        writer.WriteString(InfoRecord.VALID_ON)
        writer.WriteEndElement()
        writer.WriteStartElement("VALID_TO")
        writer.WriteString(InfoRecord.VALID_TO)
        writer.WriteEndElement()
        writer.WriteStartElement("COND")
        writer.WriteString(InfoRecord.COND)
        writer.WriteEndElement()
        writer.WriteStartElement("PRICE")
        writer.WriteString(InfoRecord.PRICE)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub createNode_MaterialMaster(ByVal materialMaster As MaterialMaster, ByVal writer As XmlTextWriter)
        ''TOADVANCE: advancing this code into more small line by using reflections and array
        writer.WriteStartElement("Item")
        writer.WriteStartElement("MATERIAL")
        writer.WriteString(materialMaster.MATERIAL)
        writer.WriteEndElement()
        writer.WriteStartElement("IND_SECTOR")
        writer.WriteString(materialMaster.IND_SECTOR)
        writer.WriteEndElement()
        writer.WriteStartElement("MATL_TYPE")
        writer.WriteString(materialMaster.IND_SECTOR)
        writer.WriteEndElement()
        writer.WriteStartElement("BASIC_VIEW")
        writer.WriteString(materialMaster.BASIC_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("SALES_VIEW")
        writer.WriteString(materialMaster.SALES_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("PURCHASE_VIEW")
        writer.WriteString(materialMaster.PURCHASE_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("MRP_VIEW")
        writer.WriteString(materialMaster.MRP_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("STORAGE_VIEW")
        writer.WriteString(materialMaster.STORAGE_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("ACCOUNT_VIEW")
        writer.WriteString(materialMaster.ACCOUNT_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("COST_VIEW")
        writer.WriteString(materialMaster.COST_VIEW)
        writer.WriteEndElement()
        writer.WriteStartElement("MATL_DESC")
        writer.WriteString(materialMaster.MATL_DESC)
        writer.WriteEndElement()
        writer.WriteStartElement("LANGU")
        writer.WriteString(materialMaster.LANGU)
        writer.WriteEndElement()
        writer.WriteStartElement("MATL_GROUP")
        writer.WriteString(materialMaster.MATL_GROUP)
        writer.WriteEndElement()
        writer.WriteStartElement("BASE_UOM")
        writer.WriteString(materialMaster.BASE_UOM)
        writer.WriteEndElement()
        writer.WriteStartElement("DIVISION")
        writer.WriteString(materialMaster.DIVISION)
        writer.WriteEndElement()
        writer.WriteStartElement("PLANT")
        writer.WriteString(materialMaster.PLANT)
        writer.WriteEndElement()
        writer.WriteStartElement("TRANS_GRP")
        writer.WriteString(materialMaster.TRANS_GRP)
        writer.WriteEndElement()
        writer.WriteStartElement("LOADINGGRP")
        writer.WriteString(materialMaster.LOADINGGRP)
        writer.WriteEndElement()
        writer.WriteStartElement("AVAILCHECK")
        writer.WriteString(materialMaster.AVAILCHECK)
        writer.WriteEndElement()
        writer.WriteStartElement("PROFIT_CTR")
        writer.WriteString(materialMaster.PROFIT_CTR)
        writer.WriteEndElement()
        writer.WriteStartElement("PUR_GROUP")
        writer.WriteString(materialMaster.PUR_GROUP)
        writer.WriteEndElement()
        writer.WriteStartElement("MANU_MAT")
        writer.WriteString(materialMaster.MANU_MAT)
        writer.WriteEndElement()
        writer.WriteStartElement("MRP_TYPE")
        writer.WriteString(materialMaster.MRP_TYPE)
        writer.WriteEndElement()
        writer.WriteStartElement("MRP_CTRLER")
        writer.WriteString(materialMaster.MRP_CTRLER)
        writer.WriteEndElement()
        writer.WriteStartElement("LOTSIZEKEY")
        writer.WriteString(materialMaster.LOTSIZEKEY)
        writer.WriteEndElement()
        writer.WriteStartElement("REORDER_PT")
        writer.WriteString(materialMaster.REORDER_PT)
        writer.WriteEndElement()
        writer.WriteStartElement("ROUND_VAL")
        writer.WriteString(materialMaster.ROUND_VAL)
        writer.WriteEndElement()
        writer.WriteStartElement("MAX_STOCK")
        writer.WriteString(materialMaster.MAX_STOCK)
        writer.WriteEndElement()
        writer.WriteStartElement("PLND_DELRY")
        writer.WriteString(materialMaster.PLND_DELRY)
        writer.WriteEndElement()
        writer.WriteStartElement("PROC_TYPE")
        writer.WriteString(materialMaster.PROC_TYPE)
        writer.WriteEndElement()
        writer.WriteStartElement("SAFETY_STK")
        writer.WriteString(materialMaster.SAFETY_STK)
        writer.WriteEndElement()
        writer.WriteStartElement("BACKFLUSH")
        writer.WriteString(materialMaster.BACKFLUSH)
        writer.WriteEndElement()
        writer.WriteStartElement("ISS_ST_LOC")
        writer.WriteString(materialMaster.ISS_ST_LOC)
        writer.WriteEndElement()
        writer.WriteStartElement("SLOC_EXPRC")
        writer.WriteString(materialMaster.SLOC_EXPRC)
        writer.WriteEndElement()
        writer.WriteStartElement("PERIOD_IND")
        writer.WriteString(materialMaster.PERIOD_IND)
        writer.WriteEndElement()
        writer.WriteStartElement("PLANT_D")
        writer.WriteString(materialMaster.PLANT_D)
        writer.WriteEndElement()
        writer.WriteStartElement("STGE_LOC")
        writer.WriteString(materialMaster.STGE_LOC)
        writer.WriteEndElement()
        writer.WriteStartElement("VAL_AREA")
        writer.WriteString(materialMaster.VAL_AREA)
        writer.WriteEndElement()
        writer.WriteStartElement("PRICE_CTRL")
        writer.WriteString(materialMaster.PRICE_CTRL)
        writer.WriteEndElement()
        writer.WriteStartElement("VAL_CLASS")
        writer.WriteString(materialMaster.VAL_CLASS)
        writer.WriteEndElement()
        writer.WriteStartElement("MOVING_PR")
        writer.WriteString(materialMaster.MOVING_PR)
        writer.WriteEndElement()
        writer.WriteStartElement("ORIG_MAT")
        writer.WriteString(materialMaster.ORIG_MAT)
        writer.WriteEndElement()
        writer.WriteStartElement("MFR_PRTNUMB")
        writer.WriteString(materialMaster.MFR_PRTNUMB)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub createNodeIAPrice(ByVal pSupplier As String, ByVal pMaterial As String, ByVal pPrice As String, ByVal pValid As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Item")
        writer.WriteStartElement("Supplier_Code")
        writer.WriteString(pSupplier)
        writer.WriteEndElement()
        writer.WriteStartElement("Material_No")
        writer.WriteString(pMaterial)
        writer.WriteEndElement()
        writer.WriteStartElement("IA_Price")
        writer.WriteString(pPrice)
        writer.WriteEndElement()
        writer.WriteStartElement("Valid_Date")
        writer.WriteString(pValid)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub createNodePRResponse(ByVal pPRNo As String, ByVal pStatus As String, ByVal pErrorMessage As String, ByVal pProcessDate As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Item")
        writer.WriteStartElement("PR_Number")
        writer.WriteString(pPRNo)
        writer.WriteEndElement()
        writer.WriteStartElement("Status")
        writer.WriteString(pStatus)
        writer.WriteEndElement()
        writer.WriteStartElement("Error_Message")
        writer.WriteString(pErrorMessage)
        writer.WriteEndElement()
        writer.WriteStartElement("Process_Date")
        writer.WriteString(pProcessDate)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub createNodeIAMI_PressPartLocalComponent(ByVal pYear As String, ByVal pMonth As String, ByVal pMaterialFinishUnit As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Item")
        writer.WriteStartElement("Year")
        writer.WriteString(pYear)
        writer.WriteEndElement()
        writer.WriteStartElement("Month")
        writer.WriteString(pMonth)
        writer.WriteEndElement()
        writer.WriteStartElement("MaterialFinishUnit")
        writer.WriteString(pMaterialFinishUnit)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub createNodeIAMI_CMHistory(ByVal pMaterialFinishUnit As String, ByVal pMaterialFinishUnitDesc As String, ByVal pSubstitutionOldVariant As String, _
                                         ByVal pYear As String, ByVal pMonth As String, ByVal pCurr As String, ByVal pIAMILocalParts As Double, ByVal pMIILocalParts As Double, _
                                         ByVal pMIIAssyFree As Double, ByVal pIAMILP_OrderMultiSupplier As Double, ByVal pIAMILP_ExchangeRate As Double, _
                                         ByVal pIAMILP_BOMCorrection As Double, ByVal pIAMILP_AddLocalization As Double, ByVal pIAMILP_ProjectImprovement As Double, _
                                         ByVal pIAMILP_CostReduction_PriceAdjustment As Double, ByVal pIAMILP_CostReduction_Others As Double, ByVal pIAMILP_Process_Adjustment As String, _
                                         ByVal pIAMILP_MaterialAdjustment As Double, ByVal pMIILP_OrderMultiSupplier As Double, ByVal pMIILP_ExchangeRate As Double, _
                                         ByVal pMIILP_BOMCorrection As Double, ByVal pMIILP_AddLocalization As Double, ByVal pMIILP_ProjectImprovement As Double, _
                                         ByVal pMIILP_CostReduction_PriceAdjustment As Double, ByVal pMIILP_CostReduction_Others As Double, ByVal pMIILP_Process_Adjustment As String, _
                                         ByVal pMIILP_MaterialAdjustment As Double, ByVal pMIIAF_ProcessAdjustment As Double, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Item")
        writer.WriteStartElement("MaterialFinishUnit")
        writer.WriteString(pMaterialFinishUnit)
        writer.WriteEndElement()
        writer.WriteStartElement("MaterialFUDescription")
        writer.WriteString(pMaterialFinishUnitDesc)
        writer.WriteEndElement()
        writer.WriteStartElement("SubstitutionOldVariant")
        writer.WriteString(pSubstitutionOldVariant)
        writer.WriteEndElement()
        writer.WriteStartElement("Year")
        writer.WriteString(pYear)
        writer.WriteEndElement()
        writer.WriteStartElement("Period")
        writer.WriteString(pMonth)
        writer.WriteEndElement()
        writer.WriteStartElement("Currency")
        writer.WriteString(pCurr)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP")
        writer.WriteString(pIAMILocalParts)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILocalParts")
        writer.WriteString(pMIILocalParts)
        writer.WriteEndElement()
        writer.WriteStartElement("MIIAssyFee")
        writer.WriteString(pMIIAssyFree)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_OrderMultiSupplier")
        writer.WriteString(pIAMILP_OrderMultiSupplier)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_ExchangeRate")
        writer.WriteString(pIAMILP_ExchangeRate)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_BOMCorrection")
        writer.WriteString(pIAMILP_BOMCorrection)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_AddLocalization")
        writer.WriteString(pIAMILP_AddLocalization)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_ProjectImprovement")
        writer.WriteString(pIAMILP_ProjectImprovement)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_CostReduction_PriceAdjustment")
        writer.WriteString(pIAMILP_CostReduction_PriceAdjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_CostReduction_Others")
        writer.WriteString(pIAMILP_CostReduction_Others)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_Process_Adjustment")
        writer.WriteString(pIAMILP_Process_Adjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("IAMILP_MaterialAdjustment")
        writer.WriteString(pIAMILP_MaterialAdjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_OrderMultiSupplier")
        writer.WriteString(pMIILP_OrderMultiSupplier)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_ExchangeRate")
        writer.WriteString(pMIILP_ExchangeRate)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_BOMCorrection")
        writer.WriteString(pMIILP_BOMCorrection)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_AddLocalization")
        writer.WriteString(pMIILP_AddLocalization)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_ProjectImprovement")
        writer.WriteString(pMIILP_ProjectImprovement)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_CostReduction_PriceAdjustment")
        writer.WriteString(pMIILP_CostReduction_PriceAdjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_CostReduction_Others")
        writer.WriteString(pMIILP_CostReduction_Others)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_ProcessAdjustment")
        writer.WriteString(pMIILP_Process_Adjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("MIILP_MaterialAdjustment")
        writer.WriteString(pMIILP_MaterialAdjustment)
        writer.WriteEndElement()
        writer.WriteStartElement("MIIAF_ProcessAdjustment")
        writer.WriteString(pMIIAF_ProcessAdjustment)
        writer.WriteEndElement()

        writer.WriteEndElement()
    End Sub

    Private Sub createNodeXMLGRToDMMS(ByVal pReceivingNumber As String, ByVal pReceivingDate As String, ByVal pPRNumber As String, _
                                         ByVal pPONumber As String, ByVal pPODate As String, ByVal pIABudgetNo As String, ByVal pItemCode As String, _
                                         ByVal pQty As String, ByVal pPrice As String, ByVal pCurrency As String, _
                                         ByVal pCreateDate As String, ByVal pCreateBy As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Item")
        writer.WriteStartElement("ReceivingNumber")
        writer.WriteString(pReceivingNumber)
        writer.WriteEndElement()
        writer.WriteStartElement("ReceivingDate")
        writer.WriteString(pReceivingDate)
        writer.WriteEndElement()
        writer.WriteStartElement("PRNumber")
        writer.WriteString(pPRNumber)
        writer.WriteEndElement()
        writer.WriteStartElement("PONumber")
        writer.WriteString(pPONumber)
        writer.WriteEndElement()
        writer.WriteStartElement("PODate")
        writer.WriteString(pPODate)
        writer.WriteEndElement()
        writer.WriteStartElement("IABudgetNo")
        writer.WriteString(pIABudgetNo)
        writer.WriteEndElement()
        writer.WriteStartElement("ItemCode")
        writer.WriteString(pItemCode)
        writer.WriteEndElement()
        writer.WriteStartElement("Qty")
        writer.WriteString(pQty)
        writer.WriteEndElement()
        writer.WriteStartElement("Price")
        writer.WriteString(pPrice)
        writer.WriteEndElement()
        writer.WriteStartElement("Currency")
        writer.WriteString(pCurrency)
        writer.WriteEndElement()
        writer.WriteStartElement("CreateDate")
        writer.WriteString(pCreateDate)
        writer.WriteEndElement()
        writer.WriteStartElement("CreateBy")
        writer.WriteString(pCreateBy)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Function generateXMLIAPrice() As Boolean
        Dim iStatus As Boolean = True
        Try
            Dim supplierCode As String = ""
            Dim materialNo As String = ""
            Dim IAPrice As String = ""
            Dim validDate As String = ""
            Dim fromEmail As String = ""
            Dim iRow As Integer

            Dim statusSend As Boolean = False

            ConStr = Builder.ConnectionString
            Dim dt As DataTable
            dt = ClsInterfaceSettingDB.getData(ConStr)

            If dt.Rows.Count > 0 Then
                gs_IAPricePath = Trim(dt.Rows(0).Item("XML_IAPrice_Path"))
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for IA Price " & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If
            dt.Reset()

            dt = ClsInterfaceSettingDB.getDataIAPrice(ConStr)

            If dt.Rows.Count > 0 Then
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] START Interface Generate IA Price!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                Dim path As String = gs_IAPricePath + "\" + "IA PRICE " & Format(Now, "yyyyMMdd HHmmss") & ".xml"
                Dim writer As New XmlTextWriter(path, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Indentation = 2
                writer.WriteStartElement("IA_PRICEOUTPUT")
                For iRow = 0 To dt.Rows.Count - 1
                    supplierCode = Trim(dt.Rows(iRow).Item("Supplier_Code") & "")
                    materialNo = Trim(dt.Rows(iRow).Item("Material_No") & "")
                    IAPrice = Trim(dt.Rows(iRow).Item("FinalPrice") & "")
                    validDate = Trim(dt.Rows(iRow).Item("Valid_Date").ToString & "")

                    createNodeIAPrice(supplierCode, materialNo, IAPrice, validDate, writer)
                Next iRow

                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()

                ' Log for Success Sending Email 
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Interface Process SAP Item Success !" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                statusSend = True
            End If

            If statusSend = False Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] No Data IA Price to interface ! " & vbCrLf
                up_gridProcess(Rtb1, 3, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If

        Catch ex As Exception
            'up_ShowMessage(Err.Number, txtMsg, clsGlobal.MsgTypeEnum.ErrorMsg, Err.Description)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error Interface Process SAP Item... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            iStatus = False
            Return iStatus
        End Try

        Return iStatus
    End Function

    Private Function generateXMLIAMI_PressPartLocalComponent() As Boolean
        Dim iStatus As Boolean = True

        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim year As String = ""
            Dim month As String = ""
            Dim materialFinishUnit As String = ""
            Dim iRow As Integer

            Dim statusSend As Boolean = False

            ConStr = Builder.ConnectionString
            Dim dt As DataTable
            dt = ClsInterfaceSettingDB.getData(ConStr)

            If dt.Rows.Count > 0 Then
                gs_IAMI_PressPartXMLToSAPPath = Trim(dt.Rows(0).Item("IAMI_PressPart_XMLToSAP_Path"))
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for IAMI Press Part & Local Component XML to SAP!" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If
            dt.Reset()

            dt = getDataIAMI_PressPartLocalComponent_XMLToSAP(ConStr)

            If dt.Rows.Count > 0 Then
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] START Interface Generate IAMI Press Part & Local Component XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "IAMI_PressPartsLocalComponent_Req_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"
                'Dim path As String = gs_IAMI_PressPartXMLToSAPPath + "\IAMI_PressPartsLocalComponent_Req_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"

                Dim xmlFilePath As String = tempLocalPath & tempLocalFile

                Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())

                writer.WriteStartDocument()
                writer.Indentation = 2
                writer.WriteStartElement("ns0:MT_EPIC_PressPartLocalComponent_Req")
                writer.WriteStartAttribute("xmlns:ns0")
                writer.WriteString("http://iamiepic")
                writer.WriteStartElement("FT_INPUT")
                For iRow = 0 To dt.Rows.Count - 1
                    year = Trim(dt.Rows(iRow).Item("Year") & "")
                    month = Trim(dt.Rows(iRow).Item("Month") & "")
                    materialFinishUnit = Trim(dt.Rows(iRow).Item("MaterialFinishUnit") & "")

                    createNodeIAMI_PressPartLocalComponent(year, month, materialFinishUnit, writer)
                Next iRow

                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()

                'Log for Success Sending Email 
                '-----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Interface Process IAMI Press Part & Local Component XML To SAP Success !" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                'update status request
                For iRow = 0 To dt.Rows.Count - 1
                    year = Trim(dt.Rows(iRow).Item("Year") & "")
                    month = Trim(dt.Rows(iRow).Item("Month") & "")
                    materialFinishUnit = Trim(dt.Rows(iRow).Item("MaterialFinishUnit") & "")

                    update_CM_CostHistory_Header_StatusRequest(ConStr, year, month, materialFinishUnit)
                Next iRow

                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, gs_IAMI_PressPartXMLToSAPPath & "\" & tempLocalFile, True)

                statusSend = True
            End If

            If statusSend = False Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] No Data IAMI Press Part & Local Component XML To SAP! " & vbCrLf
                up_gridProcess(Rtb1, 6, 0, ls_StepProcess)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Generate IAMI Press Part & Local Component XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                iStatus = False
                Return iStatus
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Generate IAMI Press Part & Local Component XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                Return iStatus
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error Interface Process IAMI Press Part & Local Component XML To SAP... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Generate IAMI Press Part & Local Component XML To SAP!" & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            iStatus = False
            Return iStatus
        End Try

        Return iStatus
    End Function

    Private Function generateXMLCMHistory() As Boolean
        Dim iStatus As Boolean = True

        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim year As String = ""
            Dim month As String = ""
            Dim materialFinishUnit As String = ""
            Dim materialFinishUnitDesc As String = ""
            Dim substitutionOldVariant As String = ""

            Dim iRow As Integer
            Dim iRowDetail As Integer

            Dim statusSend As Boolean = False

            ConStr = Builder.ConnectionString
            Dim dt As DataTable
            Dim dtDetail As DataTable

            dt = ClsInterfaceSettingDB.getData(ConStr)

            If dt.Rows.Count > 0 Then
                gs_CostHistoryXMLToSAPPath = Trim(dt.Rows(0).Item("CostHistory_XMLToSAP_Path"))
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for IAMI Cost History XML to SAP!" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If
            dt.Reset()

            dt = getDataCM_Status_XMLToSAP(ConStr)

            If dt.Rows.Count > 0 Then
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] START Interface IAMI Cost History XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "cost_history_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"
                'Dim path As String = gs_CostHistoryXMLToSAPPath + "\cost_history_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"

                Dim xmlFilePath As String = tempLocalPath & tempLocalFile

                Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())

                writer.WriteStartDocument()
                writer.Indentation = 2
                writer.WriteStartElement("ns0:MT_CostHistoryMToSAP")
                writer.WriteStartAttribute("xmlns:ns0")
                writer.WriteString("http://iamiepic")
                writer.WriteStartElement("Table")

                For iRow = 0 To dt.Rows.Count - 1
                    year = Trim(dt.Rows(iRow).Item("Year") & "")
                    month = Trim(dt.Rows(iRow).Item("Periode") & "")
                    materialFinishUnit = Trim(dt.Rows(iRow).Item("VariantCode") & "")
                    materialFinishUnitDesc = Trim(dt.Rows(iRow).Item("VariantName") & "")
                    substitutionOldVariant = Trim(dt.Rows(iRow).Item("SubstitutionOldVariant") & "")

                    dtDetail = getDataCM_Detail_XMLToSAP(ConStr, year, month, materialFinishUnit, materialFinishUnitDesc, substitutionOldVariant)

                    Dim materialFinishUnitDetail As String = ""
                    Dim materialFinishUnitDescDetail As String = ""
                    Dim substitutionOldVariantDetail As String = ""
                    Dim yearDetail As String = ""
                    Dim monthDetail As String = ""
                    Dim currency As String = ""
                    Dim iamiLocalParts As Double = 0
                    Dim miiLocalParts As Double = 0
                    Dim miiAssyFree As Double = 0
                    Dim iamiLP_OrderMultiSupplier As Double = 0
                    Dim iamiLP_ExchangeRate As Double = 0
                    Dim iamiLP_BOMCorrection As Double = 0
                    Dim iamiLP_AddLocalization As Double = 0
                    Dim iamiLP_ProjectImprovement As Double = 0
                    Dim iamiLP_CostReduction_PriceAdjustment As Double = 0
                    Dim iamiLP_CostReduction_Others As Double = 0
                    Dim iamiLP_Process_Adjustment As Double = 0
                    Dim iamiLP_Material_Adjustment As Double = 0
                    Dim miiLP_OrderMultiSupplier As Double = 0
                    Dim miiLP_ExchangeRate As Double = 0
                    Dim miiLP_BOMCorrection As Double = 0
                    Dim miiLP_AddLocalization As Double = 0
                    Dim miiLP_ProjectImprovement As Double = 0
                    Dim miiLP_CostReduction_PriceAdjustment As Double = 0
                    Dim miiLP_CostReduction_Others As Double = 0
                    Dim miiLP_Process_Adjustment As Double = 0
                    Dim miiLP_Material_Adjustment As Double = 0
                    Dim miiAF_Process_Adjustment As Double = 0

                    For iRowDetail = 0 To dtDetail.Rows.Count - 1
                        materialFinishUnitDetail = Trim(dtDetail.Rows(iRowDetail).Item("MaterialFinishUnit") & "")
                        materialFinishUnitDescDetail = Trim(dtDetail.Rows(iRowDetail).Item("MaterialFUDescription") & "")
                        substitutionOldVariantDetail = Trim(dtDetail.Rows(iRowDetail).Item("SubstitutionOldVariant") & "")
                        yearDetail = Trim(dtDetail.Rows(iRowDetail).Item("Year") & "")
                        monthDetail = Trim(dtDetail.Rows(iRowDetail).Item("Month") & "")
                        currency = Trim(dtDetail.Rows(iRowDetail).Item("Currency") & "")
                        iamiLocalParts = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILocalParts") & "")
                        miiLocalParts = CDbl(dtDetail.Rows(iRowDetail).Item("MIILocalParts") & "")
                        miiAssyFree = CDbl(dtDetail.Rows(iRowDetail).Item("MIIAssyFee") & "")
                        iamiLP_OrderMultiSupplier = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_OrderMultiSupplier") & "")
                        iamiLP_ExchangeRate = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_ExchangeRate") & "")
                        iamiLP_BOMCorrection = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_BOMCorrection") & "")
                        iamiLP_AddLocalization = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_AddLocalization") & "")
                        iamiLP_ProjectImprovement = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_ProjectImprovement") & "")
                        iamiLP_CostReduction_PriceAdjustment = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_CostReduction_PriceAdjustment") & "")
                        iamiLP_CostReduction_Others = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_CostReduction_Others") & "")
                        iamiLP_Process_Adjustment = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_Process_Adjustment") & "")
                        iamiLP_Material_Adjustment = CDbl(dtDetail.Rows(iRowDetail).Item("IAMILP_MaterialAdjustment") & "")
                        miiLP_OrderMultiSupplier = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_OrderMultiSupplier") & "")
                        miiLP_ExchangeRate = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_ExchangeRate") & "")
                        miiLP_BOMCorrection = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_BOMCorrection") & "")
                        miiLP_AddLocalization = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_AddLocalization") & "")
                        miiLP_ProjectImprovement = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_ProjectImprovement") & "")
                        miiLP_CostReduction_PriceAdjustment = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_CostReduction_PriceAdjustment") & "")
                        miiLP_CostReduction_Others = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_CostReduction_Others") & "")
                        miiLP_Process_Adjustment = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_ProcessAdjustment") & "")
                        miiLP_Material_Adjustment = CDbl(dtDetail.Rows(iRowDetail).Item("MIILP_MaterialAdjustment") & "")
                        miiAF_Process_Adjustment = CDbl(dtDetail.Rows(iRowDetail).Item("MIIAF_ProcessAdjustment") & "")

                        createNodeIAMI_CMHistory(materialFinishUnitDetail, materialFinishUnitDescDetail, substitutionOldVariantDetail, yearDetail, monthDetail, _
                                             currency, iamiLocalParts, miiLocalParts, miiAssyFree, iamiLP_OrderMultiSupplier, iamiLP_ExchangeRate, iamiLP_BOMCorrection, _
                                             iamiLP_AddLocalization, iamiLP_ProjectImprovement, iamiLP_CostReduction_PriceAdjustment, iamiLP_CostReduction_Others, _
                                             iamiLP_Process_Adjustment, iamiLP_Material_Adjustment, miiLP_OrderMultiSupplier, miiLP_ExchangeRate, miiLP_BOMCorrection, _
                                             miiLP_AddLocalization, miiLP_ProjectImprovement, miiLP_CostReduction_PriceAdjustment, miiLP_CostReduction_Others, _
                                             miiLP_Process_Adjustment, miiLP_Material_Adjustment, miiAF_Process_Adjustment, writer)
                    Next iRowDetail
                    dtDetail.Reset()
                Next iRow

                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()

                ' Log for Success Sending Email 
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Interface Process IAMI Cost History XML To SAP Success !" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                'update status request
                For iRow = 0 To dt.Rows.Count - 1
                    year = Trim(dt.Rows(iRow).Item("Year") & "")
                    month = Trim(dt.Rows(iRow).Item("Periode") & "")
                    materialFinishUnit = Trim(dt.Rows(iRow).Item("VariantCode") & "")

                    update_CM_CostHistory_Header_Status(ConStr, year, month, materialFinishUnit)
                Next iRow

                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, gs_CostHistoryXMLToSAPPath & "\" & tempLocalFile, True)

                statusSend = True
            End If

            If statusSend = False Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] No Data IAMI Cost History XML To SAP! " & vbCrLf
                up_gridProcess(Rtb1, 6, 0, ls_StepProcess)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface IAMI Cost History XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                iStatus = False
                Return iStatus
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface IAMI Cost History XML To SAP!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                Return iStatus
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error Interface Process IAMI Cost History XML To SAP... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface IAMI Cost History XML To SAP!" & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            iStatus = False
            Return iStatus
        End Try

        Return iStatus
    End Function

    Private Function generateXMLGRToDMMS() As Boolean
        Dim iStatus As Boolean = True

        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim ReceivingNumber As String = ""
            Dim ReceivingDate As String = ""
            Dim PRNumber As String = ""
            Dim PONumber As String = ""
            Dim PODate As String = ""
            Dim IABudgetNo As String = ""
            Dim ItemCode As String = ""
            Dim Qty As String = ""
            Dim Currency As String = ""
            Dim Price As String = ""
            Dim CreateDate As String = ""
            Dim CreateBy As String = ""

            Dim iRow As Integer

            Dim statusSend As Boolean = False

            ConStr = Builder.ConnectionString
            Dim dt As DataTable

            dt = ClsInterfaceSettingDB.getData(ConStr)

            If dt.Rows.Count > 0 Then
                gs_GoodsReceiveXMLToDMMSPath = Trim(dt.Rows(0).Item("GoodsReceive_XMLToDMMS_Path"))
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for Goods Receive XML to DMMS!" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If
            dt.Reset()

            If gs_GoodsReceiveXMLToDMMSPath = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for Goods Receive XML to DMMS!" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If

            dt = getDataGR_XMLToDMMS(ConStr)

            If dt.Rows.Count > 0 Then
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] START Interface Goods Receive XML To DMMS!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "DMMS_GR_Send_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"
                'Dim path As String = gs_GoodsReceiveXMLToDMMSPath + "\DMMS_GR_Send_" & Format(Now, "yyyyMMdd_HHmmss_fff") & ".xml"

                Dim xmlFilePath As String = tempLocalPath & tempLocalFile

                Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())

                writer.WriteStartDocument()
                writer.Indentation = 2
                writer.WriteStartElement("ns0:MT_DMMS_GR_Send")
                writer.WriteStartAttribute("xmlns:ns0")
                writer.WriteString("http://iamiepic")
                writer.WriteStartElement("Table")

                For iRow = 0 To dt.Rows.Count - 1
                    ReceivingNumber = Trim(dt.Rows(iRow).Item("ReceivingNumber") & "")
                    ReceivingDate = CDate(dt.Rows(iRow).Item("ReceivingDate")).ToString("yyyy-MM-dd")
                    PRNumber = Trim(dt.Rows(iRow).Item("PRNumber") & "")
                    PONumber = Trim(dt.Rows(iRow).Item("PONumber") & "")
                    PODate = CDate(dt.Rows(iRow).Item("PODate")).ToString("yyyy-MM-dd")
                    IABudgetNo = Trim(dt.Rows(iRow).Item("IABudgetNo") & "")
                    ItemCode = Trim(dt.Rows(iRow).Item("ItemCode") & "")
                    Qty = CInt(dt.Rows(iRow).Item("Qty"))
                    Price = CDbl(dt.Rows(iRow).Item("Price"))
                    Currency = Trim(dt.Rows(iRow).Item("Currency") & "")
                    CreateDate = CDate(dt.Rows(iRow).Item("CreateDate")).ToString("yyyy-MM-dd HH:mm:ss")
                    CreateBy = Trim(dt.Rows(iRow).Item("CreatedBy") & "")

                    createNodeXMLGRToDMMS(ReceivingNumber, ReceivingDate, PRNumber, PONumber, PODate, _
                                             IABudgetNo, ItemCode, Qty, Price, Currency, CreateDate, CreateBy, writer)
                Next iRow

                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()

                ' Log for Success Sending Email 
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Interface Process Goods Receive XML To DMMS Success !" & vbCrLf
                up_gridProcess(Rtb1, 3, 0, ls_StepProcess)

                'update status request
                For iRow = 0 To dt.Rows.Count - 1
                    ReceivingNumber = Trim(dt.Rows(iRow).Item("ReceivingNumber") & "")
                    PONumber = Trim(dt.Rows(iRow).Item("PONumber") & "")
                    ItemCode = Trim(dt.Rows(iRow).Item("ItemCode") & "")

                    update_GR_TransferDate(ConStr, ReceivingNumber, PONumber, ItemCode)
                Next iRow

                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, gs_GoodsReceiveXMLToDMMSPath & "\" & tempLocalFile, True)

                statusSend = True
            End If

            If statusSend = False Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] No Data Good Receive XML To DMMS! " & vbCrLf
                up_gridProcess(Rtb1, 6, 0, ls_StepProcess)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Goods Receive XML To DMMS!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                iStatus = False
                Return iStatus
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Goods Receive XML To DMMS!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                Return iStatus
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error Interface Process Goods Receive XML To DMMS... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Goods Receive XML To DMMS!" & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            iStatus = False
            Return iStatus
        End Try

        Return iStatus
    End Function


    Public Shared Function ResetSchedule(pConStr As String)
        Dim sql As String = ""
        Dim retValue As Integer = 0

        Try
            Using con As New SqlConnection(pConStr)
                con.Open()

                sql = " UPDATE SendEmail_Schedule  " & vbCrLf & _
                          " SET [Time] = CONVERT(char(10),SYSDATETIME(),120) + ' '+ CONVERT(char(8),[Time],108)  " & vbCrLf & _
                          " WHERE [Time] < SYSDATETIME() " & vbCrLf & _
                          "  " & vbCrLf & _
                          " UPDATE SendEmail_Schedule  " & vbCrLf & _
                          " SET [Time] = CONVERT(char(10),DATEADD(HOUR,1,SYSDATETIME()),120) + ' '+ CONVERT(char(8),[Time],108)  " & vbCrLf & _
                          " WHERE [Time] < SYSDATETIME() " & vbCrLf & _
                          "  "

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandType = CommandType.Text
                retValue = cmd.ExecuteNonQuery()

                Return retValue
            End Using
        Catch ex As Exception
            'Throw New Exception("UPDATE DATA ERROR : " & ex.Message)
            Return retValue
        End Try

    End Function

    Private Sub settext(ByVal s As String, ByVal textcontrol As System.Windows.Forms.TextBox)
        If textcontrol.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf settext)
            Me.Invoke(d, New Object() {s, textcontrol})
        Else
            textcontrol.Text = s
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub up_Interface_EPIC1_SupplierManagement_Process()
        ConStr = Builder.ConnectionString
        Dim dtInterface_Setting As DataTable
        dtInterface_Setting = ClsInterfaceSettingDB.getData(ConStr)

        If dtInterface_Setting.Rows.Count > 0 Then
            gs_JSONPath = Trim(dtInterface_Setting.Rows(0).Item("JSON_Path"))
            gs_JSONPathBackup = Trim(dtInterface_Setting.Rows(0).Item("JSON_PathBackup"))
        Else
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for EPIC1 - Supplier Management  " & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Exit Sub
        End If

        Dim di As New IO.DirectoryInfo(gs_JSONPath & "\")
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.json")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            jsonfilename = fi.Name
            jsonfilenameLocation = fi.DirectoryName
            Dim jsonString As String = "{'results':" + File.ReadAllText(jsonfilenameLocation + "\" + jsonfilename) + "} "
            Dim json As JObject = JObject.Parse(jsonString)

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Interface EPIC1 - Supplier Management ... file name ==> " & jsonfilename & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            ' Variable Get data From JSon Parse
            Dim JsonSupplier_Code As JToken
            Dim JsonSupplier_Name As JToken
            Dim JsonAddress As JToken
            Dim JsonPhone As JToken
            Dim JsonFax As JToken
            Dim JsonCountry As JToken
            Dim JsonPresident_Director As JToken
            Dim JsonPresident_Director_Email As JToken
            Dim JsonVice_President_Director As JToken
            Dim JsonVice_President_Director_Email As JToken
            Dim JsonMarketing_Director As JToken
            Dim JsonMarketing_Director_Email As JToken
            Dim JsonMarketing_General_Manager As JToken
            Dim JsonMarketing_General_Manager_Email As JToken
            Dim JsonMarketing_Manager As JToken
            Dim JsonMarketing_Manager_Email As JToken
            Dim JsonPlant_Director As JToken
            Dim JsonPlant_Director_Email As JToken
            Dim JsonPlant_Manager As JToken
            Dim JsonPlant_Manager_Email As JToken
            Dim JsonMarketing_Sales As JToken
            Dim JsonMarketing_Sales_Email As JToken
            Dim JsonEmail As JToken
            Dim JsonPeriod_Code As JToken
            Dim JsonSupplierType As JToken
            Dim JsonCertification As JToken
            Dim JsonQualityAudit_Score As JToken
            Dim JsonDeliveryAudit_Score As JToken
            Dim JsonEHS_Value As JToken
            Dim JsonPPM As JToken
            Dim JsonMother_Company As JToken
            Dim JsonProduct_Name As JToken
            Dim JsonOEM_Customer As JToken

            ' Variable Data Convert from json
            Dim data_Supplier_Code As String
            Dim data_Supplier_Name As String
            Dim data_Address As String
            Dim data_Phone As String
            Dim data_Fax As String
            Dim data_Country As String
            Dim data_President_Director As String
            Dim data_President_Director_Email As String
            Dim data_Vice_President_Director As String
            Dim data_Vice_President_Director_Email As String
            Dim data_Marketing_Director As String
            Dim data_Marketing_Director_Email As String
            Dim data_Marketing_General_Manager As String
            Dim data_Marketing_General_Manager_Email As String
            Dim data_Marketing_Manager As String
            Dim data_Marketing_Manager_Email As String
            Dim data_Plant_Director As String
            Dim data_Plant_Director_Email As String
            Dim data_Plant_Manager As String
            Dim data_Plant_Manager_Email As String
            Dim data_Marketing_Sales As String
            Dim data_Marketing_Sales_Email As String
            Dim data_Email As String
            Dim data_Period_Code As String
            Dim data_SupplierType As String
            Dim data_Certification As String
            Dim data_QualityAudit_Score As String
            Dim data_DeliveryAudit_Score As String
            Dim data_EHS_Value As String
            Dim data_PPM As String
            Dim data_Mother_Company As String
            Dim data_Product_Name As String
            Dim data_OEM_Customer As String

            For Each Row In json("results").ToList()
                JsonSupplier_Code = Row("Supplier_Code")
                JsonSupplier_Name = Row("Supplier_Name")
                JsonAddress = Row("Address")
                JsonPhone = Row("Phone")
                JsonFax = Row("Fax")
                JsonCountry = Row("Country")
                JsonPresident_Director = Row("President_Director")
                JsonPresident_Director_Email = Row("President_Director_Email")
                JsonVice_President_Director = Row("Vice_President_Director")
                JsonVice_President_Director_Email = Row("Vice_President_Director_Email")
                JsonMarketing_Director = Row("Marketing_Director")
                JsonMarketing_Director_Email = Row("Marketing_Director_Email")
                JsonMarketing_General_Manager = Row("Marketing_General_Manager")
                JsonMarketing_General_Manager_Email = Row("Marketing_General_Manager_Email")
                JsonMarketing_Manager = Row("Marketing_Manager")
                JsonMarketing_Manager_Email = Row("Marketing_Manager_Email")
                JsonPlant_Director = Row("Plant_Director")
                JsonPlant_Director_Email = Row("Plant_Director_Email")
                JsonPlant_Manager = Row("Plant_Manager")
                JsonPlant_Manager_Email = Row("Plant_Manager_Email")
                JsonMarketing_Sales = Row("Marketing_Sales")
                JsonMarketing_Sales_Email = Row("Marketing_Sales_Email")
                JsonEmail = Row("Email")
                JsonPeriod_Code = Row("Period_Code")
                JsonSupplierType = Row("SupplierType")
                JsonCertification = Row("Certification")
                JsonQualityAudit_Score = Row("QualityAudit_Score")
                JsonDeliveryAudit_Score = Row("DeliveryAudit_Score")
                JsonEHS_Value = Row("EHS_Value")
                JsonPPM = Row("PPM")
                JsonMother_Company = Row("Mother_Company")
                JsonProduct_Name = Row("Product_Name")
                JsonOEM_Customer = Row("OEM_Customer")

                data_Supplier_Code = IIf(String.IsNullOrEmpty(DirectCast(JsonSupplier_Code, JValue).Value), "", DirectCast(JsonSupplier_Code, JValue).Value)
                data_Supplier_Name = IIf(String.IsNullOrEmpty(DirectCast(JsonSupplier_Name, JValue).Value), "", DirectCast(JsonSupplier_Name, JValue).Value)
                data_Address = IIf(String.IsNullOrEmpty(DirectCast(JsonAddress, JValue).Value), "", DirectCast(JsonAddress, JValue).Value)
                data_Phone = IIf(String.IsNullOrEmpty(DirectCast(JsonPhone, JValue).Value), "", DirectCast(JsonPhone, JValue).Value)
                data_Fax = IIf(String.IsNullOrEmpty(DirectCast(JsonFax, JValue).Value), "", DirectCast(JsonFax, JValue).Value)
                data_Country = IIf(String.IsNullOrEmpty(DirectCast(JsonCountry, JValue).Value), "", DirectCast(JsonCountry, JValue).Value)
                data_President_Director = IIf(String.IsNullOrEmpty(DirectCast(JsonPresident_Director, JValue).Value), "", DirectCast(JsonPresident_Director, JValue).Value)
                data_President_Director_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonPresident_Director_Email, JValue).Value), "", DirectCast(JsonPresident_Director_Email, JValue).Value)
                data_Vice_President_Director = IIf(String.IsNullOrEmpty(DirectCast(JsonVice_President_Director, JValue).Value), "", DirectCast(JsonVice_President_Director, JValue).Value)
                data_Vice_President_Director_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonVice_President_Director_Email, JValue).Value), "", DirectCast(JsonVice_President_Director_Email, JValue).Value)
                data_Marketing_Director = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Director, JValue).Value), "", DirectCast(JsonMarketing_Director, JValue).Value)
                data_Marketing_Director_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Director_Email, JValue).Value), "", DirectCast(JsonMarketing_Director_Email, JValue).Value)
                data_Marketing_General_Manager = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_General_Manager, JValue).Value), "", DirectCast(JsonMarketing_General_Manager, JValue).Value)
                data_Marketing_General_Manager_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_General_Manager_Email, JValue).Value), "", DirectCast(JsonMarketing_General_Manager_Email, JValue).Value)
                data_Marketing_Manager = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Manager, JValue).Value), "", DirectCast(JsonMarketing_Manager, JValue).Value)
                data_Marketing_Manager_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Manager_Email, JValue).Value), "", DirectCast(JsonMarketing_Manager_Email, JValue).Value)
                data_Plant_Director = IIf(String.IsNullOrEmpty(DirectCast(JsonPlant_Director, JValue).Value), "", DirectCast(JsonPlant_Director, JValue).Value)
                data_Plant_Director_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonPlant_Director_Email, JValue).Value), "", DirectCast(JsonPlant_Director_Email, JValue).Value)
                data_Plant_Manager = IIf(String.IsNullOrEmpty(DirectCast(JsonPlant_Manager, JValue).Value), "", DirectCast(JsonPlant_Manager, JValue).Value)
                data_Plant_Manager_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonPlant_Manager_Email, JValue).Value), "", DirectCast(JsonPlant_Manager_Email, JValue).Value)
                data_Marketing_Sales = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Sales, JValue).Value), "", DirectCast(JsonMarketing_Sales, JValue).Value)
                data_Marketing_Sales_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonMarketing_Sales_Email, JValue).Value), "", DirectCast(JsonMarketing_Sales_Email, JValue).Value)
                data_Email = IIf(String.IsNullOrEmpty(DirectCast(JsonEmail, JValue).Value), "", DirectCast(JsonEmail, JValue).Value)
                data_Period_Code = IIf(String.IsNullOrEmpty(DirectCast(JsonPeriod_Code, JValue).Value), "", DirectCast(JsonPeriod_Code, JValue).Value)
                data_SupplierType = IIf(String.IsNullOrEmpty(DirectCast(JsonSupplierType, JValue).Value), "", DirectCast(JsonSupplierType, JValue).Value)
                data_Certification = IIf(String.IsNullOrEmpty(DirectCast(JsonCertification, JValue).Value), "", DirectCast(JsonCertification, JValue).Value)
                data_QualityAudit_Score = IIf(String.IsNullOrEmpty(DirectCast(JsonQualityAudit_Score, JValue).Value), "0", DirectCast(JsonQualityAudit_Score, JValue).Value)
                data_DeliveryAudit_Score = IIf(String.IsNullOrEmpty(DirectCast(JsonDeliveryAudit_Score, JValue).Value), "0", DirectCast(JsonDeliveryAudit_Score, JValue).Value)
                data_EHS_Value = IIf(String.IsNullOrEmpty(DirectCast(JsonEHS_Value, JValue).Value), "0", DirectCast(JsonEHS_Value, JValue).Value)
                data_PPM = IIf(String.IsNullOrEmpty(DirectCast(JsonPPM, JValue).Value), "0", DirectCast(JsonPPM, JValue).Value)
                data_Mother_Company = IIf(String.IsNullOrEmpty(DirectCast(JsonMother_Company, JValue).Value), "", DirectCast(JsonMother_Company, JValue).Value)
                data_Product_Name = IIf(String.IsNullOrEmpty(DirectCast(JsonProduct_Name, JValue).Value), "", DirectCast(JsonProduct_Name, JValue).Value)
                data_OEM_Customer = IIf(String.IsNullOrEmpty(DirectCast(JsonOEM_Customer, JValue).Value), "", DirectCast(JsonOEM_Customer, JValue).Value)


                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload EPIC1 - Supplier Management ==> Supplier :  " & data_Supplier_Name & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update Data Master Supplier
                EPIC1SupplierTransaction(ConStr, data_Supplier_Code, data_Supplier_Name, data_Address, data_Phone, data_Fax, data_Country, data_President_Director, data_President_Director_Email, data_Vice_President_Director, data_Vice_President_Director_Email, data_Marketing_Director, data_Marketing_Director_Email, data_Marketing_General_Manager, data_Marketing_General_Manager_Email, data_Marketing_Manager, data_Marketing_Manager_Email, data_Plant_Director, data_Plant_Director_Email, data_Plant_Manager, data_Plant_Manager_Email, data_Marketing_Sales, data_Marketing_Sales_Email, data_Email, data_Period_Code, data_SupplierType, data_Certification, data_QualityAudit_Score, data_DeliveryAudit_Score, data_EHS_Value, data_PPM, data_Mother_Company, data_Product_Name, data_OEM_Customer)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload EPIC1 - Supplier Management ==> Supplier :  " & data_Supplier_Name & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next
            'MOVE FILE
            up_JSONFile_Copy(gs_JSONPath & "\", gs_JSONPathBackup & "\", jsonfilename)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface EPIC1 - Supplier Management ... file name ==> " & jsonfilename & " and move file " & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

        Next
    End Sub

    Private Sub up_Interface_ESS_Budget_Process()
        ConStr = Builder.ConnectionString
        Dim dtInterface_Setting As DataTable
        dtInterface_Setting = ClsInterfaceSettingDB.getData(ConStr)

        If dtInterface_Setting.Rows.Count > 0 Then
            gs_XMLPath = Trim(dtInterface_Setting.Rows(0).Item("XML_IABudget_Path"))
            gs_XMLPathBackup = Trim(dtInterface_Setting.Rows(0).Item("XML_IABudget_PathBackup"))
        Else
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for ESS - Budget " & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Exit Sub
        End If

        Dim di As New IO.DirectoryInfo(gs_XMLPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            xmlfilename = fi.Name
            xmlfilenameLocation = fi.DirectoryName

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Interface ESS - Budget ... file name ==> " & xmlfilename & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)
            Dim IA_Number As String
            Dim IA_date As String
            Dim Departemen_ID As String
            Dim Departemen_Name As String
            Dim Cost_Center As String
            Dim IO_Number As String
            Dim GL_Account As String
            Dim NPK As String
            Dim Employee_Name As String
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("IA_OUTPUT").Elements("Item")
            For Each XEL1 As XElement In XMLMaterials

                IA_Number = XEL1.Element("IA_Number").Value
                IA_date = XEL1.Element("IA_date").Value
                Departemen_ID = XEL1.Element("Departemen_ID").Value
                Departemen_Name = XEL1.Element("Departemen_Name").Value
                Cost_Center = XEL1.Element("Cost_Center").Value
                IO_Number = XEL1.Element("IO_Number").Value
                GL_Account = XEL1.Element("GL_Account").Value
                NPK = XEL1.Element("NPK").Value
                Employee_Name = XEL1.Element("Employee_Name").Value

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload ESS - Budget ==> IA Number :  " & IA_Number & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update ESS - Budget
                ESSBudgetTransaction(ConStr, IA_Number, IA_date, Departemen_ID, Departemen_Name, Cost_Center, IO_Number, GL_Account, NPK, Employee_Name)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload ESS - Budget ==> IA Number :  " & IA_Number & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next

            'MOVE FILE
            up_XMLFile_Move(gs_XMLPath & "\", gs_XMLPathBackup & "\", xmlfilename)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface ESS - Budget ... file name ==> " & xmlfilename & " and move file " & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
        Next
    End Sub

    Private Sub up_Interface_SAP_Process()
        ConStr = Builder.ConnectionString
        Dim dtInterface_Setting As DataTable
        dtInterface_Setting = ClsInterfaceSettingDB.getData(ConStr)

        If dtInterface_Setting.Rows.Count > 0 Then
            gs_XMLPath = Trim(dtInterface_Setting.Rows(0).Item("XML_SAP_Path"))
            gs_XMLPathBackup = Trim(dtInterface_Setting.Rows(0).Item("XML_SAP_PathBackup"))
        Else
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for SAP " & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Exit Sub
        End If

        Dim di As New IO.DirectoryInfo(gs_XMLPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            xmlfilename = fi.Name
            xmlfilenameLocation = fi.DirectoryName

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Interface SAP ... file name ==> " & xmlfilename & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)
            '------- Upload SAP PO -------------'

            Dim PONumber As String
            Dim PODate As String
            Dim IABudgetNo As String
            Dim PRNumber As String
            Dim ItemCode As String
            Dim Qty As String
            Dim Currency As String
            Dim POPrice As String
            Dim POAmount As String
            Dim CreateDate As String
            Dim Createdby As String

            Dim XMLPO As IEnumerable(Of XElement) = doc.Root.Elements("PO_OUTPUT").Elements("Item")
            For Each XEL1 As XElement In XMLPO
                PONumber = XEL1.Element("PONumber").Value
                PODate = XEL1.Element("PODate").Value
                IABudgetNo = XEL1.Element("IABudgetNo").Value
                PRNumber = XEL1.Element("PRNumber").Value
                ItemCode = XEL1.Element("ItemCode").Value
                Qty = XEL1.Element("Qty").Value
                Currency = XEL1.Element("Currency").Value
                POPrice = XEL1.Element("POPrice").Value
                POAmount = XEL1.Element("POAmount").Value
                CreateDate = XEL1.Element("CreateDate").Value
                Createdby = XEL1.Element("Createdby").Value

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload SAP - Send PO ==> PO Number :  " & PONumber & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update ESS - Budget
                SAPPOTransaction(ConStr, PONumber, PODate, IABudgetNo, PRNumber, ItemCode, Qty, Currency, POPrice, POAmount, CreateDate, Createdby)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload SAP - Send PO ==> PO Number :  " & PONumber & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next

            '------------------- end Process Upload SAP PO -------------------------'

            '------- Upload SAP ITEM -------------'
            Dim Material_No As String
            Dim SAP_No As String
            Dim Description As String
            Dim UOM As String

            Dim XMLITEM As IEnumerable(Of XElement) = doc.Root.Elements("FT_OUTPUT").Elements("Item")
            For Each XEL1 As XElement In XMLITEM

                Material_No = XEL1.Element("Material_No").Value
                SAP_No = XEL1.Element("SAP_No").Value
                Description = XEL1.Element("Description").Value
                UOM = XEL1.Element("UOM").Value


                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload SAP - Item ==> Material Number :  " & Material_No & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update ESS - Budget
                SAP_Item_Transaction(ConStr, Material_No, SAP_No, Description, UOM)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload SAP - Item ==> Material Number :  " & Material_No & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next

            '------------------- end Process Upload SAP ITEM -------------------------'

            '------- Upload SAP GR -------------'
            Dim ReceivingNo As String
            Dim ReceivingDate As String
            Dim PONo As String
            Dim Item_Code As String
            Dim Qty_Item As String
            Dim Status As String
            Dim CreateUser As String

            Dim XMLGR As IEnumerable(Of XElement) = doc.Root.Elements("GR_OUTPUT").Elements("Item")
            For Each XEL1 As XElement In XMLGR

                ReceivingNo = XEL1.Element("RecivingNumber").Value
                ReceivingDate = XEL1.Element("ReceivingDate").Value
                PONo = XEL1.Element("PONumber").Value
                Item_Code = XEL1.Element("ItemCode").Value
                Qty_Item = XEL1.Element("Qty").Value
                Status = XEL1.Element("Status").Value
                CreateDate = XEL1.Element("CreateDate").Value
                CreateUser = XEL1.Element("CreateUser").Value

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload SAP - GR ==> Receiving Number :  " & ReceivingNo & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update ESS - Budget
                SAP_GR_Transaction(ConStr, ReceivingNo, ReceivingDate, PONo, Item_Code, Qty_Item, Status, CreateDate, CreateUser)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload SAP - GR ==> Receiving Number :  " & ReceivingNo & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next
            '------- end Process Upload SAP GR -------------'

            'MOVE FILE
            up_XMLFile_Move(gs_XMLPath & "\", gs_XMLPathBackup & "\", xmlfilename)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface SAP ... file name ==> " & xmlfilename & " and move file " & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
        Next
    End Sub

    Private Sub up_Interface_IAMI_PressPart_LocalComponent_Process()
        ConStr = Builder.ConnectionString
        Dim dt As DataTable
        dt = ClsInterfaceSettingDB.getData(ConStr)

        If dt.Rows.Count > 0 Then
            gs_IAMI_PressPartPath = Trim(dt.Rows(0).Item("IAMI_PressPart_Path"))
            gs_IAMI_PressPartPathBackup = Trim(dt.Rows(0).Item("IAMI_PressPart_PathBackup"))
        Else
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for IAMI Press Part & Local Component" & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Exit Sub
        End If

        Dim di As New IO.DirectoryInfo(gs_IAMI_PressPartPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            xmlfilename = fi.Name
            xmlfilenameLocation = fi.DirectoryName

            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Interface IAMI Press Part & Local Component ... file name ==> " & xmlfilename & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)

            Dim year As String
            Dim month As String
            Dim materialNoFinishedUnit As String
            Dim substitutionOldVariant As String
            Dim pressPart As String
            Dim localComponent As String
            Dim messageError As String

            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_OUTPUT").Elements("Item")
            For Each XEL1 As XElement In XMLMaterials

                year = XEL1.Element("Year").Value
                month = XEL1.Element("Month").Value
                materialNoFinishedUnit = XEL1.Element("MaterialFinishUnit").Value
                substitutionOldVariant = XEL1.Element("SubstitutionOldVariant").Value
                pressPart = XEL1.Element("PressPart").Value
                localComponent = XEL1.Element("LocalComponent").Value
                messageError = XEL1.Element("MessageError").Value

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload IAMI Press Part & Local Component ==> Year : " & year & ", Month : " & month & ", Material No. Finished Unit : " & materialNoFinishedUnit & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                ' Process Insert or Update IAMI PressPart & LocalComponent XML From SAP
                update_IF_IAMI_PressPartLocalComponent_XMLFromSAP(ConStr, year, month, materialNoFinishedUnit, substitutionOldVariant, pressPart, localComponent, messageError)

                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload IAMI Press Part & Local Component ==> Year : " & year & ", Month : " & month & ", Material No. Finished Unit : " & materialNoFinishedUnit & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
            Next

            'MOVE FILE
            up_XMLFile_Move(gs_IAMI_PressPartPath & "\", gs_IAMI_PressPartPathBackup & "\", xmlfilename)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface IAMI Press Part & Local Component ... file name ==> " & xmlfilename & " and move file " & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
        Next
    End Sub

    Private Function generateXMLPR_Response(ByVal pFileName As String, ByVal pFolderOut As String) As Boolean
        Dim iStatus As Boolean = True
        Try
            Dim sPRNo As String = ""
            Dim sStatus As String = ""
            Dim sErrorMessage As String = ""
            Dim sProcessCode As String = ""
            Dim iRow As Integer

            Dim statusSend As Boolean = False

            Dim dt As DataTable
            dt = getData_IF_PR_DMMS_RESP(ConStr, pFileName)

            If dt.Rows.Count > 0 Then
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] START Interface Generate PR Response!" & vbCrLf
                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                Dim path As String = pFolderOut + "\" + pFileName.Split(".")(0).ToString + "." + pFileName.Split(".")(1).ToString + "." + pFileName.Split(".")(2).ToString + ".RESP.xml"

                Dim writer As New XmlTextWriter(path, New UpperCaseUTF8Encoding())

                writer.WriteStartDocument()
                writer.Indentation = 2
                writer.WriteStartElement("ns0:MT_DMMS_PR_Resp")
                writer.WriteStartAttribute("xmlns:ns0")
                writer.WriteString("http://iamiepic")
                writer.WriteStartElement("Table")

                For iRow = 0 To dt.Rows.Count - 1
                    sPRNo = Trim(dt.Rows(iRow).Item("PR_Number") & "")
                    sStatus = Trim(dt.Rows(iRow).Item("Status") & "")
                    sErrorMessage = Trim(dt.Rows(iRow).Item("ErrorMessage") & "")
                    sProcessCode = CDate(dt.Rows(iRow).Item("ProcessDate")).ToString("yyyy-MM-dd HH:mm:ss")

                    createNodePRResponse(sPRNo, sStatus, sErrorMessage, sProcessCode, writer)
                Next iRow

                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()

                ' Log for Success Sending Email 
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Interface Generate PR Response Success !" & vbCrLf
                up_gridProcess(Rtb1, 3, 0, ls_StepProcess)

                statusSend = True
            End If

            If statusSend = False Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] No Data PR Response to interface ! " & vbCrLf
                up_gridProcess(Rtb1, 6, 0, ls_StepProcess)
                iStatus = False
                Return iStatus
            End If

        Catch ex As Exception
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error Interface Generate PR Response... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            iStatus = False
            Return iStatus
        End Try

        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] END Interface Generate PR Response!" & vbCrLf
        up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

        Return iStatus
    End Function

    Private Sub up_Interface_PurchaseRequest_Process()
        Try

            ConStr = Builder.ConnectionString
            Dim dt As DataTable
            dt = ClsInterfaceSettingDB.getData(ConStr)

            If dt.Rows.Count > 0 Then
                gs_PurchaseRequestPath = Trim(dt.Rows(0).Item("XML_PurchaseRequest_Path"))
                gs_PurchaseRequestPathBackup = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathBackup"))
                gs_PurchaseRequestPathFailed = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathFailed"))
                gs_PurchaseRequestPathOut = Trim(dt.Rows(0).Item("XML_PurchaseRequest_PathOut"))
            Else
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for Purchase Request" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Exit Sub
            End If
            dt.Reset()

            If gs_PurchaseRequestPath = "" And gs_PurchaseRequestPathBackup = "" And gs_PurchaseRequestPathFailed = "" And gs_PurchaseRequestPathOut = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please Setting Interface for Purchase Request" & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Exit Sub
            End If

            Dim di As New IO.DirectoryInfo(gs_PurchaseRequestPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
            Dim fi As IO.FileInfo
            Dim jmlFile As Integer = aryFi.Length
            Dim respStatus As Boolean = True

            If jmlFile > 0 Then
                'delete IF PR DMMS RESP
                delete_IF_PR_DMMS_RESP(ConStr)

                'delete temporary data IF PR XML From SAP yang lebih dari 30 hari
                delete_IF_PR_XMLFromDMMS_ProcessDate(ConStr)

                Dim errorStatus As Boolean
                Dim errorMessage As String
                For Each fi In aryFi
                    xmlfilename = fi.Name
                    xmlfilenameLocation = fi.DirectoryName

                    ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Interface DMMS Purchase Request ... file name ==> " & xmlfilename & vbCrLf
                    up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                    'check IF History --> SystemFrom DMMS, Module PR 
                    If check_DMMS_PR_IF_History(ConStr, xmlfilename) = True Then
                        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... file name ==> " & xmlfilename & " is already exist " & vbCrLf
                        up_gridProcess(Rtb1, 2, 0, ls_StepProcess)

                        'MOVE FAILED FILE
                        up_XMLFile_Move(gs_PurchaseRequestPath & "\", gs_PurchaseRequestPathFailed & "\", xmlfilename)
                        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface DMMS Purchase Request ... file name ==> " & xmlfilename & " and move file " & vbCrLf
                        up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
                    Else
                        Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)

                        Dim sFileName As String = ""
                        Dim sPRNumber As String = ""
                        Dim sPRDate As String = ""
                        Dim sBudgetCode As String = ""
                        Dim sPRTypeCode As String = ""
                        Dim sDepartmentCode As String = ""
                        Dim sSectionCode As String = ""
                        Dim sCostCenter As String = ""
                        Dim sProject As String = ""
                        Dim sUrgentStatus As String = ""
                        Dim sUrgentNote As String = ""
                        Dim sReqPOIssueDate As String = ""
                        Dim sMaterialNo As String = ""
                        Dim sMaterialNoDMMS As String = ""
                        Dim sQty As String = ""
                        Dim sRemarks As String = ""
                        Dim sUpdateUser As String = ""
                        Dim sUpdateDate As String = ""

                        errorStatus = False
                        errorMessage = ""

                        Dim XMLHeader As IEnumerable(Of XElement) = doc.Root.Elements("Table").Elements("Item")
                        For Each XEHeader As XElement In XMLHeader
                            sFileName = xmlfilename
                            sPRNumber = XEHeader.Element("Code").Value
                            sPRDate = XEHeader.Element("Date").Value
                            sBudgetCode = XEHeader.Element("Budget").Value
                            sPRTypeCode = XEHeader.Element("Type").Value
                            sDepartmentCode = XEHeader.Element("Department").Value
                            sSectionCode = XEHeader.Element("Section").Value
                            sCostCenter = XEHeader.Element("CostCenter").Value
                            sProject = XEHeader.Element("Project").Value
                            sUrgentStatus = XEHeader.Element("Urgent").Value
                            sUrgentNote = XEHeader.Element("UrgentNote").Value
                            sReqPOIssueDate = XEHeader.Element("IssueDate").Value
                            sMaterialNo = XEHeader.Element("MaterialNo").Value
                            sMaterialNoDMMS = XEHeader.Element("MaterialNoDMMS").Value
                            sQty = XEHeader.Element("Qty").Value
                            sRemarks = XEHeader.Element("Remarks").Value
                            sUpdateUser = XEHeader.Element("User").Value
                            sUpdateDate = XEHeader.Element("Updated").Value

                            '#1 Validate - PR Number
                            If sPRNumber = "" Then
                                errorStatus = True
                                errorMessage = "PR Number is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#2 Validate - PR Date
                            If sPRDate = "" Then
                                errorStatus = True
                                errorMessage = "PR Date is invalid!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            Else
                                Dim tempDate As String
                                tempDate = Mid(sPRDate, 1, 10)
                                If checkDate(tempDate) = False Then
                                    errorStatus = True
                                    errorMessage = "PR Date is invalid!"
                                    ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                    up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                    Exit For
                                End If
                            End If

                            '#3 Validate - Budget Code
                            If sBudgetCode = "" Then
                                errorStatus = True
                                errorMessage = "Budget Code is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#4 Validate - PR Type Code
                            If sPRTypeCode = "" Then
                                errorStatus = True
                                errorMessage = "PR Type Code is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#5 Validate - Department Code
                            If sDepartmentCode = "" Then
                                errorStatus = True
                                errorMessage = "Department Code is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#6 Validate - Cost Center
                            If sCostCenter = "" Then
                                errorStatus = True
                                errorMessage = "Cost Center is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#7 Validate - Material No.
                            If sMaterialNo = "" Then
                                errorStatus = True
                                errorMessage = "Material No. is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#8 Validate - Material No. DMMS
                            If sMaterialNoDMMS = "" Then
                                errorStatus = True
                                errorMessage = "Material No. DMMS is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#9 Validate - Qty
                            If sQty = "" Then
                                errorStatus = True
                                errorMessage = "Qty is invalid!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            Else
                                '#10 Validate - Qty
                                If CInt(sQty) = 0 Then
                                    errorStatus = True
                                    errorMessage = "Qty is invalid!"
                                    ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                    up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                    Exit For
                                End If
                            End If

                            '#11 Validate - Update User
                            If sUpdateUser = "" Then
                                errorStatus = True
                                errorMessage = "User is empty!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#12 Validate - Update Date
                            If sUpdateDate = "" Then
                                errorStatus = True
                                errorMessage = "Update Date is invalid!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            Else
                                Dim tempDate As String
                                Dim tempTime As String
                                tempDate = Mid(sUpdateDate, 1, 10)
                                tempTime = Mid(sUpdateDate, 12, 8)
                                If checkDate(tempDate + " " + tempTime) = False Then
                                    errorStatus = True
                                    errorMessage = "Update Date is invalid!"
                                    ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                    up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                    Exit For
                                End If
                            End If

                            '#13 Validate - Check existing PR Number
                            If check_PR_Header(ConStr, sPRNumber) = True Then
                                errorStatus = True
                                errorMessage = "PR Number " & sPRNumber & " is already registered!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#14 Validate - Check existing Item Material
                            If check_Mst_Item(ConStr, sMaterialNo) = False Then
                                errorStatus = True
                                errorMessage = "Material No. " & sMaterialNo & " is not registered in Item Master!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            '#15 Validate - Check existing user id
                            If check_UserSetup(ConStr, sUpdateUser) = False Then
                                errorStatus = True
                                errorMessage = "User Id " & sUpdateUser & " is not registered in User Setup!"
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... " & errorMessage & ", file name ==> " & xmlfilename & " " & vbCrLf
                                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                                Exit For
                            End If

                            'Insert XML From DMMS to Temporary
                            insert_IF_PR_XMLFromDMMS(ConStr, sFileName, sPRNumber, Mid(sPRDate, 1, 10), sBudgetCode, sPRTypeCode, sDepartmentCode, _
                                                     sSectionCode, sCostCenter, sProject, sUrgentStatus, sUrgentNote, Mid(sReqPOIssueDate, 1, 10), _
                                                     sMaterialNo, sMaterialNoDMMS, sQty, sRemarks, Mid(sUpdateDate, 1, 10) + " " + Mid(sUpdateDate, 12, 8), sUpdateUser)
                        Next

                        If errorStatus = True Then
                            delete_IF_PR_XMLFromDMMS(ConStr, xmlfilename)
                            update_DMMS_PR_IF_History(ConStr, xmlfilename, "NG", errorMessage)
                            insert_DMMS_PR_IF_RESP(ConStr, xmlfilename, sPRNumber, "NG", errorMessage)

                            'MOVE FAILED FILE
                            up_XMLFile_Move(gs_PurchaseRequestPath & "\", gs_PurchaseRequestPathFailed & "\", xmlfilename)
                            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface DMMS Purchase Request ... file name ==> " & xmlfilename & " and move file " & vbCrLf
                            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                            'PR DMMS Reponse
                            generateXMLPR_Response(xmlfilename, gs_PurchaseRequestPathOut)

                            'Send Email
                            sendEmail_DMMS_PR(xmlfilename)
                        Else
                            errorStatus = False
                            errorMessage = ""

                            Try
                                'Get PR Number from XML DMMS Temporary
                                Dim dtPR As DataTable
                                Dim dtPRDetail As DataTable
                                Dim iRow As Integer
                                Dim iRowDetail As Integer
                                dtPR = getData_IF_PR_XMLFromDMMS(ConStr, xmlfilename)
                                If dtPR.Rows.Count > 0 Then
                                    For iRow = 0 To dtPR.Rows.Count - 1
                                        sPRNumber = dtPR.Rows(iRow)("PR_Number").ToString

                                        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start upload Purchase Request ==> PR Number : " & sPRNumber & vbCrLf
                                        up_gridProcess(Rtb1, 3, 0, ls_StepProcess)

                                        dtPRDetail = getData_IF_PR_XMLFromDMMS_Detail(ConStr, xmlfilename, sPRNumber)

                                        Dim tempPRNumber As String = ""
                                        For iRowDetail = 0 To dtPRDetail.Rows.Count - 1
                                            sPRNumber = dtPRDetail.Rows(iRowDetail)("PR_Number").ToString
                                            sPRDate = CDate(dtPRDetail.Rows(iRowDetail)("PR_Date")).ToString("yyyy-MM-dd")
                                            sBudgetCode = dtPRDetail.Rows(iRowDetail)("Budget_Code").ToString
                                            sPRTypeCode = dtPRDetail.Rows(iRowDetail)("PRType_Code").ToString
                                            sDepartmentCode = dtPRDetail.Rows(iRowDetail)("Department_Code").ToString
                                            sSectionCode = dtPRDetail.Rows(iRowDetail)("Section_Code").ToString
                                            sCostCenter = dtPRDetail.Rows(iRowDetail)("CostCenter").ToString
                                            sProject = dtPRDetail.Rows(iRowDetail)("Project").ToString
                                            sUrgentStatus = dtPRDetail.Rows(iRowDetail)("Urgent_Status").ToString
                                            sUrgentNote = dtPRDetail.Rows(iRowDetail)("Urgent_Note").ToString
                                            sReqPOIssueDate = CDate(dtPRDetail.Rows(iRowDetail)("Req_POIssueDate")).ToString("yyyy-MM-dd")
                                            sMaterialNo = dtPRDetail.Rows(iRowDetail)("Material_No").ToString
                                            sMaterialNoDMMS = dtPRDetail.Rows(iRowDetail)("Material_No_DMMS").ToString
                                            sQty = dtPRDetail.Rows(iRowDetail)("Qty").ToString
                                            sRemarks = dtPRDetail.Rows(iRowDetail)("Remarks").ToString
                                            sUpdateDate = CDate(dtPRDetail.Rows(iRowDetail)("UpdateDate")).ToString("yyyy-MM-dd HH:mm:ss")
                                            sUpdateUser = dtPRDetail.Rows(iRowDetail)("UpdateUser").ToString

                                            If tempPRNumber <> sPRNumber Then
                                                tempPRNumber = dtPRDetail.Rows(iRowDetail)("PR_Number").ToString
                                                'insert header
                                                insert_PR_Header(ConStr, sPRNumber, sPRDate, sBudgetCode, sPRTypeCode, sDepartmentCode, _
                                                                sSectionCode, sCostCenter, sProject, sUrgentStatus, sUrgentNote, sReqPOIssueDate, _
                                                                sUpdateDate, sUpdateUser)

                                                'insert detail
                                                insert_PR_Detail(ConStr, sPRNumber, sMaterialNo, sMaterialNoDMMS, sQty, sRemarks, sUpdateDate, sUpdateUser)

                                                'insert approval
                                                insert_PR_Approval(ConStr, sPRNumber, sDepartmentCode, sUrgentStatus)

                                                'update IF PR XML From DMMS
                                                update_IF_PR_XMLFromDMMS_Process(ConStr, xmlfilename, sPRNumber, sMaterialNo)

                                                'insert IF DMMS PR
                                                insert_DMMS_PR_IF_RESP(ConStr, xmlfilename, sPRNumber, "OK", errorMessage)
                                            Else
                                                'insert detail
                                                insert_PR_Detail(ConStr, sPRNumber, sMaterialNo, sMaterialNoDMMS, sQty, sRemarks, sUpdateDate, sUpdateUser)

                                                'update IF PR XML From DMMS
                                                update_IF_PR_XMLFromDMMS_Process(ConStr, xmlfilename, sPRNumber, sMaterialNo)
                                            End If
                                        Next

                                        ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End upload Purchase Request ==> PR Number : " & sPRNumber & vbCrLf
                                        up_gridProcess(Rtb1, 3, 0, ls_StepProcess)
                                    Next
                                End If
                                dtPR.Reset()
                            Catch ex As Exception
                                errorStatus = True
                                errorMessage = "up_Interface_PurchaseRequest_Process --> Error: " & ex.Message
                            End Try

                            If errorStatus = True Then
                                delete_IF_PR_XMLFromDMMS_All(ConStr, xmlfilename)
                                delete_IF_PR_XMLFromDMMS(ConStr, xmlfilename)
                                update_DMMS_PR_IF_History(ConStr, xmlfilename, "NG", errorMessage)
                                insert_DMMS_PR_IF_RESP(ConStr, xmlfilename, sPRNumber, "NG", errorMessage)

                                'MOVE FAILED FILE
                                up_XMLFile_Move(gs_PurchaseRequestPath & "\", gs_PurchaseRequestPathFailed & "\", xmlfilename)
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface DMMS Purchase Request ... file name ==> " & xmlfilename & " and move file " & vbCrLf
                                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                                'PR DMMS Reponse
                                generateXMLPR_Response(xmlfilename, gs_PurchaseRequestPathOut)

                                'Send Email
                                sendEmail_DMMS_PR(xmlfilename)
                            Else
                                update_DMMS_PR_IF_History(ConStr, xmlfilename, "OK", errorMessage)

                                'MOVE BACKUP FILE
                                up_XMLFile_Move(gs_PurchaseRequestPath & "\", gs_PurchaseRequestPathBackup & "\", xmlfilename)
                                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface DMMS Purchase Request ... file name ==> " & xmlfilename & " and move file " & vbCrLf
                                up_gridProcess(Rtb1, 1, 0, ls_StepProcess)

                                'PR DMMS Reponse
                                generateXMLPR_Response(xmlfilename, gs_PurchaseRequestPathOut)

                                'Send Email
                                sendEmail_DMMS_PR(xmlfilename)
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End process Interface DMMS Purchase Request ... " & ex.Message() & vbCrLf
            up_gridProcess(Rtb1, 1, 0, ls_StepProcess)
        End Try

    End Sub

    Private Shared Function customCertValidation(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Function FormatMultipleEmailAddresses(ByVal emailAddresses As String) As String
        Dim delimiters = {","c, ";"c}
        Dim addresses = emailAddresses.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join(",", addresses)
    End Function

    Private Function sendEmail_DMMS_PR(ByVal pFileName As String) As Boolean
        Try
            Dim sPRNo As String = ""
            Dim sStatus As String = ""
            Dim sErrorMessage As String = ""
            Dim sProcessDate As String = ""

            Dim statusSend As Boolean = False
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
                gs_EmailFrom_DMMS_PR = ""
                gs_EmailTo_DMMS_PR = ""
                gs_Subject_DMMS_PR = ""
            End If

            If gs_SMTPAddress_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's SMTP Address is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            If gs_SMTPPort_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's SMTP Port is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            If gs_SMTPUser_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's SMTP User is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            If gs_SMTPPassword_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's SMTP Password is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            If gs_EmailFrom_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's Email From is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            If gs_EmailTo_DMMS_PR = "" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Mailer's Email To is not found." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End If

            Dim countPR As Integer = 0
            Dim e_mail As New MailMessage()
            Dim iRow As Integer = 0

            Try
                Dim Smtp_Server As New SmtpClient
                Dim strFileName As String = ""

                e_mail.From = New MailAddress(gs_EmailFrom_DMMS_PR)
                e_mail.To.Add(FormatMultipleEmailAddresses(gs_EmailTo_DMMS_PR))

                e_mail.Subject = gs_Subject_DMMS_PR + " " + Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
                e_mail.IsBodyHtml = True
                e_mail.Body = "<br>"
                e_mail.Body = e_mail.Body & "Dear Mr./Mrs. " & gs_EmailTo_DMMS_PR & ", " & "<br>" & "<br>"
                e_mail.Body = e_mail.Body & "We would like to inform you that there are list of Purchase Request's Status." & "<br><br>"

                Dim dtPR As DataTable
                dtPR = getData_IF_PR_DMMS_RESP(ConStr, pFileName)
                countPR = dtPR.Rows.Count
                If dtPR.Rows.Count > 0 Then
                    For iRow = 0 To dtPR.Rows.Count - 1
                        If dtPR.Rows(iRow)("Status").ToString.Trim = "OK" Then
                            e_mail.Body = e_mail.Body & "PR No. " & dtPR.Rows(iRow)("PR_Number").ToString.Trim & ", Status: " & dtPR.Rows(iRow)("Status").ToString.Trim & ", Process Date: " & CDate(dtPR.Rows(iRow)("ProcessDate")).ToString("yyyy-MM-dd HH:mm:ss") & "<br>"
                        Else
                            e_mail.Body = e_mail.Body & "PR No. " & dtPR.Rows(iRow)("PR_Number").ToString.Trim & ", Status: " & dtPR.Rows(iRow)("Status").ToString.Trim & ", Process Date: " & CDate(dtPR.Rows(iRow)("ProcessDate")).ToString("yyyy-MM-dd HH:mm:ss") & ", Error: " & dtPR.Rows(iRow)("ErrorMessage").ToString.Trim & "<br>"
                        End If
                    Next
                End If

                e_mail.Body = e_mail.Body & "<br>"
                e_mail.Body = e_mail.Body & "Best Regards, " & "<br>"
                e_mail.Body = e_mail.Body & "E-Procurement Information Center (EPIC) System" & "<br>"

                Smtp_Server.Host = gs_SMTPAddress_DMMS_PR
                Smtp_Server.Port = gs_SMTPPort_DMMS_PR
                Smtp_Server.Credentials = New Net.NetworkCredential(gs_SMTPUser_DMMS_PR, gs_SMTPPassword_DMMS_PR)
                Smtp_Server.EnableSsl = False
                ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                Smtp_Server.Send(e_mail)

                ' Log for Success Sending Email 
                ' -----------------------------
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Sending Email ... " & countPR & " to " & gs_EmailTo_DMMS_PR & " ... Success !" & vbCrLf
                up_gridProcess(Rtb1, 3, 0, ls_StepProcess)

                statusSend = True
                ' -----------------------------

            Catch ex As SmtpException
                e_mail.Dispose()
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error send email. SMTP Error." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            Catch ex As ArgumentOutOfRangeException
                e_mail.Dispose()
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error send email. Check SMTP Port Number." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            Catch Ex As InvalidOperationException
                e_mail.Dispose()
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error send email. Check SMTP Port Number." & vbCrLf
                up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
                Return False
            End Try
        Catch ex As Exception
            'up_ShowMessage(Err.Number, txtMsg, clsGlobal.MsgTypeEnum.ErrorMsg, Err.Description)
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Error send email... " & " [" & Err.Number & "]" & Err.Description & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Return False
        End Try

        Return True
    End Function

    Private Sub up_JSONFile_Copy(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal jsonfilename As String)
        Try
            Dim di As New IO.DirectoryInfo(pPathSource)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.json")
            Dim fi As IO.FileInfo = Nothing

            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            My.Computer.FileSystem.MoveFile(pPathSource & jsonfilename, pPathDestination & jsonfilename, True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub up_XMLFile_Move(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal xmlfilename As String)
        Try
            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            My.Computer.FileSystem.MoveFile(pPathSource & xmlfilename, pPathDestination & xmlfilename, True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub up_StartScheduler()

        Dim dtNextProcess As DataTable

        dtNextProcess = ClsInterfaceSettingDB.getNextProcess(ConStr)

        If dtNextProcess.Rows.Count > 0 Then
            dtpNext.Value = dtNextProcess.Rows(0).Item("Time")
        Else
            ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Automatic Process Error ... Please set Next Process Time " & vbCrLf
            up_gridProcess(Rtb1, 2, 0, ls_StepProcess)
            Exit Sub
        End If

        Try
            If btnStart.Text.Trim = "S&TART" Then
                ls_StepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] Start Batch Process ... " & vbCrLf
                up_gridProcess(Rtb1, 4, 0, ls_StepProcess)
                EnableControl(False)
                btnStart.Text = "S&TOP"

                Timer1.Enabled = True
                dtpLast.Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

            Else
                ls_StepProcess = "" & vbCrLf

                ls_StepProcess = ls_StepProcess & "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] End Batch Process ... " & vbCrLf
                up_gridProcess(Rtb1, 4, 0, ls_StepProcess)

                EnableControl(True)
                btnStart.Text = "S&TART"

                Timer1.Enabled = False
            End If
        Catch ex As Exception
            Gb.up_ShowMessage("Scheduler error: " & ex.Message.ToString, txtMsg, clsGlobal.MsgTypeEnum.ErrorMsg)
        End Try
    End Sub

    Private Sub up_gridProcess(ByVal rtbox As RichTextBox, ByVal pColor As Integer, ByVal pPos As Integer, ByVal pMsg As String, Optional UseTime As Boolean = False)
        With rtbox
            Select Case pColor
                Case 1 : .SelectionColor = Color.Black   ' Hitam
                Case 2 : .SelectionColor = Color.Red     ' Merah
                Case 3 : .SelectionColor = Color.Green   ' Hijau
                Case 4 : .SelectionColor = Color.Blue    ' Biru
                Case 5 : .SelectionColor = Color.Gray    ' Abu
                Case 6 : .SelectionColor = Color.Orange  ' Orange
                Case 7 : .SelectionColor = Color.Purple  ' Ungu
            End Select
            Dim CurrentTime As String = Format(Date.Now, "HH:mm:ss")
            Dim Message As String
            If UseTime Then
                Message = "[" & CurrentTime & "] " & pMsg
            Else
                Message = pMsg
            End If
            .AppendText(Space(pPos) & Message & vbCrLf)
            .ScrollToCaret()
        End With
    End Sub

    Private Sub EnableControl(Enable As Boolean)
        btnProcess.Enabled = Enable
        btnClose.Enabled = Enable

        Application.DoEvents()
    End Sub

    Private Function uf_AddSpace(ByVal pDuration As String, ByVal pSpace As Integer) As String
        Return Space(pSpace - pDuration.Length) & pDuration
    End Function

    Private Function checkDate(input As String) As Boolean
        Dim result As DateTime
        Return DateTime.TryParse(input, result)
    End Function

    Private Function WriteToEventLog(ByVal Entry As String, Optional ByVal AppName As String = "SSP", _
        Optional ByVal EventType As EventLogEntryType = EventLogEntryType.Information, _
        Optional ByVal LogName As String = "Application", Optional ByVal pID As Integer = 1000) As Boolean


        Dim objEventLog As New EventLog()

        Try
            'Register the App as an Event Source
            If Not EventLog.SourceExists(AppName) Then
                EventLog.CreateEventSource(AppName, LogName)
            End If

            objEventLog.Source = AppName

            'WriteEntry is overloaded; this is one
            'of 10 ways to call it
            objEventLog.WriteEntry(Entry, EventType, pID, CShort(0))
            Return True
        Catch Ex As Exception

            Return False

        End Try

    End Function
#End Region

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        DtNowDisplay.Value = Format(Now(), "dd MMM yyyy HH:mm:ss")
    End Sub

    Private Sub btnSchedule_Click(sender As System.Object, e As System.EventArgs) Handles btnSchedule.Click
        frmSchedule.ShowDialog()
    End Sub

    Private Sub btnSetting_Click(sender As System.Object, e As System.EventArgs) Handles btnSetting.Click
        frmSetting.ShowDialog()
    End Sub
End Class

Public Class UpperCaseUTF8Encoding
    Inherits UTF8Encoding

    Public Overrides ReadOnly Property WebName As String
        Get
            Return MyBase.WebName.ToUpper()
        End Get
    End Property

    Public Shared ReadOnly Property UpperCaseUTF8 As UpperCaseUTF8Encoding
        Get

            If upperCaseUtf8Encoding Is Nothing Then
                upperCaseUtf8Encoding = New UpperCaseUTF8Encoding()
            End If

            Return upperCaseUtf8Encoding
        End Get
    End Property

    Private Shared upperCaseUtf8Encoding As UpperCaseUTF8Encoding = Nothing
End Class

#Region "Class Model"
Public Class MaterialMaster
    Public MATERIAL As String = ""
    Public IND_SECTOR As String = ""
    Public MATL_TYPE As String = ""
    Public BASIC_VIEW As String = ""
    Public SALES_VIEW As String = ""
    Public PURCHASE_VIEW As String = ""
    Public MRP_VIEW As String = ""
    Public STORAGE_VIEW As String = ""
    Public ACCOUNT_VIEW As String = ""
    Public COST_VIEW As String = ""
    Public MATL_DESC As String = ""
    Public LANGU As String = ""
    Public MATL_GROUP As String = ""
    Public BASE_UOM As String = ""
    Public DIVISION As String = ""
    Public PLANT As String = ""
    Public TRANS_GRP As String = ""
    Public LOADINGGRP As String = ""
    Public AVAILCHECK As String = ""
    Public PROFIT_CTR As String = ""
    Public PUR_GROUP As String = ""
    Public MANU_MAT As String = ""
    Public MRP_TYPE As String = ""
    Public MRP_CTRLER As String = ""
    Public LOTSIZEKEY As String = ""
    Public REORDER_PT As String = ""
    Public ROUND_VAL As String = ""
    Public MAX_STOCK As String = ""
    Public PLND_DELRY As String = ""
    Public PROC_TYPE As String = ""
    Public SAFETY_STK As String = ""
    Public BACKFLUSH As String = ""
    Public ISS_ST_LOC As String = ""
    Public SLOC_EXPRC As String = ""
    Public PERIOD_IND As String = ""
    Public PLANT_D As String = ""
    Public STGE_LOC As String = ""
    Public VAL_AREA As String = ""
    Public PRICE_CTRL As String = ""
    Public VAL_CLASS As String = ""
    Public MOVING_PR As String = ""
    Public ORIG_MAT As String = ""
    Public MFR_PRTNUMB As String = ""
End Class

#End Region


'******************************************************************************************************************
'System name    : ADM PIS Karawang
'Class          : clsGlobal
'Overview       : 
'   class containing global functions that used in the entire solution
'******************************************************************************************************************
'Created by     : Ari
'Create Date    : 20-Dec-2011
'Modify History :
'Remarks        : 
'******************************************************************************************************************

Imports C1.Win
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing

Public Class clsGlobal

#Region "Declaration"
    Dim nRetry As Integer
    Dim ConStr As String
    Dim UserID As String
    Const AppID As String = "P01"
    Dim dtMsg As DataTable
    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Enum MsgTypeEnum
        InformationMsg = 0
        ErrorMsg = 1
    End Enum

    Public Enum MsgButtonEnum
        OKOnly = 0
        OKCancel = 1
        YesNo = 3
        YesNoCancel = 4
    End Enum

    Public Enum MsgResultEnum
        Cancel = 0
        OK = 1
        No = 2
        Yes = 3
    End Enum

    Public Enum MsgIDEnum
        InvalidUserIDorPassword = 1
        PleaseInput_XX = 2
        PleaseInputValid_XX = 3
        VINisNotFoundInWOS = 4
        XX_IsNotFound = 5
        DataIsAlreadyUsedIn_XX = 6
        InvalidFileFormat = 7
        InvalidFile = 8
        NoDataToImport = 9
        InvalidDateFormat = 10
        VINDoesNotMatchWithChasSMSNo = 11
        VINAlreadySuspend = 12
        VINAlreadyCancelSuspend = 13
        VINAlreadyJigIn = 14
        DataAlreadyExists = 15
        YouDontHavePrivilegeToUpdate = 16
        YouDontHavePrivilegeToReprint = 17
        DataIsNotFound = 18
        PleaseSelect_XX = 19
        VINAlreadyScrap = 20
        NoDataToDelete = 21
        NoDataToSave = 22
        VINisAlreadyActiveForPlanJiginDate = 23
        NoDataToCopy = 24
        StartTimeMustbeMoreThanPreviousEndTime = 25
        EndTimeMustBeMoreThanStartTime = 26
        BreakTimeMustBeWithinStartAndEndShift = 27
        DayShiftMustNotOverlapWithNightShift = 29
        OvertimeMustNotOverlapWithStartAndEndOfShift = 30
        WorkingHourCodeHasNotBeenSet = 33
        ProductionDateHasAlreadyPassed = 34
        VINLengthMustBe17 = 35
        NoDataChanges = 37
        VINHasJustBeenScannedHere = 40

        InvalidTPRoute = 48
        XX_Failed = 98
        DatabaseConnectionFailed = 99

        LoadDataSuccessfull = 101
        InsertDataSuccesfull = 102
        UpdateDataSuccesfull = 103
        DeleteDataSuccesfull = 104
        PrintDataSuccessfull = 105
        ExportToExcelSuccessfull = 106
        VINSuspendSuccessfull = 107
        VINCancelSuspendSuccessfull = 108

        XX_Successful = 121

    End Enum
#End Region

    ''' <summary>
    ''' Get all data from Message table and store it in a datatable.
    ''' This function is called once when the class is initialized.
    ''' Later this datatable is accessed to get messages without having to connect to database.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub up_GetMsg()
        Try
            dtMsg = New DataTable
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                Dim cmd As New SqlCommand("select * from Message", Cn)
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dtMsg)
            End Using
        Catch ex As Exception
            If ex.Message.Contains(ConnectionErrorMsg) Or ex.Message.Contains(TransportErrorMsg) Then
                Throw New Exception("Database connection failed")
            Else
                Throw ex
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Get form backcolor
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Formbackcolor As Color
        Get
            Return Color.White
        End Get
    End Property

    ''' <summary>
    ''' Get grid backcolor
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GridBackcolor As Color
        Get
            Return Color.LemonChiffon
        End Get
    End Property

    ''' <summary>
    ''' Message box prompt when user wants to back to Sub Menu.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConfirmSubMenu() As MsgBoxResult
        Return MsgBox("Are you sure you want to close ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Sub Menu")
    End Function

    ''' <summary>
    ''' Message box prompt when user wants to delete data.
    ''' </summary>
    ''' <param name="pItemDeleted"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConfirmDelete(Optional ByVal pItemDeleted As String = "this data") As MsgBoxResult
        Return MsgBox("Are you sure you want to delete " & pItemDeleted & " ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Delete")
    End Function

    Public Function ConfirmActivate() As MsgBoxResult
        Return MsgBox("This data already exists, Are You Want To ReActivate This Data?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Activate")
    End Function

    Public Function ConfirmSave() As MsgBoxResult
        Return MsgBox("Are you sure to save this data?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Save")
    End Function

    Public Function ConfirmUpdate() As MsgBoxResult
        Return MsgBox("Are you sure you want to update this data?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Update")
    End Function

    ''' <summary>
    ''' Message box prompt when user wants to cancel operation.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConfirmCancel() As MsgBoxResult
        Return MsgBox("Are you sure you want to cancel this data without any changes ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + vbDefaultButton2, "Cancel")
    End Function

    ''' <summary>
    ''' Returns true if the selected date is working day.
    ''' </summary>
    ''' <param name="pDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function uf_IsWorkingDay(ByVal pDate As Date) As Boolean
        Dim Q As String = _
            "select WorkingType from P_WHMst W inner join P_CalendarMst C on W.WHCode = C.WHCode " & vbCrLf & _
            "where C.CalendarDate = '" & Format(pDate, "yyyy-MM-dd") & "' "
        Dim dt As DataTable = uf_GetDataTable(Q)
        If dt.Rows.Count = 0 Then
            Return False
        ElseIf dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Get production date and shift for current time.
    ''' </summary>
    ''' <param name="pShopCode"></param>
    ''' <param name="pProdDate"></param>
    ''' <param name="pShift"></param>
    ''' <remarks></remarks>
    Public Sub GetDateShift(ByVal pShopCode As String, Optional ByRef pProdDate As Date = Nothing, Optional ByRef pShift As String = "")
        Using Cn As New SqlConnection(ConStr)
            Cn.Open()
            Dim cmd As New SqlCommand("select * from uf_GetDateShift('" & pShopCode & "') ", Cn)
            Dim rd As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SingleRow)
            If rd.Read Then
                pProdDate = rd(0)
                pShift = rd(2)
            End If
            rd.Close()
        End Using
    End Sub

    ''' <summary>
    ''' Show data in a new window where user can select a row
    ''' </summary>
    ''' <param name="pQuery">
    ''' Query to show data in the grid
    ''' </param>
    ''' <param name="pTitle">
    ''' The text that will be displayed in the form title bar
    ''' </param>
    ''' <param name="pColumnIndex">
    ''' Index of the column that this function will return.
    ''' Use -1 to retrieve all columns, separated by vbTab character
    ''' Then use "Split" function to transform this string into an array
    ''' </param>    
    ''' <param name="pFindColumnIndex">
    ''' Default column to be used in filtering
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    ''' <summary>
    ''' Fill combobox with data
    ''' </summary>
    ''' <param name="pCbo">The name of the combo box</param>
    ''' <param name="pTable">Table name(s), can contain JOIN or UNION statement</param>
    ''' <param name="pField">Field name(s)</param>
    ''' <param name="pColWidth">Width of the first column</param>
    ''' <param name="pDropDownWidth">Width of all colums</param>
    ''' <param name="pAll">Determine whether 'ALL' is displayed in the selection</param>
    ''' <param name="pShowHeaders">Determine whether column header is displayed</param>
    ''' <param name="pOrderBy">Fields to sort the query</param>
    ''' <param name="pAsterisk">Determine whether asterisk (*) is displayed in the selection</param>
    ''' <remarks></remarks>
 

    Public Function uf_GetPaperSource(ByVal m_prtName As String, _
    ByVal m_PrtTray As String) As System.Drawing.Printing.PaperSource
        Dim m_PrintDocument As New PrintDocument
        Dim m_PaperSources As System.Drawing.Printing.PaperSource = Nothing

        For Each strPrinter In PrinterSettings.InstalledPrinters
            m_PrintDocument.PrinterSettings.PrinterName = strPrinter
            If m_prtName = m_PrintDocument.PrinterSettings.PrinterName Then
                For Each m_PaperSources In m_PrintDocument.PrinterSettings.PaperSources
                    If m_PaperSources.SourceName.ToString.Trim = m_PrtTray Then
                        m_PrintDocument.DefaultPageSettings.PaperSource = m_PaperSources
                        Exit For
                    End If
                Next
            End If
        Next

        Return m_PaperSources
    End Function

    ''' <summary>
    ''' Open connection to database.
    ''' </summary>
    ''' <param name="pCon">Connection object</param>
    ''' <param name="pConStr">Connection string from application setting</param>
    ''' <param name="pRetryInterval">Time interval that the system should try to reconnect, if previous attempt failed</param>
    ''' <param name="pMaxRetry">Maximum retry if connection failed</param>
    ''' <param name="pErrNumber"></param>
    ''' <param name="pErrMsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function uf_OpenCon(ByVal pCon As SqlConnection, ByVal pConStr As String, Optional ByVal pRetryInterval As Integer = 0, Optional ByVal pMaxRetry As Integer = 0, Optional ByRef pErrNumber As String = "", Optional ByRef pErrMsg As String = "") As Integer
        Do
            Try
                Dim cf As New clsConfig
                pMaxRetry = cf.MaxRetry
                pRetryInterval = cf.RetryInterval

                pCon.ConnectionString = pConStr
                pCon.Open()
                nRetry = 0
                Return 0
            Catch ex As SqlException
                nRetry = nRetry + 1
                If nRetry > pMaxRetry Then
                    nRetry = 0
                    pErrNumber = ex.Number
                    pErrMsg = ex.Message
                    Return -1
                End If
                Dim IntervalSeconds As Integer = 1000 * pRetryInterval
                Threading.Thread.Sleep(IntervalSeconds)
            Catch ex As Exception
                nRetry = nRetry + 1
                If nRetry > pMaxRetry Then
                    pErrNumber = 0
                    pErrMsg = ex.Message
                    Return -1
                End If

            End Try
        Loop
        Return 0
    End Function

    ''' <summary>
    ''' Convert the value into string.
    ''' </summary>
    ''' <param name="Value">
    ''' Value to convert.
    ''' </param>
    ''' <returns>
    ''' Returns string.
    ''' If the value is Nothing or Null, returns empty string.
    ''' </returns>
    ''' <remarks></remarks>
    Public Function uf_Text(ByVal Value As Object) As String
        Dim result As String = ""
        If Not (Value Is Nothing Or IsDBNull(Value)) Then
            result = Value.ToString.TrimEnd
        End If
        Return result
    End Function

    Public Function uf_Val(ByVal Value As Object) As Double
        Dim result As Double = 0
        If Not (Value Is Nothing Or IsDBNull(Value)) Then
            result = Val(Value)
        End If
        Return result
    End Function

    Public Function uf_GetDataSet(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataSet
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)
            da = Nothing
            Return ds
        Else
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds)
                da = Nothing
                Return ds
            End Using
        End If
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function

    Public Function uf_ExecuteSql(ByVal Query As String, ByVal pCon As SqlConnection, Optional ByVal pTrans As SqlTransaction = Nothing) As Integer
        Dim cmd As New SqlCommand(Query, pCon)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        Dim RowAff As Long
        RowAff = cmd.ExecuteNonQuery
        Return RowAff
    End Function

    Public Function uf_ExecuteScalar(ByVal Query As String, ByVal pCon As SqlConnection, Optional ByVal pTrans As SqlTransaction = Nothing) As Object
        Dim cmd As New SqlCommand(Query, pCon)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        Dim Result As Object
        Result = cmd.ExecuteScalar
        Return Result
    End Function

    ''' <summary>
    ''' Gets message description from database based on message ID.
    ''' </summary>
    ''' <param name="pMsgID">
    ''' Message ID.
    ''' </param>
    ''' <returns>
    ''' Returns message description.
    ''' </returns>
    ''' <remarks></remarks>
    Public Function uf_ConvertMsg(ByVal pMsgID As String) As String
        Try
            'If Not IsNumeric(pMsgID) Then
            '    Return pMsgID
            'End If
            Dim Msg As String = ""
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pMsgID)
            If row IsNot Nothing AndAlso row(1) IsNot Nothing Then
                Msg = row(1).ToString.TrimEnd
            End If
            Return Msg
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Sub WriteLog(ByVal pLogID As String)
        Dim q As String = "Select "

    End Sub

    Public Sub up_GetMenu(ByVal pFormName As String, ByRef pMenuID As String, ByRef pMenuDesc As String)
        Try
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                Dim q As String = "Select Menu_ID, Menu_Desc from Gen_SMS_Menu where Menu_Name = '" & pFormName & "' "
                Dim cmd As New SqlCommand(q, Cn)
                Dim rd As SqlDataReader = cmd.ExecuteReader
                If rd.Read Then
                    pMenuID = rd(0).ToString.TrimEnd
                    pMenuDesc = rd(1).ToString.TrimEnd
                Else
                    pMenuID = ""
                    pMenuDesc = ""
                End If
                rd.Close()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function uf_FormTitle(ByVal pFormName As String, Optional ByRef pMenuID As String = "") As String
        Using Cn As New SqlConnection(ConStr)
            Dim ds As DataSet = uf_GetDataSet("Select * from Gen_SMS_Menu where Menu_Name = '" & pFormName & "' ", Cn)
            Dim dt As DataTable = ds.Tables(0)
            If dt.Rows.Count = 0 Then
                Return ""
            Else
                pMenuID = dt.Rows(0)("Menu_ID").ToString.Trim
                Return "[" & pMenuID & "] " & dt.Rows(0)("Menu_Desc").ToString
            End If
        End Using
    End Function

    ''' <summary>
    ''' Show message in the selected control.
    ''' </summary>
    ''' <param name="pMessage">
    ''' Message to show.
    ''' </param>
    ''' <param name="ptxtMsg">
    ''' Control to show the message.
    ''' </param>
    ''' <param name="pMsgType">
    ''' Sets the color of the message text. ErrorMessage = Red, NormalMessage = Blue.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub up_ShowMsg(ByVal pMessage As String, ByVal ptxtMsg As Control, Optional ByVal pMsgType As MsgTypeEnum = MsgTypeEnum.InformationMsg)
        If pMsgType = MsgTypeEnum.InformationMsg Then
            ptxtMsg.ForeColor = Color.Blue
        Else
            ptxtMsg.ForeColor = Color.Red
        End If
        ptxtMsg.Text = uf_ConError(pMessage)
    End Sub

    Private Function uf_ConError(ByVal strMsg As String) As String
        If strMsg.Contains(ConnectionErrorMsg) Or strMsg.Contains(TransportErrorMsg) Then
            Return "Database connection error"
        Else
            Return strMsg
        End If
    End Function

    ''' <summary>
    ''' Gets message from database based on message ID and displays it in the selected control.
    ''' </summary>
    ''' <param name="pMsgID">
    ''' Message ID
    ''' </param>
    ''' <param name="ptxtMsg">
    ''' The control to show the message.
    ''' </param>
    ''' <param name="pMsgType">
    ''' Sets the color of the message text. ErrorMessage = Red, NormalMessage = Blue.
    ''' </param>
    ''' <param name="pParameters">
    ''' Parameters of the message. 
    ''' If the message contains '%%', this value can be replaced with parameters.
    ''' Check the Message Table in the database.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub up_ShowMessage(ByVal pMsgID As String, ByVal ptxtMsg As Control, Optional ByVal pMsgType As MsgTypeEnum = MsgTypeEnum.InformationMsg, Optional ByVal pParameters As String = "")
        Dim iserror As Boolean
        'If pMsgID >= 100 Then
        '    iserror = False
        'ElseIf pMsgID >= 1 Then
        '    iserror = True

        If pMsgType = MsgTypeEnum.InformationMsg Then
            iserror = False
        Else
            iserror = True
        End If
        If iserror Then
            ptxtMsg.ForeColor = Color.Red
        Else
            ptxtMsg.ForeColor = Color.Blue
        End If
        Dim Message As String = uf_ConvertMsg(pMsgID)
        Dim MsgFound As Boolean = False
        If Message = "" Then
            Message = pMsgID
        Else
            MsgFound = True
        End If
        Dim i As Integer, Position As Long
        Dim Parameters() As String
        Parameters = Split(pParameters, "|")
        If UBound(Parameters) <> -1 Then
            Position = InStr(1, Message, "%%")
            Do While Position > 0
                Message = Left(Message, Position - 1) & Parameters(i) & Mid(Message, Position + 2, Len(Message) - Position)
                Position = InStr(1, Message, "%%")
                i = i + 1
            Loop
        Else
            Message = Replace(Message, "%", "")
        End If
        If MsgFound Then
            ptxtMsg.Text = "[" & pMsgID & "] " & Message
        Else
            ptxtMsg.Text = Message
        End If
    End Sub

    Public Function uf_MsgWindow(ByVal Prompt As String, ByVal Buttons As MsgButtonEnum, Optional ByVal Title As String = "")
        Dim frm As New frmMsg(Prompt, Buttons, Title)

        frm.ShowDialog()
        Return frm.MsgResult
        frm.Dispose()
    End Function


    Public Function uf_UserShop(ByVal pUser As String, Optional ByVal pShopName As String = "") As String
        Dim result As String
        Dim dt As New DataTable
        dt = uf_GetDataTable("Select U.ShopCode, S.ShopName from UserSetup U inner join P_ShopMst S on U.ShopCode = S.ShopCode where U.UserID = '" & pUser & "'")
        result = dt.Rows(0)(0).ToString.TrimEnd
        pShopName = dt.Rows(0)(1).ToString.TrimEnd
        Return result
    End Function

    ''' <summary>Check if the value is already used in another table</summary>
    ''' <param name="pTableList">Table name, can be more than one. Separate with comma</param>
    ''' <param name="pFieldList">Field name, can be more than one. Separate with comma</param>
    ''' <param name="pValueList">The number of values must match the number of fields</param>
    ''' <param name="pCriteria">Additional criteria</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function uf_Used(ByVal pTableList As String, ByVal pFieldList As String, ByVal pValueList As String, Optional ByVal pCriteria As String = "") As String
        Dim q As String
        Dim pTableArray() As String = Split(pTableList, ",")
        Dim pTable As String
        Dim nTable As Integer
        Dim iTable As Integer
        Dim TableName As String = ""
        Dim pFieldArray() As String = Split(pFieldList, ",")
        Dim pField As String
        Dim nField As Integer
        Dim iField As Integer
        Dim pValueArray() As String = Split(pValueList, ",")
        Dim pValue As String
        Dim nValue As Integer

        nTable = UBound(pTableArray)
        nField = UBound(pFieldArray)
        nValue = UBound(pValueArray)
        For iTable = 0 To nTable
            pTable = pTableArray(iTable)
            q = "Select * from " & pTable & " where "
            For iField = 0 To nField
                pField = pFieldArray(iField)
                If iField <= nValue Then
                    pValue = pValueArray(iField)
                Else
                    pValue = "Null"
                End If
                q = q & pField & " = '" & Replace(pValue, "'", "''") & "' "
                If iField < nField Then
                    q = q & "and "
                End If
            Next
            If pCriteria <> "" Then
                q = q & pCriteria
            End If
            Dim dt As DataTable = uf_GetDataTable(q)
            If dt.Rows.Count > 0 Then
                TableName = TableName + pTable + ", "
            End If
        Next
        If TableName.Length > 2 Then
            TableName = TableName.Substring(0, Len(TableName) - 2)
        End If
        Return TableName
    End Function

    ''' <summary>
    ''' Gets user privilege for read, write, irregular, and gets admin status for the selected User ID and Menu ID.
    ''' </summary>
    ''' <param name="pUserID">
    ''' User ID
    ''' </param>
    ''' <param name="pMenuID">
    ''' Menu ID
    ''' </param>
    ''' <param name="pAllowRead">
    ''' Gets read privilege.
    ''' </param>
    ''' <param name="pAllowWrite">
    ''' Gets write privilege.
    ''' </param>
    ''' <param name="pAllowDelete">
    ''' Gets Delete privilege.
    ''' </param>
    ''' <param name="pAllowApprove">
    ''' Gets Approve privilege.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub up_GetPrivilege(ByVal pUserID As String, ByVal pMenuID As String, Optional ByRef pAllowRead As Integer = 0, Optional ByRef pAllowWrite As Integer = 0, Optional ByRef pAllowDelete As Integer = 0, Optional ByRef pAllowApprove As Integer = 0)
        Dim q As String = "SELECT top 1 AllowRead, AllowWrite, AllowIrregular, AdminStatus " & vbCrLf & _
                          " from UserSetup U left join UserPrivilege P on P.UserID = U.UserID and P.AppID = U.AppID " & vbCrLf & _
                          "where U.AppID = '" & AppID & "' and U.UserID = '" & pUserID & "' " & vbCrLf & _
                          "and (P.MenuID = '" & pMenuID & "' or U.AdminStatus = 1) " & vbCrLf

        q = " SELECT TOP 1 AllowRead = P.Allow_Access_Flg, AllowWrite = P.Allow_Update_Flg, AllowDelete = P.Allow_delete_Flg, AllowApprove = P.Allow_Approve_Flg " & vbCrLf & _
            "   FROM Gen_User_Setup U  " & vbCrLf & _
            "        LEFT JOIN Gen_User_Privilege P on P.User_ID = U.User_ID  " & vbCrLf & _
            "  WHERE U.User_ID = '" & pUserID & "' AND P.Menu_ID = '" & pMenuID & "' "

        Dim ds As DataSet = uf_GetDataSet(q)
        If ds.Tables(0).Rows.Count > 0 Then
            If ds.Tables(0).Rows(0).Item(0) = True Then pAllowRead = 1 Else pAllowRead = 0
            If ds.Tables(0).Rows(0).Item(1) = True Then pAllowWrite = 1 Else pAllowWrite = 0
            If ds.Tables(0).Rows(0).Item(2) = True Then pAllowDelete = 1 Else pAllowDelete = 0
            If ds.Tables(0).Rows(0).Item(3) = True Then pAllowApprove = 1 Else pAllowApprove = 0
        Else
            pAllowRead = 0
            pAllowWrite = 0
            pAllowDelete = 0
            pAllowApprove = 0
        End If
    End Sub

    Public Function GetServerDate() As String
        Dim retValue As String = Now.ToString()
        Dim sqlConn As New SqlConnection(ConStr)
        sqlConn.Open()

        Dim sqlDA As New SqlDataAdapter("SELECT ServerDate = GETDATE()", sqlConn)
        Dim ds As New DataSet

        sqlDA.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            retValue = ds.Tables(0).Rows(0).Item("ServerDate").ToString()
        End If

        Return retValue
    End Function

    ''' <summary>
    ''' This function is used to get Max Length from the table.
    ''' </summary>
    ''' <param name="pTableName">Table Name</param>
    ''' <param name="pFieldName">Field Name</param>
    ''' <returns>return Max Length by Field Name</returns>
    ''' <remarks></remarks>
    Public Function GetMaxLength(ByVal pTableName As String, ByVal pFieldName As String) As Integer
        Dim retValue As Integer = 0
        Dim sqlConn As New SqlConnection(ConStr)
        sqlConn.Open()

        Dim sqlDA As New SqlDataAdapter("SELECT max_length FROM sys.columns WHERE object_id IN (SELECT object_id FROM sys.tables WHERE name = '" & pTableName & "') AND name = '" & pFieldName & "'", sqlConn)
        Dim ds As New DataSet

        sqlDA.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            retValue = ds.Tables(0).Rows(0).Item("max_length")
        End If

        Return retValue
    End Function
End Class



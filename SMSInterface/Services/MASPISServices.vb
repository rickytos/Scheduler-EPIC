Imports System.Xml
Imports System.IO

Public Class MASPISServices

#Region "DECLARATION"
    Protected Shared titleServices As String = "PANonReguler_Quotation"
    Protected Shared quotNameReq As String = "QuotationMasPis"
    Protected Shared quotName As String = "Quotation"

    Protected Shared infoRecordNameReq As String = "InfoRecordMasPis"
    Protected Shared infoRecordName As String = "InfoRecord"


    Private Shared xmlPath_req As String = ""
    Private Shared xmlPathBackup_req As String = ""

    Private Shared xmlPath_res As String = ""
    Private Shared xmlPathBackup_res As String = ""

    Private Shared xmlIrecordPath_req As String = ""

    Private Shared xmlIrecordPath_res As String = ""
    Private Shared xmlIrecordPathBackup_res As String = ""

    Private conStr As String = ""

    Private Enum ProcessType
        Response
        CostTarget
    End Enum
    'Private generalPath As GeneralTextPath

#End Region

    Public Sub New()
        'generalPath = Config.SetSetting()
    End Sub

    '########################################################################################################################################
    '  Generate Quotation XML
    '########################################################################################################################################
    Public Shared Function GenQuotationMasPis(ByVal ConStr As String, ByVal rtbox As RichTextBox) As Boolean
        Dim iStatus As Boolean = True
        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim statussend As Boolean = False

            Dim pathResult As CallProcedureResult = ClsInterfaceSettingDB.getData()

            If String.IsNullOrEmpty(pathResult.ResultMessage) Then
                xmlIrecordPath_req = Trim(pathResult.DataResult.Rows(0).Item("PAInfoRecordMaspis_ReqPath"))
            Else
                Throw New ArgumentException(" Automatic process error ... please setting " & titleServices & " Path ")
            End If

            Dim result As CallProcedureResult = SPQuotationMasPis.SAPQuotationRequest()

            If Not String.IsNullOrEmpty(result.ResultMessage) Then
                Throw New ArgumentException(" An error when getting the data request, error descriptions : " & result.ErrorMessage)
            End If

            If result.DataResult.Rows.Count > 0 Then
                ShowMessage(rtbox, StaticValue.STATUSINFO, " START Generate " & titleServices & " XML To SAP!")

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "In" & quotNameReq & "_" & Format(Now, "yyyymmdd_hhmmss_mss") & ".xml"

                'writer.WriteStartElement("ns0:MT_EPIC2_Quotation_Req")
                'writer.WriteStartAttribute("xmlns:ns0")
                'writer.WriteString("http://iamiepic2")
                'writer.WriteStartElement("FI_FILEXML")
                'writer.WriteEndElement()
                'writer.WriteStartElement("FT_QUOTATION_REQUEST")

                Dim writeElement = New XmlStartElement() With {.StartElementXML = "ns0:MT_EPIC2_Quotation_Req", .StartAttribute = "xmlns:ns0", .WriteString = "http://iamiepic2", .StartElementXMLFile = String.Empty, .StartElemtnXMLRequest = "FT_QUOTATION_REQUEST"}

                Dim writeResp As ResponseStatus = WriteToXML(tempLocalPath & tempLocalFile, result.DataResult, writeElement)

                If writeResp.StatusResp = False Then
                    ShowMessage(rtbox, 2, "START Generate " & titleServices & " XML To SAP!")
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Interface Process " & titleServices & " XML To SAP Success !")

                '=====================================
                ' Update MasPis data
                '=====================================

                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, xmlIrecordPath_req & "\" & tempLocalFile, True)

                statussend = True
            End If

            If statussend = False Then

                ShowMessage(rtbox, 6, " No Data " & titleServices & " XML To SAP ! ")

                ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

                iStatus = False
                Return iStatus
            Else
                ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

                Return iStatus
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ShowMessage(rtbox, 2, " Error Interface Process " & titleServices & " XML To SAP... " & " [" & Err.Number & "]" & Err.Description)

            ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

            iStatus = False
            Return iStatus
        End Try
        Return iStatus
    End Function

    '########################################################################################################################################
    '  Quotaiton To Get XM From SAP
    '########################################################################################################################################
    Public Shared Sub MASPISProcess(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("PAQuotResPath"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("PAQuotResBackupPath"))
            Else
                ShowMessage(rtbox, StaticValue.STATUSERROR, " Automatic Process Error ... Please Setting Interface for " & titleServices & " - Response Path ")
                Exit Sub
            End If

            ' Call to check for response or cost target files
            Dim fileList As FileInfo() = IsResponseOrCostTarget()
            If fileList IsNot Nothing AndAlso fileList.Length > 0 Then
                Check(fileList, rtbox, ConStr)
            Else
                ShowMessage(rtbox, StaticValue.STATUSINFO, "No relevant XML files found.")
            End If

        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub


    '########################################################################################################################################
    '  Generate Info Record for maspis
    '########################################################################################################################################

    Public Shared Function GetInfoRecordMaspis(ByVal ConStr As String, ByVal rtbox As RichTextBox) As Boolean
        Dim iStatus As Boolean = True
        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim statussend As Boolean = False

            Dim pathResult As CallProcedureResult = ClsInterfaceSettingDB.getData()

            If String.IsNullOrEmpty(pathResult.ResultMessage) Then
                xmlPath_req = Trim(pathResult.DataResult.Rows(0).Item("PAQuotReqPath"))
            Else
                Throw New ArgumentException(" Automatic process error ... please setting " & titleServices & " Path ")
            End If


            Dim result As CallProcedureResult = SPQuotationMasPis.SAPInfoRecordRequest()

            If Not String.IsNullOrEmpty(result.ResultMessage) Then
                Throw New ArgumentException(" An error when getting the data request, error descriptions : " & result.ErrorMessage)
            End If

            If result.DataResult.Rows.Count > 0 Then
                ShowMessage(rtbox, StaticValue.STATUSINFO, " START Generate " & titleServices & " XML To SAP!")

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "In" & infoRecordNameReq & "_" & Format(Now, "yyyymmdd_hhmmss_mss") & ".xml"

                Dim writeElement = New XmlStartElement() With {.StartElementXML = "ns0:MT_EPIC2_InfoRecord_Req", .StartAttribute = "xmlns:ns0", .WriteString = "http://iamiepic2", .StartElementXMLFile = String.Empty, .StartElemtnXMLRequest = "FT_INFORECORD_REQUEST"}

                Dim writeResp As ResponseStatus = WriteToXML(tempLocalPath & tempLocalFile, result.DataResult, writeElement)

                If writeResp.StatusResp = False Then
                    ShowMessage(rtbox, 2, "START Generate " & titleServices & " XML To SAP!")
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Interface Process " & titleServices & " XML To SAP Success !")

                '=====================================
                ' Update MasPis data
                '=====================================

                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, xmlPath_req & "\" & tempLocalFile, True)

                statussend = True
            End If

            If statussend = False Then

                ShowMessage(rtbox, 6, " No Data " & titleServices & " XML To SAP ! ")

                ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

                iStatus = False
                Return iStatus
            Else
                ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

                Return iStatus
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ShowMessage(rtbox, 2, " Error Interface Process " & titleServices & " XML To SAP... " & " [" & Err.Number & "]" & Err.Description)

            ShowMessage(rtbox, StaticValue.STATUSINFO, " END Interface Generate " & titleServices & " XML To SAP!")

            iStatus = False
            Return iStatus
        End Try
        Return iStatus
    End Function

    '########################################################################################################################################
    '  Getting Response Info Record for maspis
    '########################################################################################################################################

    Public Shared Sub InfoRecordResp(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("PAQuotResPath"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("PAQuotResBackupPath"))
            Else
                ShowMessage(rtbox, StaticValue.STATUSERROR, " Automatic Process Error ... Please Setting Interface for " & titleServices & " - Response Path ")
                Exit Sub
            End If

            ' Call to check for response or cost target files
            Dim fileList As FileInfo() = IsResponseOrCostTarget()
            If fileList IsNot Nothing AndAlso fileList.Length > 0 Then
                Check(fileList, rtbox, ConStr)
            Else
                ShowMessage(rtbox, StaticValue.STATUSINFO, "No relevant XML files found.")
            End If

        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub


#Region "EXTENSIONS"

    Private Shared Function IsResponseOrCostTarget() As FileInfo()
        Dim di As New IO.DirectoryInfo(xmlPath_res)

        Dim costTarget As FileInfo() = di.GetFiles("*OutCostTarget*.xml")
        Dim QuotationMaspis As FileInfo() = di.GetFiles("*OutQuotationMaspis*.xml")

        Return costTarget.Concat(QuotationMaspis).ToArray()

    End Function

    Public Shared Sub Check(aryFi As FileInfo(), ByVal rtbox As RichTextBox, ByVal conStr As String)
        If aryFi.Length = 0 Then
            ShowMessage(rtbox, StaticValue.STATUSINFO, "No files to process.")
            Return
        End If

        Dim responseMsg As String = "OutQuotationMaspis"
        Dim costMsg As String = "OutCostTarget"

        For Each fi As FileInfo In aryFi
            Dim xmlfilename = fi.Name.ToLower()
            Dim xmlfilenameLocation = fi.DirectoryName



            If xmlfilename.Contains(responseMsg.ToLower()) Then
                WriteToDBResponse(conStr, xmlfilenameLocation, xmlfilename, rtbox) ' Call the SP for QuotationMasPis
            ElseIf xmlfilename.Contains(costMsg.ToLower()) Then
                WriteToDBCostTarget(conStr, xmlfilenameLocation, xmlfilename, rtbox) ' Call the SP for CostTarget
            End If

            ' Move the file
            Utils.MoveXMLFile(xmlPath_res & "\", xmlPathBackup_res & "\", xmlfilename)
            ShowMessage(rtbox, StaticValue.STATUSINFO, "End process Interface " & titleServices & " ... file name ==> " & xmlfilename & " and move file ")
        Next
    End Sub
    Private Shared Function WriteToDBResponse(ByVal Constr As String, ByVal xmlfilenameLocation As String, ByVal xmlfilename As String, ByVal rtbox As RichTextBox)
        Try
            ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start Interface {0} Response ==> {1}", titleServices, xmlfilename))

            Dim doc As XDocument = XDocument.Load(System.IO.Path.Combine(xmlfilenameLocation, xmlfilename))
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_QUOTATION_RESPONSE").Elements("item")

            Dim quotNo = XMLMaterials.First().Element("QUOTATION").Value

            Dim isAlreadyResponded = SPQuotationMasPis.IsAlreadyGetResp(quotNo)

            If isAlreadyResponded Then
                ShowMessage(rtbox, StaticValue.STATUSWRNING, String.Format("{0} with quotatiion no {1} Is already responsed is null or empty and will be skipped", quotNameReq, quotNo))
                Return False
            End If

            ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start process {0} Response ==> Quotation Number : {1}", quotNameReq, quotNo))

            ' Process each item in the XML
            For Each XEL1 As XElement In XMLMaterials
                Dim QuotationRes As New QuotationMasPisRes()
                PopulateQuotationRes(XEL1, QuotationRes)

                Dim sapMaterial = QuotationRes.MATERIAL
                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start material SAP response ==> Part Number : {0}", sapMaterial))

                Dim isExists = SPQuotationMasPis.IsMtrlResExists(quotNo, QuotationRes.MATERIAL)

                If isExists Then
                    ShowMessage(rtbox, StaticValue.STATUSWRNING, String.Format("For Material SAP : {0} is null or empty and will be skipped", sapMaterial))
                    Continue For
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Insert {0} For SAP Number: {1}", quotNo, sapMaterial))

                Dim resultResponse = SPQuotationMasPis.SAPQuotationResponse(QuotationRes)
                Dim isSuccessUploaded = resultResponse.BooleanResult

                If isSuccessUploaded Then
                    ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Success upload {0} Response ==> {1}", quotNameReq, sapMaterial))
                Else
                    Throw New ArgumentNullException(String.Format("Error on : {0} ", resultResponse.ErrorMessage))
                End If


                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("End upload {0} Response ==> {1}", quotNameReq, sapMaterial))
            Next
            Return True
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, String.Format("End upload {0} Response ==> ERROR: {1}, STACKTRACE: {2}", quotNameReq, ex.Message, ex.StackTrace))
            Return False
        End Try
    End Function

    Private Shared Function WriteToDBCostTarget(ByVal Constr As String, ByVal xmlfilenameLocation As String, ByVal xmlfilename As String, ByVal rtbox As RichTextBox)
        Try
            ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start Interface {0} COSTTARGET with Response ==> {1}", titleServices, xmlfilename))

            Dim doc As XDocument = XDocument.Load(System.IO.Path.Combine(xmlfilenameLocation, xmlfilename))
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_COSTTARGET_RESPONSE").Elements("item")

            Dim quotNo = XMLMaterials.First().Element("QUOTATION").Value
            Dim isAlreadyResponded = SPQuotationMasPis.IsAlreadyGetResp(quotNo)

            If isAlreadyResponded Then
                ShowMessage(rtbox, StaticValue.STATUSWRNING, String.Format("Cost Target with quotatiion no {1} Is already responsed is null or empty and will be skipped", quotNameReq, quotNo))
                Return False
            End If

            ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start process {0} Response ==> Quotation Number : {1}", quotNameReq, quotNo))

            ' Process each item in the XML
            For Each XEL1 As XElement In XMLMaterials
                Dim CostTarget As New QuotationMaspisCostTarget()
                PopulateQuotationRes(XEL1, CostTarget)

                Dim sapMaterial = CostTarget.MATERIAL
                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Start material SAP response ==> Part Number : {0}", sapMaterial))

                'Dim isExists = SPQuotationMasPis.IsMtrlResExists(quotNo, CostTarget.MATERIAL)

                'If isExists Then
                '    ShowMessage(rtbox, StaticValue.STATUSWRNING, String.Format("For Material SAP : {0} is null or empty and will be skipped", sapMaterial))
                '    Continue For
                'End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Insert {0} For SAP Number: {1}", quotNo, sapMaterial))

                Dim resultResponse = SPQuotationMasPis.SAPCostTarget(CostTarget)
                Dim isSuccessUploaded = resultResponse.BooleanResult

                If isSuccessUploaded Then
                    ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("Success upload {0} Response ==> {1}", quotNameReq, sapMaterial))
                Else
                    Throw New ArgumentNullException(String.Format("Error on : {0} ", resultResponse.ErrorMessage))
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, String.Format("End upload {0} Response ==> {1}", quotNameReq, sapMaterial))
            Next
            Return True
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, String.Format("End upload {0} Response ==> ERROR: {1}, STACKTRACE: {2}", quotNameReq, ex.Message, ex.StackTrace))
            Return False
        End Try
    End Function


    Private Shared Function PopulateQuotationRes(Of T)(ByVal XEL1 As XElement, ByRef propParam As T) As T
        For Each prop In GetType(T).GetProperties()
            ' Get the XML element by property name
            Dim xmlElement As XElement = XEL1.Element(prop.Name)

            ' Check if the element exists and has a value
            If xmlElement IsNot Nothing AndAlso xmlElement.Value IsNot Nothing Then
                prop.SetValue(propParam, xmlElement.Value, Nothing)
            End If
        Next
        Return propParam
    End Function

    Private Shared Function WriteToXML(ByVal xmlFilePath As String, ByVal dtMasPis As DataTable, ByVal xmlPath As XmlStartElement) As ResponseStatus
        Try
            Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())
            writer.Formatting = Xml.Formatting.Indented

            writer.WriteStartDocument()
            writer.Indentation = 2
            writer.WriteStartElement(xmlPath.StartElementXML)
            writer.WriteStartAttribute(xmlPath.StartAttribute)
            writer.WriteString(xmlPath.WriteString)

            If Not String.IsNullOrEmpty(xmlPath.StartElemtnXMLRequest) Then
                writer.WriteStartElement(xmlPath.StartElementXMLFile)
                writer.WriteEndElement()
            End If
           
            writer.WriteStartElement(xmlPath.StartElemtnXMLRequest)

            Mapper.Map(Of QuotationMasPisReq)(dtMasPis, writer, "item")

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()

            Return New ResponseStatus(True, "Success")
        Catch ex As Exception
            Return New ResponseStatus(False, ex.Message)
        End Try
    End Function
#End Region


    Private Shared Sub ShowMessage(ByVal rtbox As RichTextBox, ByVal Color As Integer, Optional ByVal Message As String = "")
        Dim lsStepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] " & Message & vbCrLf
        GridProcess.upGridProcess(rtbox, Color, 0, lsStepProcess)
    End Sub

End Class

Public Class XmlStartElement
    Public Property StartElementXML As String
    Public Property StartAttribute As String
    Public Property WriteString As String
    Public Property StartElementXMLFile As String
    Public Property StartElemtnXMLRequest As String

End Class
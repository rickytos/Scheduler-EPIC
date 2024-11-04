Imports System.Xml
Imports System.IO
Imports System.Reflection

Public Class CostHistoryServices

#Region "DECLARATION"
    Protected Shared titleServices As String = "Cost History Services"
    Protected costHistNameReq As String = "CostHistory"
    Protected costHistName As String = "Cost History"


    Private xmlPath_req As String = ""
    Private xmlPathBackup_req As String = ""

    Private xmlPath_res As String = ""
    Private xmlPathBackup_res As String = ""

    Private conStr As String = ""

    'Private generalPath As GeneralTextPath

#End Region

    Public Sub New()
        'generalPath = Config.SetSetting()
    End Sub


    'Process EPIC to SAP request GenReqXML
    Public Function GenXMLForSAP(ByVal Constr As String, ByVal rtbox As RichTextBox) As Boolean
        Dim iStatus As Boolean = True
        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        Try
            Dim statussend As Boolean = False
            Dim dtPath = ClsInterfaceSettingDB.getData(Constr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_req = Trim(dtPath.Rows(0).Item("IF_CostHistory_GenXMLToSAP_Path"))
            Else
                Throw New ArgumentException(" Automatic process error ... please setting " & titleServices & " Path ")
            End If

            Dim genXMLSAPcostHistory = SPCostHistory.SAPCostHistoryRequest()

            If genXMLSAPcostHistory.ErrorMessage <> "" Then
                ShowMessage(rtbox, 2, " Error Message : " & genXMLSAPcostHistory.ErrorMessage)
                iStatus = False
                Return iStatus
            End If

            If genXMLSAPcostHistory.DataResult.Rows.Count > 0 Then
                ShowMessage(rtbox, StaticValue.STATUSINFO, " START Generate " & titleServices & " XML To SAP!")

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "In" & costHistNameReq & "_" & Format(Now, "yyyymmdd_hhmmss_mss") & ".xml"

                Dim writeResp As ResponseStatus = WriteToXML(Of CostHistorySAPReq)(tempLocalPath & tempLocalFile, genXMLSAPcostHistory.DataResult, "ns0:MT_EPIC2_XMLPriceAndQuotaReq")

                If writeResp.StatusResp = False Then
                    ShowMessage(rtbox, 2, "START Generate " & titleServices & " XML To SAP!")
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Interface Process " & titleServices & " XML To SAP Success !")

                '=====================================
                ' Update CostHistory data
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

    'Process EPIC getData from SAP 
    Public Sub GetRespFromSAP(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("IF_CostHistory_SAPResp_PriceQty_Path"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("IF_CostHistory_SAPResp_PriceQty_BackupPath"))
            Else
                ShowMessage(rtbox, StaticValue.STATUSERROR, " Automatic Process Error ... Please Setting Interface for Material Master - Response Path ")
                Exit Sub
            End If

            Dim di As New IO.DirectoryInfo(xmlPath_res)
            Dim getFileString As String = "Out" & costHistNameReq & ""
            Dim aryFi As IO.FileInfo() = di.GetFiles("*QuotationMasPis*.xml")

            Dim fi As IO.FileInfo
            Dim j As Integer
            Dim jmlFile As Integer = aryFi.Length
            For Each fi In aryFi
                Dim xmlfilename = fi.Name
                Dim xmlfilenameLocation = fi.DirectoryName

                titleServices = "SAP Data BOM " + titleServices

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start Interface " & titleServices & " Response ==> " & xmlfilename)

                Dim parameterArray As New ParamWriteDB With {
                                        .Constr = ConStr,
                                        .XMLFileNameLocation = xmlfilenameLocation,
                                        .XMLFileName = xmlfilename,
                                        .NameServices = costHistName,
                                        .NameRequest = costHistNameReq,
                                        .KeyField = "VariantName"
                                    }

                WriteToDB(Of CostHistorySAPRes)(parameterArray, rtbox, AddressOf SPCostHistory.SAPCostHistoryResponse)

                'MOVEFILE
                Utils.MoveXMLFile(xmlPath_res & "\", xmlPathBackup_res & "\", xmlfilename)
                ShowMessage(rtbox, StaticValue.STATUSINFO, " End process Interface " & titleServices & " ... file name ==> " & xmlfilename & " and move file ")

            Next

        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub

    'Process EPIC getData From KDPL
    Public Sub GetRespFromKDPL(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try
            SPCostHistory.KDPLCostHistoryResponse()
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub

    Private Shared Sub ShowMessage(ByVal rtbox As RichTextBox, ByVal Color As Integer, Optional ByVal Message As String = "")
        Dim lsStepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] " & Message & vbCrLf
        GridProcess.upGridProcess(rtbox, Color, 0, lsStepProcess)
    End Sub

    Private Function WriteToDB(Of T)(ByVal parameterArray As ParamWriteDB, ByVal rtbox As RichTextBox, ByVal FuncStatement As Func(Of T, CallProcedureResult)) As Boolean
        Try
            Dim nameServices = parameterArray.NameServices

            Dim doc As XDocument = XDocument.Load(parameterArray.XMLFileNameLocation + "\" + parameterArray.XMLFileName)
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_COSTHISTORY").Elements("item")
            For Each XEL1 As XElement In XMLMaterials

                Dim CostHistoryData = Activator.CreateInstance(Of T)()

                For Each prop In GetType(T).GetProperties()
                    Dim propName As String = prop.Name
                    Dim propValue As Object = XEL1.Element(propName).Value

                    If propValue IsNot Nothing Then

                        prop.SetValue(CostHistoryData, propValue, Nothing)
                    End If

                Next

                Dim propInfo As PropertyInfo = GetType(T).GetProperty(parameterArray.KeyField)

                Dim propExistingValue = propInfo.GetValue(CostHistoryData, Nothing)

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start upload " & parameterArray.NameRequest & " Response ==> " & nameServices & " Number :  " & propInfo.Name)

                If propInfo Is Nothing Or propExistingValue = "" Then
                    ShowMessage(rtbox, StaticValue.STATUSWRNING, " " & parameterArray.NameRequest & " For SAP Number :  " & propInfo.Name & " Null or empty it will be skipped")
                Else
                    ShowMessage(rtbox, StaticValue.STATUSINFO, " Insert " & parameterArray.NameRequest & " For SAP Number :  " & propInfo.Name & " with " & nameServices & " : " & propInfo.Name)
                End If

                Dim result = FuncStatement(CostHistoryData)
                If result.ErrorMessage <> "" Then
                    ShowMessage(rtbox, StaticValue.STATUSERROR, " Error Messags : " & result.ErrorMessage)
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, "End upload " & parameterArray.NameRequest & " Response ==> " & nameServices & " Number :  " & propInfo.Name)
            Next
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "Info Record Response => ERROR: " + ex.Message + " " + ex.StackTrace)
        End Try
    End Function

    Private Shared Function WriteToXML(Of T As New)(ByVal xmlFilePath As String, ByVal dtMasPis As DataTable, Optional ByVal HeaderText As String = Nothing) As ResponseStatus
        Try
            Dim headerTag As String = "ns0:MT_EPIC2_CreateCostHistoryReq"

            If HeaderText Is Nothing Then
                headerTag = HeaderText
            End If

            Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())
            writer.Formatting = Xml.Formatting.Indented

            writer.WriteStartDocument()
            writer.Indentation = 2
            writer.WriteStartElement(headerTag)
            writer.WriteStartAttribute("xmlns:ns0")
            writer.WriteString("http://iamiepic2")
            writer.WriteStartElement("FI_FILEXML")
            writer.WriteEndElement()
            writer.WriteStartElement("FT_COSTHISTORY")

            Mapper.Map(Of T)(dtMasPis, writer, "item")

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()

            Return New ResponseStatus(True, "Success")
        Catch ex As Exception
            Return New ResponseStatus(False, ex.Message)
        End Try
    End Function


End Class

Public Class ParamWriteDB
    Public Property Constr As String = ""
    Public Property XMLFileNameLocation As String = ""
    Public Property XMLFileName As String = ""
    Public Property NameRequest As String = ""
    Public Property NameServices As String = ""
    Public Property KeyField As String = ""
End Class

Imports System.Xml
Imports System.IO

Public Class MASPISServices

#Region "DECLARATION"
    Protected Shared titleServices As String = "PANonReguler_Quotation"
    Protected Shared quotNameReq As String = "QuotationMasPis"
    Protected Shared quotName As String = "Quotation"


    Private Shared xmlPath_req As String = ""
    Private Shared xmlPathBackup_req As String = ""

    Private Shared xmlPath_res As String = ""
    Private Shared xmlPathBackup_res As String = ""

    Private conStr As String = ""

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

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_req = Trim(dtPath.Rows(0).Item("IF_Maspis_XMLtoSAP_Path"))
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

                Dim writeResp As ResponseStatus = WriteToXML(tempLocalPath & tempLocalFile, result.DataResult)

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
    '  Quotaiton To Get XM From SAP
    '########################################################################################################################################
    Public Sub MASPISProcess(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("IF_MasPis_RespSAP_Path"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("IF_MasPis_RespSAP_BackupPath"))
            Else
                ShowMessage(rtbox, StaticValue.STATUSERROR, " Automatic Process Error ... Please Setting Interface for " & titleServices & " - Response Path ")
                Exit Sub
            End If

            Dim di As New IO.DirectoryInfo(xmlPath_res)
            Dim getFileString As String = "Out" & quotNameReq & ""
            Dim aryFi As IO.FileInfo() = di.GetFiles("*QuotationMasPis*.xml")

            Dim fi As IO.FileInfo
            Dim j As Integer
            Dim jmlFile As Integer = aryFi.Length
            For Each fi In aryFi
                Dim xmlfilename = fi.Name
                Dim xmlfilenameLocation = fi.DirectoryName

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start Interface " & titleServices & " Response ==> " & xmlfilename)
                WriteToDB(ConStr, xmlfilenameLocation, xmlfilename, rtbox)

                'MOVEFILE
                Utils.MoveXMLFile(xmlPath_res & "\", xmlPathBackup_res & "\", xmlfilename)
                ShowMessage(rtbox, StaticValue.STATUSINFO, " End process Interface " & titleServices & " ... file name ==> " & xmlfilename & " and move file ")

            Next

        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub


    Private Function WriteToDB(ByVal Constr As String, ByVal xmlfilenameLocation As String, ByVal xmlfilename As String, ByVal rtbox As RichTextBox)
        Try
            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_QUOTATIONMASPIS_RESP").Elements("item")
            For Each XEL1 As XElement In XMLMaterials

                Dim QuotationRes As New QuotationMasPisRes

                For Each prop In GetType(QuotationMasPisRes).GetProperties()
                    Dim propName As String = prop.Name
                    Dim propValue As Object = XEL1.Element(propName).Value

                    If propValue IsNot Nothing Then
                        prop.SetValue(QuotationRes, propValue, Nothing)
                    End If

                Next

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start upload " & quotNameReq & " Response ==> " & quotName & " Number :  " & QuotationRes.QUOTATION)

                If QuotationRes.QUOTATION Is Nothing Or QuotationRes.QUOTATION = "" Then
                    ShowMessage(rtbox, StaticValue.STATUSWRNING, " " & quotNameReq & " For SAP Number :  " & QuotationRes.QUOTATION & " Null or empty it will be skipped")
                Else
                    ShowMessage(rtbox, StaticValue.STATUSINFO, " Insert " & quotNameReq & " For SAP Number :  " & QuotationRes.QUOTATION & " with " & quotName & " : " & QuotationRes.QUOTATION)
                End If

                SPQuotationMasPis.SAPQuotationResponse(QuotationRes)

                ShowMessage(rtbox, StaticValue.STATUSINFO, "End upload " & quotNameReq & " Response ==> " & quotName & " Number :  " & QuotationRes.QUOTATION)
            Next
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "Info Record Response => ERROR: " + ex.Message + " " + ex.StackTrace)
        End Try
    End Function

    Private Shared Function WriteToXML(ByVal xmlFilePath As String, ByVal dtMasPis As DataTable) As ResponseStatus
        Try
            Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())
            writer.Formatting = Xml.Formatting.Indented

            writer.WriteStartDocument()
            writer.Indentation = 2
            writer.WriteStartElement("ns0:MT_EPIC2_Quotation_Req")
            writer.WriteStartAttribute("xmlns:ns0")
            writer.WriteString("http://iamiepic2")
            writer.WriteStartElement("FI_FILEXML")
            writer.WriteEndElement()
            writer.WriteStartElement("FT_QUOTATION")

            Mapper.Map(Of QuotationMasPisReq)(dtMasPis, writer, "item")

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()

            Return New ResponseStatus(True, "Success")
        Catch ex As Exception
            Return New ResponseStatus(False, ex.Message)
        End Try
    End Function

    Private Shared Sub ShowMessage(ByVal rtbox As RichTextBox, ByVal Color As Integer, Optional ByVal Message As String = "")
        Dim lsStepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] " & Message & vbCrLf
        GridProcess.upGridProcess(rtbox, Color, 0, lsStepProcess)
    End Sub

End Class

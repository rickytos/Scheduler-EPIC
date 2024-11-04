Imports System.IO
Imports System.Xml
Imports System.Reflection

Public Class InfoRecordServices
#Region "DECLARATION"
    Protected Shared titleServices As String = "Interface Info Record"
    Protected Shared titleRes As String = "Info Record"


    Private Shared xmlPath_req As String = ""
    Private Shared xmlPathBackup_req As String = ""

    Private Shared xmlPath_res As String = ""
    Private Shared xmlPathBackup_res As String = ""

    Private Shared conStr As String = ""

    'Private generalPath As GeneralTextPath

#End Region

    Public Sub New()
        'generalPath = Config.SetSetting()
    End Sub

    '########################################################################################################################################
    '  Generate Quotation XML
    '########################################################################################################################################
    Public Shared Sub generateXMLInfoRecord(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Dim iStatus As Boolean = True
        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        ShowMessage(rtbox, StaticValue.STATUSINFO, "Start " & titleServices & " XML To SAP !")

        Try
            Dim statussend As Boolean = False

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_req = Trim(dtPath.Rows(0).Item("IR_XML_Req_Path"))
            Else
                Throw New ArgumentException(" Automatic process error ... please setting " & titleServices & " Path ")
            End If

            Dim genInfoRecord = SPInfoRecord.getInfoRecordXML()

            If genInfoRecord.ErrorMessage <> "" Then
                Throw New ArgumentException(" Error Message: " & genInfoRecord.ErrorMessage)
            End If

            If genInfoRecord.DataResult.Rows.Count > 0 Then
                ShowMessage(rtbox, StaticValue.STATUSINFO, " START Generate InfoRecord XML To SAP!")

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If

                tempLocalFile = "InInforecord_" & Format(Now, "yyyymmdd_hhmmss_mss") & ".xml"

                Dim writeResp As ResponseStatus = WriteToXML(tempLocalPath & tempLocalFile, genInfoRecord.DataResult)
                If writeResp.StatusResp = False Then
                    ShowMessage(rtbox, 2, "START Generate Material Master XML To SAP!")
                End If

                'Log for Success Sending Email 
                '-----------------------------
                ShowMessage(rtbox, StaticValue.STATUSINFO, " Interface Process Info Record XML To SAP Success !")


                For iRow = 0 To genInfoRecord.DataResult.Rows.Count - 1
                    Dim materialNoSAP = Trim(genInfoRecord.DataResult.Rows(iRow).Item("MATNR") & "")
                    Dim supplier = Trim(genInfoRecord.DataResult.Rows(iRow).Item("VENDOR") & "")

                    Dim newParameter As New Dictionary(Of String, Object)
                    newParameter.Add("MaterialNumbSAP", materialNoSAP)
                    newParameter.Add("StatusHeader", supplier)

                    Dim infoRecordServ = SPInfoRecord.UpdateInfoRecordRequest(newParameter)

                    If infoRecordServ.ErrorMessage <> "" Then
                        ShowMessage(rtbox, StaticValue.STATUSERROR, infoRecordServ.ErrorMessage)
                    End If

                Next iRow


                'Move XML File
                My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, xmlPath_req & "\" & tempLocalFile, True)

            Else
                ShowMessage(rtbox, 6, " No Data " & titleRes & " XML To SAP ! ")
            End If

        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ShowMessage(rtbox, StaticValue.STATUSERROR, " Error " & titleServices & " XML To SAP... " & " [" & Err.Number & "]" & Err.Description)

        End Try

        ShowMessage(rtbox, StaticValue.STATUSINFO, "End " & titleServices & " XML To SAP !")
    End Sub

    '########################################################################################################################################
    ''  Interface InfoRecord
    '########################################################################################################################################

    Public Shared Sub InterfaceRecordProcess(ByVal ConStr As String, ByVal rtbox As RichTextBox)
        Try
            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("IR_XML_Res_Path"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("IR_XML_Res_BackupPath"))
            Else
                ShowMessage(rtbox, 2, " Automatic Process Error ... Please Setting Interface for Info Record - Response Path ")
                Exit Sub
            End If

            Dim di As New IO.DirectoryInfo(xmlPath_res)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*OutInforecord*.xml")

            Dim fi As IO.FileInfo
            Dim j As Integer
            Dim jmlFile As Integer = aryFi.Length
            For Each fi In aryFi
                Dim xmlfilename = fi.Name
                Dim xmlfilenameLocation = fi.DirectoryName

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start Interface Info Record Response ==> " & xmlfilename)

                WriteToDB(ConStr, xmlfilenameLocation, xmlfilename, rtbox)

                'MOVE FILE
                Utils.MoveXMLFile(xmlPath_res & "\", xmlPathBackup_res & "\", xmlfilename)
                ShowMessage(rtbox, StaticValue.STATUSINFO, " End process Interface Info Record ... file name ==> " & xmlfilename & " and move file ")
            Next
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "" + ex.Message + ex.StackTrace)
        End Try
    End Sub


    Private Shared Function WriteToDB(ByVal Constr As String, ByVal xmlfilenameLocation As String, ByVal xmlfilename As String, ByVal rtbox As RichTextBox)
        Try
            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)
            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_INFOREC_RESPONSE").Elements("item")
            For Each XEL1 As XElement In XMLMaterials

                Dim InfoRecordResp As New InfoRecordRes

                For Each prop In GetType(InfoRecordRes).GetProperties()
                    Dim propName As String = prop.Name
                    Dim propValue As Object = XEL1.Element(propName).Value

                    If propValue IsNot Nothing Then
                        prop.SetValue(InfoRecordResp, propValue, Nothing)
                    End If
                Next

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start upload Info Record Response ==> Material SAP Number :  " & InfoRecordResp.INFNR)

                If InfoRecordResp.INFNR = "" Then
                    ShowMessage(rtbox, StaticValue.STATUSWRNING, " InfoRecord For SAP Number :  " & InfoRecordResp.MATNR & " Null or empty it will be skipped")
                Else
                    ShowMessage(rtbox, StaticValue.STATUSINFO, " Insert InfoRecord For SAP Number :  " & InfoRecordResp.MATNR & " with inforecord : " & InfoRecordResp.MATNR)
                End If

                Dim result = SPInfoRecord.UpdateSAPInfoRecordResponse(InfoRecordResp)

                If result.ErrorMessage <> "" Then
                    ShowMessage(rtbox, StaticValue.STATUSERROR, "ErrorMessage" & result.ErrorMessage)
                End If

                ShowMessage(rtbox, StaticValue.STATUSINFO, "End upload Info Record Response ==> Material SAP Number :  " & InfoRecordResp.MATNR)
            Next
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "Info Record Response => ERROR: " + ex.Message + " " + ex.StackTrace)
        End Try
    End Function

    Private Shared Function WriteToXML(ByVal xmlFilePath As String, ByVal dtInfoRecord As DataTable) As ResponseStatus
        Try
            Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())
            writer.Formatting = Xml.Formatting.Indented

            writer.WriteStartDocument()
            writer.Indentation = 2
            writer.WriteStartElement("ns0:MT_EPIC2_CreateInfoRecord_req")
            writer.WriteStartAttribute("xmlns:ns0")
            writer.WriteString("http://iamiepic2")
            writer.WriteStartElement("FI_FILEXML")
            writer.WriteEndElement()
            writer.WriteStartElement("FT_INFOREC")

            Mapper.Map(Of InfoRecordReq)(dtInfoRecord, writer, "item")

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

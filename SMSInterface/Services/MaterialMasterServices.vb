Imports System.Xml
Imports System.IO

Public Class MaterialMasterServices
#Region "DECLARATION"
    Protected Shared titleServices As String = "Interface Process Material Master"

    Private Shared xmlPath_req As String = Nothing
    Private Shared xmlPathBackup_req As String = ""

    Private Shared xmlPath_res As String = ""
    Private Shared xmlPathBackup_res As String = ""

    Private conStr As String = ""

    'Private generalPath As GeneralTextPath
    Private generalPath As String


#End Region

    Public Sub New()
        'generalPath = Config.SetSetting()

    End Sub

    '########################################################################################################################################
    '  Generate Material Master XML
    '########################################################################################################################################
    Public Shared Sub generateXMLMaterialMaster(ByVal ConStr As String, ByVal rtbox As RichTextBox)

        Dim iStatus As Boolean = True
        Dim tempLocalPath As String = ""
        Dim tempLocalFile As String = ""

        ShowMessage(rtbox, StaticValue.STATUSINFO, "Start " & titleServices & " XML To SAP !")

        Try

            Dim year As String = ""
            Dim month As String = ""

            Dim statussend As Boolean = False

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)
            Dim xmlPath_req As String = Nothing

            If dtPath.Rows.Count > 0 Then
                xmlPath_req = Trim(dtPath.Rows(0).Item("MaterialMasterXML_ToSAP_Path"))
            Else
                Throw New ArgumentException(" Automatic process error ... please setting " & titleServices & " Path")
            End If

            Dim genMaterialMaster As DataTable = SPMaterialMaster.GetMaterialMasterXML()

            Dim Count As Integer
            Count = genMaterialMaster.Rows.Count

            If genMaterialMaster.Rows.Count > 0 Then
                ShowMessage(rtbox, StaticValue.STATUSINFO, "Processing Generate " & "Material Master" & " XML To SAP...")

                tempLocalPath = Application.StartupPath & "\Import\"

                If Not System.IO.Directory.Exists(tempLocalPath) Then
                    System.IO.Directory.CreateDirectory(tempLocalPath)
                End If


                tempLocalFile = "InMaterial_" & Format(Now, "yyyymmdd_hhmmss_mss") & ".xml"

                Dim writeResp = WriteToXML(tempLocalPath & tempLocalFile, genMaterialMaster)

                If Not writeResp.StatusResp = True Then
                    Throw New ArgumentException(writeResp.StatusMessage)
                End If

                'Log for Success Sending Email 
                '-----------------------------

                For iRow = 0 To genMaterialMaster.Rows.Count - 1

                    Dim materialNo = Trim(genMaterialMaster.Rows(iRow).Item("MANU_MAT") & "")
                    Dim materialMaster As New MaterialMasterReq
                    materialMaster.MATERIAL = materialNo

                    Dim updaateRequest As Boolean = SPMaterialMaster.UPDMaterialMasterReq(materialMaster, 1)
                    If updaateRequest = False Then
                        Throw New ArgumentException("Failed to update material master status")
                    End If
                Next iRow

                ShowMessage(rtbox, StaticValue.STATUSSUCCESS, titleServices & " XML To SAP Success!")

                'Move XML File
                Try
                    My.Computer.FileSystem.MoveFile(tempLocalPath & tempLocalFile, xmlPath_req & "\" & tempLocalFile, True)
                Catch ex As Exception
                    Throw New ArgumentException("Failed to move file")
                End Try
            Else
                ShowMessage(rtbox, StaticValue.STATUSWRNING, "No Data " & "Material Master" & " XML To SAP ! ")
                ShowMessage(rtbox, StaticValue.STATUSINFO, "End " & titleServices & " XML To SAP !")
                Return
            End If


        Catch ex As Exception
            For Each deleteFile In Directory.GetFiles(tempLocalPath, tempLocalFile, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            ShowMessage(rtbox, StaticValue.STATUSERROR, " Error Interface Process Material Master XML To SAP... " & " [" & Err.Number & "]" & Err.Description)

        End Try

        ShowMessage(rtbox, StaticValue.STATUSINFO, "End " & titleServices & " XML To SAP !")
    End Sub

    '########################################################################################################################################
    '  Response Material Master
    '########################################################################################################################################

    Public Shared Sub InterfaceMaterialMasterProcess(ByVal ConStr As String, ByVal rtbox As RichTextBox)

        ShowMessage(rtbox, StaticValue.STATUSINFO, "Start " & titleServices & " XML To SAP !")

        Try

            Dim dtPath = ClsInterfaceSettingDB.getData(ConStr)

            If dtPath.Rows.Count > 0 Then
                xmlPath_res = Trim(dtPath.Rows(0).Item("MaterialMasterXML_Resp_Pth"))
                xmlPathBackup_res = Trim(dtPath.Rows(0).Item("MaterialMasterXML_Resp_BackupPth"))
            Else
                Throw New ArgumentException("Automatic Process Error ... Please Setting Interface for Material Master - Response Path")
            End If

            Dim di As New IO.DirectoryInfo(xmlPath_res)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*OutMaterial*.xml")

            Dim fi As IO.FileInfo
            Dim j As Integer
            Dim jmlFile As Integer = aryFi.Length
            For Each fi In aryFi
                Dim xmlfilename = fi.Name
                Dim xmlfilenameLocation = fi.DirectoryName

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Processing Response ==> " & xmlfilename & " Please wait...")

                WriteToDB(ConStr, xmlfilenameLocation, xmlfilename, rtbox)

                'MOVE FILE
                Dim messageError = Utils.MoveXMLFile(xmlPath_res & "\", xmlPathBackup_res & "\", xmlfilename)

                If messageError IsNot Nothing Then
                    Throw New ArgumentException("Failed to move file " & messageError)
                End If
            Next
        Catch ex As Exception

            ShowMessage(rtbox, StaticValue.STATUSERROR, " Error Interface Process Material Master XML To SAP... " & " [" & Err.Number & "]" & Err.Description)

        End Try

        ShowMessage(rtbox, StaticValue.STATUSINFO, "End " & titleServices & " XML To SAP !")
    End Sub

    Private Shared Function WriteToXML(ByVal xmlFilePath As String, ByVal dtMMaster As DataTable) As ResponseStatus
        Try
            Dim writer As New XmlTextWriter(xmlFilePath, New UpperCaseUTF8Encoding())
            writer.Formatting = Xml.Formatting.Indented

            writer.WriteStartDocument()
            writer.Indentation = 2
            writer.WriteStartElement("ns0:MT_EPIC2_MaterialMaster_Req")
            writer.WriteStartAttribute("xmlns:ns0")
            writer.WriteString("http://iamiepic2")
            ''TOADV: add this into the createNode_MaterialMaster by using an ARRAY
            writer.WriteStartElement("FI_FILEXML")
            writer.WriteEndElement()
            writer.WriteStartElement("FT_MATERIAL")

            'here
            Mapper.Map(Of MaterialMasterReq)(dtMMaster, writer, "item")

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()

            Return New ResponseStatus(True, "Success")
        Catch ex As Exception

            Return New ResponseStatus(False, ex.Message)
        End Try
    End Function
    Private Shared Sub WriteToDB(ByVal ConStr As String, ByVal xmlfilenameLocation As String, ByVal xmlfilename As String, ByVal rtbox As RichTextBox)
        Try
            Dim doc As XDocument = XDocument.Load(xmlfilenameLocation + "\" + xmlfilename)

            Dim XMLMaterials As IEnumerable(Of XElement) = doc.Root.Elements("FT_MATERIAL_RESPONSE").Elements("item")
            For Each XEL1 As XElement In XMLMaterials

                Dim MaterialMasterResp As New MaterialMasterRes

                For Each prop In GetType(MaterialMasterRes).GetProperties()
                    Dim propName As String = prop.Name
                    Dim propValue As Object = XEL1.Element(propName).Value

                    If propValue IsNot Nothing Then
                        prop.SetValue(MaterialMasterResp, propValue, Nothing)
                    End If
                Next

                ShowMessage(rtbox, StaticValue.STATUSINFO, " Start upload Material Master Response ==> MANU MAT Number :  " & MaterialMasterResp.MANU_MAT)

                Dim respStatus = SPMaterialMaster.SAPMaterialMasterResp(ConStr, MaterialMasterResp)

                ShowMessage(rtbox, respStatus.StatusInt, "End upload Material Master Response ==> " & respStatus.StatusMessage)


                ShowMessage(rtbox, StaticValue.STATUSINFO, " End upload Material Master Response ==> MANU MAT Number :  " & MaterialMasterResp.MANU_MAT)
            Next
        Catch ex As Exception
            ShowMessage(rtbox, StaticValue.STATUSERROR, "Material Master Response => ERROR: " + ex.Message + " " + ex.StackTrace)
        End Try
    End Sub

    Private Shared Sub ShowMessage(ByVal rtbox As RichTextBox, ByVal Color As Integer, Optional ByVal Message As String = "")
        Dim lsStepProcess = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] " & Message & vbCrLf
        GridProcess.upGridProcess(rtbox, Color, 0, lsStepProcess)
    End Sub
End Class

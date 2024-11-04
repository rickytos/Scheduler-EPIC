Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient

Public Class Form1
    Dim jsonfilename As String = ""
    Dim jsonfilenameLocation As String = ""
    Dim xmlfilename As String = ""
    Dim xmlfilenameLocation As String = ""
    Protected ConStr As String


    Private Sub json()
        ConStr = Builder.ConnectionString
        Dim dtInterface_Setting As DataTable
        dtInterface_Setting = ClsInterfaceSettingDB.getData(ConStr)

        If dtInterface_Setting.Rows.Count > 0 Then
            gs_JSONPath = Trim(dtInterface_Setting.Rows(0).Item("JSON_Path"))
            gs_JSONPathBackup = Trim(dtInterface_Setting.Rows(0).Item("JSON_PathBackup"))
        Else
            gs_JSONPath = ""
            gs_JSONPathBackup = ""

        End If

        Dim di As New IO.DirectoryInfo(gs_JSONPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.json")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            jsonfilename = fi.Name
            jsonfilenameLocation = fi.DirectoryName
            Dim jsonString As String = "{'results':" + File.ReadAllText(jsonfilenameLocation + jsonfilename) + "} "
            Dim json As JObject = JObject.Parse(jsonString)

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

                data_Supplier_Code = DirectCast(JsonSupplier_Code, JValue).Value
                data_Supplier_Name = DirectCast(JsonSupplier_Name, JValue).Value
                data_Address = DirectCast(JsonAddress, JValue).Value
                data_Phone = DirectCast(JsonPhone, JValue).Value
                data_Fax = DirectCast(JsonFax, JValue).Value
                data_Country = DirectCast(JsonCountry, JValue).Value
                data_President_Director = DirectCast(JsonPresident_Director, JValue).Value
                data_President_Director_Email = DirectCast(JsonPresident_Director_Email, JValue).Value
                data_Vice_President_Director = DirectCast(JsonVice_President_Director, JValue).Value
                data_Vice_President_Director_Email = DirectCast(JsonVice_President_Director_Email, JValue).Value
                data_Marketing_Director = DirectCast(JsonMarketing_Director, JValue).Value
                data_Marketing_Director_Email = DirectCast(JsonMarketing_Director_Email, JValue).Value
                data_Marketing_General_Manager = DirectCast(JsonMarketing_General_Manager, JValue).Value
                data_Marketing_General_Manager_Email = DirectCast(JsonMarketing_General_Manager_Email, JValue).Value
                data_Marketing_Manager = DirectCast(JsonMarketing_Manager, JValue).Value
                data_Marketing_Manager_Email = DirectCast(JsonMarketing_Manager_Email, JValue).Value
                data_Plant_Director = DirectCast(JsonPlant_Director, JValue).Value
                data_Plant_Director_Email = DirectCast(JsonPlant_Director_Email, JValue).Value
                data_Plant_Manager = DirectCast(JsonPlant_Manager, JValue).Value
                data_Plant_Manager_Email = DirectCast(JsonPlant_Manager_Email, JValue).Value
                data_Marketing_Sales = DirectCast(JsonMarketing_Sales, JValue).Value
                data_Marketing_Sales_Email = DirectCast(JsonMarketing_Sales_Email, JValue).Value
                data_Email = DirectCast(JsonEmail, JValue).Value
                data_Period_Code = DirectCast(JsonPeriod_Code, JValue).Value
                data_SupplierType = DirectCast(JsonSupplierType, JValue).Value
                data_Certification = DirectCast(JsonCertification, JValue).Value
                data_QualityAudit_Score = DirectCast(JsonQualityAudit_Score, JValue).Value
                data_DeliveryAudit_Score = DirectCast(JsonDeliveryAudit_Score, JValue).Value
                data_EHS_Value = DirectCast(JsonEHS_Value, JValue).Value
                data_PPM = DirectCast(JsonPPM, JValue).Value
                data_Mother_Company = DirectCast(JsonMother_Company, JValue).Value
                data_Product_Name = DirectCast(JsonProduct_Name, JValue).Value
                data_OEM_Customer = DirectCast(JsonOEM_Customer, JValue).Value
            Next
            'MOVE FILE
            up_JSONFile_Copy(gs_JSONPath & "\", gs_JSONPathBackup & "\", jsonfilename)

        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim di As New IO.DirectoryInfo("D:\xml")
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
        Dim fi As IO.FileInfo
        Dim j As Integer
        Dim jmlFile As Integer = aryFi.Length
        For Each fi In aryFi
            xmlfilename = fi.Name
            xmlfilenameLocation = fi.DirectoryName


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

            Next

        Next
    End Sub
  
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
End Class
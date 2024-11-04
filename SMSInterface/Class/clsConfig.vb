Imports System.Data.SqlClient
Imports System.Xml

Public Class clsConfig

#Region "Declaration"

    Private builder As SqlConnectionStringBuilder
    Private OSMbuilder As SqlConnectionStringBuilder
    Private PPCbuilder As SqlConnectionStringBuilder
    Private JKEDbuilder As SqlConnectionStringBuilder
    Private JPJbuilder As SqlConnectionStringBuilder

    'Used for SMS Database
    Private m_Server As String
    Private m_Database As String
    Private m_User As String
    Private m_Password As String
    Private m_CommandTimeout As Integer
    Private m_DatabaseTimeout As Integer
    Private m_MaxRetry As Integer
    Private m_RetryInterval As Integer
    Private m_WinMode As String
    Private m_ConnectionString As String
    Private m_LinkServer As String

    Private ls_path As String
    Dim NewEnryption As New clsDESEncryption("TOS")
#End Region

#Region "Properties"
    Public Property Server As String
        Get
            Return m_Server
        End Get
        Set(ByVal value As String)
            m_Server = value
        End Set
    End Property

    Public Property Database As String
        Get
            Return m_Database
        End Get
        Set(ByVal value As String)
            m_Database = value
        End Set
    End Property

    Public Property User As String
        Get
            Return m_User
        End Get
        Set(ByVal value As String)
            m_User = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return m_Password
        End Get
        Set(ByVal value As String)
            m_Password = value
        End Set
    End Property

    Public Property CommandTimeout As Integer
        Get
            Return m_CommandTimeout
        End Get
        Set(ByVal value As Integer)
            m_CommandTimeout = value
        End Set
    End Property

    Public Property DatabaseTimeout As Integer
        Get
            Return m_DatabaseTimeout
        End Get
        Set(ByVal value As Integer)
            m_DatabaseTimeout = value
        End Set
    End Property

    Public Property WinMode As String
        Get
            Return m_WinMode
        End Get
        Set(ByVal value As String)
            m_WinMode = value
        End Set
    End Property

    Public Property MaxRetry As Integer
        Get
            Return m_MaxRetry
        End Get
        Set(ByVal value As Integer)
            m_MaxRetry = value
        End Set
    End Property

    Public Property RetryInterval As Integer
        Get
            Return m_RetryInterval
        End Get
        Set(ByVal value As Integer)
            m_RetryInterval = value
        End Set
    End Property

#End Region

#Region "Function"
    Public Function AddSlash(ByVal Path As String) As String
        Dim Result As String = Path
        If Path.EndsWith("\") = False Then
            Result = Result + "\"
        End If
        Return Result
    End Function

    Public Function updateConfigBody(ByVal pDest As String, ByVal pProdDate As String, ByVal pExeTime As String) As Boolean

        updateConfigBody = True

        Try
            Dim MyXML As New XmlDocument()

            MyXML.Load(ls_path)

            Dim m_Dest As XmlNode = MyXML.SelectSingleNode("/Connection/BODY/Destination")
            If m_Dest IsNot Nothing Then
                m_Dest.ChildNodes(0).InnerText = NewEnryption.EncryptData(pDest)
            End If

            Dim m_ProdDate As XmlNode = MyXML.SelectSingleNode("/Connection/BODY/ProductionDate")
            If m_ProdDate IsNot Nothing Then
                m_ProdDate.ChildNodes(0).InnerText = NewEnryption.EncryptData(pProdDate)
            End If

            Dim m_ExeTime As XmlNode = MyXML.SelectSingleNode("/Connection/BODY/ExecuteTime")
            If m_ExeTime IsNot Nothing Then
                m_ExeTime.ChildNodes(0).InnerText = NewEnryption.EncryptData(pExeTime)
            End If

            MyXML.Save(ls_path)
        Catch ex As Exception
            Return False
            Throw ex
        End Try
    End Function

#End Region

#Region "Initialization"
    ''' <summary>
    ''' Open config file and store the value in local variables.
    ''' </summary>
    ''' <param name="pConfigFile"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal pConfigFile As String = "config.xml")
        'Dim Ret As String = Space(1500)

        ls_path = AddSlash(My.Application.Info.DirectoryPath) & pConfigFile

        If Not My.Computer.FileSystem.FileExists(ls_path) Then
            Throw New Exception("Config file is not found")
        End If

        'Check XML file is empty or not
        If Trim(IO.File.ReadAllText(ls_path).Length) = 0 Then Exit Sub

        'Load XML file
        Dim document = XDocument.Load(ls_path)
        'Dim newString = NewEnryption.EncryptData("IAMI_EPIC_2_20240823")
        'Console.WriteLine(newString)


        '--------------------------------- SMSDB Setting ------------------------------------------------------
        Dim SMSDB = document.Descendants("SMSDB").FirstOrDefault()
        If Not IsNothing(SMSDB) Then
            If Not IsNothing(SMSDB.Element("ServerName")) Then m_Server = NewEnryption.DecryptData(SMSDB.Element("ServerName").Value)
            If Not IsNothing(SMSDB.Element("Database")) Then m_Database = NewEnryption.DecryptData(SMSDB.Element("Database").Value)
            If Not IsNothing(SMSDB.Element("UserID")) Then m_User = NewEnryption.DecryptData(SMSDB.Element("UserID").Value)
            If Not IsNothing(SMSDB.Element("Password")) Then m_Password = NewEnryption.DecryptData(SMSDB.Element("Password").Value)
            If Not IsNothing(SMSDB.Element("WinMode")) Then m_WinMode = NewEnryption.DecryptData(SMSDB.Element("WinMode").Value)
            If Not IsNothing(SMSDB.Element("CommandTimeOut")) Then m_CommandTimeout = IIf(IsNumeric(NewEnryption.DecryptData(SMSDB.Element("CommandTimeOut").Value)) = True, NewEnryption.DecryptData(SMSDB.Element("CommandTimeOut").Value), 0)
            If Not IsNothing(SMSDB.Element("DatabaseTimeOut")) Then m_DatabaseTimeout = IIf(IsNumeric(NewEnryption.DecryptData(SMSDB.Element("DatabaseTimeOut").Value)) = True, NewEnryption.DecryptData(SMSDB.Element("DatabaseTimeOut").Value), 0) 'NewEnryption.DecryptData(SMSDB.Element("DatabaseTimeOut").Value)
            If Not IsNothing(SMSDB.Element("MaxRetry")) Then m_MaxRetry = IIf(IsNumeric(NewEnryption.DecryptData(SMSDB.Element("MaxRetry").Value)) = True, NewEnryption.DecryptData(SMSDB.Element("MaxRetry").Value), 0) 'NewEnryption.DecryptData(SMSDB.Element("MaxRetry").Value)
            If Not IsNothing(SMSDB.Element("RetryInterval")) Then m_RetryInterval = IIf(IsNumeric(NewEnryption.DecryptData(SMSDB.Element("RetryInterval").Value)) = True, NewEnryption.DecryptData(SMSDB.Element("RetryInterval").Value), 0) 'NewEnryption.DecryptData(SMSDB.Element("RetryInterval").Value)
            If Not IsNothing(SMSDB.Element("LinkServer")) Then m_LinkServer = NewEnryption.DecryptData(SMSDB.Element("LinkServer").Value)
            If m_Server = "" Or m_Database = "" Then
                Throw New Exception("Application setting is empty")
            End If

            builder = New SqlConnectionStringBuilder
            builder.DataSource = m_Server
            builder.InitialCatalog = m_Database
            builder.UserID = m_User
            builder.Password = m_Password
            builder.IntegratedSecurity = m_WinMode = "win"

            m_ConnectionString = builder.ConnectionString
        End If

    End Sub
#End Region

End Class

Imports System.IO
Imports System.Text
Imports System.Data.SqlClient

Public Class ClsInterfaceSettingDB
    Inherits DataBaseServices
    Public Shared Function getDataSMTP_DMMS_PR(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = " select  isnull(SMTP_Address, '') as SMTP_Address, isnull(SMTP_Port, '') as SMTP_Port, " & vbCrLf & _
                "           isnull(SMTP_User, '') as SMTP_User, isnull(SMTP_Password, '') as SMTP_Password, " & vbCrLf & _
                "           isnull(FromEmail, '') as FromEmail, isnull(ReceiptEmail, '') as ReceiptEmail, " & vbCrLf & _
                "           isnull(Subject, '') as Subject " & vbCrLf & _
                "   from    SMTP_Settings " & vbCrLf & _
                "   where   SystemFrom = 'DMMS' and " & vbCrLf & _
                "           Module = 'PR' "

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getData(pConStr As String) As DataTable
        Try
            Dim results = CallProcedures("GetEmailSettings")

            If Not String.IsNullOrEmpty(results.ResultMessage) Then
                Throw New ArgumentException(results.ResultMessage)
            End If


            Return results.DataResult

        Catch ex As Exception
            Console.WriteLine(ex.Message)

            Return New DataTable()
        End Try
    End Function

    Public Shared Function getData() As CallProcedureResult

        Return CallProcedures("GetEmailSettings")

    End Function

    Public Shared Function getDataApprovalSchedule(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_DataApproval"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataIAPrice(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_GET_IA_PRICE_FOR_INTERFACE"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getDataApprovalInterval(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = "exec sp_DataApproval_Interval"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function

    Public Shared Function getNextProcess(pConStr As String) As DataTable
        Dim sql As String
        Using con As New SqlConnection(pConStr)
            con.Open()

            sql = " SELECT TOP 1 Convert(char(8),[Time],108) NextProcess,* " & vbCrLf & _
                    " 	FROM sendEmail_Schedule " & vbCrLf & _
                    " 	ORDER BY [Time]  "

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Return dt
        End Using
    End Function


End Class

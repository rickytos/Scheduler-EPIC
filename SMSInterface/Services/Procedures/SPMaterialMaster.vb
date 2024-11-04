Imports System.Data.SqlClient
Imports System.Text
Imports System.Reflection

Public Class SPMaterialMaster
    Public Shared Function SAPMaterialMasterResp(ByVal pConStr As String, ByVal MaterialMasterResp As MaterialMasterRes) As ResponseStatus
        Dim sql As String
        Dim printMessage As New StringBuilder()

        Try
            Using con As New SqlConnection(pConStr)
                AddHandler con.InfoMessage, Sub(sender, e)
                                                For Each message As SqlError In e.Errors
                                                    printMessage.AppendLine("Print message: " & message.Message)
                                                Next
                                            End Sub

                con.Open()

                sql = "Upd_MaterialMaster_SAPnumber"

                Dim cmd As New SqlCommand(sql, con)
                cmd.CommandText = sql
                cmd.Connection = con
                cmd.CommandType = CommandType.StoredProcedure

                Dim properties As PropertyInfo() = GetType(MaterialMasterRes).GetProperties()

                For Each prop As PropertyInfo In properties
                    Dim paramName As String = "@" & prop.Name
                    Dim propValue As Object = prop.GetValue(MaterialMasterResp, Nothing)

                    cmd.Parameters.AddWithValue(paramName, If(propValue, DBNull.Value))
                Next

                cmd.ExecuteNonQuery()

                If printMessage.ToString().Contains("|") Then

                    Dim ResultMsg() As String = printMessage.ToString().Split("|")
                    If ResultMsg(0) = "00" Then
                        Return New ResponseStatus(True, ResultMsg(1), StaticValue.STATUSSUCCESS)
                    Else
                        Return New ResponseStatus(False, ResultMsg(1), StaticValue.STATUSERROR)
                    End If

                Else
                    Return New ResponseStatus(True, "Update Triggered !!", StaticValue.STATUSINFO)
                End If

                
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Return New ResponseStatus(False, ex.Message, 2)
        End Try
    End Function

    Public Shared Function GetMaterialMasterXML() As DataTable

        Dim ConStr = Builder.ConnectionString

        Dim sql As String
        Using con As New SqlConnection(ConStr)
            con.Open()

            sql = "exec SP_MaterialMaster_XMLtoPath"

            Dim cmd As New SqlCommand(sql, con)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)


            Return dt
        End Using

    End Function
    Public Shared Function UPDMaterialMasterReq(ByVal materialMaster As MaterialMasterReq, ByVal StatusHeader As Integer) As Integer
        Dim cmd As SqlCommand
        Dim sql As String
        Dim i As Integer

        Dim ConStr = Builder.ConnectionString

        Using con As New SqlConnection(ConStr)
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


End Class

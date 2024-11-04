Imports System.Data.SqlClient
Imports System.Reflection

Public Class DataBaseServices
    Inherits DataBaseHelper
    Public Shared Function GetConnectionString() As String
        Return Builder.ConnectionString
    End Function
    Protected Shared Function CallProcedures(ByVal procedureName As String) As CallProcedureResult
        Try
            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = procedureName

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataTable
                da.Fill(ds)

                Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataTable
            Return New CallProcedureResult With {.DataResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    Protected Shared Function CallProcedures(ByVal parameterArray As ParamCallFunction(Of String())) As CallProcedureResult
        Try
            If parameterArray.Parameter.Count < 2 Then
                Throw New Exception("Need 2 Argument When using Function with String Array Generic")
            End If

            'If parameterArray.ParamStrings.Count < 2 Then
            '    Throw New Exception("Need 2 Argument When using Function with String Array Generic")
            'End If

            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = parameterArray.StoreProcedure

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim propName As String = parameterArray.Parameter(0)
                Dim paramValue As String = parameterArray.Parameter(1)

                cmd.Parameters.AddWithValue("@" + propName, paramValue)

                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataTable
                da.Fill(ds)

                Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataTable
            result.Columns.Add("No Column")
            Return New CallProcedureResult With {.DataResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    ''' <summary>
    ''' Calls a stored procedure with the provided parameters and returns the result. Parameter are Dictionary
    ''' </summary>
    ''' <param name="parameterArray">The object that contains the parameters, and optionally parameters to exclude.</param>
    ''' <returns>A <see cref="CallProcedureResult"/> containing the result of the stored procedure or an error message if an exception occurs.</returns>
    Protected Shared Function CallProcedures(ByVal parameterArray As ParamCallFunction(Of IDictionary(Of String, Object))) As CallProcedureResult

        Try
            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = parameterArray.StoreProcedure

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim propertyParam = parameterArray.Parameter
                For Each paramKey In propertyParam.Keys
                    Dim propName = "@" & paramKey.ToString()
                    Dim propValue = propertyParam.Item(paramKey)

                    Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(paramKey)

                    If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                        cmd.Parameters.AddWithValue(propName, propValue)
                    End If
                Next

                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataTable
                da.Fill(ds)

                Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataSet
            Return New CallProcedureResult With {.DataSetResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    Protected Shared Function CallProcedures(ByVal parameterArray As ParamCallFunction(Of Dictionary(Of String, Object)), ByVal dbManager As DataBaseHelper) As CallProcedureResult

        Try

            Dim sql As String = parameterArray.StoreProcedure

            Dim cmd As New SqlCommand(sql, dbManager.SqlConnection, dbManager.SqlTransaction) ' Using the same transaction
            cmd.CommandType = CommandType.StoredProcedure

            Dim propertyParam = parameterArray.Parameter
            For Each paramKey In propertyParam.Keys
                Dim propName = "@" & paramKey.ToString()
                Dim propValue = propertyParam.Item(paramKey)

                Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(paramKey)

                If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                    cmd.Parameters.AddWithValue(propName, propValue)
                End If
            Next

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataTable
            da.Fill(ds)

            Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}
        Catch ex As Exception
            Dim result As New DataSet
            Return New CallProcedureResult With {.DataSetResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    'Getting more than two datatables
    Protected Shared Function CallProcedureSets(ByVal parameterArray As ParamCallFunction(Of IDictionary(Of String, Object))) As CallProcedureResult

        Try

            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = parameterArray.StoreProcedure

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim propertyParam = parameterArray.Parameter
                For Each paramKey In propertyParam.Keys
                    Dim propName = "@" & paramKey.ToString()
                    Dim propValue = propertyParam.Item(paramKey)

                    Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(paramKey)

                    If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                        cmd.Parameters.AddWithValue(propName, propValue)
                    End If
                Next

                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet()
                da.Fill(ds)

                Return New CallProcedureResult With {.DataSetResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataSet
            Return New CallProcedureResult With {.DataSetResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    'Using database transaction
    Protected Shared Function CallProcedures(ByVal parameterArray As ParamCallFunction(Of IDictionary(Of String, Object)), ByVal dbManager As DataBaseHelper) As CallProcedureResult
        Try
            Dim sql As String = parameterArray.StoreProcedure

            Dim cmd As New SqlCommand(sql, dbManager.SqlConnection, dbManager.SqlTransaction) ' Using the same transaction
            cmd.CommandType = CommandType.StoredProcedure

            ' Add parameters
            Dim propertyParam = parameterArray.Parameter
            For Each paramKey In propertyParam.Keys
                Dim propName = "@" & paramKey.ToString()
                Dim propValue = propertyParam.Item(paramKey)

                Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(paramKey)

                If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                    cmd.Parameters.AddWithValue(propName, propValue)
                End If
            Next

            ' Execute the command
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet()
            da.Fill(ds)

            Return New CallProcedureResult With {.DataSetResult = ds, .ErrorMessage = ""}
        Catch ex As Exception
            ' Log or handle error and return a failed result
            Dim result As New DataSet
            Return New CallProcedureResult With {.DataSetResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function


    ''' <summary>
    ''' Calls a stored procedure with the provided parameters and returns the result. Motstly need generic object for example ar T and paramter
    ''' Dictionary
    ''' </summary>
    ''' <typeparam name="T">The type of the parameter object that contains the properties matching the stored procedure parameters.</typeparam>
    ''' <param name="parameterArray">The object that contains the stored procedure name, parameters, and optionally parameters to exclude.</param>
    ''' <returns>A <see cref="CallProcedureResult"/> containing the result of the stored procedure or an error message if an exception occurs.</returns>
    Protected Shared Function CallProcedures(Of T)(ByVal parameterArray As ParamCallFunction(Of T), Optional ByVal outputParams As String = Nothing) As CallProcedureResult
        Try
            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = parameterArray.StoreProcedure
                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim props As PropertyInfo() = GetType(T).GetProperties()

                For Each prop In props
                    Dim propName = "@" + prop.Name
                    Dim propValue = prop.GetValue(parameterArray.Parameter, Nothing)

                    Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(prop.Name)

                    ' Add the parameter to the command if it is not in the exclude list and has a value
                    If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                        cmd.Parameters.AddWithValue(propName, propValue)
                    End If
                Next

                Dim ds As New DataTable

                If outputParams IsNot Nothing Then
                    Dim outputParam As New SqlParameter("@" + outputParams, SqlDbType.NVarChar, 40)
                    outputParam.Direction = ParameterDirection.Output
                    cmd.Parameters.Add(outputParam)

                    conn.Open()
                    cmd.ExecuteNonQuery()
                    conn.Close()

                    Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = "", .OutputString = cmd.Parameters("@" + outputParams).Value.ToString()}
                End If

                Dim da As New SqlDataAdapter(cmd)

                da.Fill(ds)

                Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataTable
            Return New CallProcedureResult With {.DataResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function


    'Using Database Transaction
    Protected Shared Function CallProcedures(Of T)(ByVal parameterArray As ParamCallFunction(Of T), ByVal dbManager As DataBaseHelper) As CallProcedureResult

        Try
            Dim sql As String = parameterArray.StoreProcedure

            Dim cmd As New SqlCommand(sql, dbManager.SqlConnection, dbManager.SqlTransaction)
            cmd.CommandType = CommandType.StoredProcedure

            Dim props As PropertyInfo() = GetType(T).GetProperties()

            For Each prop In props
                Dim propName = "@" + prop.Name
                Dim propValue = prop.GetValue(parameterArray.Parameter, Nothing)

                Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(prop.Name)

                ' Add the parameter to the command if it is not in the exclude list and has a value
                If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                    cmd.Parameters.AddWithValue(propName, propValue)
                End If

            Next
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataTable
            da.Fill(ds)

            Return New CallProcedureResult With {.DataResult = ds, .ErrorMessage = ""}

        Catch ex As Exception
            Dim result As New DataTable
            Return New CallProcedureResult With {.DataResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    'Getting more than two datatables
    Protected Shared Function CallProcedureSets(Of T)(ByVal parameterArray As ParamCallFunction(Of T)) As CallProcedureResult

        Try
            Using conn As New SqlConnection(GetConnectionString)
                Dim sql As String = parameterArray.StoreProcedure

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim props As PropertyInfo() = GetType(T).GetProperties()

                For Each prop In props
                    Dim propName = "@" + prop.Name
                    Dim propValue = prop.GetValue(parameterArray.Parameter, Nothing)

                    Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(prop.Name)

                    ' Add the parameter to the command if it is not in the exclude list and has a value
                    If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                        cmd.Parameters.AddWithValue(propName, propValue)
                    End If

                Next
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet()
                da.Fill(ds)

                Return New CallProcedureResult With {.DataSetResult = ds, .ErrorMessage = ""}
            End Using
        Catch ex As Exception
            Dim result As New DataTable
            Return New CallProcedureResult With {.DataResult = result, .ErrorMessage = ex.Message}
        End Try
    End Function

    Protected Shared Function CallProceduresBatchProcess(Of T)(ByVal parameterArray As ParamCallFunction(Of T)) As CallProcedureResult
        Dim dbHelper As New DataBaseHelper
        Dim transaction As SqlTransaction = Nothing

        Try
            Using conn As New SqlConnection(GetConnectionString)

                Dim sql As String = parameterArray.StoreProcedure

                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandType = CommandType.StoredProcedure

                Dim listBatch As List(Of T) = parameterArray.BatchUpdate

                For Each item In listBatch

                    Dim props As PropertyInfo() = GetType(T).GetProperties()

                    For Each prop In props
                        Dim propName = "@" + prop.Name
                        Dim propValue = prop.GetValue(item, Nothing)

                        Dim isParExcludeExists As Boolean = parameterArray.ExcludeParameter IsNot Nothing AndAlso parameterArray.ExcludeParameter.Contains(prop.Name)

                        If Not isParExcludeExists AndAlso propValue IsNot Nothing Then
                            cmd.Parameters.AddWithValue(propName, propValue)
                        End If
                    Next
                Next
                cmd.ExecuteNonQuery()

                Return New CallProcedureResult With {.ResultMessage = ""}
            End Using
        Catch ex As Exception
            Return (New CallProcedureResult With {.ErrorMessage = ex.Message})
        End Try
    End Function
End Class

#Region "PARAMETER CLASS"
Public Class ParamCallFunction(Of T)
    Public StoreProcedure As String = ""
    Public Parameter As T
    Public ParamStrings As String()
    Public BatchUpdate As List(Of T)
    Public ExcludeParameter As String() = {}
End Class

Public Class CallProcedureResult
    Public ResultCode As Integer
    Public DataResult As DataTable
    Public DataSetResult As DataSet
    Public ResultMessage As String = String.Empty
    Public ErrorMessage As String = String.Empty
    Public OutputString As String = String.Empty
End Class
#End Region

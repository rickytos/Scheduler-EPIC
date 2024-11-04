Imports System.Data.SqlClient

Public Class DataBaseHelper
    Implements IDisposable
    Public Shared cnString As String = ""
    Private conn As SqlConnection
    Private transaction As SqlTransaction
    Private transactionActive As Boolean = False ' Flag to track if a transaction is active

    Public Sub New()
        cnString = Builder.ConnectionString
        conn.Open()
        transaction = conn.BeginTransaction()
        transactionActive = True ' Mark transaction as active
    End Sub

    ' Expose the transaction and connection for other methods to use
    Public ReadOnly Property SqlConnection() As SqlConnection
        Get
            Return conn
        End Get
    End Property

    Public ReadOnly Property SqlTransaction() As SqlTransaction
        Get
            Return transaction
        End Get
    End Property

    ' Commit the transaction
    Public Sub Commit()
        If transactionActive Then
            transaction.Commit()
            transactionActive = False ' Mark transaction as no longer active
        End If
    End Sub

    ' Rollback the transaction
    Public Sub Rollback()
        If transactionActive Then
            transaction.Rollback()
            transactionActive = False ' Mark transaction as no longer active
        End If
    End Sub

    ' Dispose pattern to clean up resources
    Public Sub Dispose() Implements IDisposable.Dispose
        If transaction IsNot Nothing Then
            transaction.Dispose()
        End If

        If conn IsNot Nothing Then
            conn.Dispose()
        End If
    End Sub

    ' Property to check if a transaction is active
    Public ReadOnly Property IsTransactionActive() As Boolean
        Get
            Return transactionActive
        End Get
    End Property
End Class

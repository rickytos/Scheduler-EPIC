Public Class Utils
    Public Shared Function MoveXMLFile(ByVal pPathSource As String, ByVal pPathDestination As String, ByVal xmlfilename As String) As String
        Try
            If Not System.IO.Directory.Exists(pPathSource) Then
                System.IO.Directory.CreateDirectory(pPathSource)
            End If

            My.Computer.FileSystem.MoveFile(pPathSource & xmlfilename, pPathDestination & xmlfilename, True)
        Catch ex As Exception
            Return ex.Message
        End Try

        Return Nothing
    End Function
End Class

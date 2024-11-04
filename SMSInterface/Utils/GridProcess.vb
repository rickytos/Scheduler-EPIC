Public Class GridProcess
    Public Shared Sub upGridProcess(ByVal rtbox As RichTextBox, ByVal pColor As Integer, ByVal pPos As Integer, ByVal pMsg As String, Optional UseTime As Boolean = False)
        With rtbox
            Select Case pColor
                Case 1 : .SelectionColor = Color.Black   ' Hitam
                Case 2 : .SelectionColor = Color.Red     ' Merah
                Case 3 : .SelectionColor = Color.Green   ' Hijau
                Case 4 : .SelectionColor = Color.Blue    ' Biru
                Case 5 : .SelectionColor = Color.Gray    ' Abu
                Case 6 : .SelectionColor = Color.Orange  ' Orange
                Case 7 : .SelectionColor = Color.Purple  ' Ungu
            End Select
            Dim CurrentTime As String = Format(Date.Now, "HH:mm:ss")
            Dim Message As String
            If UseTime Then
                Message = "[" & CurrentTime & "] " & pMsg
            Else
                Message = pMsg
            End If
            .AppendText(Space(pPos) & Message & vbCrLf)
            .ScrollToCaret()
        End With
    End Sub


End Class

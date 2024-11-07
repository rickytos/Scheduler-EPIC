Public Class Mapper
    Public Shared Sub Map(Of T As New)(dt As DataTable, ByVal writer As Xml.XmlTextWriter, ByVal startElement As String)
        Try

            For iRow = 0 To dt.Rows.Count - 1
                Dim newModel As New T()
                writer.WriteStartElement(startElement)
                For Each column As DataColumn In dt.Columns
                    Dim propertyName As String = column.ColumnName

                    writer.WriteStartElement(propertyName)
                    Dim value = dt.Rows(iRow)(propertyName)

                    If Not Convert.IsDBNull(value) AndAlso Not String.IsNullOrEmpty(value.ToString()) Then
                        writer.WriteString(Trim(value.ToString()))
                    Else
                        writer.WriteString("")
                    End If
                    writer.WriteEndElement()
                Next

                writer.WriteEndElement()
            Next iRow
        Catch ex As Exception
            Console.WriteLine(ex.Message + ex.StackTrace)
        End Try
    End Sub

    Public Shared Sub Map(Of T As New)(ByVal instances As T, ByVal writer As Xml.XmlTextWriter, ByVal startElement As String)
        Try
            writer.WriteStartElement(startElement)
            For Each prop In GetType(T).GetProperties()
                Dim propName As String = prop.Name
                Dim propValue As String = GetType(T).GetProperty(propName).GetValue(instances, Nothing)

                If propValue IsNot Nothing Then
                    writer.WriteStartElement(propName)
                    writer.WriteString(propValue)
                    writer.WriteEndElement()
                End If
            Next
            writer.WriteEndElement()

        Catch ex As Exception
            Console.WriteLine(ex.Message + ex.StackTrace)
        End Try
    End Sub

    Public Shared Function MapClass(Of T As New)(ByVal xmlElement As XElement) As T
        Dim newClass As New T

        For Each prop In GetType(T).GetProperties()
            Dim propName As String = prop.Name
            Dim propValue As String = xmlElement.Element(propName).Value

            If propValue IsNot Nothing Then
                'Dim convertedValue As Object = Convert.ChangeType(propValue, prop.PropertyType)
                prop.SetValue(newClass, prop.PropertyType, Nothing)
            End If
        Next

        Return newClass
    End Function
    Public Shared Function MapClass(Of T As New)(ByVal dataTable As DataTable) As List(Of T)
        Dim classList As New List(Of T)()

        For Each row As DataRow In dataTable.Rows
            Dim classObj As New T()

            For Each prop In GetType(T).GetProperties()
                Dim propName As String = prop.Name

                If dataTable.Columns.Contains(propName) Then
                    Dim propValue As Object = row(propName)

                    prop.SetValue(classObj, propValue, Nothing)

                End If
            Next

            classList.Add(classObj)
        Next

        Return classList
    End Function

End Class
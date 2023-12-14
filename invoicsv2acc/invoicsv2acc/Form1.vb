Imports Microsoft.VisualBasic.FileIO
Imports System.Data.OleDb
Imports System.IO

Public Class Form1

    Dim csvFilePath As String = "C:\Users\Pacleb\Desktop\temporary\invoice1.csv"
    Dim cnt As Integer = 0
    Dim vHead As New ClassQto
    Public Function readByRow()

        Dim rec As New ArrayList

        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            Dim marker As Integer = 0

            While Not parser.EndOfData

                Dim fields As String() = parser.ReadFields()

                If fields IsNot Nothing AndAlso fields.Length > 0 Then
                    Dim rowList As New ArrayList

                    For Each field In fields

                        'rec.Add(field)
                        If marker = 0 Then

                            marker = 1
                            Continue For

                        End If

                        'MsgBox(field & " ")
                        If field = Nothing Then

                            rec.Add(Nothing)
                            Continue While

                        End If
                        'rowList.Add(Convert.ChangeType(field, GetType(Object)))
                        'rec.Add(convert2any(field))
                        'rec.Add(ProcessCSV(field))
                        'rec.Add(field)

                    Next


                End If

            End While

        End Using

        Return rec
    End Function

    Public Sub readByCol()

        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            Dim counter As Integer = 0
            Dim headerFields As String() = parser.ReadFields()
            Dim HF As New ClassQto
            If headerFields IsNot Nothing AndAlso headerFields.Length > 0 Then
                'read the header

                For Each headerField In headerFields

                    'MsgBox(headerField & " ")
                    HF.addCount(counter)
                    counter = counter + 1
                    HF.LAdd(headerField)

                Next

                InsertDb(HF)

            End If

        End Using

    End Sub
    Private Sub InsertDb(ByRef data As ClassQto)

        Dim accessConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Pacleb\Desktop\temporary\invoice.accdb;Persist Security Info=False;"

        Using accessConn As New OleDbConnection(accessConnStr)
            Dim rec2 As New ArrayList
            Dim count As Integer
            accessConn.Open()
            rec2 = readByRow()
            count = rec2.Count
            Label1.Text = "Transferring data of " + rec2.Count
            'ProgressBar1.Step 
            For i = 0 To data.WCount

                'Dim accessCmd As New OleDbCommand("INSERT INTO invoice(Field1, Field2, Field3) VALUES (@Field1, @Field2, @Field3)", accessConn)
                Dim accessCmd As New OleDbCommand("INSERT INTO invoice( " + data.LString(i) + ") VALUES (@Field)", accessConn)
                Dim k As Integer = 0
                'readByRow()
                For j = 0 To rec2.Count

                    If k = data.WCount Then
                        Exit For
                    End If

                    accessCmd.Parameters.AddWithValue("@Field", rec2(j))
                    Label1.Text = "Transfering data " + j + " of " + count
                    'accessCmd.Parameters.AddWithValue("@Field" + i, data.LString(i))
                    'accessCmd.Parameters.AddWithValue("@Field" + i, data.LString(i))
                Next

                accessCmd.ExecuteNonQuery()

            Next

            accessConn.Close()

        End Using

        MessageBox.Show("Data Transfer completed.")

    End Sub
    Private Sub convert2any(ByVal data As String)
        Try

            Dim targetTypes As Type() = {GetType(String), GetType(Integer), GetType(Double), GetType(Boolean), GetType(Single)}
            Dim convertedValue As Object = Nothing

            'For Each targetType As Type In targetTypes
            '  convertedValue = Convert.ChangeType(data, targetType)
            '   MessageBox.Show("Converted Value ({" & targetType.Name & "}): {" & convertedValue & "}")
            'Next

            'Dim dataType As Type = targetTypes(i)

            'If dataType = GetType(Integer) Then
            'Integer.TryParse(fields(i), convertedValue)
            'ElseIf dataType = GetType(Double) Then
            'Double.TryParse(fields(i), convertedValue)
            'ElseIf dataType = GetType(String) Then
            'convertedValue = fields(i)
            'End If

            'Return convertedValue
        Catch ex As Exception

            MessageBox.Show("Error converting the String to one or more data types.")

        End Try
    End Sub
    'Sub ProcessCSV(ByVal filePath As String, ByVal delimiter As String, ByVal dataTypes() As Type)
    ' Read and parse the CSV file
    '    Using parser As New TextFieldParser(filePath)
    '       parser.TextFieldType = FieldType.Delimited
    '      parser.SetDelimiters(delimiter)

    '     While Not parser.EndOfData
    'Dim fields() As String = parser.ReadFields()

    '           For i As Integer = 0 To fields.Length - 1
    'Dim dataType As Type = dataTypes(i)
    'Dim parsedValue As Object = Nothing

    '               If dataType = GetType(Integer) Then
    '                  Integer.TryParse(fields(i), parsedValue)
    '             ElseIf dataType = GetType(Double) Then
    '                Double.TryParse(fields(i), parsedValue)
    '           ElseIf dataType = GetType(String) Then
    '              parsedValue = fields(i)
    '         End If

    ' Do something with parsedValue (according to data type)
    '        Console.WriteLine("Column {i + 1}: {parsedValue} ({dataType.Name})")
    '   Next
    ' End While
    ' End Using
    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'readByCol()
        'readByRow()

        'Dim filePath As String = "C:\Path\To\Your\File.csv" ' Replace with your CSV file path
        Dim delimiter As String = "," ' Change the delimiter if necessary

        ' Call the function to determine data types
        Dim dataTypes() As Type = DetermineDataTypes(csvFilePath, delimiter)

        ' Display the determined data types
        For i As Integer = 0 To dataTypes.Length - 1
            'Console.WriteLine($"Column {i + 1}: {GetDataTypeName(dataTypes(i))}")
            MsgBox("Column {" & "i + 1 " & "}: {GetDataTypeName(dataTypes(i))}")
        Next

        'Console.ReadLine()

    End Sub

    Function DetermineDataTypes(ByVal filePath As String, ByVal delimiter As String) As Type()

        File.AppendAllText(filePath, """")

        Dim reader As New TextFieldParser(filePath)
        reader.TextFieldType = FieldType.Delimited
        reader.SetDelimiters(delimiter)

        Dim headers As String() = reader.ReadFields() ' Read headers to determine the number of columns
        Dim dataTypes(headers.Length - 1) As Type

        ' Initialize all data types as String by default
        For i As Integer = 0 To dataTypes.Length - 1
            dataTypes(i) = GetType(String)
        Next

        ' Loop through rows to determine data types
        While Not reader.EndOfData
            Dim fields() As String = reader.ReadFields()

            For i As Integer = 0 To fields.Length - 1
                Dim value As String = fields(i).Trim()

                ' Check if it's an Integer
                Dim intValue As Integer
                If Integer.TryParse(value, intValue) Then
                    If Type.GetTypeCode(dataTypes(i)) > TypeCode.Int32 Then
                        dataTypes(i) = GetType(Integer)
                    End If
                End If

                ' Check if it's a Single
                Dim singleValue As Single
                If Single.TryParse(value, singleValue) Then
                    If Type.GetTypeCode(dataTypes(i)) > TypeCode.Single Then
                        dataTypes(i) = GetType(Single)
                    End If
                End If

                ' Check if it's a Double
                Dim doubleValue As Double
                If Double.TryParse(value, doubleValue) Then
                    If Type.GetTypeCode(dataTypes(i)) > TypeCode.Double Then
                        dataTypes(i) = GetType(Double)
                    End If
                End If

                ' Check if it's a Date
                Dim dateValue As Date
                If Date.TryParse(value, dateValue) Then
                    If Type.GetTypeCode(dataTypes(i)) > TypeCode.DateTime Then
                        dataTypes(i) = GetType(Date)
                    End If
                End If

                ' Check if it's a Boolean
                Dim boolValue As Boolean
                If Boolean.TryParse(value, boolValue) Then
                    If Type.GetTypeCode(dataTypes(i)) > TypeCode.Boolean Then
                        dataTypes(i) = GetType(Boolean)
                    End If
                End If
            Next
        End While

        reader.Close()
        Return dataTypes
    End Function

    Function GetDataTypeName(ByVal type As Type) As String
        Select Case type.GetTypeCode(type)
            Case TypeCode.Int32
                Return "Integer"
            Case TypeCode.Single
                Return "Single"
            Case TypeCode.Double
                Return "Double"
            Case TypeCode.DateTime
                Return "Date"
            Case TypeCode.Boolean
                Return "Boolean"
            Case Else
                Return "String"
        End Select
    End Function
End Class

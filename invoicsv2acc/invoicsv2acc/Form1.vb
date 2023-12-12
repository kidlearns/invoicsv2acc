Imports Microsoft.VisualBasic.FileIO
Imports System.Data.OleDb

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
                        rec.Add(field)

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
    Private Function convert2any(ByVal data As String)
        Try

            Dim targetTypes As Type() = {GetType(String), GetType(Integer), GetType(Double), GetType(Boolean), GetType(Single)}
            Dim convertedValue As Object

            For Each targetType As Type In targetTypes
                convertedValue = Convert.ChangeType(data, targetType)
                MessageBox.Show("Converted Value ({" & targetType.Name & "}): {" & convertedValue & "}")
            Next

            Return convertedValue
        Catch ex As Exception

            MessageBox.Show("Error converting the String to one or more data types.")

        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'readByCol()
        readByRow()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class

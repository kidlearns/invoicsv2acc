Imports Microsoft.VisualBasic.FileIO
Imports System.Data.OleDb

Public Class Form1

    Dim csvFilePath As String = "C:\Users\Pacleb\Desktop\temporary\invoice1.csv"
    Dim cnt As Integer = 0
    Dim vHead As New ClassQto

    Public Function readByRow()

        Dim rec As ArrayList

        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            Dim marker As Integer = 0
            While Not parser.EndOfData

                If marker = 0 Then
                    marker = 1
                    Continue While
                End If

                Dim fields As String() = parser.ReadFields()

                If fields IsNot Nothing AndAlso fields.Length > 0 Then

                    For Each field In fields

                        rec.Add(field)
                        'MsgBox(field & " ")

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

            accessConn.Open()

            For i = 0 To data.WCount

                Dim accessCmd As New OleDbCommand("INSERT INTO invoice(Field1, Field2, Field3) VALUES (@Field1, @Field2, @Field3)", accessConn)

                accessCmd.Parameters.AddWithValue("@Field1", data.LString(i))
                accessCmd.Parameters.AddWithValue("@Field2", data.LString(i))
                accessCmd.Parameters.AddWithValue("@Field3", data.LString(i))

                accessCmd.ExecuteNonQuery()

            Next

            accessConn.Close()

        End Using

        MessageBox.Show("Data Transfer completed.")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        readByCol()
        'readByRow()

    End Sub
End Class

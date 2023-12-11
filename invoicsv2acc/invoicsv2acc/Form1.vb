Imports Microsoft.VisualBasic.FileIO
Public Class Form1
    Dim csvFilePath As String = "C:\Users\Pacleb\Desktop\temporary\invoice1.csv"
    Dim cnt As Integer = 0
    Dim vHead As New ClassQto

    Public Sub readByRow()

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

                        MsgBox(field & " ")

                    Next

                End If
            End While
        End Using

    End Sub

    Public Sub readByCol()

        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            Dim counter As Integer = 0
            Dim headerFields As String() = parser.ReadFields()

            If headerFields IsNot Nothing AndAlso headerFields.Length > 0 Then
                'read the header
                For Each headerField In headerFields

                    MsgBox(headerField & " ")

                Next

                'continue reading rest
                ' While Not parser.EndOfData
                'Dim fields As String() = parser.ReadFields()

                'If fields IsNot Nothing AndAlso fields.Length > 0 Then

                'For Each field In fields

                'MsgBox(field & " ")

                'Next

                'End If
                '   End While
            End If
        End Using

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class

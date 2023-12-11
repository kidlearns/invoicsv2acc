Public Class ClassQto
    Private counter As Integer = 0
    Private ListH As New ArrayList

    Public Sub addCount(ByVal inum As Integer)

        counter = inum

    End Sub

    Public Function WCount() As Integer

        Return counter

    End Function

    Public Sub LAdd(ByVal he As String)

        ListH.Add(he)

    End Sub

    Public Function LString(ByVal index As Integer)

        Return ListH.Item(index)

    End Function
End Class

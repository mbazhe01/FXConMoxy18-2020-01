Public Class CounterObj
    Private count As Integer

    Function getNext() As Integer
        count += 1
        Return count
    End Function

    Public Sub reset()
        count = 0

    End Sub

End Class

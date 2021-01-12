Public Class CounterObj
    Private count As Integer

    Function getNext() As Integer
        count += 1
        Return count
    End Function

End Class

Public Class MessageObj
    ' this object is used to provide user messaging & error mnessaging functions
    Dim screen As TextBox

    Sub New(ByVal aScreen As TextBox)
        screen = aScreen
    End Sub

    Sub errMsg01(ByVal methodName As String, ByVal errText As String)
        ' generic error reporting
        screen.Text += vbCrLf + String.Format("{0}: {1}", methodName, errText)
    End Sub

    Sub errMsg02(ByVal methodName As String, ByVal errText As String, ByVal var1 As String)
        screen.Text += vbCrLf + String.Format(methodName + ": " + errText + " for {0}", var1)
    End Sub

    Sub errMsg03(ByVal methodName As String, ByVal fileName As String)
        screen.Text += vbCrLf + String.Format(methodName + ": File {0} does not exist ", fileName)
    End Sub

    Sub errMsg04(ByVal methodName As String, ByVal errText As String, ByVal var1 As String)
        screen.Text += vbCrLf + String.Format(methodName + ": " + errText + " {0}", var1)
    End Sub

    Sub Msg01(ByVal msgText As String)
        ' generic message to a user
        screen.Text += vbCrLf + String.Format("{0}", msgText)
    End Sub

    Sub Msg02(ByVal bankName As String, ByVal recDate As Date)
        screen.Text += vbCrLf + String.Format("Reconciliation started for {0} as of {1}", bankName, recDate.ToString("MM/dd/yyyy"))
    End Sub

    Sub Msg03(ByVal msgText As String, ByVal asOfDate As Date)
        screen.Text += vbCrLf + String.Format("{0} : {1}", msgText, asOfDate.Date.ToString("MM/dd/yyyy"))
    End Sub

    Sub Msg04(ByVal recItem As String, ByVal bankItem As String, ByVal excelCol As String)
        screen.Text += vbCrLf + String.Format("Rec Item - {0} mapped to bank item - {1} (Column {2})", recItem, bankItem, excelCol)
    End Sub

End Class

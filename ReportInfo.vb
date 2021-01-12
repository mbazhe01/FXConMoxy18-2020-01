
Public Class ReportInfo

    Private m_AxysFile As String
    Private m_PositionDate As Date
    Private m_ErrMsg As String
    Private m_Portfolio As String
    Private m_AxysMacro As String
    Private m_OutputFile As String
    Private m_DBConStr As String
    Private m_WorkingDirectory As String
    Private m_varext As String  ' to pass variables to macros if neccessary

    Public Property DBConStr() As String
        Get
            Return m_DBConStr
        End Get
        Set(ByVal Value As String)
            m_DBConStr = Value
        End Set
    End Property

    Public Property OutputFile() As String
        Get
            Return m_OutputFile
        End Get
        Set(ByVal Value As String)
            m_OutputFile = Value
        End Set
    End Property

    Public Property AxysMacro() As String
        Get
            Return m_AxysMacro
        End Get
        Set(ByVal Value As String)
            m_AxysMacro = Value
        End Set
    End Property

    Public Property AxysFile() As String
        Get
            Return m_AxysFile
        End Get
        Set(ByVal Value As String)
            m_AxysFile = Value
        End Set
    End Property

    Public Property PositionDate() As Date
        Get
            Return m_PositionDate
        End Get
        Set(ByVal Value As Date)
            m_PositionDate = Value
        End Set
    End Property

    Public Property ErrMsg() As String
        Get
            Return m_ErrMsg
        End Get
        Set(ByVal Value As String)
            m_ErrMsg = Value
        End Set
    End Property

    Public Property Portfolio() As String
        Get
            Return m_Portfolio
        End Get
        Set(ByVal Value As String)
            m_Portfolio = Value
        End Set
    End Property

    Public Property WorkingDirectory() As String
        Get
            Return m_WorkingDirectory
        End Get
        Set(ByVal value As String)
            m_WorkingDirectory = value
        End Set
    End Property

    Public Property VarExt() As String
        Get
            Return m_varext
        End Get
        Set(ByVal value As String)
            m_varext = value
        End Set
    End Property

    Public Function validatePositionDate() As Integer
        Dim rtn As Integer = 0
        ' checks if the user provided the valid position date
        If Not IsDate(m_PositionDate) Then
            ErrMsg = vbCrLf + "Function validatePositionDate. Error:"
            ErrMsg = vbCrLf + String.Format("{0} is not a valid date", m_PositionDate)
            rtn = -1
        End If

        Return rtn
    End Function

    Public Function validatePortfolio() As Integer
        Dim rtn As Integer = 0
        ' checks if the user enterd portfolio
        If Not m_Portfolio.Length > 0 Then
            ErrMsg = vbCrLf + "Function validatePortfolio. Error:"
            ErrMsg = vbCrLf + "Please enter Axys group."
            rtn = -1
        End If

        Return rtn
    End Function
End Class

Option Strict On
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization


Public Class MoxyService

    Dim _moxyConn As String

    ' Default constructor
    Public Sub New(ByVal connStr As String)
        ' Initialization code here
        _moxyConn = connStr
    End Sub

    Public Function GetTradingRecapAllocations(storedProc As String, asOfDate As DateTime,
                                               Optional commandTimeoutSeconds As Integer = 60) As DataTable
        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(_moxyConn)
                Using cmd As New SqlCommand(storedProc, conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = commandTimeoutSeconds
                    cmd.Parameters.Add("@asofdate", SqlDbType.DateTime).Value = asOfDate

                    conn.Open()
                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(rdr)
                    End Using
                End Using
            End Using
        Catch ex As SqlException
            Throw New ApplicationException("DB error in GetTradingRecapAllocations.", ex)
        Catch ex As Exception
            Throw
        End Try

        Return dt
    End Function

    ' Runs the stored proc and returns strongly-typed Transaction objects
    Public Function GetTradingRecapInternationalTransactions(storedProc As String,
                                                asOfDate As Date,
                                                Optional commandTimeoutSeconds As Integer = 60) As List(Of MoxyAllocTran)
        Dim results As New List(Of MoxyAllocTran)()

        Try
            Using conn As New SqlConnection(_moxyConn)
                Using cmd As New SqlCommand(storedProc, conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = commandTimeoutSeconds
                    cmd.Parameters.Add("@asofdate", SqlDbType.DateTime).Value = asOfDate

                    conn.Open()
                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                        While rdr.Read()

                            results.Add(MapTransaction(rdr))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As SqlException
            Throw New ApplicationException($"DB error in {NameOf(GetTradingRecapInternationalTransactions)}.", ex)
        Catch e As Exception
            Throw New Exception("MoxyService Exception: " + e.Message)
        End Try

        Return results
    End Function

    ' ---- Helpers ----
    ' --- Safer ordinal lookup (case-insensitive via ADO.NET) ---
    Private Shared Function GetOrdinalSafe(r As IDataRecord, columnName As String) As Integer
        Try
            Return r.GetOrdinal(columnName) ' Tries exact then case-insensitive
        Catch ex As IndexOutOfRangeException
            Return -1
        End Try
    End Function

    Private Shared Function GetString(r As SqlDataReader, columnName As String, Optional def As String = "") As String
        Dim i As Integer = GetOrdinalSafe(r, columnName)
        If i = -1 OrElse r.IsDBNull(i) Then Return def
        Dim v As Object = r.GetValue(i)
        Return If(v IsNot Nothing, Convert.ToString(v, Globalization.CultureInfo.InvariantCulture), def)
    End Function

    Private Shared Function GetDecimal(r As SqlDataReader, columnName As String, Optional def As Decimal = 0D) As Decimal
        Dim i As Integer = GetOrdinalSafe(r, columnName)
        If i = -1 OrElse r.IsDBNull(i) Then Return def
        Dim v As Object = r.GetValue(i)
        Try
            If TypeOf v Is Decimal Then Return CDec(v)
            If TypeOf v Is Double Then Return CDec(CDbl(v))
            If TypeOf v Is Single Then Return CDec(CSng(v))
            If TypeOf v Is String Then
                    Return Decimal.Parse(CStr(v), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture)
                End If
                Return Convert.ToDecimal(v, Globalization.CultureInfo.InvariantCulture)
    Catch
            Return def
        End Try
    End Function

    ' Use an explicit default sentinel instead of Nothing to avoid ambiguity.
    Private Shared Function GetDate(r As SqlDataReader,
                                columnName As String,
                                Optional def As Date = #1/1/0001 12:00:00 AM#) As Date
        Dim i As Integer = GetOrdinalSafe(r, columnName)
        If i = -1 OrElse r.IsDBNull(i) Then Return def

        Dim v As Object = r.GetValue(i)

        ' Fast paths for native types
        If TypeOf v Is DateTime Then Return CType(v, DateTime)
        If TypeOf v Is Date Then Return CType(v, Date)

        ' Handle datetimeoffset
        If TypeOf v Is DateTimeOffset Then
            ' Pick one: .UtcDateTime, .LocalDateTime, or .DateTime (no offset).
            ' In most trading backends we want the wall-clock local date from the DB:
            Return CType(v, DateTimeOffset).DateTime
        End If

        ' Handle SqlDateTime (rare with SqlDataReader.GetValue, but safe)
        If TypeOf v Is System.Data.SqlTypes.SqlDateTime Then
            Return CType(CType(v, System.Data.SqlTypes.SqlDateTime).Value, Date)
        End If

        ' Strings or anything convertible
        If TypeOf v Is String Then
            Dim s As String = DirectCast(v, String).Trim()
            Dim dt As DateTime
            ' Try ISO first, then general invariant parse
            If DateTime.TryParseExact(s,
                                  {"o", "s", "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss"},
                                  Globalization.CultureInfo.InvariantCulture,
                                  Globalization.DateTimeStyles.AssumeLocal,
                                  dt) Then
                Return dt
            End If
            If DateTime.TryParse(s,
                             Globalization.CultureInfo.InvariantCulture,
                             Globalization.DateTimeStyles.AssumeLocal,
                             dt) Then
                Return dt
            End If
            ' As a last resort try current culture
            If DateTime.TryParse(s,
                             Globalization.CultureInfo.CurrentCulture,
                             Globalization.DateTimeStyles.AssumeLocal,
                             dt) Then
                Return dt
            End If
            Return def
        End If

        ' Fallback: attempt a generic conversion; if it fails, return default sentinel
        Try
            Return Convert.ToDateTime(v, Globalization.CultureInfo.InvariantCulture)
        Catch
            Return def
        End Try
    End Function


    ' ---- Mapping (NULL-safe, case-insensitive column lookup) ----
    Private Shared Function MapTransaction(rdr As SqlDataReader) As MoxyAllocTran
        Dim allocation As MoxyAllocTran = Nothing

        Try
            Dim formatString As String = "yyyyMMdd"
            allocation = New MoxyAllocTran()

            ' Map string and integer values
            allocation.OrderId = rdr.GetInt32(rdr.GetOrdinal("orderid"))
            allocation.PortId = rdr.GetString(rdr.GetOrdinal("PortId"))
            allocation.PortName = rdr.GetString(rdr.GetOrdinal("PortName"))
            allocation.ShortName = rdr.GetString(rdr.GetOrdinal("ShortName"))
            allocation.ISIN = rdr.GetString(rdr.GetOrdinal("ISIN"))
            allocation.Cusip = rdr.GetString(rdr.GetOrdinal("Cusip"))
            allocation.TranCode = rdr.GetString(rdr.GetOrdinal("TranCode"))

            ' Map TradeDate with error handling
            Dim tradeDateStr As String = rdr.GetString(rdr.GetOrdinal("TradeDate"))
            allocation.TradeDate = DateTime.ParseExact(tradeDateStr, formatString, System.Globalization.CultureInfo.InvariantCulture)

            ' Map decimal values
            Dim i As Integer = rdr.GetOrdinal("AllocQty")

            'allocation.AllocQty = rdr.GetDecimal(i)
            ' First, check if the value is a database NULL
            If rdr.IsDBNull(i) Then
                allocation.AllocQty = 0
            Else
                ' The value is not NULL. Now, check if it can be safely converted.
                ' This handles cases where the column is a string with non-numeric data.
                Dim valueToParse As Object = rdr.GetValue(i)

                If Not Decimal.TryParse(valueToParse.ToString(), allocation.AllocQty) Then
                    ' This block is executed if the conversion fails.
                    ' It means the value was not a valid number (e.g., "N/A", "95 ")
                    allocation.AllocQty = 0 ' Assign a default or log the error
                End If
            End If

            allocation.AllocPrice = GetDecimalValue(rdr, "AllocPrice")

            'allocation.AllocPrice = rdr.GetDecimal(rdr.GetOrdinal("AllocPrice"))
            allocation.Principal = GetDecimalValue(rdr, "Principal")

            'DELETE ME
            If allocation.PortId = "24524" Then
                allocation.Principal += 1
            End If

            allocation.Commission = GetDecimalValue(rdr, "Commission")
            allocation.SECFee = GetDecimalValue(rdr, "SECFee")
            allocation.OtherFee = GetDecimalValue(rdr, "OtherFee")
            allocation.TotalAmount = GetDecimalValue(rdr, "TotalAmount")
            allocation.Taxes = GetDecimalValue(rdr, "Taxes")

            ' Map remaining string values
            allocation.Broker = rdr.GetString(rdr.GetOrdinal("Broker"))

            ' Map final decimal and date values
            allocation.TktChrg = GetDecimalValue(rdr, "TktChrg")
            allocation.FXRate = GetDecimalValue(rdr, "FXRate")

            ' Map SettleDate with error handling
            Dim settleDateStr As String = rdr.GetString(rdr.GetOrdinal("SettleDate"))
            allocation.SettleDate = DateTime.ParseExact(settleDateStr, formatString, System.Globalization.CultureInfo.InvariantCulture)

            ' Map final string value
            allocation.SecurityCurrency = rdr.GetString(rdr.GetOrdinal("SecurityCurrency"))

            Return allocation

        Catch ex As Exception
            ' This is a more robust way to handle the error
            Dim errorMessage As String = $"Error mapping transaction with OrderId '{rdr.GetValue(rdr.GetOrdinal("orderid"))}'. A conversion failed: {ex.Message}"

            ' Re-throw the exception with the new, more descriptive message
            Throw New Exception(errorMessage, ex)
        End Try

    End Function


    Private Shared Function GetDecimalValue(ByVal rdr As SqlDataReader, ByVal columnName As String) As Decimal
        ' Get the column ordinal
        Dim i As Integer = rdr.GetOrdinal(columnName)
        Dim outputValue As Decimal

        ' Check for DBNull first, then try to parse the value
        If rdr.IsDBNull(i) Then
            outputValue = 0
        Else
            Dim valueToParse As Object = rdr.GetValue(i)

            If Not Decimal.TryParse(valueToParse.ToString(), outputValue) Then
                ' Handle the parsing failure
                ' Log this error or take other appropriate action
                ' For now, we'll assign a default of 0
                outputValue = 0
            End If
        End If

        Return outputValue
    End Function
End Class

Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.TextBox
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel


Public Class FXConManager
    Protected fName As String

    Protected screen As System.Windows.Forms.TextBox
    Protected moxyCon As String
    Protected axysCurrency As String
    Protected hdgDT As System.Data.DataTable ' hedge exposure dt
    Protected repInfo As ReportInfo
    Protected mo As MessageObj

    Sub New(ByVal aFileName As String, ByRef txtBox As System.Windows.Forms.TextBox, ByVal connStr As String, ByVal axysCurr As String)

        fName = aFileName
        screen = txtBox
        moxyCon = connStr
        axysCurrency = axysCurr
        mo = New MessageObj(txtBox)



    End Sub

    Public Function createNewGTSSFile() As Integer

        Dim file As System.IO.FileStream
        file = System.IO.File.Create(fName)

        file.Close()
    End Function

    Public Function addFXConTrades(ByVal DT As Data.DataTable) As Integer
        Dim rtn As Integer = 0
        ' chr(9) returns TAB
        Dim fw As StreamWriter
        Dim sHeader As String
        Dim sTradeLine As String
        Dim drow As DataRow
        Dim drows() As DataRow = DT.Select
        Dim cnt As Integer
        Dim newSeed As Integer
        'Dim currency As String
        Dim amountSold As Double
        Dim tmp As String
        Dim curr As String
        Dim curr2 As String
        Dim localAmount As String
        Dim usdAmount As String
        Dim bicCode As String
        Dim ConversionInstruction As String
        Dim DeliveryAgent As String = String.Empty
        Dim RecievingAgent As String = "/ABIC/UKWN|/NAME/UKWN"
        Dim RecievingAgent2 As String = String.Empty
        Dim rowNum As Integer
        Dim orderbrokerid As String
        Dim fxRate As String
        Dim crossRate As String
        Dim custID As String

        Dim firstAmt As String
        Dim secondAmt As String
        Dim firstCur As String
        Dim secondCur As String
        Dim portfolio As String

        Try

            'Pass the file path and name to the StreamWriter constructor.
            fw = File.CreateText(fName)
            fw.WriteLine("300")
            sHeader = createHeaderLine()
            fw.WriteLine(sHeader)
            For Each drow In drows

                If drow(2).ToString = "24132" Then
                    portfolio = drow(2).ToString
                End If

                rowNum = rowNum + 1
                newSeed = getTranIDSeed(moxyCon)

                If Not newSeed > 0 Then
                    screen.Text += vbCrLf + "Could not generate new seed number"
                    Return -1
                End If

                bicCode = drow("BicCode").ToString
                If Not bicCode.Length > 0 Then
                    screen.Text += vbCrLf + "No BIC Code found for " + drow("orderbrokerid")
                End If

                firstAmt = drow("amount")

                'If getDisplayDirect(drow("sectype").ToString) = "y" Then
                If Len(drow("crossrate")) > 0 Then
                    amountSold = drow("amount")
                Else
                    amountSold = drow("amount") * drow("tradefxrate")
                End If

                'Else
                'amountSold = drow("amount") / drow("tradefxrate")
                'End If
                ' when tran code is ss (Short Sell) swap localAmount and usdAmount

                localAmount = Math.Round(drow("amount"), 2).ToString()

                ' MB 3/4/2021
                usdAmount = Math.Round(amountSold, 2).ToString

                Dim secType As String = drow("sectype")

                secType = secType.Substring(2, 2)

                curr = getCurrency(drow("sectype"))
                If curr = Nothing Then
                    screen.Text += vbCrLf + "Unable to extract currency from security type: " + drow("sectype") + " Row#: " + rowNum.ToString + vbCrLf
                End If

                If Len(drow("tradecurrency").ToString) > 0 Then
                    curr2 = drow("tradecurrency").ToString
                Else
                    curr2 = "USD" ' this will change if it's non us based account
                End If

                ' If there is a cross rate apply it to local amount
                If Len(drow("crossrate")) > 0 And UCase(curr2) <> "USD" Then

                    ConversionInstruction = getConversionInstruction(curr, curr2)
                    If ConversionInstruction = "m" Then
                        ' multiply by cross rate
                        'localAmount = Math.Round(localAmount * CDbl(drow("crossrate")), 2).ToString
                        usdAmount = Trim(Math.Round(usdAmount * CDbl(drow("crossrate")), 2).ToString)
                    ElseIf ConversionInstruction = "d" Then
                        ' divide by cross rate
                        'localAmount = Math.Round(localAmount / CDbl(drow("crossrate")), 2).ToString
                        usdAmount = Trim(Math.Round(localAmount / CDbl(drow("crossrate")), 2).ToString)
                    Else
                        screen.Text += vbCrLf + "Undefine conversion instruction for " + curr + Space(1) + curr2 + Space(1) + " currencies. Row# : " + rowNum.ToString
                        Return -1
                    End If

                End If

                If drow("trancode").ToString() = "sl" Then
                    tmp = localAmount
                    localAmount = usdAmount
                    usdAmount = tmp
                    tmp = ""
                    tmp = curr
                    curr = curr2
                    curr2 = tmp
                    tmp = ""
                    'End If

                End If

                orderbrokerid = drow("orderbrokerid")

                '  Get Delivery Agent
                If getAgent(drow("orderbrokerid"), curr, DeliveryAgent) = -1 Then
                    Return -1
                Else
                    If DeliveryAgent.Length = 0 Then
                        screen.Text += vbCrLf + "Undefined delivery agent info found for " + drow("orderbrokerid") + Space(1) + curr
                        Return -1
                    End If
                End If
                'End If

                ' Get Receiving Agent
                If getAgent(drow("orderbrokerid"), curr2, RecievingAgent2) = -1 Then
                    Return -1
                Else
                    If RecievingAgent2.Length = 0 Then
                        screen.Text += vbCrLf + "Undefined receiving agent info found for " + drow("orderbrokerid") + Space(1) + curr2
                        Return -1
                    End If
                End If
                'End If

                crossRate = drow("crossrate")

                If Not Len(drow("crossrate")) > 0 Then
                    ' This could be applied only to US based accounts
                    If getDisplayDirect(drow("sectype").ToString) = "y" Then
                        fxRate = Math.Round(drow("tradefxrate"), 7).ToString()
                    Else
                        fxRate = Math.Round(1 / drow("tradefxrate"), 7).ToString()
                    End If
                Else
                    ' for non us based accounts show crossrate
                    fxRate = drow("crossrate").ToString()
                End If
                'Else
                ' for non us based accounts show crossrate
                fxRate = drow("crossrate").ToString()
                'End If

                ' get a custodian id 
                custID = Nothing

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' new calc 
                firstCur = getCurrency(drow("sectype"))
                secondCur = drow("tradecurrency").ToString
                If secondCur = "" Then
                    secondCur = "USD"
                End If

                ConversionInstruction = getConversionInstruction(firstCur, secondCur)
                If ConversionInstruction = "m" Then
                    secondAmt = Trim(Math.Round(firstAmt * CDbl(drow("crossrate")), 2).ToString)
                ElseIf ConversionInstruction = "d" Then
                    secondAmt = Trim(Math.Round(firstAmt / CDbl(drow("crossrate")), 2).ToString)
                Else
                    screen.Text += vbCrLf + "Undefine conversion instruction for " + firstCur + Space(1) + secondCur + Space(1) + " currencies. Row# : " + rowNum.ToString
                    Return -1
                End If
                If drow("trancode").ToString() = "sl" Then
                    tmp = firstAmt
                    firstAmt = secondAmt
                    secondAmt = tmp
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                cnt += 1
                sTradeLine = "TWEEUSNYXXX" + Chr(9) + bicCode + Chr(9) + "15A" + Chr(9)
                sTradeLine += newSeed.ToString() + Chr(9) + "" + Chr(9) + "NEWT" + Chr(9)
                sTradeLine += "" + Chr(9) + bicCode + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + "TWEEUSNYXXX" + Chr(9) + bicCode + Chr(9)
                sTradeLine += "" + Chr(9) + "" + Chr(9) + "15B" + Chr(9)
                sTradeLine += drow("tradedate").ToString() + Chr(9) + drow("settledate").ToString() + Chr(9) + fxRate + Chr(9)
                sTradeLine += UCase(curr) + firstAmt + Chr(9) + DeliveryAgent + Chr(9) + "" + Chr(9)
                sTradeLine += RecievingAgent + Chr(9) + UCase(curr2) + secondAmt + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + RecievingAgent2 + Chr(9) + "/ABIC/" + bicCode + "|/NAME/UKWN" + Chr(9)
                sTradeLine += "15C" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                If Trim(custID) <> "RTST" Then
                    sTradeLine += "/GLCID/" + drow("portfolio").ToString
                Else
                    sTradeLine += "/GLCID/" + drow("portfolio").ToString + "/SPOT/"
                End If
                sTradeLine += Chr(9) + "" + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                sTradeLine += "-"
                If (firstCur <> secondCur) Then
                    fw.WriteLine(sTradeLine)
                End If
                sTradeLine = ""
            Next
            fw.WriteLine("<<" + cnt.ToString + ">>")
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
        Finally
            'Close the file.
            fw.Close()
        End Try
        Return rtn
    End Function
    Private Function createFooterLine(ByVal numOfTrades As Integer) As String
        Return "<<" + numOfTrades.ToString + ">>"
    End Function
    Protected Function createHeaderLine() As String
        ' Note: chr(34) returns double quote

        Dim sHeader As String

        sHeader = "FromBIC" + Chr(9) + "ToBIC" + Chr(9) + "New Sequence 15A" + Chr(9)
        sHeader += "Sender's Reference" + Chr(9) + "Related Reference" + Chr(9) + "Type of operation" + Chr(9)
        sHeader += "Scope of operation" + Chr(9) + "Common Reference" + Chr(9) + "Block Trade Indicator" + Chr(9)
        sHeader += "Split Settlement" + Chr(9) + "Party A" + Chr(9) + "Party B" + Chr(9)
        sHeader += "Fund or Beneficiary" + Chr(9) + "Terms and Conditions" + Chr(9) + "New Sequence 15B" + Chr(9)
        sHeader += "Trade Date" + Chr(9) + "Value Date" + Chr(9) + "Exchange Rate" + Chr(9)
        sHeader += "CurrencyAmount Bought" + Chr(9) + "Delivery Agent" + Chr(9) + "Intermediary" + Chr(9)
        sHeader += "Receiving Agent" + Chr(9) + "CurrencyAmount Sold" + Chr(9) + "Delivery Agent" + Chr(9)
        sHeader += "Intermediary" + Chr(9) + "Receiving Agent" + Chr(9) + "Beneficiary Institution" + Chr(9)
        sHeader += "New Sequence 15C" + Chr(9) + "Contact Information" + Chr(9) + "Dealing Method" + Chr(9)
        sHeader += "Dealing Branch Party A" + Chr(9) + "Dealing Branch Party B" + Chr(9) + "Broker Id" + Chr(9)
        sHeader += "Broker's Commission" + Chr(9) + "Counterparty's Reference" + Chr(9) + "Broker's Reference" + Chr(9)
        sHeader += "Sender to Receiver" + Chr(9) + "New Sequence 15D" + Chr(9) + "Buy/Sell Indicator" + Chr(9)
        sHeader += "CurrencyAmount" + Chr(9) + "Delivery Agent" + Chr(9) + "Intermediary" + Chr(9)
        sHeader += "Receiving Agent" + Chr(9) + "Beneficiary Institution" + Chr(9) + "Number of Settlements"


        Return sHeader
    End Function

    Public Shared Function getTranIDSeed(ByVal connStr As String) As Integer
        Dim Conn As New SqlConnection(connStr)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetTranIDSeed", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet

        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue

        DA.SelectCommand = Cmd
        Try
            Conn.Open()
            DA.Fill(DSet, "seeds")

        Catch ex As Exception
            Throw New Exception("Failed to retrieve new seed." + vbCrLf)

            Return -1
        Finally
            Conn.Close()
        End Try

        Return RetValue.Value
    End Function

    Protected Function getCurrency(ByVal secType As String) As String
        Dim currency As String

        Select Case Strings.Right(secType, 2)
            Case "ca"
                currency = "CAD"
            Case "dk"
                currency = "DKK"
            Case "hk"
                currency = "HKD"
            Case "jp"
                currency = "JPY"
            Case "mx"
                currency = "MXN"
            Case "my"
                currency = "MYR"
            Case "nz"
                currency = "NZD"
            Case "no"
                currency = "NOK"
            Case "sg"
                currency = "SGD"
            Case "se"
                currency = "SEK"
            Case "ch"
                currency = "CHF"
            Case "th"
                currency = "THB"
            Case "gb"
                currency = "GBP"
            Case "eu"
                currency = "EUR"
            Case "za"
                currency = "ZAR"
            Case "au"
                currency = "AUD"
            Case "cs"
                currency = "CZK"
            Case "bs"
                currency = "BSD"
            Case "lu"
                currency = "LUF"
            Case "vg"
                currency = "VGD"
            Case "kr"
                currency = "KRW"
            Case "ar"
                currency = "ARS"
            Case "br"
                currency = "BRL"
            Case "cl"
                currency = "CLP"
            Case "cn"
                currency = "CNY"
            Case "eg"
                currency = "EGP"
            Case "gr"
                currency = "GRD"
            Case "hu"
                currency = "HUF"
            Case "in"
                currency = "INR"
            Case "ph"
                currency = "PHP"
            Case "pl"
                currency = "PLZ"
            Case "tr"
                currency = "TRL"
            Case "us"
                currency = "USD"
            Case "il"
                currency = "ILS"
            Case "tw"
                currency = "TWN"


        End Select

        Return currency
    End Function

    Protected Function getDisplayDirect(ByVal secType As String) As String
        Dim disp As String
        Select Case Strings.Right(secType, 2)
            Case "ca"
                disp = "n"
            Case "dk"
                disp = "n"
            Case "hk"
                disp = "n"
            Case "jp"
                disp = "n"
            Case "mx"
                disp = "n"
            Case "my"
                disp = "n"
            Case "nz"
                disp = "y"
            Case "no"
                disp = "n"
            Case "sg"
                disp = "n"
            Case "se"
                disp = "n"
            Case "ch"
                disp = "n"
            Case "th"
                disp = "n"
            Case "gb"
                disp = "y"
            Case "eu"
                disp = "y"
            Case "za"
                disp = "n"
            Case "au"
                disp = "y"
            Case "cs"
                disp = "n"
            Case "bs"
                disp = "y"
            Case "lu"
                disp = "n"
            Case "vg"
                disp = "n"
            Case "kr"
                disp = "n"
            Case "ar"
                disp = "n"
            Case "br"
                disp = "n"
            Case "cl"
                disp = "n"
            Case "cn"
                disp = "n"
            Case "eg"
                disp = "n"
            Case "gr"
                disp = "n"
            Case "hu"
                disp = "n"
            Case "in"
                disp = "n"
            Case "ph"
                disp = "n"
            Case "pl"
                disp = "n"
            Case "tr"
                disp = "n"
            Case "us"
                disp = "n"
            Case "hr"
                disp = "n"
            Case "il"
                disp = "n"
        End Select

        Return disp
    End Function

    Protected Function flipFXRate(ByVal curr As String) As String
        Dim disp As String
        Select Case UCase(curr)
            Case "CAD"
                disp = "n"
            Case "DKK"
                disp = "n"
            Case "HKD"
                disp = "n"
            Case "JPY"
                disp = "n"
            Case "MNX"
                disp = "n"
            Case "MYR"
                disp = "n"
            Case "NZD"
                disp = "y"
            Case "NOK"
                disp = "n"
            Case "SGD"
                disp = "n"
            Case "SEK"
                disp = "n"
            Case "CHF"
                disp = "n"
            Case "THB"
                disp = "n"
            Case "GBP"
                disp = "y"
            Case "EUR"
                disp = "y"
            Case "ZAR"
                disp = "n"
            Case "AUD"
                disp = "y"
            Case "CZK"
                disp = "n"
            Case "BSD"
                disp = "y"
            Case "LUF"
                disp = "n"
            Case "VGD"
                disp = "n"
            Case "KRW"
                disp = "n"
            Case "ARS"
                disp = "n"
            Case "BRL"
                disp = "n"
            Case "CLP"
                disp = "n"
            Case "CNY"
                disp = "n"
            Case "EGP"
                disp = "n"
            Case "GRD"
                disp = "n"
            Case "HUF"
                disp = "n"
            Case "INR"
                disp = "n"
            Case "PHP"
                disp = "n"
            Case "PLZ"
                disp = "n"
            Case "TRL"
                disp = "n"
            Case "USD"
                disp = "n"
            Case "HRD"
                disp = "n"
            Case "ILS"
                disp = "n"
            Case Else
                disp = "n"
        End Select

        Return disp
    End Function

    Public Shared Function validateTRNFile(ByVal fName As String, ByVal screen As System.Windows.Forms.TextBox) As Boolean
        ' validates TRN files before import

        Try
            If Not File.Exists(fName) Then
                screen.Text += vbCrLf + "File " + fName + " not found"
                Return False
            End If

            If New FileInfo(fName).Length = 0 Then
                screen.Text += vbCrLf + "File " + fName + " is empty"
                Return False

            End If
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
        End Try

        Return True
    End Function

    Public Function createAxysFXTradesFile(ByVal inFile As String, ByVal outFile As String) As Integer
        ' this functions reads through HTML file produced by Axys fxconn.mac , parse it and
        ' creates a text output file in GTSS format
        Dim portCode As String
        Dim tranCode As String
        Dim curr As String
        Dim tradeDate As String
        Dim settleDate As String
        Dim settleDate2 As Date
        Dim localAmount As String
        Dim usdAmount As String
        Dim fxRate As String
        Dim trTag As String  ' Table Row Tag
        Dim fw As StreamWriter
        Dim sHeader As String
        Dim sTradeLine As String
        Dim cnt As Integer
        Dim newSeed As Integer
        Dim tmp As String
        Dim curr2 As String
        Dim bicCode As String = ""
        Dim brokerCode As String
        Dim deliveryAgent As String
        Dim receivingAgent As String = "/NETS/"
        'Dim receivingAgent As String = "/ABIC/UKWN|/NAME/UKWN"
        Dim recievingAgent2 As String
        Dim fixingDate As String

        'Dim row As DataRow
        'Dim rows() As DataRow

        trTag = "S0-Detail"
        Try
            Dim sr As StreamReader = New StreamReader(inFile)
            Dim line As String

            fw = File.CreateText(fName)
            fw.WriteLine("300")
            sHeader = createHeaderLine()
            fw.WriteLine(sHeader)
            Do
                line = sr.ReadLine()
                If line Is Nothing Then Exit Do
                If line.IndexOf(trTag) > 0 Then
                    'new table row dound: now read columns
                    line = sr.ReadLine()

                    Do While line.IndexOf("<td align") > 0
                        'new data column found
                        'line = sr.ReadLine 'data line #1
                        line = sr.ReadLine 'data line #2
                        If extractString(line, portCode) = -1 Then
                            Return -1
                        End If
                        line = sr.ReadLine 'data line #3
                        line = sr.ReadLine 'data line #4
                        If extractString(line, tranCode) = -1 Then Return -1
                        line = sr.ReadLine 'data line #5
                        line = sr.ReadLine 'data line #6
                        If extractString(line, curr) = -1 Then Return -1
                        line = sr.ReadLine 'data line #7
                        line = sr.ReadLine 'data line #8
                        If extractString(line, tradeDate) = -1 Then Return -1
                        tradeDate = formatDate(tradeDate)
                        line = sr.ReadLine 'data line #9
                        line = sr.ReadLine 'data line #10
                        If extractString(line, settleDate) = -1 Then Return -1
                        settleDate2 = CDate(settleDate)
                        settleDate = formatDate(settleDate)
                        line = sr.ReadLine  'data line #11
                        If extractString(line, localAmount) = -1 Then Return -1
                        localAmount = formatAmount(localAmount)
                        line = sr.ReadLine  'data line #12
                        If extractString(line, usdAmount) = -1 Then Return -1
                        usdAmount = formatAmount(usdAmount)
                        line = sr.ReadLine  'data line #13
                        If extractString(line, fxRate) = -1 Then Return -1
                        ' Axys can not pick up the right exchange rate.
                        ' We have to calculate it
                        fxRate = Math.Round(CDbl(localAmount) / CDbl(usdAmount), 9).ToString("##0.000000000")
                        If flipFXRate(curr) = "y" Then
                            fxRate = (1 / CDbl(fxRate)).ToString("##0.000000000")
                        End If
                        line = sr.ReadLine 'data line #14
                        If extractString(line, brokerCode) = -1 Then Return -1

                        System.Windows.Forms.Application.DoEvents()
                        line = sr.ReadLine

                        ' add a trade to FC Trades file
                        newSeed = getTranIDSeed(moxyCon)

                        ' when tran code is ss (Short Sell) swap localAmount and usdAmount
                        curr2 = axysCurrency

                        If tranCode = "ss" Then
                            tmp = localAmount
                            localAmount = usdAmount
                            usdAmount = tmp
                            tmp = ""
                            tmp = curr
                            curr = curr2
                            curr2 = tmp
                            tmp = ""
                        End If

                        ' exclude KRW & HRK  ==> UCase(curr) = "KRW" Or 
                        If UCase(curr) = "HRK" Then
                            GoTo Endofloop
                        End If

                        If getBICCode(brokerCode, bicCode) = -1 Then
                            Return -1
                        End If

                        '  Get Delivery Agent
                        If getAgent(brokerCode, curr, deliveryAgent) = -1 Then
                            Return -1
                        Else
                            If deliveryAgent.Length = 0 Then
                                screen.Text += vbCrLf + "Undefined delivery agent info found for " + brokerCode + Space(1) + curr
                                Return -1
                            End If
                        End If

                        ' Get Receiving Agent
                        If getAgent(brokerCode, curr2, recievingAgent2) = -1 Then
                            Return -1
                        Else
                            If recievingAgent2.Length = 0 Then
                                screen.Text += vbCrLf + "Undefined recieving agent info found for " + brokerCode + Space(1) + curr2
                                Return -1
                            End If
                        End If

                        'Get Fixing Date
                        If UCase(curr) = "KRW" Or UCase(curr2) = "KRW" Then
                            fixingDate = getFixingDate("KRW", settleDate2)
                            If fixingDate = Nothing Then fixingDate = String.Empty

                        End If

                        cnt += 1

                        If Not newSeed > 0 Then Return -1
                        sTradeLine = "TWEEUSNYXXX" + Chr(9) + bicCode + Chr(9) + "15A" + Chr(9)
                        sTradeLine += newSeed.ToString() + Chr(9) + "" + Chr(9) + "NEWT" + Chr(9)
                        sTradeLine += "" + Chr(9) + Strings.Left(bicCode, 6) + "TWEEUS" + cnt.ToString() + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + "TWEEUSNYXXX" + Chr(9) + bicCode + Chr(9)
                        sTradeLine += "" + Chr(9) + fixingDate + Chr(9) + "15B" + Chr(9)
                        sTradeLine += tradeDate + Chr(9) + settleDate + Chr(9) + fxRate + Chr(9)
                        sTradeLine += UCase(curr) + localAmount + Chr(9) + deliveryAgent + Chr(9) + "" + Chr(9)
                        sTradeLine += receivingAgent + Chr(9) + UCase(curr2) + usdAmount + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + recievingAgent2 + Chr(9) + "/ABIC/" + bicCode + "/NAME/UKWN" + Chr(9)
                        sTradeLine += "15C" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "/GLCID/" + Strings.Left(portCode, 5) + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
                        sTradeLine += "-"

                        fw.WriteLine(sTradeLine)
                        sTradeLine = ""
                        fixingDate = Nothing

Endofloop:          Loop

                End If

            Loop Until line Is Nothing
            sr.Close()
            fw.WriteLine("<<" + cnt.ToString + ">>")
            fw.Close()

            screen.Text += vbCrLf + "Finished writing to " + fName

        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message + vbCrLf
            screen.Text += "Error --> Portfolio: " + portCode + Space(1) + "TranCode:" + tranCode + vbCrLf
            screen.Text += "Currency 1:" + curr + Space(1) + "Currency 2:" + curr2 + vbCrLf
            screen.Text += "Currency 1:" + curr + Space(1) + "Currency 2:" + curr2 + vbCrLf
            screen.Text += "Local Amount:" + localAmount.ToString + Space(1) + "USD Amount:" + usdAmount.ToString + vbCrLf
            screen.Text += "FX rate:" + fxRate.ToString + Space(1) + "Broker Code:" + brokerCode + vbCrLf
            screen.Text += "Bic Code:" + bicCode + Space(1) + "Delivery Agent:" + deliveryAgent + vbCrLf
            screen.Text += "Recieving Agent:" + recievingAgent2.ToString


        Finally


        End Try


        Return 1
    End Function

    Protected Function extractString(ByVal srcString As String, ByRef strVal As String) As Integer
        ' This function extracts value contained between data tags
        Dim posStart As Integer
        Dim posEnd As Integer

        'validate
        If Not srcString.IndexOf("<td") > 0 Then

            screen.Text += vbCrLf + "Could not extract value from" + srcString
            Return -1
        End If

        'extract
        posStart = srcString.IndexOf(">")
        posEnd = srcString.IndexOf("</td>")
        strVal = Mid(srcString, posStart + 2, posEnd - posStart - 1) ' index starts with 0 position


        Return 1
    End Function

    Protected Function formatAmount(ByVal amountStr As String) As String
        ' This function removes comma from amount and rounds it to 2 digit
        Dim num As String = ""
        Try
            num = Math.Round(CDbl(amountStr), 2).ToString("###.00")
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
        End Try

        Return num
    End Function

    Protected Function formatAmount02(ByVal amountStr As String) As String
        ' This function removes comma from amount and rounds it to 2 digit
        Dim num As String = ""
        Try
            num = Math.Round(CDbl(amountStr), 2).ToString("#,##0")
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
        End Try

        Return num
    End Function

    Protected Function formatDate(ByVal dateStr As String) As String
        Dim str As String

        str = CDate(dateStr).ToString("yyyyMMdd")

        Return str
    End Function

    Public Shared Function getISOCurrency(ByVal cur As String, ByVal moxyCon As String) As String
        ' converts 2 char currency to 3 char iso currency
        Dim isoCur As String = String.Empty
        Dim Conn As New SqlConnection(moxyCon)
        Try

            Dim Cmd As SqlCommand = New SqlCommand("usp_GetISOCurrency", Conn)
            Dim DA As SqlDataAdapter = New SqlDataAdapter
            Dim DSet As New DataSet
            Cmd.CommandType = CommandType.StoredProcedure
            Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
            RetValue.Direction = ParameterDirection.ReturnValue
            Dim currency As SqlParameter = Cmd.Parameters.Add("@cur", SqlDbType.VarChar)
            currency.Direction = ParameterDirection.Input
            currency.Value = cur
            DA.SelectCommand = Cmd
            Conn.Open()
            DA.Fill(DSet, "iso")
            Dim DTable As System.Data.DataTable = DSet.Tables("iso")
            If DTable.Rows.Count = 0 Then
                ' no portfolio found
                Throw New Exception(vbCrLf + "FUNCTION getISOCurrency: usp_GetISOCurrency: No ISO currency found for:  " + cur)

            ElseIf DTable.Rows.Count > 1 Then
                Throw New Exception("usp_GetISOCurrency: Too many ISO currencies found for:  " + cur)

            Else
                Dim drow As DataRow
                Dim drows() As DataRow = DTable.Select
                For Each drow In drows
                    isoCur = drow(0).ToString
                Next

            End If

        Catch ex As Exception
            Throw New Exception("getISOCurrency: Failed to retrieve ISO currency from Moxy for " + cur + vbCrLf)
        Finally
            Conn.Close()
        End Try

        Return isoCur
    End Function

    Protected Function getBICCode(ByVal brokerCode As String, ByRef bicCode As String) As Integer
        ' this function returns FX connect BIC code for a broker
        Dim rtn As Integer = 0

        Dim Conn As New SqlConnection(moxyCon)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetBrokerSwiftCode", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet

        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim broker As SqlParameter = Cmd.Parameters.Add("@brokercode", SqlDbType.VarChar)
        broker.Direction = ParameterDirection.Input
        broker.Value = brokerCode
        DA.SelectCommand = Cmd
        Try
            Conn.Open()
            DA.Fill(DSet, "swiftcodes")
            Dim DTable As System.Data.DataTable = DSet.Tables("swiftcodes")
            If DTable.Rows.Count = 0 Then
                ' no portfolio found
                screen.Text += vbCrLf + "FUNCTION getBICCode: usp_GetBrokerSwiftCode: No broker swift code found for  " + brokerCode
                rtn = -1
            ElseIf DTable.Rows.Count > 1 Then
                screen.Text += "usp_GetBrokerSwiftCode: Too many broker swift codes found for  " + brokerCode
                rtn = -1
            Else
                Dim drow As DataRow
                Dim drows() As DataRow = DTable.Select
                For Each drow In drows
                    bicCode = drow(0).ToString
                Next
                rtn = 1
            End If

        Catch ex As Exception
            screen.Text += "getBICCode: Failed to retrieve swift code from Moxy table for " + brokerCode + vbCrLf
            screen.Text += ex.Message
            Return -1
        Finally
            Conn.Close()
        End Try


        Return rtn
    End Function

    Protected Function getCustodianID(ByVal portfolio As String) As String
        ' this function retrieves custodian ID for the portfolio from Moxy

        Dim rtn As String = ""
        Dim Conn As New SqlConnection(moxyCon)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetCustodianID", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet

        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim port As SqlParameter = Cmd.Parameters.Add("@portfolio", SqlDbType.VarChar)
        port.Direction = ParameterDirection.Input
        port.Value = portfolio

        DA.SelectCommand = Cmd
        Try
            Conn.Open()
            DA.Fill(DSet, "custodian")
            Dim DTable As System.Data.DataTable = DSet.Tables("custodian")
            If DTable.Rows.Count = 0 Then
                ' no portfolio found
                screen.Text += "usp_GetCustodianID: No custodian ID for " + portfolio
                Return rtn
            Else
                Dim drow As DataRow
                Dim drows() As DataRow = DTable.Select
                For Each drow In drows
                    rtn = drow("CustID").ToString
                Next

            End If

        Catch ex As Exception
            screen.Text += "getCustodianID: Failed to retrieve custodian ID from Moxy table for" + portfolio + vbCrLf
            screen.Text += ex.Message
            Return -1
        Finally
            Conn.Close()
        End Try


        Return rtn
    End Function

    Protected Function getConversionInstruction(ByVal cur1 As String, ByVal cur2 As String) As String
        ' this function return the instruction how to apply cross rate
        ' going from the base currency to local currency: multiply or divide
        Dim rtn As String = String.Empty

        Dim Conn As New SqlConnection(moxyCon)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetConversionInstructionFXConnect", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet

        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim BaseCur As SqlParameter = Cmd.Parameters.Add("@cur1", SqlDbType.VarChar)
        BaseCur.Direction = ParameterDirection.Input
        BaseCur.Value = cur1
        Dim LocalCur As SqlParameter = Cmd.Parameters.Add("@cur2", SqlDbType.VarChar)
        LocalCur.Direction = ParameterDirection.Input
        LocalCur.Value = cur2

        DA.SelectCommand = Cmd
        Try
            Conn.Open()
            DA.Fill(DSet, "instruction")
            Dim DTable As System.Data.DataTable = DSet.Tables("instruction")
            If DTable.Rows.Count = 0 Then
                ' no trades for the date
                screen.Text += "getConversionInstruction: No conversion instruction from " + cur1 + " to " + cur2
                Return -1
            Else
                Dim drow As DataRow
                Dim drows() As DataRow = DTable.Select
                For Each drow In drows
                    rtn = drow("ConversionInstruction").ToString
                Next

            End If



        Catch ex As Exception
            screen.Text += "Failed to retrieve conversion instruction from Moxy table." + vbCrLf
            screen.Text += ex.Message
            Return -1
        Finally
            Conn.Close()
        End Try


        Return rtn
    End Function

    Protected Function getSwiftBrokerCode(ByVal adventBroker As String) As String
        'This function translates advent broker code to swift broker code
        Dim swiftBroker As String = String.Empty
        Dim Conn As New SqlConnection(moxyCon)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetSwiftBrokerCode", Conn)
        Dim DR As SqlDataReader

        Dim DSet As New DataSet

        Try

            Conn.Open()
            Cmd.CommandType = CommandType.StoredProcedure
            Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
            RetValue.Direction = ParameterDirection.ReturnValue
            Dim ParmBrokerCode As SqlParameter = Cmd.Parameters.Add("@adventbroker", SqlDbType.VarChar)
            ParmBrokerCode.Direction = ParameterDirection.Input
            ParmBrokerCode.Value = adventBroker

            DR = Cmd.ExecuteReader()

            While DR.Read()
                If DR.GetString(0) = Nothing Then
                    screen.Text += "FXConManager.getSwiftBrokerCode: No swift broker code found for " + swiftBroker
                    Return -1
                End If

                swiftBroker = DR.GetString(0)
            End While

            DR.Close()

        Catch ex As Exception
            screen.Text += "FXConManger.getSwiftBrokerCode: Failed to retrieve swift broker code from Moxy tb_brokercodes table for boker " + adventBroker + vbCrLf
            screen.Text += ex.Message
            swiftBroker = Nothing

        Finally
            If Not DR Is Nothing Then
                DR.Close()

            End If
            Conn.Close()

        End Try

        Return swiftBroker
    End Function

    Protected Function getAgent(ByVal brokerCode As String, ByVal buyCurrency As String, ByRef aAgent As String) As Integer
        'This function gets swift code and account name for the broker and
        'generates DeliveryAgent field in BuyMT300 file
        Dim rtn As Integer = 0

        Dim broker As String
        Dim Conn As New SqlConnection(moxyCon)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetSwiftCode", Conn)
        Dim DR As SqlDataReader
        Dim DSet As New DataSet
        Dim bankName As String

        Try



            broker = getSwiftBrokerCode(brokerCode)



            Conn.Open()
            Cmd.CommandType = CommandType.StoredProcedure
            Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
            RetValue.Direction = ParameterDirection.ReturnValue
            Dim ParmBrokerCode As SqlParameter = Cmd.Parameters.Add("@brokercode", SqlDbType.VarChar)
            ParmBrokerCode.Direction = ParameterDirection.Input
            ParmBrokerCode.Value = broker
            Dim ParmCurrency As SqlParameter = Cmd.Parameters.Add("@currency", SqlDbType.Char, 3)
            ParmCurrency.Direction = ParameterDirection.Input
            ParmCurrency.Value = buyCurrency

            DR = Cmd.ExecuteReader()

            While DR.Read()
                If DR.GetString(0) = Nothing Then
                    screen.Text += "No swift code found for " + brokerCode + Space(1) + buyCurrency
                    Return -1
                End If

                If DR.GetString(1) = Nothing Then
                    screen.Text += "No bank name found for " + brokerCode + Space(1) + buyCurrency
                    Return -1
                End If
                bankName = DR.GetString(1)

                If DR.GetString(2) = Nothing Then
                    screen.Text += "No account info found for " + brokerCode + Space(1) + buyCurrency
                    Return -1
                End If

                aAgent = "/ABIC/" + DR.GetString(0) + "|/NAME/" + DR.GetString(1) + "|/ACCT/" + DR.GetString(2)
            End While

            DR.Close()

        Catch ex As Exception
            screen.Text += "Failed to retrieve swift instruction from Moxy table for boker " + brokerCode + Space(1) + buyCurrency + " Stored Proc: " + Cmd.CommandText + vbCrLf
            screen.Text += ex.Message
            rtn = -1
        Finally

            Conn.Close()
        End Try

        If aAgent Is Nothing Then
            screen.Text += "Failed to retrieve swift instruction from Moxy table for boker " + brokerCode + Space(1) + buyCurrency + " Stored Proc: " + Cmd.CommandText + vbCrLf
            aAgent = "?"
            rtn = -1
        End If

        Return rtn
    End Function

    Protected Function getHoliday(ByVal currency As String, ByVal checkDate As Date) As Integer
        ' This function checks MoxyHolyday table if the argument checkDate is a Holiday for the specified
        ' currency, returns true if it's a holiday

        Dim holydayFlag As Integer

        Try
            Dim conn As New SqlConnection(moxyCon)
            Dim cmd As SqlCommand = New SqlCommand("usp_GetHolyday", conn)
            cmd.CommandType = CommandType.StoredProcedure
            Dim RetValue As SqlParameter = cmd.Parameters.Add("RetValue", SqlDbType.Int)
            RetValue.Direction = ParameterDirection.ReturnValue
            Dim curr As SqlParameter = cmd.Parameters.Add("@currency", SqlDbType.VarChar)
            curr.Direction = ParameterDirection.Input
            curr.Value = currency
            Dim theDate As SqlParameter = cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
            theDate.Direction = ParameterDirection.Input
            theDate.Value = Convert.ToDateTime(checkDate)
            Dim Rtn As SqlParameter = cmd.Parameters.Add("@rtn", SqlDbType.Int)
            Rtn.Direction = ParameterDirection.Output

            conn.Open()
            cmd.ExecuteNonQuery()

            holydayFlag = cmd.Parameters("@rtn").Value
            If holydayFlag = -1 Then
                ' calendar for specified currency & year does not exist
                screen.Text += vbCrLf + "Function getHoliday: No records in MoxyHoliday for " + currency + " " + Year(checkDate).ToString + vbCrLf
            End If
            conn.Close()
        Catch ex As Exception
            screen.Text += vbCrLf + "Function getHoliday: Failed to retrieve conversion instruction from Moxy table." + vbCrLf
            screen.Text += ex.Message
            Return -1

        End Try

        Return holydayFlag

    End Function

    Protected Function getFixingDate(ByVal cur As String, ByVal testDate As Date) As String
        'This function calculates the Fixing date as two business days before the settle date
        Dim wday As Integer
        Dim bdays As Integer = 2
        Dim cnt As Integer = 0

        Do While cnt < bdays
            testDate = DateAdd(DateInterval.Day, -1, testDate)
            wday = Weekday(testDate)
            If wday >= 2 And wday <= 6 Then
                ' weekdays from Monday to Friday
                Select Case getHoliday(cur, testDate)
                    Case 1
                        ' This is a holiday continue to other date
                    Case 0
                        ' Not a holiday
                        cnt += 1
                    Case -1
                        ' No holiday calendar for the currency
                        testDate = Nothing
                        Exit Do
                End Select

            End If
        Loop

        Return testDate.ToString("yyyyMMdd")
    End Function


    Public Function getFundTradingRecapAllFunds(ByVal asOfDate As DateTime, ByVal includeDTCCConfirms As Boolean) As Integer
        Dim rtn As Integer = 0
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer = 0
        Dim confirmFile As String
        ' Call the function to get the list of transactions.
        'Dim confirmsList As List(Of DTCCTransaction)
        Dim dtccTradesList As List(Of TradeAllocation)
        Dim dtccTradeCnt As Integer = 0
        Try

            Dim DTCCFolder As String = ReadConfigSetting("DTCCConfirmsFolder")

            confirmFile = FindConfirmFile(DTCCFolder, asOfDate)

            If confirmFile = String.Empty And includeDTCCConfirms Then
                screen.Text += vbCrLf + "No DTCC confirm file found for " + asOfDate.ToString("MM/dd/yyyy")
                If includeDTCCConfirms Then
                    Return -1
                    'Else
                    '    screen.Text += vbCrLf + "Will create Fund Trading Recap without DTCC confirms"
                End If
            Else
                If includeDTCCConfirms Then
                    screen.Text += vbCrLf + "DTCC Confirm file found: " + confirmFile
                End If
            End If
            'confirmsList = ReadTransactionsFromFile(confirmFile, asOfDate)
            If includeDTCCConfirms Then
                dtccTradesList = ReadTradeAllocationsFromFile(confirmFile)
                dtccTradesList = dtccTradesList _
                .OrderBy(Function(t) If(t.AcctID, ""), StringComparer.OrdinalIgnoreCase) _
                .ThenBy(Function(t) If(t.AllocSettleCurr, ""), StringComparer.OrdinalIgnoreCase) _
                .ThenBy(Function(t) If(t.BS, ""), StringComparer.OrdinalIgnoreCase) _
                .ToList()
            End If


            If File.Exists(fName) Then
                File.Delete(fName)
            End If


            Dim myReader As SqlDataReader = ExecuteFundTradesRecapAllFundsReader(asOfDate, moxyCon)

            ' Create a list to hold the transaction objects from Moxy.
            Dim transactions As New List(Of Transaction)()
            While myReader.Read()

                ' Create a new Transaction object and populate it from the reader using the constructor
                Dim transaction As New Transaction(
                        myReader.GetValue(3).ToString,
                        myReader.GetValue(5).ToString,
                        myReader.GetValue(1).ToString,
                        Convert.ToDecimal(myReader.GetValue(6)),
                        Convert.ToDecimal(myReader.GetValue(11)),
                        myReader.GetValue(12).ToString,
                        myReader.GetValue(19).ToString,
                        CDate(myReader.GetValue(4)),
                        CDate(myReader.GetValue(17)),
                        Convert.ToDecimal(myReader.GetValue(22)),
                        myReader.GetValue(18).ToString,
                        myReader.GetValue(23).ToString,
                        Convert.ToDecimal(myReader.GetValue(8)),
                        Convert.ToDecimal(myReader.GetValue(24)),
                        myReader.GetValue(15).ToString,
                        myReader.GetValue(2).ToString,
                        myReader.GetValue(20).ToString,
                        myReader.GetValue(16).ToString, ' sedol
                        myReader.GetValue(25).ToString(),' reflow flag,
                        myReader.GetValue(26).ToString() ' lot selection method
                    )

                ' Add the new Transaction object to the list.
                transactions.Add(transaction)
            End While

            ' create a new excel file
            Dim oXL As New Excel.Application
            Dim theWorkbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(1, 1) = "Trade Date"
            ' Sets the cell format to text
            worksheet.Cells(1, 2).NumberFormat = "@"
            worksheet.Cells(1, 2) = asOfDate.ToString("MM/dd/yyyy") & Space(5)

            worksheet.Cells(2, 1) = "Authorized by"
            worksheet.Cells(2, 2) = "AM"

            Dim co As New CounterObj()
            worksheet.Cells(5, co.getNext()) = "Buy/Sell"
            worksheet.Cells(5, co.getNext()) = "SHS"
            worksheet.Cells(5, co.getNext()) = "Description"
            worksheet.Cells(5, co.getNext()) = "Price"
            worksheet.Cells(5, co.getNext()) = "Net Amount"
            worksheet.Cells(5, co.getNext()) = "Broker"
            worksheet.Cells(5, co.getNext()) = "Posted T+1"
            worksheet.Cells(5, co.getNext()) = "BNY Initial"
            worksheet.Cells(5, co.getNext()) = "Fund"
            worksheet.Cells(5, co.getNext()) = "Identifier"
            worksheet.Cells(5, co.getNext()) = "Security Currency"
            worksheet.Cells(5, co.getNext()) = "Trade Date"
            worksheet.Cells(5, co.getNext()) = "Settle Date"
            worksheet.Cells(5, co.getNext()) = "Net Fees"
            worksheet.Cells(5, co.getNext()) = "Broker Name"
            worksheet.Cells(5, co.getNext()) = "Cusip"
            worksheet.Cells(5, co.getNext()) = "Commission"
            worksheet.Cells(5, co.getNext()) = "Other Fees"
            worksheet.Cells(5, co.getNext()) = "Reflow Flag"
            worksheet.Cells(5, co.getNext()) = "Lot Selection Method"
            worksheet.Cells(5, co.getNext()) = "Port Id"
            worksheet.Cells(5, co.getNext()) = "ISIN"

            rowNum = 7


            transactions = transactions _
            .OrderBy(Function(t) If(t.PortId, ""), StringComparer.OrdinalIgnoreCase) _
            .ThenBy(Function(t) If(t.SecurityCurrency, ""), StringComparer.OrdinalIgnoreCase) _
            .ThenBy(Function(t) If(t.TranCode, ""), StringComparer.OrdinalIgnoreCase) _
            .ToList()

            For Each transaction As Transaction In transactions

                Dim co2 As New CounterObj()
                ' Now assign the Transaction object's properties to the worksheet cells
                worksheet.Cells(rowNum, co2.getNext()) = transaction.TranCode
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Quantity
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Description
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Price

                Dim col As Integer = co2.getNext()
                WriteDecimal(worksheet, rowNum, col, transaction.NetAmount)
                'worksheet.Cells(rowNum, co2.getNext()) = transaction.NetAmount
                worksheet.Cells(rowNum, co2.getNext()) = transaction.BrokerId
                worksheet.Cells(rowNum, co2.getNext()) = String.Empty
                worksheet.Cells(rowNum, co2.getNext()) = String.Empty

                'worksheet.Cells(rowNum, co2.getNext()) = transaction.SecurityCurrency
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Fund

                ' prevent excel to drop leading zeros
                Dim colNum As Int16 = co2.getNext()
                worksheet.Cells.NumberFormat = "@"
                worksheet.Cells(rowNum, colNum) = transaction.Sedol

                worksheet.Cells(rowNum, co2.getNext()) = transaction.SecurityCurrency
                worksheet.Cells(rowNum, co2.getNext()) = transaction.TradeDate
                worksheet.Cells(rowNum, co2.getNext()) = transaction.SettleDate

                worksheet.Cells(rowNum, co2.getNext()) = transaction.NetFees
                worksheet.Cells(rowNum, co2.getNext()) = transaction.BrokerName

                worksheet.Cells(rowNum, co2.getNext()) = transaction.Cusip
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Commission
                worksheet.Cells(rowNum, co2.getNext()) = transaction.OtherFees
                worksheet.Cells(rowNum, co2.getNext()) = transaction.ReflowFlag
                worksheet.Cells(rowNum, co2.getNext()) = transaction.LotSelectionMethod
                worksheet.Cells(rowNum, co2.getNext()) = transaction.PortId
                worksheet.Cells(rowNum, co2.getNext()) = transaction.Isin




                rowNum += 1
            Next


            'myReader.Close()

            rowNum += 1

            ' DTCC confirms

            rowNum += 1
            If dtccTradesList IsNot Nothing AndAlso dtccTradesList.Count > 0 AndAlso includeDTCCConfirms Then
                worksheet.Cells(rowNum, 1) = "DTCC Confirms"
                rowNum += 1

                Dim fundMap As Dictionary(Of String, String) = GetTweedyFunds()

                For Each transaction As TradeAllocation In dtccTradesList
                    Dim fund As String
                    Try
                        fund = fundMap(transaction.AcctID)
                        dtccTradeCnt += 1
                    Catch ex As Exception
                        Continue For
                    End Try
                    Dim match As TranMatchObj = IsTransactionInAllocationList(transaction, transactions)
                    Dim co3 As New CounterObj()
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.BS
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.QtyAlloc
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.Security
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.Price
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.NetCashAmount
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.ExecBroker
                    worksheet.Cells(rowNum, co3.getNext()) = String.Empty
                    worksheet.Cells(rowNum, co3.getNext()) = String.Empty
                    worksheet.Cells(rowNum, co3.getNext()) = fund
                    worksheet.Cells(rowNum, co3.getNext()) = "Identifier"
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.AllocSettleCurr

                    Dim col As Integer = co3.getNext()
                    Dim cell As Excel.Range = CType(worksheet.Cells(rowNum, col), Excel.Range)

                    cell.NumberFormat = "General"
                    cell.Value2 = transaction.TradeDate.ToOADate()
                    cell.NumberFormat = "mm/dd/yyyy"  ' or "m/d/yyyy"

                    'worksheet.Cells(rowNum, co3.getNext()) = transaction.TradeDate.ToShortDateString()
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.SettleDate
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.Fees + transaction.Comm  ' this is sum of all fees from DTCC
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.ExecBroker
                    worksheet.Cells(rowNum, co3.getNext()) = "N/A"
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.Comm
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.ChargesTaxesFeesAmount1 + transaction.ChargesTaxesFeesAmount2 + transaction.ChargesTaxesFeesAmount3
                    worksheet.Cells(rowNum, co3.getNext()) = "No"
                    worksheet.Cells(rowNum, co3.getNext()) = "High Cost"
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.AcctID
                    worksheet.Cells(rowNum, co3.getNext()) = transaction.SecCode

                    If match.IsMatched Then
                        worksheet.Range(worksheet.Cells(rowNum, 1), worksheet.Cells(rowNum, 2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        worksheet.Range(worksheet.Cells(rowNum, 2), worksheet.Cells(rowNum, 2)).Interior.TintAndShade = 0.6
                        worksheet.Range(worksheet.Cells(rowNum, 4), worksheet.Cells(rowNum, 5)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        worksheet.Range(worksheet.Cells(rowNum, 11), worksheet.Cells(rowNum, 14)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        worksheet.Range(worksheet.Cells(rowNum, 21), worksheet.Cells(rowNum, 22)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        worksheet.Range(worksheet.Cells(rowNum, 22), worksheet.Cells(rowNum, 22)).Interior.TintAndShade = 0.6
                    Else
                        If match.QtyDiff = False Then
                            worksheet.Range(worksheet.Cells(rowNum, 2), worksheet.Cells(rowNum, 2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                        End If
                    End If

                        rowNum += 1
                Next
            End If

            rowNum += 1
            If dtccTradeCnt > 0 And transactions.Count >= 0 And includeDTCCConfirms Then
                If dtccTradeCnt = transactions.Count Then
                    worksheet.Cells(rowNum, 2) = "Trades Counts Match: " & dtccTradeCnt & " : " & transactions.Count
                Else
                    worksheet.Cells(rowNum, 2) = "Trades Counts Do Not Match: " & dtccTradeCnt & " : " & transactions.Count
                End If
            End If

            ' save generated Excel file
            worksheet.Columns.AutoFit()
            ' After writing values/formats for that column:
            Dim netAmtCol As Integer = 5  ' example
            AutoFitColumnWithMin(worksheet, netAmtCol, minWidth:=16)  ' ~16 chars wide

            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()

            myReader.Close()
            'Conn.Close()

            Return 1
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
            Return -1
        End Try

    End Function


    ' Auto-fit a single column, then enforce a minimum width and add a small padding.
    <CLSCompliant(False)>
    Public Function AutoFitColumnWithMin(ws As Excel.Worksheet,
                                         col As Integer,
                                         minWidth As Double,
                                         Optional padding As Double = 0.8) As Boolean
        Dim c As Excel.Range = Nothing
        Try
            c = ws.Columns(col)
            c.WrapText = False
            c.ShrinkToFit = False
            c.Columns.AutoFit()

            Dim fitted As Double = c.ColumnWidth
            Dim finalWidth As Double = Math.Max(fitted + padding, minWidth)
            c.ColumnWidth = finalWidth
            Return True
        Catch
            Return False
        Finally
            If c IsNot Nothing Then Marshal.ReleaseComObject(c)
        End Try
    End Function

    ' Write a Decimal as a NUMBER and show many decimals (avoids display rounding).
    ' Example format default: "###0.00##########" (2 fixed + up to 10 optional)
    <CLSCompliant(False)>
    Public Function WriteDecimal(ws As Excel.Worksheet,
                                 row As Integer,
                                 col As Integer,
                                 amount As Decimal) As Boolean
        Dim cell As Excel.Range = Nothing

        Try
            cell = CType(ws.Cells(row, col), Excel.Range)

            ' Clear any prior fixed-decimal format (e.g., "0.0") and force up to 6 decimals
            cell.NumberFormat = "General"
            cell.NumberFormat = "###0.######"   ' thousands + up to 6 decimals

            ' Excel stores numerics as Double; write explicitly as Double
            Dim Val As Double = CDbl(amount)
            cell.Value2 = Val


            Return True
        Catch
            Return False
        Finally
            If cell IsNot Nothing Then Marshal.ReleaseComObject(cell)
        End Try
    End Function

    ''' <summary>
    ''' Checks if a Transaction object is present in a list of TradeAllocation objects based on key properties.
    ''' </summary>
    ''' <param name="transaction">The Transaction object to search for.</param>
    ''' <param name="allocationList">The list of TradeAllocation objects to search in.</param>
    ''' <returns>True if a match is found, otherwise False.</returns>
    Public Function IsTransactionInMoxyAllocationList(ByVal transaction As TradeAllocation, ByVal allocationList As List(Of MoxyAllocTran)) As TranMatchObj
        ' transaction <-- from DTCC
        ' allocationList <-- from Moxy
        ' TODO: There is no other fee for 10/02/2025 in Moxy?

        Dim Result As TranMatchObj = New TranMatchObj()
        Dim netFees As Decimal = transaction.Comm + transaction.Fees + transaction.CommAmount1 + transaction.CommAmount2 + transaction.CommAmount3

        ' It looks litke in the confirm file net fees are summed up in Fees column
        netFees = transaction.Fees + transaction.Comm

        Dim tranCode As String = transaction.BS.Trim().ToLower()
        If tranCode.Contains("buy") Then
            tranCode = "buy"
        End If
        Dim netCashAmountTolerance As Decimal = 0.01D

        If transaction.AllocSettleCurr.Equals("JPY") Or transaction.AllocSettleCurr.Equals("KRW") Then
            netCashAmountTolerance = 1D
        End If

        For Each allocation As MoxyAllocTran In allocationList

            If Not (transaction.SecCode.Equals(allocation.ISIN) Or transaction.SecCode.Equals(allocation.Cusip)) Then
                Continue For
            End If

            Dim qtyDiff As Decimal = Math.Abs(transaction.QtyAlloc - allocation.AllocQty)
            If qtyDiff > 0.0 Then
                Result.QtyDiff = False
            End If
            Dim amtDiff As Decimal = Math.Abs(transaction.TrdAmt - allocation.Principal)
            If amtDiff > netCashAmountTolerance Then
                Result.PrincipalDiff = True
            End If
            Dim tmp As Boolean = False
            Dim quantity As Decimal = CDec(allocation.AllocQty)
            Dim priceDiff As Decimal = Math.Abs(transaction.Price - allocation.AllocPrice)
            Dim netFeesDiff As Decimal = Math.Abs(netFees - allocation.Commission)

            Dim acctIdMatch As Boolean = String.Equals(transaction.AcctID, allocation.PortId, StringComparison.OrdinalIgnoreCase)
            Dim tranCodeMatch As Boolean = String.Equals(tranCode, allocation.TranCode, StringComparison.OrdinalIgnoreCase)

            Dim securityMatch As Boolean = String.Equals(transaction.SecCode, allocation.ISIN, StringComparison.OrdinalIgnoreCase)
            If Not securityMatch Then
                securityMatch = String.Equals(transaction.SecCode, allocation.Cusip, StringComparison.OrdinalIgnoreCase)
            End If

            Dim settleCurMatch As Boolean = String.Equals(transaction.AllocSettleCurr, allocation.SecurityCurrency.Trim, StringComparison.OrdinalIgnoreCase)

            If acctIdMatch And
               tranCodeMatch And
               securityMatch And
               qtyDiff = 0.0D And
               amtDiff < netCashAmountTolerance And
                priceDiff < 0.01D And    'netFeesDiff < 0.01D And
                String.Equals(transaction.AllocSettleCurr, allocation.SecurityCurrency.Trim, StringComparison.OrdinalIgnoreCase) Then

                Result.IsMatched = True
                Return Result
            End If
        Next

        ' If the loop completes without finding a match, the transaction is not in the list.
        Return Result

    End Function



    ''' <summary>
    ''' Checks if a Transaction object is present in a list of TradeAllocation objects based on key properties.
    ''' </summary>
    ''' <param name="transaction">The Transaction object to search for.</param>
    ''' <param name="allocationList">The list of TradeAllocation objects to search in.</param>
    ''' <returns>True if a match is found, otherwise False.</returns>
    Public Function IsTransactionInAllocationList(ByVal transaction As TradeAllocation, ByVal allocationList As List(Of Transaction)) As TranMatchObj


        Dim Result As TranMatchObj = New TranMatchObj()
        Dim netFees As Decimal = transaction.Comm + transaction.Fees + transaction.CommAmount1 + transaction.CommAmount2 + transaction.CommAmount3

        ' It looks litke in the confirm file net fees are summed up in Fees column
        netFees = transaction.Fees + transaction.Comm

        Dim tranCode As String = transaction.BS.Trim().ToLower()
        If tranCode.Contains("buy") Then
            tranCode = "buy"
        End If
        Dim netCashAmountTolerance As Decimal = 0.01D

        If transaction.AllocSettleCurr.Equals("JPY") Or transaction.AllocSettleCurr.Equals("KRW") Then
            netCashAmountTolerance = 1D
        End If

        For Each allocation As Transaction In allocationList

            If Not (transaction.SecCode.Equals(allocation.Isin) Or transaction.SecCode.Equals(allocation.Cusip)) Then
                Continue For
            End If

            Dim qtyDiff As Decimal = Math.Abs(transaction.QtyAlloc - allocation.Quantity)
            If qtyDiff > 0.0 Then
                Result.QtyDiff = False
            End If
            Dim amtDiff As Decimal = Math.Abs(transaction.NetCashAmount - allocation.NetAmount)
            Dim tmp As Boolean = False
            Dim quantity As Decimal = CDec(allocation.Quantity)
            Dim priceDiff As Decimal = Math.Abs(transaction.Price - allocation.Price)
            Dim netFeesDiff As Decimal = Math.Abs(netFees - allocation.NetFees)

            Dim acctIdMatch As Boolean = String.Equals(transaction.AcctID, allocation.PortId, StringComparison.OrdinalIgnoreCase)
            Dim tranCodeMatch As Boolean = String.Equals(tranCode, allocation.TranCode, StringComparison.OrdinalIgnoreCase)


            Dim securityMatch As Boolean = String.Equals(transaction.SecCode, allocation.Isin, StringComparison.OrdinalIgnoreCase)
            If Not securityMatch Then
                securityMatch = String.Equals(transaction.SecCode, allocation.Cusip, StringComparison.OrdinalIgnoreCase)
            End If

            Dim settleCurMatch As Boolean = String.Equals(transaction.AllocSettleCurr, allocation.SecurityCurrency.Trim, StringComparison.OrdinalIgnoreCase)

            If acctIdMatch And
               tranCodeMatch And
               securityMatch And
               qtyDiff = 0.0D And
               amtDiff < netCashAmountTolerance And
                priceDiff < 0.01D And
                  netFeesDiff < 0.01D And
                String.Equals(transaction.AllocSettleCurr, allocation.SecurityCurrency.Trim, StringComparison.OrdinalIgnoreCase) Then

                Result.IsMatched = True
                Return Result
            End If
        Next

        ' If the loop completes without finding a match, the transaction is not in the list.
        Return Result

    End Function

    ' This function searches for a confirmation file in a given folder based on a date.
    ' It expects the file to be named in the format "Combined Match Agreed Allocations YYYY-Mon-DD ####.csv".
    Public Function FindConfirmFile(ByVal dtccFolder As String, ByVal asOfDate As Date) As String
        Try
            ' Format the date to match the "YYYY-Mon-DD" pattern in the file name.
            ' Example: 2025-Aug-25
            Dim datePart As String = asOfDate.ToString("yyyy-MMM-dd")

            ' Construct the expected file name prefix.
            Dim searchPattern As String = "Combined Match Agreed Allocations " & datePart & "*.csv"

            ' Use Directory.GetFiles to find all files in the folder that match the pattern.
            ' SearchOption.TopDirectoryOnly means it will not search subfolders.
            Dim matchingFiles As String() = Directory.GetFiles(dtccFolder, searchPattern, SearchOption.TopDirectoryOnly)

            ' Check if any matching files were found.
            If matchingFiles.Length > 0 Then
                ' If multiple files are found, sort them in descending order to get the latest one.
                ' We'll use a more robust method by extracting the time part and sorting on that.
                Dim latestFile As String = matchingFiles.OrderByDescending(Function(f) GetTimeFromFileName(f)).First()

                ' Return the full path of the latest file.
                Return latestFile
            Else
                ' If no file is found, return an empty string.
                Return String.Empty
            End If

        Catch ex As Exception
            ' Catch any potential exceptions (e.g., folder not found) and throw a new one
            ' with a descriptive message.
            Throw New Exception("FindConfirmFile Exception: " & ex.Message, ex)
        End Try
    End Function

    ' This helper function extracts the time from the file name.
    Private Function GetTimeFromFileName(ByVal fileName As String) As String
        Try
            ' A regular expression to find the four-digit time stamp.
            ' It looks for a sequence of four digits at the end of the filename
            ' just before the ".csv" extension.
            Dim timeRegex As New Regex("\d{4}(?=\.csv$)")
            Dim match As Match = timeRegex.Match(fileName)

            If match.Success Then
                Return match.Value
            Else
                ' If no time is found, return an empty string or a default value.
                Return String.Empty
            End If
        Catch ex As Exception
            ' Catch any potential exceptions and return an empty string.
            ' This prevents the main function from crashing if the regex fails.
            Return String.Empty
        End Try
    End Function

    ' This function reads a pipe-delimited file, finds the header, and
    ' returns a list of DTCCTransaction objects.
    Public Function ReadTransactionsFromFile(ByVal filePath As String, ByVal asOfDate As Date) As List(Of DTCCTransaction)
        Dim transactions As New List(Of DTCCTransaction)()

        Try
            ' Read all lines from the file.
            Dim lines As String() = File.ReadAllLines(filePath)
            Dim headerIndex As Integer = -1

            ' Find the header row by searching for "Exec Broker" and "Asset Class".
            For i As Integer = 0 To lines.Length - 1
                If lines(i).Contains("Exec Broker") AndAlso lines(i).Contains("Asset Class") Then
                    headerIndex = i
                    Exit For
                End If
            Next

            ' If the header is not found, return an empty list.
            If headerIndex = -1 Then
                Return transactions
            End If

            ' Start reading transactions from the row after the header.
            For i As Integer = headerIndex + 1 To lines.Length - 1
                Dim line As String = lines(i)
                If String.IsNullOrWhiteSpace(line) Then Continue For

                ' Remove double quotes from the line before parsing.
                line = line.Replace("""", String.Empty)

                Dim fields As String() = line.Split("|"c)

                ' Check if the line has a valid number of fields.
                ' The header has 42 columns, so we expect at least that many.
                If fields.Length >= 42 Then
                    Dim price As Decimal = 0
                    Decimal.TryParse(fields(5), price)

                    Dim allocQty As Decimal = 0
                    Decimal.TryParse(fields(8), allocQty)

                    Dim trdAmt As Decimal = 0
                    Decimal.TryParse(fields(9), trdAmt)

                    Dim chargesTaxesFeesAmount1 As Decimal = 0
                    Decimal.TryParse(fields(11), chargesTaxesFeesAmount1)

                    Dim chargesTaxesFeesAmount2 As Decimal = 0
                    Decimal.TryParse(fields(13), chargesTaxesFeesAmount2)

                    Dim chargesTaxesFeesAmount3 As Decimal = 0
                    Decimal.TryParse(fields(15), chargesTaxesFeesAmount3)

                    Dim netCashAmount As Decimal = 0
                    Decimal.TryParse(fields(22), netCashAmount)

                    Dim settleAmount As Decimal = 0
                    Decimal.TryParse(fields(23), settleAmount)

                    Dim comm As Decimal = 0
                    Decimal.TryParse(fields(41), comm)

                    ' Create a new DTCCTransaction object with the parsed values.
                    Dim transaction As New DTCCTransaction(
                    fields(0), ' Exec Broker
                    fields(1), ' Asset Class Sec Code
                    fields(3), ' Security
                    fields(4), ' TranCode
                    price,
                    fields(5), ' AcctID
                    allocQty,
                    trdAmt,
                    chargesTaxesFeesAmount1,
                    chargesTaxesFeesAmount2,
                    chargesTaxesFeesAmount3,
                    netCashAmount,
                    settleAmount,
                    fields(24), ' Alloc Settle Curr
                    comm,
                    asOfDate
                )

                    ' Add the object to the list.
                    transactions.Add(transaction)
                End If
            Next

        Catch ex As Exception
            ' Re-throw a more descriptive exception.
            Throw New Exception("Error reading transactions from file: " & ex.Message, ex)
        End Try

        Return transactions
    End Function


    ' This function reads a pipe-delimited file and automatically maps columns to a TradeAllocation object.
    Public Function ReadTradeAllocationsFromFile(ByVal filePath As String) As List(Of TradeAllocation)
        Dim allocations As New List(Of TradeAllocation)()

        Try
            Dim lines As String() = File.ReadAllLines(filePath)
            Dim headerIndex As Integer = -1

            For i As Integer = 0 To lines.Length - 1
                If lines(i).Contains("Exec Broker") AndAlso lines(i).Contains("Asset Class") Then
                    headerIndex = i
                    Exit For
                End If
            Next

            If headerIndex = -1 Then
                Return allocations
            End If

            ' Sanitize header to create a map of sanitized names to column indices.
            Dim headerLine As String = lines(headerIndex).Replace("""", String.Empty)
            Dim rawHeaderFields As String() = headerLine.Split("|"c)

            Dim columnMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For i As Integer = 0 To rawHeaderFields.Length - 1

                ' Remove spaces, slashes, and dashes, then check if it's not empty.
                Dim sanitizedName As String = rawHeaderFields(i).Replace(" ", "").Replace("/", "").Replace("-", "")
                If i = 2 Then
                    Dim test = sanitizedName
                End If
                If Not String.IsNullOrWhiteSpace(sanitizedName) Then
                    columnMap(sanitizedName) = i
                End If
            Next

            Dim properties As PropertyInfo() = GetType(TradeAllocation).GetProperties()

            For i As Integer = headerIndex + 1 To lines.Length - 1
                Dim line As String = lines(i)
                If String.IsNullOrWhiteSpace(line) Then Continue For

                line = line.Replace("""", String.Empty)
                Dim fields As String() = line.Split("|"c)

                If fields.Length >= rawHeaderFields.Length Then
                    Dim newAlloc As New TradeAllocation()

                    ' Loop through the properties of the TradeAllocation class
                    For Each prop As PropertyInfo In properties
                        If prop.Name = "SecCode" Then
                            Dim propName = prop.Name
                        End If
                        If columnMap.ContainsKey(prop.Name) Then
                            Dim columnIndex As Integer = columnMap(prop.Name)
                            Dim rawValue As String = fields(columnIndex)

                            Try

                                ' Convert the value to the property's type and assign it.
                                Select Case prop.PropertyType.FullName
                                    Case GetType(Decimal).FullName
                                        Dim dVal As Decimal = 0
                                        Decimal.TryParse(rawValue, dVal)
                                        prop.SetValue(newAlloc, dVal)
                                    Case GetType(Date).FullName
                                        Dim dDate As Date
                                        If Date.TryParse(rawValue, dDate) Then
                                            prop.SetValue(newAlloc, dDate)
                                        End If
                                    Case Else ' Default to string
                                        prop.SetValue(newAlloc, rawValue)
                                End Select
                            Catch ex As Exception
                                ' Log the error without stopping the loop.
                                ' For production, you might want more detailed logging.
                                mo.errMsg01("ReadTradeAllocationsFromFile", $"Error converting value '{rawValue}' for property '{prop.Name}': {ex.Message}")

                            End Try
                        End If
                    Next
                    allocations.Add(newAlloc)
                End If
            Next

            If allocations.Count > 0 Then
                ' has items


            End If

        Catch ex As Exception
            Throw New Exception("Error reading transactions from file: " & ex.Message, ex)
        End Try

        Return allocations
    End Function

    ' This function reads Tweedy funds from app.config.
    Public Function GetTweedyFunds() As Dictionary(Of String, String)

        Dim fundMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Try

            Dim fund1 As String = ReadConfigSetting("55090")
            fundMap("55090") = fund1

            Dim fund2 As String = ReadConfigSetting("55093")
            fundMap("55093") = fund2

            Dim fund3 As String = ReadConfigSetting("55095")
            fundMap("55095") = fund3

            Dim fund4 As String = ReadConfigSetting("55096")
            fundMap("55096") = fund4


        Catch ex As Exception
            Throw New Exception("s from file: " & ex.Message, ex)
        End Try

        Return fundMap
    End Function

    Public Function SortBySecurityThenAcctId(ByVal allocations As List(Of TradeAllocation)) _
                                         As List(Of TradeAllocation)
        If allocations Is Nothing OrElse allocations.Count <= 1 Then Return allocations

        Return allocations _
        .OrderBy(Function(a) If(a?.Security, ""), StringComparer.OrdinalIgnoreCase) _
        .ThenBy(Function(a) If(a?.AcctID, ""), StringComparer.OrdinalIgnoreCase) _
        .ToList()
    End Function




    Public Function getPortfolioRecapInternational(ByVal asOfDate As DateTime) As Integer
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer
        Dim sec As String = ""

        Try
            If File.Exists(fName) Then
                File.Delete(fName)
            End If

            Dim Conn As New SqlConnection(moxyCon)
            Dim Cmd As New SqlCommand("usp_TradeingRecapInernationalDaily", Conn)
            Cmd.CommandType = CommandType.StoredProcedure
            Dim PortDate As New SqlParameter
            PortDate = Cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
            PortDate.Direction = ParameterDirection.Input
            PortDate.Value = asOfDate

            Conn.Open()
            Dim myReader As SqlDataReader = Cmd.ExecuteReader()

            ' create a new excel file
            Dim oXL As New Excel.Application
            Dim theWorkbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(1, 1) = "Trade Date"
            worksheet.Cells(1, 2) = asOfDate.ToString("MM/dd/yyyy") + Space(5)

            worksheet.Cells(2, 1) = "Authorized by"
            worksheet.Cells(2, 2) = "AM"
            worksheet.Cells(3, 1) = "Daily International Trade Execution Report"
            worksheet.UsedRange.Font.Size = 12
            worksheet.Cells(5, 1) = "Security"
            worksheet.Cells(5, 2) = "Trancode"
            worksheet.Cells(5, 3) = "Qty"
            worksheet.Cells(5, 4) = "Price"
            worksheet.Cells(5, 5) = "Broker"
            worksheet.Cells(5, 6) = "Broker ID"

            worksheet.Range("A5", "F5").Font.Bold = True

            'worksheet.Cells(5, 5) = "Net Amount"
            'worksheet.Cells(5, 6) = "Broker"
            'worksheet.Cells(5, 7) = "Posted T/D"
            'worksheet.Cells(5, 8) = "Posted T+1"
            'worksheet.Cells(5, 9) = "PFPC Initial"

            rowNum = 6
            ' Column order in DataReader:
            '   0 - shortname, 1 - isin, 2 - trancode,
            '   3 - tradedate, 4 - qty

            While myReader.Read()
                worksheet.Cells(rowNum, 1) = myReader.GetValue(0).ToString ' security
                worksheet.Cells(rowNum, 4) = Format(myReader.GetValue(1), "##,##0.00000")       ' avg price 
                worksheet.Cells(rowNum, 2) = myReader.GetValue(2).ToString   ' tran code
                worksheet.Cells(rowNum, 3) = Format(myReader.GetValue(4), "##,##0")        ' order qty
                worksheet.Cells(rowNum, 5) = myReader.GetValue(5).ToString ' broker
                worksheet.Cells(rowNum, 6) = myReader.GetValue(6).ToString ' broker id
                rowNum += 1

            End While

            myReader.Close()
            rowNum += 5

            ' Get the allocations
            worksheet.Cells(rowNum, 1) = "International Trading Allocation"
            rowNum += 1

            worksheet.Cells(rowNum, 1) = "Security"
            worksheet.Cells(rowNum, 2) = "Portfolio ID"
            worksheet.Cells(rowNum, 3) = "Portfolio"
            worksheet.Cells(rowNum, 4) = "Tran Code"
            worksheet.Cells(rowNum, 5) = "Qty"
            worksheet.Cells(rowNum, 6) = "Price"
            worksheet.Cells(rowNum, 7) = "Broker ID"

            worksheet.Range("A" + rowNum.ToString, "G" + rowNum.ToString).Font.Bold = True

            rowNum += 1
            Dim Conn2 As New SqlConnection(moxyCon)
            Dim Cmd2 As New SqlCommand("usp_TradingRecapAllocationInternational", Conn)
            Cmd2.CommandType = CommandType.StoredProcedure
            Dim AllocDate As New SqlParameter
            AllocDate = Cmd2.Parameters.Add("@asofdate", SqlDbType.DateTime)
            AllocDate.Direction = ParameterDirection.Input
            AllocDate.Value = asOfDate

            Conn2.Open()
            Dim allocReader As SqlDataReader = Cmd2.ExecuteReader()
            While allocReader.Read()
                If sec <> allocReader.GetValue(3).ToString Then
                    rowNum += 1
                    worksheet.Cells(rowNum, 1) = allocReader.GetValue(3).ToString ' security
                End If
                sec = allocReader.GetValue(3).ToString

                worksheet.Cells(rowNum, 2) = allocReader.GetString(1) ' portfolio id
                worksheet.Cells(rowNum, 3) = allocReader.GetValue(2).ToString ' portfolio
                worksheet.Cells(rowNum, 4) = allocReader.GetValue(5).ToString   ' tran code
                worksheet.Cells(rowNum, 5) = Format(allocReader.GetValue(7), "##,##0")        ' order qty
                worksheet.Cells(rowNum, 6) = Format(allocReader.GetValue(8), "##,##0.00000")       ' avg price 
                worksheet.Cells(rowNum, 7) = Format(allocReader.GetString(14))       ' broker id 

                rowNum += 1

            End While

            myReader.Close()
            rowNum += 5

            ' save generated Excel file
            worksheet.Rows.Font.Size = 11
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()
            Conn.Close()
            Conn2.Close()

            Return 1

        Catch ex As Exception
            screen.Text += "getPortfolioRecapInternational: " + vbCrLf + ex.Message
            Return -1
        End Try
    End Function



    Public Function getPortfolioRecapInternationalDTCC(ByVal asOfDate As DateTime) As Integer
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer
        Dim sec As String = ""
        Dim confirmFile As String = ""
        Dim dtccTradesList As List(Of TradeAllocation)
        Dim dtccTradeCnt As Integer = 0
        Dim dtccRecapFolder As String = ""
        Dim oXL As Excel.Application = Nothing
        Dim theWorkbook As Excel.Workbook = Nothing
        Dim worksheet As Excel.Worksheet = Nothing
        Try

            dtccRecapFolder = ReadConfigSetting("RecapFolderDTCC")

            fName = AddDTCCToFilename(fName)

            Dim fileNoDir As String = Path.GetFileName(fName)

            fName = Path.Combine(dtccRecapFolder, fileNoDir)

            If File.Exists(fName) Then
                File.Delete(fName)
            End If

            Dim moxy As New MoxyService(moxyCon)
            Dim storedProc As String = ReadConfigSetting("TradingRecapAllocationsInternationalDTCCSP")

            ' 1. Get the allocations from Moxy
            'Dim dt As System.Data.DataTable = moxy.GetTradingRecapAllocations(storedProc, asOfDate)

            Dim items As List(Of MoxyAllocTran) = moxy.GetTradingRecapInternationalTransactions(storedProc, asOfDate)

            screen.Text += vbCrLf

            Dim DTCCFolder As String = ReadConfigSetting("DTCCConfirmsFolder")

            ' 2. Find the DTCC confirm file for the given date
            confirmFile = FindConfirmFile(DTCCFolder, asOfDate)

            If confirmFile = String.Empty Then
                mo.errMsg01("getPortfolioRecapInternationalDTCC", "No DTCC confirm file found for " + asOfDate.ToString("MM/dd/yyyy"))
                'screen.Text += vbCrLf + "No DTCC confirm file found for " + asOfDate.ToString("MM/dd/yyyy")
                Return -1
            End If

            ' 3. Read the DTCC confirm file into a list of TradeAllocation objects
            dtccTradesList = ReadTradeAllocationsFromFile(confirmFile)

            If dtccTradesList.Count > 0 Then
                dtccTradesList = SortBySecurityThenAcctId(dtccTradesList)
            End If

            ' 4. Create a new Excel file and write the data
            ' create a new excel file
            oXL = New Excel.Application

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(1, 1) = "Trade Date"
            worksheet.Cells(1, 2) = asOfDate.ToString("MM/dd/yyyy") + Space(5)

            Dim cnt As CounterObj = New CounterObj

            rowNum = 2

            worksheet.Cells(rowNum, 1) = "Moxy Allocation"

            rowNum = 3

            worksheet.Cells(rowNum, cnt.getNext) = "Order Id"
            worksheet.Cells(rowNum, cnt.getNext) = "PortId"
            worksheet.Cells(rowNum, cnt.getNext) = "Port Name"
            worksheet.Cells(rowNum, cnt.getNext) = "Security"
            worksheet.Cells(rowNum, cnt.getNext) = "Security Currency"
            worksheet.Cells(rowNum, cnt.getNext) = "TranCode"
            worksheet.Cells(rowNum, cnt.getNext) = "Trade Date"
            worksheet.Cells(rowNum, cnt.getNext) = "Settle Date"
            worksheet.Cells(rowNum, cnt.getNext) = "Alloc Qty"
            worksheet.Cells(rowNum, cnt.getNext) = "Alloc Price"
            worksheet.Cells(rowNum, cnt.getNext) = "Principal"
            worksheet.Cells(rowNum, cnt.getNext) = "Commission"
            worksheet.Cells(rowNum, cnt.getNext) = "SEC Fee"
            worksheet.Cells(rowNum, cnt.getNext) = "Other Fee"
            worksheet.Cells(rowNum, cnt.getNext) = "Tkt Charge"
            worksheet.Cells(rowNum, cnt.getNext) = "Taxes"
            worksheet.Cells(rowNum, cnt.getNext) = "ISIN"
            worksheet.Cells(rowNum, cnt.getNext) = "Cusip"
            worksheet.Cells(rowNum, cnt.getNext) = "Broker"

            cnt.reset()

            rowNum = 4
            For Each tran As MoxyAllocTran In items
                worksheet.Cells(rowNum, cnt.getNext) = tran.OrderId
                worksheet.Cells(rowNum, cnt.getNext) = tran.PortId
                worksheet.Cells(rowNum, cnt.getNext) = tran.PortName
                worksheet.Cells(rowNum, cnt.getNext) = tran.ShortName
                worksheet.Cells(rowNum, cnt.getNext) = tran.SecurityCurrency
                worksheet.Cells(rowNum, cnt.getNext) = tran.TranCode
                worksheet.Cells(rowNum, cnt.getNext) = tran.TradeDate
                worksheet.Cells(rowNum, cnt.getNext) = tran.SettleDate
                worksheet.Cells(rowNum, cnt.getNext) = tran.AllocQty
                worksheet.Cells(rowNum, cnt.getNext) = tran.AllocPrice
                worksheet.Cells(rowNum, cnt.getNext) = tran.Principal
                worksheet.Cells(rowNum, cnt.getNext) = tran.Commission
                worksheet.Cells(rowNum, cnt.getNext) = tran.SECFee
                worksheet.Cells(rowNum, cnt.getNext) = tran.OtherFee
                worksheet.Cells(rowNum, cnt.getNext) = tran.TktChrg
                worksheet.Cells(rowNum, cnt.getNext) = tran.Taxes
                worksheet.Cells(rowNum, cnt.getNext) = tran.ISIN
                worksheet.Cells(rowNum, cnt.getNext) = tran.Cusip
                worksheet.Cells(rowNum, cnt.getNext) = tran.Broker

                rowNum += 1
                cnt.reset()
            Next

            worksheet.Cells(rowNum, 1) = "DTCC Allocation"
            cnt.reset()
            rowNum += 1

            For Each alloc As TradeAllocation In dtccTradesList
                cnt.getNext() ' skip first col
                worksheet.Cells(rowNum, cnt.getNext) = alloc.AcctID
                worksheet.Cells(rowNum, cnt.getNext) = "N/A"
                worksheet.Cells(rowNum, cnt.getNext) = alloc.Security
                worksheet.Cells(rowNum, cnt.getNext) = alloc.AllocSettleCurr
                worksheet.Cells(rowNum, cnt.getNext) = alloc.BS
                'Trade Date	Settle Date	Alloc Qty	Alloc Price	Principal	Commission	SEC Fee	Other Fee	Tkt Charge	Taxes	Total Amount	ISIN	Cusip	Broker
                worksheet.Cells(rowNum, cnt.getNext) = alloc.TradeDate
                worksheet.Cells(rowNum, cnt.getNext) = alloc.SettleDate
                worksheet.Cells(rowNum, cnt.getNext) = alloc.QtyAlloc
                worksheet.Cells(rowNum, cnt.getNext) = alloc.Price
                worksheet.Cells(rowNum, cnt.getNext) = alloc.TrdAmt
                worksheet.Cells(rowNum, cnt.getNext) = alloc.Comm
                worksheet.Cells(rowNum, cnt.getNext) = "N/A"
                worksheet.Cells(rowNum, cnt.getNext) = alloc.ChargesTaxesFeesAmount1
                worksheet.Cells(rowNum, cnt.getNext) = "N/A"
                worksheet.Cells(rowNum, cnt.getNext) = alloc.Fees
                worksheet.Cells(rowNum, cnt.getNext) = alloc.SecCode
                worksheet.Cells(rowNum, cnt.getNext) = ""
                worksheet.Cells(rowNum, cnt.getNext) = alloc.ExecBroker

                Dim match As TranMatchObj = IsTransactionInMoxyAllocationList(alloc, items)
                If match.IsMatched Then
                    worksheet.Range(worksheet.Cells(rowNum, 1), worksheet.Cells(rowNum, 2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                    worksheet.Range(worksheet.Cells(rowNum, 2), worksheet.Cells(rowNum, 2)).Interior.TintAndShade = 0.6
                    worksheet.Range(worksheet.Cells(rowNum, 4), worksheet.Cells(rowNum, 5)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                    worksheet.Range(worksheet.Cells(rowNum, 11), worksheet.Cells(rowNum, 14)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                    'worksheet.Range(worksheet.Cells(rowNum, 21), worksheet.Cells(rowNum, 22)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                    worksheet.Range(worksheet.Cells(rowNum, 22), worksheet.Cells(rowNum, 22)).Interior.TintAndShade = 0.6
                Else
                    If match.QtyDiff = True Then

                        worksheet.Range(worksheet.Cells(rowNum, 2), worksheet.Cells(rowNum, 2)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                    End If
                    If match.PrincipalDiff = True Then
                        worksheet.Range(worksheet.Cells(rowNum, 11), worksheet.Cells(rowNum, 11)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                    End If
                End If
                rowNum += 1
                cnt.reset()
            Next

            rowNum += 1
            If dtccTradeCnt > 0 And items.Count >= 0 Then
                If dtccTradeCnt = items.Count Then
                    worksheet.Cells(rowNum, 2) = "Trades Counts Match: " & dtccTradeCnt & " : " & items.Count
                Else
                    worksheet.Cells(rowNum, 2) = "Trades Counts Do Not Match: " & dtccTradeCnt & " : " & items.Count
                End If
            End If


            ' save generated Excel file
            worksheet.Rows.Font.Size = 11
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            'theWorkbook.Close(False, objOpt, objOpt)

            screen.Text += "Created Portfolio recap with DTCC file :  " + fName + vbCrLf

            Return 1
        Catch ex As Exception
            screen.Text += "getPortfolioRecapInternationalDTCC: " + vbCrLf + ex.Message
            Return -1
        Finally
            ' 1. Cleanup Excel Objects
            If Not Worksheet Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Worksheet)
                Worksheet = Nothing
            End If

            If Not theWorkbook Is Nothing Then
                ' Close the workbook before releasing it
                theWorkbook.Close(False, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(theWorkbook)
                theWorkbook = Nothing
            End If

            If Not oXL Is Nothing Then
                ' Quit the Excel application
                oXL.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)
                oXL = Nothing
            End If
        End Try
    End Function


    Public Function AddDTCCToFilename(ByVal fileName As String)

        If String.IsNullOrEmpty(fileName) Then
            Return ""
        End If

        ' 1. get the directory, filename without extension, and extension
        Dim directory As String = Path.GetDirectoryName(fileName)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(fileName)
        Dim extension As String = Path.GetExtension(fileName)

        ' 2. append "_DTCC" to the filename
        Dim newName As String = nameWithoutExt & "_DTCC" & extension
        ' 3. combine them back into a full path
        Dim newFileName As String = Path.Combine(directory, newName)
        Return newFileName


    End Function


    Public Function getMoxyExport(ByVal asOfDate As DateTime) As Integer
        Dim rtn As Integer = 0
        Dim outfile As StreamWriter

        Try
            If File.Exists(fName) Then
                File.Delete(fName)
            End If
            Dim Conn As New SqlConnection(moxyCon)
            Dim Cmd As New SqlCommand("usp_GetMoxyExport", Conn)
            Cmd.CommandType = CommandType.StoredProcedure
            Dim PortDate As New SqlParameter
            PortDate = Cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
            PortDate.Direction = ParameterDirection.Input
            PortDate.Value = asOfDate

            Dim da As New SqlDataAdapter(Cmd)
            Dim dt As New System.Data.DataTable()
            da.Fill(dt)


            Dim sb As New StringBuilder()

            For i As Integer = 0 To dt.Columns.Count - 1
                sb.Append(dt.Columns(i).ColumnName + Chr(9))
            Next
            sb.Append(Environment.NewLine)

            For j As Integer = 0 To dt.Rows.Count - 1
                For k As Integer = 0 To dt.Columns.Count - 1
                    ' Chr(9) is a TAB charatcter
                    sb.Append(dt.Rows(j)(k).ToString() + Chr(9))
                Next
                sb.Append(Environment.NewLine)
            Next

            outfile = File.CreateText(fName)
            outfile.Write(sb.ToString())
            outfile.Close()
            Conn.Close()

        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
            Return -1
        End Try
        Return rtn
    End Function


    Public Function getOffshoreFundsRecapsFromPortia(ByVal aStartDate As String, ByVal aEndDate As String, ByVal aPortiaConStr As String) As Integer
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer
        Try

            Dim ufunc As String = ReadConfigSetting("OffshoreFundsRecapFN")

            ' Connection string
            Dim connectionString As String = aPortiaConStr


            ' Query to execute the table-valued function
            Dim query As String = $"SELECT * FROM {ufunc}(@start_date, @end_date) ORDER BY TradeDate, PortId"

            ' Parameters
            Dim startDate As Date = aStartDate
            Dim endDate As Date = aEndDate

            ' Use a connection to execute the query
            Using connection As New SqlConnection(connectionString)
                Try
                    ' Open the connection
                    connection.Open()

                    ' Create a command object
                    Using command As New SqlCommand(query, connection)
                        ' Add parameters
                        command.Parameters.AddWithValue("@start_date", startDate)
                        command.Parameters.AddWithValue("@end_date", endDate)

                        ' Execute the command and read the data
                        Using reader As SqlDataReader = command.ExecuteReader()
                            ' Check if there are rows
                            If reader.HasRows Then

                                ' create a new excel file
                                Dim oXL As New Excel.Application
                                Dim theWorkbook As Excel.Workbook
                                Dim worksheet As Excel.Worksheet

                                theWorkbook = oXL.Workbooks.Add(objOpt)
                                worksheet = theWorkbook.ActiveSheet

                                ' create header
                                worksheet.Cells(1, 1) = "Date Range"
                                worksheet.Cells(1, 2) = aStartDate + " -- " + aEndDate

                                rowNum = 3

                                worksheet.Cells(rowNum, 1) = "Order ID"
                                worksheet.Cells(rowNum, 2) = "Portfolio"
                                worksheet.Cells(rowNum, 3) = "Security"
                                worksheet.Cells(rowNum, 4) = "Symbol"
                                worksheet.Cells(rowNum, 5) = "Tran Code"
                                worksheet.Cells(rowNum, 6) = "Trade Date"
                                worksheet.Cells(rowNum, 7) = "Qty"
                                worksheet.Cells(rowNum, 8) = "Price"
                                worksheet.Cells(rowNum, 9) = "Order TimeStamp"
                                worksheet.Cells(rowNum, 10) = "Order TimeStamp Time Zone"

                                ' Loop through the rows
                                rowNum = 4

                                While reader.Read()

                                    worksheet.Cells(rowNum, 1) = reader("orderid").ToString ' orderid
                                    worksheet.Cells(rowNum, 2) = reader("PortId").ToString   ' portid
                                    worksheet.Cells(rowNum, 3) = reader("shortname").ToString   ' sec name
                                    worksheet.Cells(rowNum, 4) = reader("symbol").ToString   ' symbol
                                    worksheet.Cells(rowNum, 5) = reader("trancode").ToString   ' tran code
                                    worksheet.Cells(rowNum, 6) = reader("TradeDate").ToString   ' trade date
                                    worksheet.Cells(rowNum, 7) = Format(reader("quantity"), "##,##0")        ' order qty
                                    worksheet.Cells(rowNum, 8) = Format(reader("price"), "##,##0.##")        ' price
                                    worksheet.Cells(rowNum, 9) = reader("OrderTimeStamp").ToString ' time stamp
                                    worksheet.Cells(rowNum, 10) = reader("OrderTimeStampTimeZone").ToString   ' time zone

                                    rowNum += 1


                                End While

                                reader.Close()

                                ' save generated Excel file
                                worksheet.Rows.Font.Size = 11
                                worksheet.Columns.AutoFit()
                                theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
                                theWorkbook.Close(False, objOpt, objOpt)

                                Marshal.ReleaseComObject(worksheet)
                                Marshal.ReleaseComObject(theWorkbook)
                                Marshal.ReleaseComObject(oXL)
                                worksheet = Nothing
                                theWorkbook = Nothing
                                oXL = Nothing

                            Else
                                MsgBox("No rows found.", vbOK, "Execute Result")
                            End If
                        End Using

                    End Using
                Catch ex As Exception
                    ' Handle exceptions
                    MsgBox("Error: " & ex.Message)
                End Try
            End Using



            Return 1
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
            Return -1
        End Try

    End Function

    Public Function GetOffshoreFundsRecapsFromPortiaRef(
    ByVal aStartDate As String,
    ByVal aEndDate As String,
    ByVal aPortiaConStr As String,
    ByVal aFilePath As String) As Integer

        ' Validate and parse dates
        Dim startDate As Date
        Dim endDate As Date
        If Not Date.TryParse(aStartDate, startDate) OrElse Not Date.TryParse(aEndDate, endDate) Then
            Throw New ArgumentException("Invalid start or end date")
        End If

        ' Read function name from configuration
        Dim ufunc As String = ReadConfigSetting("OffshoreFundsRecapFN")
        If String.IsNullOrWhiteSpace(ufunc) Then
            Throw New InvalidOperationException("Function name not found in configuration")
        End If

        ' SQL query
        Dim query As String = $"SELECT * FROM {ufunc}(@start_date, @end_date) ORDER BY TradeDate, PortId"

        ' Initialize Excel application
        Dim oXL As Excel.Application = Nothing
        Dim theWorkbook As Excel.Workbook = Nothing
        Dim worksheet As Excel.Worksheet = Nothing

        Try
            ' Open SQL connection and execute query
            Using connection As New SqlConnection(aPortiaConStr)
                connection.Open()
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@start_date", startDate)
                    command.Parameters.AddWithValue("@end_date", endDate)

                    Using reader As SqlDataReader = command.ExecuteReader()
                        If Not reader.HasRows Then
                            MsgBox("No rows found.", vbOK, "Execute Result")
                            Return 0
                        End If

                        Try
                            For i As Integer = 0 To reader.FieldCount - 1
                                Dim colName As String = reader.GetName(i)
                                Console.WriteLine($"Column {i + 1}: {colName}")
                            Next
                        Catch ex As Exception
                            screen.AppendText($"Error: {ex.Message}" + Environment.NewLine)
                        End Try


                        ' Create Excel workbook
                        oXL = New Excel.Application()
                        theWorkbook = oXL.Workbooks.Add()
                        worksheet = theWorkbook.ActiveSheet

                        ' Write header
                        WriteExcelHeader(worksheet, aStartDate, aEndDate)

                        ' Write data rows
                        Dim rowNum As Integer = 4
                        While reader.Read()
                            WriteExcelRow(worksheet, rowNum, reader)
                            rowNum += 1
                        End While

                        ' Format and save Excel file
                        FormatAndSaveExcel(worksheet, theWorkbook, aFilePath)
                    End Using
                End Using
            End Using

            Return 1

        Catch ex As Exception
            MsgBox("Error: " & ex.Message & vbCrLf & ex.StackTrace, vbCritical, "Error")
            Return -1

        Finally
            ' Release Excel COM objects
            If worksheet IsNot Nothing Then Marshal.ReleaseComObject(worksheet)
            If theWorkbook IsNot Nothing Then Marshal.ReleaseComObject(theWorkbook)
            If oXL IsNot Nothing Then
                oXL.Quit()
                Marshal.ReleaseComObject(oXL)
            End If

            worksheet = Nothing
            theWorkbook = Nothing
            oXL = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    ' Helper function to write Excel header
    Private Sub WriteExcelHeader(ByVal worksheet As Excel.Worksheet, ByVal aStartDate As String, ByVal aEndDate As String)
        worksheet.Cells(1, 1) = "Date Range"
        worksheet.Cells(1, 2) = $"{aStartDate} -- {aEndDate}"
        worksheet.Cells(3, 1) = "Order ID"
        worksheet.Cells(3, 2) = "Portfolio"
        worksheet.Cells(3, 3) = "Security"
        worksheet.Cells(3, 4) = "Symbol"
        worksheet.Cells(3, 5) = "Tran Code"
        worksheet.Cells(3, 6) = "Trade Date"
        worksheet.Cells(3, 7) = "Qty"
        worksheet.Cells(3, 8) = "Price"
        worksheet.Cells(3, 9) = "Order TimeStamp"
        worksheet.Cells(3, 10) = "Order TimeStamp Time Zone"
    End Sub

    ' Helper function to write a row of data
    Private Sub WriteExcelRow(ByVal worksheet As Excel.Worksheet, ByVal rowNum As Integer, ByVal reader As SqlDataReader)
        worksheet.Cells(rowNum, 1) = reader("orderid").ToString()
        worksheet.Cells(rowNum, 2) = reader("PortId").ToString()
        worksheet.Cells(rowNum, 3) = reader("shortname").ToString()
        worksheet.Cells(rowNum, 4) = reader("symbol").ToString()
        worksheet.Cells(rowNum, 5) = reader("trancode").ToString()
        worksheet.Cells(rowNum, 6) = reader("TradeDate").ToString()
        worksheet.Cells(rowNum, 7) = Format(reader("quantity"), "##,##0")
        worksheet.Cells(rowNum, 8) = Format(reader("price"), "##,##0.##")
        worksheet.Cells(rowNum, 9) = reader("OrderTimeStamp").ToString()
        worksheet.Cells(rowNum, 10) = reader("OrderTimeStampTimeZone").ToString()
    End Sub

    ' Helper function to format and save Excel file
    Private Sub FormatAndSaveExcel(ByVal worksheet As Excel.Worksheet, ByVal theWorkbook As Excel.Workbook, ByVal filePath As String)
        worksheet.Rows.Font.Size = 11
        worksheet.Columns.AutoFit()

        Dim fName As String = filePath
        theWorkbook.SaveAs(fName)
        theWorkbook.Close(False)
    End Sub

    Public Function getOffshoreFundsRecap(ByVal aStartDate As String, ByVal aEndDate As String) As Integer
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer
        Dim sec As String = ""

        Try
            If File.Exists(fName) Then
                File.Delete(fName)
            End If

            Dim Conn As New SqlConnection(moxyCon)
            Dim Cmd As New SqlCommand("usp_OffshreFundsRecap", Conn)
            Cmd.CommandType = CommandType.StoredProcedure

            Dim StartDate As New SqlParameter
            StartDate = Cmd.Parameters.Add("@startdate", SqlDbType.DateTime)
            StartDate.Direction = ParameterDirection.Input
            StartDate.Value = aStartDate

            Dim EndDate As New SqlParameter
            EndDate = Cmd.Parameters.Add("@enddate", SqlDbType.DateTime)
            EndDate.Direction = ParameterDirection.Input
            EndDate.Value = aEndDate

            Conn.Open()
            Dim myReader As SqlDataReader = Cmd.ExecuteReader()


            ' create a new excel file
            Dim oXL As New Excel.Application
            Dim theWorkbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(1, 1) = "Date Range"
            worksheet.Cells(1, 2) = aStartDate + " -- " + aEndDate

            rowNum = 3

            worksheet.Cells(rowNum, 1) = "Order ID"
            worksheet.Cells(rowNum, 2) = "Portfolio"
            worksheet.Cells(rowNum, 3) = "Security"
            worksheet.Cells(rowNum, 4) = "Symbol"
            worksheet.Cells(rowNum, 5) = "Tran Code"
            worksheet.Cells(rowNum, 6) = "Trade Date"
            worksheet.Cells(rowNum, 7) = "Qty"
            worksheet.Cells(rowNum, 8) = "Price"
            worksheet.Cells(rowNum, 9) = "Order TimeStamp"
            worksheet.Cells(rowNum, 10) = "Order TimeStamp Time Zone"





            rowNum = 4


            While myReader.Read()
                worksheet.Cells(rowNum, 1) = myReader.GetValue(0).ToString ' orderid
                'worksheet.Cells(rowNum, 4) = Format(myReader.GetValue(1), "##,##0.00000")       ' avg price 
                worksheet.Cells(rowNum, 2) = myReader.GetValue(1).ToString   ' portid
                worksheet.Cells(rowNum, 3) = myReader.GetValue(2).ToString   ' sec name
                worksheet.Cells(rowNum, 4) = myReader.GetValue(3).ToString   ' symbol
                worksheet.Cells(rowNum, 5) = myReader.GetValue(4).ToString   ' tran code
                worksheet.Cells(rowNum, 6) = myReader.GetValue(5).ToString   ' trade date
                worksheet.Cells(rowNum, 7) = Format(myReader.GetValue(6), "##,##0")        ' order qty
                worksheet.Cells(rowNum, 8) = Format(myReader.GetValue(7), "##,##0.##")        ' price
                worksheet.Cells(rowNum, 9) = myReader.GetValue(8).ToString ' time satmp
                worksheet.Cells(rowNum, 10) = myReader.GetValue(9).ToString   ' time zome

                rowNum += 1

            End While
            'worksheet.Range("All").Font.Size = 10
            myReader.Close()

            ' save generated Excel file
            worksheet.Rows.Font.Size = 11
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()
            Conn.Close()

            Return 1
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
            Return -1
        End Try

    End Function


    Public Function getPortfolioRecap(ByVal asOfDate As DateTime) As Integer
        ' if more accounts will require daily recaps,
        ' add parameter - account

        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer
        Dim sec As String = ""

        Try
            If File.Exists(fName) Then
                File.Delete(fName)
            End If

            Dim Conn As New SqlConnection(moxyCon)
            Dim Cmd As New SqlCommand("usp_TradeingRecapDaily", Conn)
            Cmd.CommandType = CommandType.StoredProcedure
            Dim PortDate As New SqlParameter
            PortDate = Cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
            PortDate.Direction = ParameterDirection.Input
            PortDate.Value = asOfDate

            Conn.Open()
            Dim myReader As SqlDataReader = Cmd.ExecuteReader()

            ' create a new excel file
            Dim oXL As New Excel.Application
            Dim theWorkbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(1, 1) = "Trade Date"
            worksheet.Cells(1, 2) = asOfDate.ToString("MM/dd/yyyy") + Space(5)

            worksheet.Cells(2, 1) = "Authorized by"
            worksheet.Cells(2, 2) = "AM"
            worksheet.Cells(3, 1) = "Daily Domestic Trade Execution Report"

            worksheet.Cells(5, 1) = "Security"
            worksheet.Cells(5, 2) = "Trancode"
            worksheet.Cells(5, 3) = "Qty"
            worksheet.Cells(5, 4) = "Price"
            worksheet.Cells(5, 5) = "Broker"
            worksheet.Cells(5, 6) = "Cusip"
            worksheet.Cells(5, 7) = "Reflow Flag"
            worksheet.Cells(5, 8) = "Lot Selection Method"


            worksheet.Range("A5", "H5").Font.Bold = True

            'worksheet.Cells(5, 5) = "Net Amount"
            'worksheet.Cells(5, 6) = "Broker"
            'worksheet.Cells(5, 7) = "Posted T/D"
            'worksheet.Cells(5, 8) = "Posted T+1"
            'worksheet.Cells(5, 9) = "PFPC Initial"

            rowNum = 6
            ' Column order in DataReader:
            '   0 - shortname, 1 - isin, 2 - trancode,
            '   3 - tradedate, 4 - qty

            While myReader.Read()
                worksheet.Cells(rowNum, 1) = myReader.GetValue(0).ToString ' security
                worksheet.Cells(rowNum, 4) = Format(myReader.GetValue(1), "##,##0.00000")       ' avg price 
                worksheet.Cells(rowNum, 2) = myReader.GetValue(2).ToString   ' tran code
                worksheet.Cells(rowNum, 3) = Format(myReader.GetValue(4), "##,##0")        ' order qty
                worksheet.Cells(rowNum, 5) = myReader.GetValue(5).ToString ' broker
                worksheet.Cells(rowNum, 6) = myReader.GetValue(6).ToString   ' cusip
                worksheet.Cells(rowNum, 7) = myReader.GetValue(7).ToString   ' reflow
                worksheet.Cells(rowNum, 8) = myReader.GetValue(8).ToString   ' lot selection method
                rowNum += 1

            End While
            'worksheet.Range("All").Font.Size = 10
            myReader.Close()
            rowNum += 5

            ' Get the allocations
            worksheet.Cells(rowNum, 1) = "Domestic Trading Allocation"
            rowNum += 1

            worksheet.Cells(rowNum, 1) = "Security"
            worksheet.Cells(rowNum, 2) = "Portfolio"
            worksheet.Cells(rowNum, 3) = "Tran Code"
            worksheet.Cells(rowNum, 4) = "Qty"
            worksheet.Cells(rowNum, 5) = "Price"

            worksheet.Range("A" + rowNum.ToString, "E" + rowNum.ToString).Font.Bold = True

            rowNum += 1
            Dim Conn2 As New SqlConnection(moxyCon)
            Dim Cmd2 As New SqlCommand("usp_TradingRecapAllocation", Conn)
            Cmd2.CommandType = CommandType.StoredProcedure
            Dim AllocDate As New SqlParameter
            AllocDate = Cmd2.Parameters.Add("@asofdate", SqlDbType.DateTime)
            AllocDate.Direction = ParameterDirection.Input
            AllocDate.Value = asOfDate

            Conn2.Open()
            Dim allocReader As SqlDataReader = Cmd2.ExecuteReader()
            While allocReader.Read()
                If sec <> allocReader.GetValue(3).ToString Then
                    rowNum += 1
                    worksheet.Cells(rowNum, 1) = allocReader.GetValue(3).ToString ' security
                End If
                sec = allocReader.GetValue(3).ToString
                worksheet.Cells(rowNum, 2) = allocReader.GetValue(2).ToString ' portfolio
                worksheet.Cells(rowNum, 3) = allocReader.GetValue(5).ToString   ' tran code
                worksheet.Cells(rowNum, 4) = Format(allocReader.GetValue(7), "##,##0")        ' order qty
                worksheet.Cells(rowNum, 5) = Format(allocReader.GetValue(8), "##,##0.00000")       ' avg price 

                rowNum += 1

            End While
            'worksheet.Range("All").Font.Size = 10
            myReader.Close()
            rowNum += 5

            ' save generated Excel file
            worksheet.Rows.Font.Size = 11
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()
            Conn.Close()
            Conn2.Close()

            Return 1
        Catch ex As Exception
            screen.Text += vbCrLf + ex.Message
            Return -1
        End Try


    End Function ' End of getPortfolioRecap

    Public Function getHedgeExposure(ByVal asOfDate As Date) As Integer
        Dim rtn As Integer = 0
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer = 0
        Dim row As DataRow

        Try
            If File.Exists(fName) Then
                File.Delete(fName)
            End If
            Dim Conn As New SqlConnection(moxyCon)
            Dim Cmd As New SqlCommand("usp_getHedgeExposure", Conn)
            screen.Text += vbCrLf + "Source procedure: " + Cmd.CommandText


            ' check the data in tb_HedgeExposure table on Moxy7
            Cmd.CommandType = CommandType.StoredProcedure

            Conn.Open()
            Dim myReader As SqlDataReader = Cmd.ExecuteReader()
            If myReader.HasRows Then
                If Me.createHdgExpCols(hdgDT) = -1 Then Return -1
                repInfo = New ReportInfo

                While myReader.Read()
                    row = hdgDT.NewRow()
                    row("pId") = myReader.GetValue(0)
                    row("bCur") = myReader.GetValue(1)
                    row("pName") = myReader.GetValue(2)
                    row("linkedCashAcct") = myReader.GetValue(3)

                    repInfo.VarExt = ""
                    Select Case UCase(row("bCur").ToString)
                        Case "CAD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\caduser"
                            repInfo.VarExt = """-l$fx ca"""
                        Case "AUD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\audxuser"
                            repInfo.VarExt = """-l$fx au"""
                        Case "GBP"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\gbpuser"
                            repInfo.VarExt = """-l$fx gb"""
                        Case "NZD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\nzduser"
                            repInfo.VarExt = """-l$fx nz"""
                        Case "EUR"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\euruser"
                            repInfo.VarExt = """-l$fx eu"""
                        Case "CHF"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\chfuser"
                            repInfo.VarExt = """-l$fx ch"""
                        Case Else
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
                            repInfo.VarExt = ""
                    End Select
                    repInfo.AxysMacro = "mbaman02.mac"
                    repInfo.Portfolio = "+@" + row("pId") + "hdg"
                    repInfo.PositionDate = asOfDate
                    repInfo.OutputFile = System.Windows.Forms.Application.StartupPath + "\" + row("pId").ToString + "pos.txt"
                    If File.Exists(repInfo.OutputFile) Then
                        File.Delete(repInfo.OutputFile)
                    End If
                    If Me.runAxysMacro(repInfo) = -1 Then Return -1
                    If Me.parseAxysPosition(repInfo.OutputFile, row, asOfDate) = -1 Then Return -1

                    Dim origTotAssts As Double = row("totAssts")

                    row("manAssts") = row("totAssts")
                    row("totAssts") = row("totAssts") - row("totL") ' this is total assts no hedges
                    row("nCash") = row("totL") + row("grCash")

                    ' get one month liabilities
                    Select Case UCase(row("bCur").ToString)
                        Case "CAD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\caduser"
                            repInfo.VarExt = """-l$fx ca#_option 3"""
                        Case "AUD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\audxuser"
                            repInfo.VarExt = """-l$fx au#_option 3"""
                        Case "GBP"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\gbpuser"
                            repInfo.VarExt = """-l$fx gb#_option 3"""
                        Case "NZD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\nzduser"
                            repInfo.VarExt = """-l$fx nz#_option 3"""
                        Case "EUR"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\euruser"
                            repInfo.VarExt = """-l$fx eu#_option 3"""
                        Case "CHF"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\chfuser"
                            repInfo.VarExt = """-l$fx ch#_option 3"""
                        Case Else
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
                            repInfo.VarExt = """-l$fx us#_option 3"""
                    End Select
                    repInfo.AxysMacro = "mbliab.mac"
                    repInfo.OutputFile = System.Windows.Forms.Application.StartupPath + "\" + row("pId").ToString + "liab.txt"
                    If File.Exists(repInfo.OutputFile) Then
                        File.Delete(repInfo.OutputFile)
                    End If
                    If Me.runAxysMacro(repInfo) = -1 Then Return -1
                    If Me.parseAxysLiability(repInfo.OutputFile, row) Then Return -1

                    ' get three month liabilities
                    Select Case UCase(row("bCur").ToString)
                        Case "CAD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\caduser"
                            repInfo.VarExt = """-l$fx ca#_option 4"""
                        Case "AUD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\audxuser"
                            repInfo.VarExt = """-l$fx au#_option 4"""
                        Case "GBP"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\gbpuser"
                            repInfo.VarExt = """-l$fx gb#_option 4"""
                        Case "NZD"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\nzduser"
                            repInfo.VarExt = """-l$fx nz#_option 4"""
                        Case "EUR"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\euruser"
                            repInfo.VarExt = """-l$fx eu#_option 4"""
                        Case "CHF"
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\chfuser"
                            repInfo.VarExt = """-l$fx ch#_option 4"""
                        Case Else
                            repInfo.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
                            repInfo.VarExt = """-l$fx us#_option 4"""
                    End Select
                    repInfo.AxysMacro = "mbliab.mac"
                    repInfo.OutputFile = System.Windows.Forms.Application.StartupPath + "\" + row("pId").ToString + "3Mliab.txt"
                    If File.Exists(repInfo.OutputFile) Then
                        File.Delete(repInfo.OutputFile)
                    End If
                    If Me.runAxysMacro(repInfo) = -1 Then Return -1
                    If Me.parseAxys3MLiability(repInfo.OutputFile, row) Then Return -1

                    ' calc hedge pct
                    row("pctH") = 1 - (row("grCash") / row("totAssts"))
                    row("pctUH") = 1 - (row("nCash") / row("totAssts"))

                    hdgDT.Rows.Add(row)
                End While
            End If
            myReader.Close()

            Me.createExclHdgRep(hdgDT, repInfo.PositionDate, fName)

        Catch ex As Exception
            screen.Text += vbCrLf + "getHedgeExposue error: " + ex.Message
            Return -1
        End Try

    End Function

    Protected Function createHdgExpCols(ByRef aDT As System.Data.DataTable) As Integer
        Dim rtn As Integer = 0
        Dim col As DataColumn
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name


        Try
            aDT = New System.Data.DataTable

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.String")
            col.ColumnName = "pId"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.String")
            col.ColumnName = "pName"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "oneML"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "threeML"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "totL"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.String")
            col.ColumnName = "bCur"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "grCash"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "nCash"
            col.Unique = False
            aDT.Columns.Add(col)


            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "manAssts"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "pctH"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "pctUH"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.Double")
            col.ColumnName = "totAssts"
            col.Unique = False
            aDT.Columns.Add(col)

            col = New DataColumn()
            col.DataType = System.Type.GetType("System.String")
            col.ColumnName = "linkedCashAcct"
            col.Unique = False
            aDT.Columns.Add(col)

        Catch ex As Exception
            screen.Text += vbCrLf + "createHedgeExpCols error: " + ex.Message
            Return -1
        End Try
    End Function


    Public Function runAxysMacro(ByVal aRepInfo As ReportInfo) As Integer
        ' Runs Axys macros
        Dim rtn As Integer = 0
        'Dim myDir As String = Application.StartupPath + "\"   ' application current directory
        Dim saveType As String = " -vx"          ' forces Axys to save macrooutput to the text file
        Dim macroName As String = repInfo.AxysMacro  ' the name of the macro in Axys 
        Dim portfolio As String = repInfo.Portfolio
        ' portfolio number in Axys
        Dim asOfDate As String = repInfo.PositionDate.ToString("MMddyy") ' positions date
        Dim axysPath As String = "\\tweedy_files\advent\Axys3\rep32.exe"
        Dim axysProc As New ProcessStartInfo(axysPath)
        Dim p As Process
        Dim fName As String = aRepInfo.OutputFile
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Try
            axysProc.Arguments = " -m" + macroName + " -p" + portfolio + saveType + " -u -b" + asOfDate + " -t" + fName
            If aRepInfo.VarExt.Length > 0 Then axysProc.Arguments += Space(1) + aRepInfo.VarExt
            ' delete old file with positions
            File.Delete(fName)
            axysProc.WorkingDirectory = repInfo.WorkingDirectory
            'axysProc.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
            'run Axys rep32 to produce positions file
            p = Process.Start(axysProc)
            While Not p.HasExited
                ' wait for the process to finish
                System.Windows.Forms.Application.DoEvents()
            End While
            If File.Exists(fName) Then
                screen.Text += vbCrLf + "Axys created positions file " + fName
            Else
                mo.errMsg04(mSig, "Axys failed to create positions file", fName)
                Return -1
            End If

            'Set the cursor to the end of the textbox.
            screen.SelectionStart = screen.TextLength
            '
            'Scroll down to the cursor position.
            screen.ScrollToCaret()
        Catch ex As Exception
            ' Let the user know what went wrong.
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function


    Public Function runAxysMacro02(ByVal aRepInfo As ReportInfo) As Integer
        ' Runs Axys macros
        Dim rtn As Integer = 0
        'Dim myDir As String = Application.StartupPath + "\"   ' application current directory
        Dim saveType As String = " -vx"          ' forces Axys to save macrooutput to the text file
        Dim macroName As String = repInfo.AxysMacro  ' the name of the macro in Axys 
        Dim portfolio As String = aRepInfo.Portfolio
        ' portfolio number in Axys
        Dim asOfDate As String = repInfo.PositionDate.ToString("MMddyy") ' positions date
        Dim axysPath As String = "\\tweedy_files\advent\Axys3\rep32.exe"
        Dim axysProc As New ProcessStartInfo(axysPath)
        Dim p As Process
        Dim fName As String = aRepInfo.OutputFile
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Try
            axysProc.Arguments = " -m" + macroName + " -p" + portfolio + saveType + " -u -b" + asOfDate + " -t" + fName
            If aRepInfo.VarExt.Length > 0 Then axysProc.Arguments += Space(1) + aRepInfo.VarExt
            ' delete old file with positions
            File.Delete(fName)
            axysProc.WorkingDirectory = repInfo.WorkingDirectory
            'axysProc.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
            'run Axys rep32 to produce positions file
            p = Process.Start(axysProc)
            While Not p.HasExited
                ' wait for the process to finish
                System.Windows.Forms.Application.DoEvents()
            End While
            If File.Exists(fName) Then
                screen.Text += vbCrLf + "Axys created positions file " + fName
            Else
                mo.errMsg04(mSig, "Axys failed to create positions file", fName)
                Return -1
            End If

            'Set the cursor to the end of the textbox.
            screen.SelectionStart = screen.TextLength
            '
            'Scroll down to the cursor position.
            screen.ScrollToCaret()
        Catch ex As Exception
            ' Let the user know what went wrong.
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function

    Protected Function parseAxysPosition(ByVal fname As String, ByRef aRow As DataRow, ByVal asOfDate As Date) As Integer
        Dim rtn As Integer = 0
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim cnt As Integer = 0
        Dim oRead As System.IO.StreamReader
        Dim line As String = Nothing
        Dim ri As ReportInfo
        Dim tmp As Double

        Try

            oRead = File.OpenText(fname)
            While oRead.Peek <> -1
                cnt += 1
                line = oRead.ReadLine()
                If cnt > 9 And line.Length > 0 Then
                    tmp = CDbl(Trim(line.Substring(107, 15)))
                    aRow("totL") = CDbl(Trim(line.Substring(107, 15)))
                    aRow("grCash") = CDbl(Trim(line.Substring(92, 14)))
                    aRow("totAssts") = CDbl(Trim(line.Substring(138, 16)))
                    Exit While ' read only one line
                End If

            End While
            oRead.Close()

            ' check for cash in linked accounts
            If Not IsDBNull(aRow("linkedCashAcct")) Then
                ri = New ReportInfo
                ri.VarExt = ""
                Select Case UCase(aRow("bCur").ToString)
                    Case "CAD"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\caduser"
                        ri.VarExt = """-l$fx ca"""
                    Case "AUD"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\audxuser"
                        repInfo.VarExt = """-l$fx au"""
                    Case "GBP"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\gbpuser"
                        ri.VarExt = """-l$fx gb"""
                    Case "NZD"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\nzduser"
                        ri.VarExt = """-l$fx nz"""
                    Case "EUR"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\euruser"
                        ri.VarExt = """-l$fx eu"""
                    Case "CHF"
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\chfuser"
                        ri.VarExt = """-l$fx ch"""
                    Case Else
                        ri.WorkingDirectory = "\\tweedy_files\advent\Axys3\users\ususon"
                        ri.VarExt = ""
                End Select
                ri.AxysMacro = "mbaman02.mac"
                ri.Portfolio = aRow("linkedCashAcct")
                If ri.Portfolio = "83583" Then
                    rtn = rtn
                End If
                ri.PositionDate = asOfDate
                ri.OutputFile = System.Windows.Forms.Application.StartupPath + "\" + ri.Portfolio + "pos.txt"
                If File.Exists(ri.OutputFile) Then
                    File.Delete(ri.OutputFile)
                End If
                If Me.runAxysMacro02(ri) = -1 Then Return -1
                If Me.parseAxysLinkedCash(ri.OutputFile, aRow) = -1 Then Return -1
            End If


        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function

    Protected Function parseAxysLinkedCash(ByVal fname As String, ByRef aRow As DataRow) As Integer
        Dim rtn As Integer = 0
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim cnt As Integer = 0
        Dim oRead As System.IO.StreamReader
        Dim line As String = Nothing
        Try
            oRead = File.OpenText(fname)
            While oRead.Peek <> -1
                cnt += 1
                line = oRead.ReadLine()
                If cnt > 9 And line.Length > 0 Then

                    If aRow("pId") = "24976" Then
                        rtn = rtn
                    End If
                    Dim tmp As Double = CDbl(Trim(line.Substring(138, 16)))
                    aRow("grCash") += CDbl(Trim(line.Substring(138, 16)))
                    aRow("TotAssts") += CDbl(Trim(line.Substring(138, 16)))

                    aRow("manAssts") = aRow("totAssts")
                    aRow("totAssts") = aRow("totAssts") - aRow("totL")  ' this is total assts no hedges
                    aRow("nCash") = aRow("totL") + aRow("grCash")
                    Exit While ' read only one line
                End If

            End While
            oRead.Close()


        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function

    Protected Function parseAxysLiability(ByVal fname As String, ByRef aRow As DataRow) As Integer
        Dim rtn As Integer = 0
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim cnt As Integer = 0
        Dim oRead As System.IO.StreamReader
        Dim line As String = Nothing

        Try

            oRead = File.OpenText(fname)
            While oRead.Peek <> -1
                cnt += 1
                line = oRead.ReadLine()
                If line.IndexOf("TOTAL") <> -1 And line.Length > 0 And cnt > 10 Then

                    aRow("oneML") = CDbl(Trim(line.Substring(120, 14)))

                    Exit While ' read only one line
                End If

            End While
            oRead.Close()
        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function

    Protected Function parseAxys3MLiability(ByVal fname As String, ByRef aRow As DataRow) As Integer
        Dim rtn As Integer = 0
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim cnt As Integer = 0
        Dim oRead As System.IO.StreamReader
        Dim line As String = Nothing
        ' three month liability
        Try

            oRead = File.OpenText(fname)
            While oRead.Peek <> -1
                cnt += 1
                line = oRead.ReadLine()
                If line.IndexOf("TOTAL") <> -1 And line.Length > 0 And cnt > 10 Then

                    aRow("threeML") = CDbl(Trim(line.Substring(120, 14)))

                    Exit While ' read only one line
                End If

            End While
            oRead.Close()
        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function

    Protected Function createExclHdgRep(ByVal aDT As System.Data.DataTable, ByVal asOfDate As Date, ByVal aExclFile As String) As Integer
        Dim rtn As Integer = 0
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim cnt As Integer = 0
        Dim objOpt As Object = System.Reflection.Missing.Value
        Dim rowNum As Integer = 1
        Dim row As DataRow
        Try
            ' create a new excel file
            Dim oXL As New Excel.Application
            Dim theWorkbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet

            theWorkbook = oXL.Workbooks.Add(objOpt)
            worksheet = theWorkbook.ActiveSheet

            ' create header
            worksheet.Cells(rowNum, 3) = "Report Date"
            worksheet.Cells(rowNum, 4) = asOfDate.ToString("MM/dd/yyyy") + Space(5)
            worksheet.Cells(rowNum, 2) = "Hedge Exposure Report"
            worksheet.Range("B" + rowNum.ToString).EntireRow.Font.Size = 10

            rowNum += 2

            worksheet.Cells(rowNum, 1) = "Port Id"
            worksheet.Cells(rowNum, 2) = "Port Name"
            worksheet.Cells(rowNum, 3) = "1 Month Liab"
            worksheet.Cells(rowNum, 4) = "3 Month Liab"
            worksheet.Cells(rowNum, 5) = "Total Liab"
            worksheet.Cells(rowNum, 6) = "Base Curr"
            worksheet.Cells(rowNum, 7) = "Gross Cash"
            worksheet.Cells(rowNum, 8) = "Net Cash"
            worksheet.Cells(rowNum, 9) = "Managed Assets"
            worksheet.Cells(rowNum, 10) = "% UnHedged"
            worksheet.Cells(rowNum, 11) = "% Hedged"
            worksheet.Cells(rowNum, 12) = "Total Assets (No Hedges) "
            worksheet.Range("B" + rowNum.ToString).EntireRow.Font.Bold = True
            'worksheet.Range("B" + rowNum.ToString).Borders.Weight = 2
            worksheet.Range("B" + rowNum.ToString).EntireRow.Borders.Weight = 2
            worksheet.Range("B" + rowNum.ToString).EntireRow.Font.Size = 10

            For Each row In aDT.Rows
                rowNum += 1
                worksheet.Cells(rowNum, 1) = row("pId")
                worksheet.Range("A" + rowNum.ToString).HorizontalAlignment = Excel.Constants.xlLeft
                worksheet.Cells(rowNum, 2) = row("pName")
                worksheet.Cells(rowNum, 3) = Me.formatAmount02(row("oneML").ToString)
                worksheet.Cells(rowNum, 4) = Me.formatAmount02(row("threeML").ToString)
                worksheet.Cells(rowNum, 5) = Me.formatAmount02(row("totL").ToString())
                worksheet.Cells(rowNum, 6) = row("bCur")
                worksheet.Range("F" + rowNum.ToString).HorizontalAlignment = Excel.Constants.xlRight

                worksheet.Cells(rowNum, 7) = Me.formatAmount02(row("grCash").ToString)
                worksheet.Cells(rowNum, 8) = Me.formatAmount02(row("nCash").ToString)
                worksheet.Cells(rowNum, 9) = Me.formatAmount02(row("ManAssts").ToString)
                worksheet.Cells(rowNum, 10) = Me.formatAmount(100 * row("pctH").ToString)
                worksheet.Cells(rowNum, 11) = Me.formatAmount(100 * row("pctUH").ToString)
                worksheet.Cells(rowNum, 12) = Me.formatAmount02(row("totAssts").ToString)
                worksheet.Range("B" + rowNum.ToString).EntireRow.Borders.Weight = 2
                worksheet.Range("B" + rowNum.ToString).EntireRow.Font.Size = 10
                worksheet.Range("B" + rowNum.ToString).EntireRow.RowHeight *= 1.2


            Next

            ' save generated Excel file
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(aExclFile, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()

        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
            rtn = -1
        End Try
        Return rtn
    End Function


    Public Function readTRNFile(fileName As String) As List(Of GTSSObj)
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Dim list As New List(Of GTSSObj)

        Try
            ' Open the file to read from.
            Dim readText() As String = File.ReadAllLines(fileName)
            Dim s As String
            For Each s In readText
                If s.Length = 0 Then Continue For
                ' Split line on comma.
                Dim parts As String() = s.Split(New Char() {","c})
                Try
                    list.Add(New GTSSObj(parts(0), parts(1), parts(3), parts(5), parts(6), parts(11), parts(8), parts(17), parts(24), moxyCon))

                Catch e As Exception
                    mo.errMsg01(mSig, e.Message)
                End Try


            Next

        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
        End Try

        Return list
    End Function

    Public Function createGTSSFile(ByVal tradesList As List(Of GTSSObj)) As String
        Dim outFile As String = vbNull
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Try
            Dim cnt As Integer = 0

            Dim list As New List(Of String)
            ' create header
            list.Add("300")
            list.Add(createHeaderLine())
            ' create body
            For Each trade As GTSSObj In tradesList
                System.Windows.Forms.Application.DoEvents()
                If trade.cashTran Or trade.excludedCurrency Then Continue For
                cnt += 1

                list.Add(createGTSSTradeLine(trade, cnt))
            Next

            list.Add(createFooterLine(cnt))
            ' save file
            'outFile = ReadConfigSetting("GTSSOUTFile")
            outFile = fName
            File.WriteAllLines(outFile, list)

        Catch ex As Exception
            mo.errMsg01(mSig, ex.Message)
        End Try

        Return outFile
    End Function

    Private Function createGTSSTradeLine(g As GTSSObj, cnt As Integer) As String
        Dim sTradeLine As String = String.Empty
        Dim mSig As String = (New System.Diagnostics.StackFrame()).GetMethod().Name
        Try
            Dim tmp As String = g.tradeDate.Insert(2, "/").Insert(5, "/")
            g.tradeDate = formatDate(tmp)
            tmp = g.settleDate.Insert(2, "/").Insert(5, "/")
            g.settleDate = formatDate(tmp)

            sTradeLine = "TWEEUSNYXXX" + Chr(9) + g.bicCode + Chr(9) + "15A" + Chr(9)
            sTradeLine += g.newSeed.ToString() + Chr(9) + "" + Chr(9) + "NEWT" + Chr(9)
            sTradeLine += "" + Chr(9) + Strings.Left(g.bicCode, 6) + "TWEEUS" + cnt.ToString() + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + "TWEEUSNYXXX" + Chr(9) + g.bicCode + Chr(9)
            sTradeLine += "" + Chr(9) + g.fixingDate + Chr(9) + "15B" + Chr(9)
            sTradeLine += g.tradeDate + Chr(9) + g.settleDate + Chr(9) + g.fxRate + Chr(9)
            sTradeLine += UCase(g.curr) + g.localAmount + Chr(9) + g.deliveryAgent + Chr(9) + "" + Chr(9)
            sTradeLine += g.receivingAgent + Chr(9) + UCase(g.curr2) + g.usdAmount + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + g.receivingAgent2 + Chr(9) + "/NETS/" + Chr(9)
            'sTradeLine += "" + Chr(9) + g.receivingAgent2 + Chr(9) + "/NETS/" + g.bicCode + Chr(9)
            ' sTradeLine += "" + Chr(9) + g.receivingAgent2 + Chr(9) + "/ABIC/" + g.bicCode + "/NAME/UKWN" + Chr(9)
            sTradeLine += "15C" + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "/GLCID/" + g.portCode + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "" + Chr(9) + "" + Chr(9) + "" + Chr(9)
            sTradeLine += "-"
        Catch ex As Exception
            Throw New Exception(vbCrLf + mSig + " " + ex.Message)
        End Try
        Return sTradeLine
    End Function

    Public Shared Function ReadConfigSetting(key As String) As String
        Dim result As String = vbNull
        Try

            Dim appSettings As Object = ConfigurationManager.AppSettings
            result = appSettings(key)
            If IsNothing(result) Then
                result = "Not found"
            End If

        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
        End Try
        Return (result)
    End Function

    Public Shared Function ExecuteFundTradesRecapAllFundsReader(asOfDate As DateTime,
                                                            moxyCon As String) As SqlDataReader
        Dim conn As New SqlConnection(moxyCon)
        Try
            Dim cmd As New SqlCommand("usp_FundTradesRecapAllFunds", conn) With {
            .CommandType = CommandType.StoredProcedure,
            .CommandTimeout = 120
        }
            cmd.Parameters.Add("@rundate", SqlDbType.DateTime).Value = asOfDate

            conn.Open()
            ' CloseConnection ensures closing the reader will close the connection
            Return cmd.ExecuteReader(CommandBehavior.CloseConnection)

        Catch ex As Exception
            conn.Dispose()
            Throw New Exception("Error in ExecuteFundTradesRecapAllFundsReader: " & ex.Message)
        End Try
    End Function


End Class

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Security.Principal
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic
Imports System.Windows.Forms
Imports System.Drawing

' Last Update: 07/17/2012
' Application is integrated with .Net 4, Excel 2007
' Last Update: 12/14/2020
'                       7/23/2025
' Application is connected to Moxy 24

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim connStr As String = ReadConfigSetting("MoxyConnection")
    Dim portiaConStr As String = ReadConfigSetting("PortiaConnection")


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.Text = "FX Connect - Moxy 24"


    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button_FXConPreAlloc As System.Windows.Forms.Button
    Friend WithEvents Button_FXConTradesAlloc As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button_ImportTRN As System.Windows.Forms.Button
    Friend WithEvents btn_AxysFCTrades As System.Windows.Forms.Button
    Friend WithEvents btnFundTrades As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ButtonHedgeExposure As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ButtonMoxyExport As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnCreateGTSSFile As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents DataView1 As System.Data.DataView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button_FXConPreAlloc = New System.Windows.Forms.Button()
        Me.Button_FXConTradesAlloc = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Button_ImportTRN = New System.Windows.Forms.Button()
        Me.btn_AxysFCTrades = New System.Windows.Forms.Button()
        Me.btnFundTrades = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataView1 = New System.Data.DataView()
        Me.ButtonHedgeExposure = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ButtonMoxyExport = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnCreateGTSSFile = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        resources.ApplyResources(Me.TextBox1, "TextBox1")
        Me.TextBox1.Name = "TextBox1"
        '
        'Button_FXConPreAlloc
        '
        resources.ApplyResources(Me.Button_FXConPreAlloc, "Button_FXConPreAlloc")
        Me.Button_FXConPreAlloc.Name = "Button_FXConPreAlloc"
        '
        'Button_FXConTradesAlloc
        '
        resources.ApplyResources(Me.Button_FXConTradesAlloc, "Button_FXConTradesAlloc")
        Me.Button_FXConTradesAlloc.Name = "Button_FXConTradesAlloc"
        '
        'DataGrid1
        '
        resources.ApplyResources(Me.DataGrid1, "DataGrid1")
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Name = "DataGrid1"
        '
        'Button_ImportTRN
        '
        resources.ApplyResources(Me.Button_ImportTRN, "Button_ImportTRN")
        Me.Button_ImportTRN.Name = "Button_ImportTRN"
        '
        'btn_AxysFCTrades
        '
        resources.ApplyResources(Me.btn_AxysFCTrades, "btn_AxysFCTrades")
        Me.btn_AxysFCTrades.Name = "btn_AxysFCTrades"
        '
        'btnFundTrades
        '
        resources.ApplyResources(Me.btnFundTrades, "btnFundTrades")
        Me.btnFundTrades.Name = "btnFundTrades"
        '
        'Button1
        '
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        '
        'ButtonHedgeExposure
        '
        resources.ApplyResources(Me.ButtonHedgeExposure, "ButtonHedgeExposure")
        Me.ButtonHedgeExposure.Name = "ButtonHedgeExposure"
        Me.ButtonHedgeExposure.UseVisualStyleBackColor = True
        '
        'Button2
        '
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ButtonMoxyExport
        '
        resources.ApplyResources(Me.ButtonMoxyExport, "ButtonMoxyExport")
        Me.ButtonMoxyExport.Name = "ButtonMoxyExport"
        Me.ButtonMoxyExport.UseVisualStyleBackColor = True
        '
        'Button3
        '
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnCreateGTSSFile
        '
        resources.ApplyResources(Me.btnCreateGTSSFile, "btnCreateGTSSFile")
        Me.btnCreateGTSSFile.Name = "btnCreateGTSSFile"
        Me.btnCreateGTSSFile.UseVisualStyleBackColor = True
        '
        'Button4
        '
        resources.ApplyResources(Me.Button4, "Button4")
        Me.Button4.Name = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Form1
        '
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.btnCreateGTSSFile)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.ButtonMoxyExport)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ButtonHedgeExposure)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnFundTrades)
        Me.Controls.Add(Me.btn_AxysFCTrades)
        Me.Controls.Add(Me.Button_ImportTRN)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.Button_FXConTradesAlloc)
        Me.Controls.Add(Me.Button_FXConPreAlloc)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Button_FXConPreAlloc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_FXConPreAlloc.Click
        TextBox1.Text = "" ' clear previous contents

        Dim Conn As New SqlConnection("Data Source=MOXY1;Initial Catalog=moxy;Integrated Security=SSPI")
        'Dim rsData As ADODB.Recordset
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetFXConTradesByDate", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet
        Dim fName As String
        Dim bNoTrades As Boolean = True
        Dim sHeaders As String = ""
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application
        Dim excelBook As Microsoft.Office.Interop.Excel.Workbook


        'file save name
        fName = "H:\FXCON\Moxy18\moxyfxexportTest.xlsx"

        ' get allocated FXCon trades from Moxy
        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim asofdate As SqlParameter = Cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
        asofdate.Direction = ParameterDirection.Input
        asofdate.Value = Today

        DA.SelectCommand = Cmd

        Try
            Conn.Open()
            DA.Fill(DSet, "trades")
            DataGrid1.SetDataBinding(DSet, "trades")
            Dim DTable As DataTable = DSet.Tables("trades")

            If DTable.Rows.Count = 0 Then
                ' no trades for the date
                TextBox1.Text = "No trades found for " + asofdate.Value
                Return
            End If

            ' save datagrid as Excel file
            Dim fi As New FileInfo(fName)


            Dim excelWorksheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim dr As DataRow
            Dim i As Integer = 1

            If fi.Exists Then
                'open existing file, clear
                excelBook = excelApp.Workbooks.Open(fName)
                excelWorksheet = CType(excelBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                excelWorksheet.UsedRange.Clear()
            Else
                ' create new file
                excelBook = excelApp.Workbooks.Add
                excelWorksheet = CType(excelBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
            End If

            'fill Excel worksheet with data
            For Each dr In DTable.Rows
                excelWorksheet.Range("A" & i.ToString).Value = dr("tradedate")
                excelWorksheet.Range("B" & i.ToString).Value = dr("settledate")
                excelWorksheet.Range("C" & i.ToString).Value = dr("portfolio")
                excelWorksheet.Range("D" & i.ToString).Value = dr("amount")
                excelWorksheet.Range("E" & i.ToString).Value = dr("trancode")
                excelWorksheet.Range("F" & i.ToString).Value = dr("sectype")
                i += 1
            Next

            ' save worksheet
            excelWorksheet.UsedRange.EntireColumn.AutoFit()
            excelApp.DisplayAlerts = False
            excelWorksheet.SaveAs(fName)
            excelApp.DisplayAlerts = True
            excelBook.Close()

            TextBox1.Text = "Created file " + fName

        Catch ex As Exception
            TextBox1.Text += vbCrLf + ex.Message
        Finally
            'excelBook.Close()
            Conn.Close()
        End Try

        Conn.Close()

    End Sub

    Private Sub Button_FXConTradesAlloc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_FXConTradesAlloc.Click
        TextBox1.Text = "" ' clear previous contents

        Dim Conn As New SqlConnection(connStr)
        'Dim rsData As ADODB.Recordset
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetFXConTradesByDateWithFX", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet
        Dim fName As String
        'Dim bNoTrades As Boolean
        'Dim sHeaders As String
        Dim allocDate As Object


        Dim outFolder As String = ReadConfigSetting("TradesAllocationFolder")

        'file save name
        fName = outFolder + "\BuyMT300a.txt"

        ' Ask user to enter the trade date
        allocDate = InputBox("Please enter the allocation date", "Request", Today())

        If allocDate = "" Then
            'user canceled 
            Return
        End If

        If Not IsDate(allocDate) Then
            MsgBox("Not a valid allocation date", MsgBoxStyle.OkOnly)
            Return
        End If


        ' get allocated FXCon trades from Moxy
        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim asofdate As SqlParameter = Cmd.Parameters.Add("@asofdate", SqlDbType.DateTime)
        asofdate.Direction = ParameterDirection.Input
        asofdate.Value = allocDate  'Today

        DA.SelectCommand = Cmd

        Conn.Open()
        DA.Fill(DSet, "trades")
        DataGrid1.SetDataBinding(DSet, "trades")
        Dim DTable As DataTable = DSet.Tables("trades")

        If DTable.Rows.Count = 0 Then
            ' no trades for the date
            TextBox1.Text = "No trades found for " + asofdate.Value
            Return
        End If

        ' create GTSS upload text file
        Dim fm As New FXConManager(fName, TextBox1, connStr, "")

        If fm.addFXConTrades(DTable) <> -1 Then

            TextBox1.Text = "Created file " + fName
        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName
        End If

        Conn.Close()

        ' Print where clause on the screen:
        TextBox1.Text += vbCrLf + "Filter settings: "
        TextBox1.Text += vbCrLf + "Group name in global, intl, globhigh, nmptr, glc "
        TextBox1.Text += vbCrLf + "Portfolio not in the table tb_FXCONN_EXCEPTIONS "
        TextBox1.Text += vbCrLf + "Security type not in csus, gsus, adus, cakr, cahr, cath, catw "
        TextBox1.Text += vbCrLf + "Broker not fx.dbcc "

    End Sub

    Private Sub Button_ImportTRN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ImportTRN.Click

        Dim fNameUSD As String
        Dim fNameCAD As String
        Dim fNameEUR As String
        Dim fnameCHF As String
        Dim fnameNZD As String
        Dim fnameGBP As String
        Dim axysPath As String
        Dim usdFlag As Boolean  ' indicates that there are USD trades to be imported
        Dim cadFlag As Boolean  ' indicates that there are CAD trades to be imported
        Dim eurFlag As Boolean  ' indicates that there are EUR trades to be imported
        Dim chfFlag As Boolean  ' indicates that there are CHF trades to be imported
        Dim nzdFlag As Boolean  ' indicates that there are NZD trades to be imported
        Dim gbpFlag As Boolean
        Dim fName As String
        Dim portfolio As String
        Dim portCodeToUse As String
        Dim baseCurrency As String
        Dim blotterPath As String = "" ' defines the user blotter
        Dim cadBlotterpath As String
        Dim eurBlotterpath As String
        Dim chfBlotterpath As String
        Dim nzdBlotterpath As String
        Dim gbpBlotterpath As String
        Dim fwUSD As StreamWriter
        Dim fwCAD As StreamWriter
        Dim fwEUR As StreamWriter
        Dim fwCHF As StreamWriter
        Dim fwNZD As StreamWriter
        Dim fwGBP As StreamWriter
        Dim srcFile As String

        srcFile = "H:\FXCON\topost1.trn"
        fName = "H:\FXCON\topost.trn"
        fNameUSD = "H:\FXCON\topostus.trn"
        fNameCAD = "H:\FXCON\topostca.trn"
        fNameEUR = "H:\FXCON\toposteu.trn"
        fnameCHF = "H:\FXCON\topostch.trn"
        fnameNZD = "H:\FXCON\topostnz.trn"
        fnameGBP = "H:\FXCON\topostgb.trn"
        axysPath = "H:\Axys3\"


        cadBlotterpath = "H:\Axys3\CAD\"
        eurBlotterpath = "H:\Axys3\EUR\"
        chfBlotterpath = "H:\Axys3\CHF\"
        nzdBlotterpath = "H:\Axys3\NZD\"
        gbpBlotterpath = "H:\Axys3\GBP\"
        ' assign the blotter path based on the windows user
        ' blotterPath = "H:\Axys3\USERS\JEANNEPR\" ' default
        Dim idWindows As WindowsIdentity = WindowsIdentity.GetCurrent
        Select Case UCase(idWindows.Name)
            Case UCase("TWEEDY\mikeba")
                blotterPath = "H:\Axys3\USERS\MIKEBA\"

            Case UCase("TWEEDY\jeannepr")
                blotterPath = "H:\Axys3\USERS\JEANNEPR\"
            Case UCase("TWEEDY\annmariema")
                blotterPath = "H:\Axys3\USERS\ANNMARIE\"

        End Select

        If blotterPath.Length = 0 Then
            TextBox1.Text += vbCrLf + "US blotter path undefined... "
            Return
        Else
            TextBox1.Text += vbCrLf + "US blotter: " + blotterPath
        End If

        Dim row As DataRow
        Dim rows() As DataRow
        Dim Conn As New SqlConnection(connStr)
        Dim Cmd As SqlCommand = New SqlCommand("usp_GetPortMap", Conn)
        Dim DA As SqlDataAdapter = New SqlDataAdapter
        Dim DSet As New DataSet
        Cmd.CommandType = CommandType.StoredProcedure
        Dim RetValue As SqlParameter = Cmd.Parameters.Add("RetValue", SqlDbType.Int)
        RetValue.Direction = ParameterDirection.ReturnValue
        Dim portcode As SqlParameter = Cmd.Parameters.Add("@portcode", SqlDbType.VarChar)
        portcode.Direction = ParameterDirection.Input
        DA.SelectCommand = Cmd

        ' read through topost.trn and put trades in to
        ' appropriate topostXX.trn file based on the base currency
        Try
            Dim sr As StreamReader = New StreamReader(srcFile)
            Dim line As String

            fwUSD = File.CreateText(fNameUSD)
            fwCAD = File.CreateText(fNameCAD)
            fwEUR = File.CreateText(fNameEUR)
            fwCHF = File.CreateText(fnameCHF)
            fwNZD = File.CreateText(fnameNZD)
            fwGBP = File.CreateText(fnameGBP)

            Do
                line = sr.ReadLine()
                If line Is Nothing Then Exit Do

                ' first 5 charecters in each line are portfolio codes
                portfolio = Microsoft.VisualBasic.Left(line, 5)
                portcode.Value = portfolio
                Conn.Open()
                DA.Fill(DSet, "maps")

                portCodeToUse = portfolio
                baseCurrency = "us"

                rows = DSet.Tables("maps").Select
                For Each row In rows
                    portCodeToUse = row("PortCodeToUse")
                    baseCurrency = Trim(row("BaseCurrency"))
                Next

                ' write a trade to appropriate blotter file based on
                ' portfolio's based currency
                Select Case baseCurrency
                    Case "ca"
                        ' replace port code with port code to use from the map table
                        line = portCodeToUse + line.Remove(0, 5)
                        fwCAD.WriteLine(line)
                        cadFlag = True
                    Case "eu"
                        line = portCodeToUse + line.Remove(0, 5)
                        fwEUR.WriteLine(line)
                        eurFlag = True
                    Case "ch"
                        line = portCodeToUse + line.Remove(0, 5)
                        fwCHF.WriteLine(line)
                        chfFlag = True
                    Case "nz"
                        line = portCodeToUse + line.Remove(0, 5)
                        fwNZD.WriteLine(line)
                        nzdFlag = True
                    Case "gb"
                        line = portCodeToUse + line.Remove(0, 5)
                        fwGBP.WriteLine(line)
                        gbpFlag = True
                    Case Else
                        fwUSD.WriteLine(line)
                        usdFlag = True
                End Select

                DSet.Clear()
                Conn.Close()
            Loop Until line Is Nothing
            sr.Close()

            fwUSD.Close()
            fwCAD.Close()
            fwEUR.Close()
            fwCHF.Close()
            fwNZD.Close()
            fwGBP.Close()
        Catch ex As Exception
            TextBox1.Text += vbCrLf + ex.Message
        End Try

        ' import into USD Axys blotter
        Dim ImexProc As New ProcessStartInfo(axysPath + "imex32.exe")
        Dim p As Process
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + blotterPath

        If usdFlag = True Then

            Try

                If FXConManager.validateTRNFile(fNameUSD, TextBox1) Then
                    File.Copy(fNameUSD, fName, True) 'Axys can import only file named topost.trn so each currency file should be moved into topost.trn

                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fNameUSD
                    TextBox1.Text += vbCrLf + "Windows User: " + idWindows.Name
                End If

            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

        ' import into CAD Axys blotter
        ImexProc.WorkingDirectory = "H:\Axys3\USERS\caduser"
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + cadBlotterpath

        If cadFlag = True Then
            Try

                If FXConManager.validateTRNFile(fNameCAD, TextBox1) Then
                    File.Copy(fNameCAD, fName, True)
                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fNameCAD
                End If

            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

        ' import into EUR Axys blotter
        ImexProc.WorkingDirectory = "H:\Axys3\USERS\euruser"
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + eurBlotterpath

        If eurFlag = True Then
            Try

                If FXConManager.validateTRNFile(fNameEUR, TextBox1) Then
                    File.Copy(fNameEUR, fName, True)

                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fNameEUR
                End If

            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

        ' import into CHF Axys blotter
        ImexProc.WorkingDirectory = "H:\Axys3\USERS\chfuser"
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + chfBlotterpath

        If chfFlag = True Then
            Try

                If FXConManager.validateTRNFile(fnameCHF, TextBox1) Then
                    File.Copy(fnameCHF, fName, True)
                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fnameCHF
                End If

            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

        ' import into NZD Axys blotter
        ImexProc.WorkingDirectory = "H:\Axys3\USERS\NZDUSER"
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + nzdBlotterpath

        If nzdFlag = True Then
            Try
                If FXConManager.validateTRNFile(fnameNZD, TextBox1) Then
                    File.Copy(fnameNZD, fName, True)
                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fnameNZD
                End If


            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

        ' import into GBP Axys blotter
        ImexProc.WorkingDirectory = "H:\Axys3\USERS\GBPUSER"
        ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + gbpBlotterpath

        If nzdFlag = True Then
            Try
                If FXConManager.validateTRNFile(fnameGBP, TextBox1) Then
                    File.Copy(fnameGBP, fName, True)
                    p = Process.Start(ImexProc)
                    While Not p.HasExited
                        ' wait for the process to finish
                        Application.DoEvents()
                    End While
                    TextBox1.Text += vbCrLf + "Finished import of " + fnameGBP
                End If

            Catch ex As Exception
                TextBox1.Text += vbCrLf + ex.Message
            End Try
        End If

    End Sub

    Private Sub btn_AxysFCTrades_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_AxysFCTrades.Click
        Dim myDir As String  ' current application directory
        Dim fName As String
        Dim saveType As String
        Dim macroName As String
        Dim groupName As String
        Dim asofdate As String
        Dim axysPath As String
        Dim fx As New FXConManager("H:\FXCON\FCTrades.txt", TextBox1, connStr, "USD")
        Dim allocDate As DateTime
        Dim allocStr As String

        ' Ask user to enter the trade date
        allocStr = InputBox("Please enter the allocation date", "Request", Today())

        If allocStr = "" Then
            Return
        End If
        allocDate = allocStr
        If Not IsDate(allocDate) Then
            MsgBox("Not a valid allocation date", MsgBoxStyle.OkOnly)
            Return
        End If

        TextBox1.Text = ""

        axysPath = "H:\Axys3\"
        myDir = CurDir() + "\"
        fName = "fxcon.htm"
        saveType = " -vh" 'Axys saves output as HTML file
        macroName = "fxconn.mac"
        groupName = "@fx"
        asofdate = allocDate.Date.ToString("MMddyy")  ' Axys date format
        Dim ImexProc As New ProcessStartInfo(axysPath + "rep32.exe")
        Dim p As Process

        ImexProc.Arguments = " -m" + macroName + " -p" + groupName + saveType + " -u -b" + asofdate + " -t" + myDir + fName

        Try

            '
            ' These are trades from US Axys
            '
            TextBox1.Text += vbCrLf + "Started export from Axys..."

            ' delete old output file if it's in the current dir
            File.Delete(myDir + fName)

            ' run macro and save output as HTML file
            p = Process.Start(ImexProc)
            While Not p.HasExited
                ' wait for the process to finish
                Application.DoEvents()
            End While
            TextBox1.Text += vbCrLf + "Finished export to " + fName
            If fx.createAxysFXTradesFile(myDir + fName, myDir + "FXTrades.txt") = -1 Then
                TextBox1.Text += vbCrLf + "Failed to create Axys FX Trades file"
            End If

            '
            ' These are trades from CAD Axys
            '
            fName = "fxconCAD.htm"
            Dim fxCAD As New FXConManager("H:\FXCON\FCTradesCAD.txt", TextBox1, connStr, "CAD")
            TextBox1.Text += vbCrLf + "Started export from CAD Axys..."
            Dim cadDir As String = "\\tweedy_files\advent\Axys3\users\caduser\"
            Dim cadProc As New ProcessStartInfo(axysPath + "rep32.exe")

            cadProc.WorkingDirectory = cadDir
            cadProc.Arguments = " -m" + macroName + " -p" + groupName + saveType + " -u -b" + asofdate + " -t" + myDir + fName
            ' delete old file with trades for CAD portfolios
            File.Delete(myDir + fName)
            ' run macro and save output as HTML file
            p = Process.Start(cadProc)
            While Not p.HasExited
                ' wait for the process to finish
                Application.DoEvents()
            End While
            TextBox1.Text += vbCrLf + "Finished export to " + fName
            If fxCAD.createAxysFXTradesFile(myDir + fName, myDir + "FXTradesCAD.txt") = -1 Then
                TextBox1.Text += vbCrLf + "Failed to create Axys FX Trades file"
            End If

            '
            ' These are trades from EUR Axys
            '
            fName = "fxconEUR.htm"
            Dim fxEUR As New FXConManager("H:\FXCON\FCTradesEUR.txt", TextBox1, connStr, "EUR")
            TextBox1.Text += vbCrLf + "Started export from EUR Axys..."
            Dim eurDir As String = "\\tweedy_files\advent\Axys3\users\euruser\"
            Dim eurProc As New ProcessStartInfo(axysPath + "rep32.exe")

            eurProc.WorkingDirectory = eurDir
            eurProc.Arguments = " -m" + macroName + " -p" + groupName + saveType + " -u -b" + asofdate + " -t" + myDir + fName
            ' delete old file with trades for CAD portfolios
            File.Delete(myDir + fName)
            ' run macro and save output as HTML file
            p = Process.Start(eurProc)
            While Not p.HasExited
                ' wait for the process to finish
                Application.DoEvents()
            End While
            TextBox1.Text += vbCrLf + "Finished export to " + fName
            If fxEUR.createAxysFXTradesFile(myDir + fName, myDir + "FXTradesEUR.txt") = -1 Then
                TextBox1.Text += vbCrLf + "Failed to create Axys FX Trades file"
            End If

            '
            ' These are trades from NZD Axys
            '
            fName = "fxconNZD.htm"
            Dim fxNZD As New FXConManager("H:\FXCON\FCTradesNZD.txt", TextBox1, connStr, "NZD")
            TextBox1.Text += vbCrLf + "Started export from NZD Axys..."
            Dim nzdDir As String = "\\tweedy_files\advent\Axys3\users\nzduser\"
            Dim nzdProc As New ProcessStartInfo(axysPath + "rep32.exe")

            nzdProc.WorkingDirectory = nzdDir
            nzdProc.Arguments = " -m" + macroName + " -p" + groupName + saveType + " -u -b" + asofdate + " -t" + myDir + fName
            ' delete old file with trades for NZD portfolios
            File.Delete(myDir + fName)
            ' run macro and save output as HTML file
            p = Process.Start(nzdProc)
            While Not p.HasExited
                ' wait for the process to finish
                Application.DoEvents()
            End While
            TextBox1.Text += vbCrLf + "Finished export to " + fName
            If fxNZD.createAxysFXTradesFile(myDir + fName, myDir + "FXTradesNZD.txt") = -1 Then
                TextBox1.Text += vbCrLf + "Failed to create Axys FX Trades file"
            End If
            '
            ' These are trades from CHF Axys
            '
            fName = "fxconCHF.htm"
            Dim fxCHF As New FXConManager("H:\FXCON\FCTradesCHF.txt", TextBox1, connStr, "CHF")
            TextBox1.Text += vbCrLf + "Started export from CHF Axys..."
            Dim chfDir As String = "\\tweedy_files\advent\Axys3\users\chfuser\"
            Dim chfProc As New ProcessStartInfo(axysPath + "rep32.exe")

            chfProc.WorkingDirectory = chfDir
            chfProc.Arguments = " -m" + macroName + " -p" + groupName + saveType + " -u -b" + asofdate + " -t" + myDir + fName
            ' delete old file with trades for CAD portfolios
            File.Delete(myDir + fName)
            ' run macro and save output as HTML file
            p = Process.Start(chfProc)
            While Not p.HasExited
                ' wait for the process to finish
                Application.DoEvents()
            End While
            TextBox1.Text += vbCrLf + "Finished export to " + fName
            If fxCHF.createAxysFXTradesFile(myDir + fName, myDir + "FXTradesCHF.txt") = -1 Then
                TextBox1.Text += vbCrLf + "Failed to create Axys FX Trades file"
            End If

        Catch ex As Exception
            TextBox1.Text += vbCrLf + ex.Message

        End Try


    End Sub


    Private Sub btnFundTrades_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFundTrades.Click
        Dim fName As String
        Dim allocDate As DateTime
        Dim allocStr As String
        Dim rtn As Integer

        TextBox1.Clear()

        ' Ask user to enter the trade date
        allocStr = InputBox("Please enter the allocation date", "Request", Today())
        If allocStr = "" Then
            Return
        End If

        allocDate = allocStr
        If Not IsDate(allocDate) Then
            MsgBox("Not a valid allocation date", MsgBoxStyle.OkOnly)
            Return
        End If

        TextBox1.Text += allocStr + vbCrLf
        Dim outFolder As String = ReadConfigSetting("FundTradingRecapFolder")
        TextBox1.Text += vbCrLf + "Output folder: " + outFolder

        fName = outFolder + "\FT" + allocDate.ToString("MMddyy") + ".xlsx"
        Dim fm As New FXConManager(fName, TextBox1, connStr, "")
        'rtn = fm.getFundTradingRecap(allocDate)

        TextBox1.Text += vbCrLf + "Generating funds recap without DTCC trades..."
        rtn = fm.getFundTradingRecapAllFunds(allocDate, False)
        If rtn <> -1 Then
            TextBox1.Text += vbCrLf + "Created file " + fName

        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName
        End If


        fName = outFolder + "\RecapsWithDTCC" + "\FTWithDTCC" + allocDate.ToString("MMddyy") + ".xlsx"
        fm = New FXConManager(fName, TextBox1, connStr, "")
        'rtn = fm.getFundTradingRecap(allocDate)
        TextBox1.Text += vbCrLf + "Generating funds recap with DTCC trades..."
        rtn = fm.getFundTradingRecapAllFunds(allocDate, True)
        If rtn <> -1 Then
            TextBox1.Text += vbNewLine & "Created file " + fName

        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName
        End If

        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.SelectionLength = 0
        TextBox1.ScrollToCaret()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text += " DB Con: " + connStr
    End Sub


    ' 
    ' Trading RECAPS
    '
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fName As String
        'Dim allocStr As String
        Dim rtn As Integer
        'Dim allocDate As DateTime? = PromptForDate("Please enter the allocation date (MM/DD/YYYY):", "Request")
        Dim allocDate As Date? = PromptForDate("Please select the allocation date :", "Request", initial:=Date.Today)

        If Not allocDate.HasValue Then
            ' user cancelled
            Exit Sub
        End If

        Dim selectedDate? As Date = allocDate.Value
        Dim selectedDateStr As String = selectedDate.Value.ToString("MMddyy")
        Dim recapFolder As String = ReadConfigSetting("RecapFolder")


        ' Trading Recap Domestic
        fName = recapFolder & "\TradingRecap_" & selectedDateStr & ".xlsx"
        Dim fm As New FXConManager(fName, TextBox1, connStr, "")
        rtn = fm.getPortfolioRecap(selectedDate)
        If rtn <> -1 Then
            TextBox1.Text = "Created file " + fName + vbCrLf
        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName + vbCrLf
        End If
        Application.DoEvents()
        'Trading Recap International
        fName = recapFolder + "\TradingRecapInternational_" + selectedDateStr + ".xlsx"
        Dim fm1 As New FXConManager(fName, TextBox1, connStr, "")
        rtn = fm1.getPortfolioRecapInternational(selectedDate)
        If rtn <> -1 Then
            TextBox1.Text += "Created file " + fName

        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName
        End If

        rtn = fm1.getPortfolioRecapInternationalDTCC(selectedDate)
        If rtn <> -1 Then
            'TextBox1.Text += "Created DTCC comparison file "

        Else
            TextBox1.Text += vbCrLf + "Failed to create DTCC comparison file "
        End If
    End Sub


    ' Reusable prompt that returns Nothing on cancel/invalid
    Private Function PromptForDate(prompt As String, title As String) As DateTime?
        Dim defaultText As String = DateTime.Today.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        Dim input As String = InputBox(prompt, title, defaultText)

        If String.IsNullOrWhiteSpace(input) Then
            ' Cancel / closed dialog
            Return Nothing
        End If

        Dim parsed As DateTime
        Dim formats() As String = {"M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd"} ' add more if you like

        If DateTime.TryParseExact(
            input.Trim(),
            formats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeLocal Or DateTimeStyles.AllowWhiteSpaces,
            parsed) Then

            Return parsed.Date   ' normalize to date-only
        End If

        MessageBox.Show("Not a valid allocation date. Please use MM/DD/YYYY.",
                    "Invalid date",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
        Return Nothing
    End Function




    Private Function PromptForDate(
    prompt As String,
    title As String,
    Optional initial As Date? = Nothing,
    Optional min As Date? = Nothing,
    Optional max As Date? = Nothing,
    Optional owner As IWin32Window = Nothing
) As Date?

        Using dlg As New Form()
            dlg.Text = If(String.IsNullOrWhiteSpace(title), "Select Date", title)
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.ShowInTaskbar = False
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.ClientSize = New Size(360, 140)

            Dim lbl As New Label() With {
            .AutoSize = False,
            .Text = If(String.IsNullOrWhiteSpace(prompt), "Please select a date:", prompt),
            .Bounds = New Rectangle(12, 12, 336, 40)
        }

            Dim dtp As New DateTimePicker() With {
            .Format = DateTimePickerFormat.Short,
            .Bounds = New Rectangle(12, 60, 150, 24)
        }
            If initial.HasValue Then dtp.Value = initial.Value
            If min.HasValue Then dtp.MinDate = min.Value
            If max.HasValue Then dtp.MaxDate = max.Value

            Dim btnOK As New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(192, 92, 75, 28)}
            Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(273, 92, 75, 28)}

            dlg.AcceptButton = btnOK
            dlg.CancelButton = btnCancel
            dlg.Controls.AddRange({lbl, dtp, btnOK, btnCancel})

            Dim result = If(owner Is Nothing, dlg.ShowDialog(), dlg.ShowDialog(owner))
            If result = DialogResult.OK Then
                Return dtp.Value.Date
            Else
                Return Nothing
            End If
        End Using
    End Function


    Private Sub ButtonHedgeExposure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHedgeExposure.Click
        Dim fName As String

        Dim allocDate As Date
        Dim rtn As Integer
        Dim allocStr As String
        ' Ask user to enter the date
        allocStr = InputBox("Please enter the hedge exposure date", "Request", Today().AddDays(-1))

        If allocStr = "" Then
            Return
        End If

        If Not IsDate(allocDate) Then
            MsgBox("Not a valid date", MsgBoxStyle.OkOnly)
            Return
        End If

        allocDate = allocStr

        fName = "H:\FXCon\HedgeExposure\HdgExposure_" + allocDate.ToString("MMddyy") + ".xlsx"
        Dim fm As New FXConManager(fName, TextBox1, connStr, "")
        rtn = fm.getHedgeExposure(allocDate)
        If rtn <> -1 Then
            TextBox1.Text += vbCrLf + "Created file " + fName
            TextBox1.Text += vbCrLf + "Use Moxy tb_HedgeExposue table to add/remove portfoios to this report." + fName

        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName
        End If

        'Set the cursor to the end of the textbox.
        TextBox1.SelectionStart = TextBox1.TextLength
        '
        'Scroll down to the cursor position.
        TextBox1.ScrollToCaret()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        ' starts help file containing Data source information 
        Help.ShowHelp(ParentForm, "FXConTradesHelp.chm")
    End Sub

    Private Sub ButtonMoxyExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonMoxyExport.Click
        Dim fName As String
        Dim allocDate As DateTime
        Dim allocStr As String
        Dim rtn As Integer

        ' Ask user to enter the trade date
        allocStr = InputBox("Please enter the allocation date", "Request", Today())

        If allocStr = "" Then
            ' user Cancel dialog
            Return
        Else
            allocDate = allocStr
        End If

        If Not IsDate(allocDate) Then
            MsgBox("Not a valid allocation date", MsgBoxStyle.OkOnly)
            Return
        End If

        ' Trading Recap Domestic
        fName = "C:\Temp\MoxyExport_" + allocDate.ToString("MMddyy") + ".tsv"
        Dim fm As New FXConManager(fName, TextBox1, connStr, "")
        rtn = fm.getMoxyExport(allocDate)
        If rtn <> -1 Then
            TextBox1.Text = "Created file " + fName + vbCrLf
        Else
            TextBox1.Text += vbCrLf + "Failed to create file " + fName + vbCrLf
        End If
        Application.DoEvents()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim fName As String
        Dim allocDate As DateTime
        Dim allocStr As String
        Dim rtn As Integer = 0
        ' Ask user to enter the trade date
        allocStr = InputBox("Please enter the allocation date", "Request", Today())
        If allocStr = "" Then
            Return

        End If
        allocDate = allocStr
        If Not IsDate(allocDate) Then
            MsgBox("Not a valid allocation date", MsgBoxStyle.OkOnly)
            Return
        End If

        fName = "P:\VBAPPS\PershingProtrak" + allocDate.ToString("MMddyy") + ".xls"

        If Not File.Exists(fName) Then
            TextBox1.Text += vbCrLf + String.Format("File {0} not found.", fName)
            Return
        End If

        Dim objOpt As Object = System.Reflection.Missing.Value
        'Dim rowNum As Integer
        Dim sec As String = ""

        Dim oXL As New Excel.Application
        Dim theWorkbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet

        Try
            theWorkbook = oXL.Workbooks.Open(fName)
            worksheet = theWorkbook.ActiveSheet

            worksheet.Columns(6).Insert()

            ' save generated Excel file
            worksheet.Columns.AutoFit()
            theWorkbook.SaveAs(fName, objOpt, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlShared, objOpt, objOpt, objOpt, objOpt, objOpt)
            theWorkbook.Close(False, objOpt, objOpt)

            oXL.Quit()


        Catch ex As Exception
            TextBox1.Text += vbCrLf + ex.Message
        End Try



        rtn = rtn

    End Sub

    '
    ' create GTSS file to send trades in MT300 format
    '
    Private Sub btnCreateGTSSFile_Click(sender As Object, e As EventArgs) Handles btnCreateGTSSFile.Click
        ' get trn file created from GTSS Excel spreadsheet by the Trading Desk

        Try
            Dim trnFile As String = ReadConfigSetting("GTSSTRNFile")
            Dim gtssFile As String = ReadConfigSetting("GTSSOUTFile")

            Dim i As Integer = gtssFile.IndexOf(".txt")
            If i = -1 Then
                MessageBox.Show("Invalid name of GTSS out file in .config file")
                Return
            Else
                gtssFile = gtssFile.Insert(i, Environment.UserName)
            End If

            FXConManager.validateTRNFile(trnFile, TextBox1)

            Dim fm As New FXConManager(gtssFile, TextBox1, connStr, "")
            Dim list As List(Of GTSSObj) = fm.readTRNFile(trnFile)
            TextBox1.AppendText(vbCrLf + "Number of trades in TRN file: " + list.Count.ToString)

            gtssFile = fm.createGTSSFile(list)
            TextBox1.AppendText(vbCrLf + "Created file : " + gtssFile)

        Catch ex As Exception
            TextBox1.Text += vbCrLf + ex.Message
        End Try

    End Sub

    '
    ' read properties of .config file
    '
    Function ReadConfigSetting(key As String) As String
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


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' offshore funds recap
        Dim frmDialog As New FormDateRange()
        Dim startDate As String
        Dim endDate As String


        ' Show testDialog as a modal dialog and determine if DialogResult = OK.
        If frmDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            ' Read the contents of testDialog's TextBox.
            'txtResult.Text = testDialog.TextBox1.Text
            startDate = frmDialog.DateTimePickerStartDate.Value.ToShortDateString
            endDate = frmDialog.DateTimePickerEndDate.Value.ToShortDateString

            If frmDialog.DateTimePickerStartDate.Value < frmDialog.DateTimePickerEndDate.Value Then
                'MessageBox.Show(startDate + " " + endDate)
                TextBox1.Text = "Selected dates: " + startDate + " " + endDate + vbCrLf
                Dim recapFolder As String = ReadConfigSetting("OffshoreRecapFolder")

                Dim fName As String = recapFolder + "OffshoreRecap_" + startDate.Replace("/", "-") + "_" + endDate.Replace("/", "-") + ".xlsx"
                Dim fm As New FXConManager(fName, TextBox1, connStr, "")
                Dim rtn As Int16 = fm.GetOffshoreFundsRecapsFromPortiaRef(startDate, endDate, portiaConStr, fName) ' Data source has been changed from Moxy to Portia
                If rtn <> -1 Then
                    TextBox1.Text = "Created file " + fName + vbCrLf
                Else
                    TextBox1.Text += vbCrLf + "Failed to create file " + fName + vbCrLf
                End If

            End If


        Else
            TextBox1.Text = "Cancelled"
        End If
        frmDialog.Dispose()

    End Sub

    Private Sub Button_FXConPreAlloc_AutoSizeChanged(sender As Object, e As EventArgs) Handles Button_FXConPreAlloc.AutoSizeChanged

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Re-center AFTER the form has finished sizing/layout
        Me.StartPosition = FormStartPosition.Manual

        Dim wa As Rectangle = Screen.FromControl(Me).WorkingArea  ' or Screen.FromPoint(Cursor.Position).WorkingArea
        Dim x As Integer = wa.Left + (wa.Width - Me.Width) \ 2
        Dim y As Integer = wa.Top + (wa.Height - Me.Height) \ 2

        Me.Location = New Point(x, y)
    End Sub
End Class

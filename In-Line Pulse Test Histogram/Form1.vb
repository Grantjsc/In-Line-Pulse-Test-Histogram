Imports System.Data.OleDb
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.UI.WebControls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports LFPHWIADBLib
Public Class Form1
    Private PICO_db As New WIADb("Data Source=BTMESSQLPROD;Initial Catalog=LFPHPICO;User ID=mesaccount;Password=superfuse;Persist Security Info=True;")
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Hioki As String
    Public HiokiRead As String
    Public Reading As String
    Public store_res(50000) As String
    Public runcount As Integer
    Public Sub labels()
        Label4.Text = "0"
        Label5.Text = "0"
        Label6.Text = "0"
        Label7.Text = "0"
        Label8.Text = "0"
        Label9.Text = "0"
        Label10.Text = "0"
        Label11.Text = "0"
        Label12.Text = "0"
        Label13.Text = "0"
        Label14.Text = "0"
        Label15.Text = "0"
        Label16.Text = "0"
        Label17.Text = "0"
        Label18.Text = "0"
        Label19.Text = "0"
        Label20.Text = "0"
        Label21.Text = "0"
        Label22.Text = "0"
        Label23.Text = "0"
        Label24.Text = "0"
        lblRange.Text = ""
    End Sub
    Public Sub labelResult()
        lblValue1.Text = "0"
        lblValue2.Text = "0"
        lblValue3.Text = "0"
        lblValue4.Text = "0"
        lblValue5.Text = "0"
        lblValue6.Text = "0"
        lblValue7.Text = "0"
        lblValue8.Text = "0"
        lblValue9.Text = "0"
        lblValue10.Text = "0"
        lblValue11.Text = "0"
        lblValue12.Text = "0"
        lblValue13.Text = "0"
        lblValue14.Text = "0"
        lblValue15.Text = "0"
        lblValue16.Text = "0"
        lblValue17.Text = "0"
        lblValue18.Text = "0"
        lblValue19.Text = "0"
        lblValue20.Text = "0"
        lblValue21.Text = "0"
        lblValue22.Text = "0"
        lblValue23.Text = "0"
    End Sub

    Public Sub ProgBar()
        Progress1.Value = 0
        Progress2.Value = 0
        Progress3.Value = 0
        Progress4.Value = 0
        Progress5.Value = 0
        Progress6.Value = 0
        Progress7.Value = 0
        Progress8.Value = 0
        Progress9.Value = 0
        Progress10.Value = 0
        Progress11.Value = 0
        Progress12.Value = 0
        Progress13.Value = 0
        Progress14.Value = 0
        Progress15.Value = 0
        Progress16.Value = 0
        Progress17.Value = 0
        Progress18.Value = 0
        Progress19.Value = 0
        Progress20.Value = 0
        Progress21.Value = 0
        Progress22.Value = 0
        Progress23.Value = 0

        'Maximum value
        Progress1.Maximum = 100
        Progress2.Maximum = 100
        Progress3.Maximum = 100
        Progress4.Maximum = 100
        Progress5.Maximum = 100
        Progress6.Maximum = 100
        Progress7.Maximum = 100
        Progress8.Maximum = 100
        Progress9.Maximum = 100
        Progress10.Maximum = 100
        Progress11.Maximum = 100
        Progress12.Maximum = 100
        Progress13.Maximum = 100
        Progress14.Maximum = 100
        Progress15.Maximum = 100
        Progress16.Maximum = 100
        Progress17.Maximum = 100
        Progress18.Maximum = 100
        Progress19.Maximum = 100
        Progress20.Maximum = 100
        Progress21.Maximum = 100
        Progress22.Maximum = 100
        Progress23.Maximum = 100
    End Sub

    Public Sub Prog_increase()
        Progress1.Maximum += 50
        Progress2.Maximum += 50
        Progress3.Maximum += 50
        Progress4.Maximum += 50
        Progress5.Maximum += 50
        Progress6.Maximum += 50
        Progress7.Maximum += 50
        Progress8.Maximum += 50
        Progress9.Maximum += 50
        Progress10.Maximum += 50
        Progress11.Maximum += 50
        Progress12.Maximum += 50
        Progress13.Maximum += 50
        Progress14.Maximum += 50
        Progress15.Maximum += 50
        Progress16.Maximum += 50
        Progress17.Maximum += 50
        Progress18.Maximum += 50
        Progress19.Maximum += 50
        Progress20.Maximum += 50
        Progress21.Maximum += 50
        Progress22.Maximum += 50
        Progress23.Maximum += 50
    End Sub

    Public Sub Prog_count()
        Reading = CDec(txtHiokiValue.Text)
        If Reading > CDec(Label4.Text) Then
            lblValue1.Text += 1
            Progress1.Value += 1
            If CDec(lblValue1.Text) >= CDec(Progress1.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label5.Text) Then
            lblValue2.Text += 1
            Progress2.Value += 1
            If CDec(lblValue2.Text) >= CDec(Progress2.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label6.Text) Then
            lblValue3.Text += 1
            Progress3.Value += 1
            If CDec(lblValue3.Text) >= CDec(Progress3.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label7.Text) Then
            lblValue4.Text += 1
            Progress4.Value += 1
            If CDec(lblValue4.Text) >= CDec(Progress4.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label8.Text) Then
            lblValue5.Text += 1
            Progress5.Value += 1
            If CDec(lblValue5.Text) >= CDec(Progress5.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label9.Text) Then
            lblValue6.Text += 1
            Progress6.Value += 1
            If CDec(lblValue6.Text) >= CDec(Progress6.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label10.Text) Then
            lblValue7.Text += 1
            Progress7.Value += 1
            If CDec(lblValue7.Text) >= CDec(Progress7.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label11.Text) Then
            lblValue8.Text += 1
            Progress8.Value += 1
            If CDec(lblValue8.Text) >= CDec(Progress8.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label12.Text) Then
            lblValue9.Text += 1
            Progress9.Value += 1
            If CDec(lblValue9.Text) >= CDec(Progress9.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Label13.Text) Then
            lblValue10.Text += 1
            Progress10.Value += 1
            If CDec(lblValue10.Text) >= CDec(Progress10.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > (CDec(Label14.Text) + splitval) Then
            lblValue11.Text += 1
            Progress11.Value += 1
            If CDec(lblValue11.Text) >= CDec(Progress11.Maximum) Then
                Prog_increase()
            End If

            '**************************************NOMINAL**************************************
            'ElseIf Reading < CDec(Label13.Text) Then
            '    lblValue13.Text += 1
            '    Progress13.Value += 1
            '    If CDec(lblValue13.Text) >= CDec(Progress13.Maximum) Then
            '        Prog_increase()
            '    End If
        ElseIf Reading >= (CDec(Label14.Text) - splitval) And Reading <= (CDec(Label14.Text) + splitval) Then
            lblValue12.Text += 1
            Progress12.Value += 1
            If CDec(lblValue12.Text) >= CDec(Progress12.Maximum) Then
                Prog_increase()
            End If
            '**************************************NOMINAL**************************************\

        ElseIf Reading < CDec(Label24.Text) Then
            lblValue23.Text += 1
            Progress23.Value += 1
            If CDec(lblValue23.Text) >= CDec(Progress23.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label23.Text) Then
            lblValue22.Text += 1
            Progress22.Value += 1
            If CDec(lblValue22.Text) >= CDec(Progress22.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label22.Text) Then
            lblValue21.Text += 1
            Progress21.Value += 1
            If CDec(lblValue21.Text) >= CDec(Progress21.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label21.Text) Then
            lblValue20.Text += 1
            Progress20.Value += 1
            If CDec(lblValue20.Text) >= CDec(Progress20.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label20.Text) Then
            lblValue19.Text += 1
            Progress19.Value += 1
            If CDec(lblValue19.Text) >= CDec(Progress19.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label19.Text) Then
            lblValue18.Text += 1
            Progress18.Value += 1
            If CDec(lblValue18.Text) >= CDec(Progress18.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label18.Text) Then
            lblValue17.Text += 1
            Progress17.Value += 1
            If CDec(lblValue17.Text) >= CDec(Progress17.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label17.Text) Then
            lblValue16.Text += 1
            Progress16.Value += 1
            If CDec(lblValue16.Text) >= CDec(Progress16.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label16.Text) Then
            lblValue15.Text += 1
            Progress15.Value += 1
            If CDec(lblValue15.Text) >= CDec(Progress15.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Label15.Text) Then
            lblValue14.Text += 1
            Progress14.Value += 1
            If CDec(lblValue14.Text) >= CDec(Progress14.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < (CDec(Label14.Text) - splitval) Then
            lblValue13.Text += 1
            Progress13.Value += 1
            If CDec(lblValue13.Text) >= CDec(Progress13.Maximum) Then
                Prog_increase()
            End If
        End If
    End Sub

    Public Sub test_result()
        Reading = CDec(txtHiokiValue.Text)
        If Reading <= nominal * 2 Then
            store_res(CDec(txtTested.Text)) = CDec(txtHiokiValue.Text)
            'runcount += 1
        End If
        runcount += 1

        If Reading <= CDec(txtHiLimit.Text) And Reading >= CDec(txtLoLimit.Text) Then
            txtResult.Text = "Good"
            txtResult.FillColor = Color.Lime
            txtHiokiValue.FillColor = Color.Lime
            txtGood.Text += 1
        End If
        If Reading > CDec(txtHiLimit.Text) Then
            If Reading > nominal * 2 Then
                txtResult.Text = "Open"
                txtResult.FillColor = Color.Red
                txtHiokiValue.FillColor = Color.Red
                txtOpen.Text += 1
            Else
                txtResult.Text = "High"
                txtResult.FillColor = Color.Red
                txtHiokiValue.FillColor = Color.Red
                txtHigh.Text += 1
            End If
        End If

        If Reading < CDec(txtLoLimit.Text) Then
            txtResult.Text = "Low"
            txtResult.FillColor = Color.Red
            txtHiokiValue.FillColor = Color.Red
            txtLow.Text += 1
        End If
        txtTested.Text += 1
        txtYield.Text = Math.Round((CDec(txtGood.Text) / CDec(txtTested.Text)) * 100, 2) & " %"
        txtRej.Text = CDec(txtHigh.Text) + CDec(txtOpen.Text) + CDec(txtLow.Text)
    End Sub

    Public maximum As Decimal
    Const limit200ohm = 18000
    Const limit20ohm = 1800
    Const limit2ohm = 180
    Const limit200mohm = 18
    Const limit20mohm = 1.8
    Const limit2mohm = 0.18

    Public Sub range()
        If maximum > limit20ohm Then
            If PICO_db.Rating >= 0.1 Then
                lblRange.Text = "R8 = 20Ω range @ 10mA Test Current"
            Else
                lblRange.Text = "R9 = 20Ω range @ 1mA Test Current"
            End If
            multiplier = 1000
        ElseIf maximum > limit2ohm Then
            If PICO_db.Rating >= 1 Then
                lblRange.Text = "R6 = 2Ω range @ 100mA Test Current"
                multiplier = 1000
            ElseIf PICO_db.Rating >= 0.1 Then
                lblRange.Text = "R7 = 2Ω range @ 10mA Test Current"
                multiplier = 1000
            Else
                lblRange.Text = "R9 = 20Ω range @ 1mA Test Current"
                multiplier = 1000
            End If
        ElseIf maximum > limit200mohm Then      ' Go to 200 mohm range
            If PICO_db.Rating >= 1 Then
                lblRange.Text = "R5 = 200mΩ range @ 100mA Test Current"
                multiplier = 1
            ElseIf PICO_db.Rating >= 10 Then
                lblRange.Text = "R4 = 200mΩ range @ 1mA Test Current"
                multiplier = 1
            Else
                lblRange.Text = "R7 = 2Ω range @ 10mA Test Current"
                multiplier = 1000
            End If
        ElseIf maximum > limit20mohm Then     ' Go to 20 mohm range
            If PICO_db.Rating >= 10 Then
                lblRange.Text = "R2 = 20mΩ range @ 1A Test Current"
                multiplier = 1
            ElseIf PICO_db.Rating >= 1 Then
                lblRange.Text = "R3 = 20mΩ range @ 100mA Test Current"
                multiplier = 1
            Else
                lblRange.Text = "R8 = 2Ω range @ 10mA Test Current"
                multiplier = 1000
            End If
        Else                          ' Go to 2 mohm range
            If PICO_db.Rating >= 10 Then
                lblRange.Text = "R1 = 2Ω range @ 1A Test Current"
                multiplier = 1
            Else
                lblRange.Text = "R3 = 20mΩ range @ 100mA Test Current"
                multiplier = 1
            End If
        End If
    End Sub
    Private Sub SerialPort1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        HiokiRead = SerialPort1.ReadLine
        Hioki = HiokiRead.Remove(HiokiRead.Length - 4)
        txtHiokiValue.Text = Hioki

        'Reading = Hioki
        'txtHiokiValue.Text = Reading
        test_result() 'GOOD/HIGH/LOW or OPEN 
        Prog_count() 'histogram counter
        statistics() 'comp od CPK and Standard Dev

    End Sub

    Private Sub SerialPort2_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort2.DataReceived

    End Sub

    Public port1 As String
    Public port2 As String
    Public Sub Get_portname()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\In-line Pulse Test_db.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfpulsetest")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From SerialPort_tb WHERE sname Like 'Serial1'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then
                port1 = Data.Rows(0).Item("portname").ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()

        End Try

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From SerialPort_tb WHERE sname Like 'Serial2'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then
                port2 = Data.Rows(0).Item("portname").ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()

        End Try
    End Sub

    Public Sub SerialName()
        SerialPort1.PortName = port1
        SerialPort2.PortName = port2
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        btnStart.Enabled = False
        btnChange.Enabled = False
        btnClear.Enabled = False
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        SerialPort1.NewLine = vbCrLf
        Get_portname()
        SerialName()
        SerialPort1.Open()
        SerialPort2.Open()

        labels()
        SendMessage(Progress1.Handle, 1040, 2, 0)
        SendMessage(Progress2.Handle, 1040, 3, 0)
        SendMessage(Progress12.Handle, 1040, 3, 0)
        SendMessage(Progress22.Handle, 1040, 3, 0)
        SendMessage(Progress23.Handle, 1040, 2, 0)
    End Sub

    Public nominal As Decimal
    Public minimum As Decimal
    Public HiLimit As Decimal
    Public LoLimit As Decimal
    Public getLoLim As Decimal
    Public getLolim2 As Decimal
    Public multiplier As Decimal '= 1000
    Public add As Decimal
    Public reduce As Decimal
    Public splitval As Decimal
    Public truemax As Decimal
    Public truemin As Decimal
    Public Eq1 As Decimal
    Private Sub txtPartNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPartNo.KeyUp
        Try
            If e.KeyCode = Keys.Enter Then
                txtLotNumber.Focus()

                If PICO_db.IsValidPartNumber(txtPartNo.Text) Then

                    Get_db() ' set Pcurrent and Pduration
                    Hioki_Basis() 'Basis of Hiokis Upper and Lower Limit
                    Hioki_set() 'Set Hiokis UCL and LCL

                    truemax = Math.Round(CDec(PICO_db.Nominal) * ((CDec(PICO_db.Maximum) / 100) + 1), 3)

                    'Eq1 = (CDec(PICO_db.Nominal) * (CDec(PICO_db.Minimum) / 100))
                    'truemin = Math.Round(CDec(PICO_db.Nominal) - Eq1, 3)

                    maximum = truemax
                    'minimum = truemin
                    range() 'mutiplier set up

                    Dim nomEQ1 As String
                    Dim nomEQ2 As String
                    nomEQ1 = UpperLimit + LowerLimit
                    nomEQ2 = nomEQ1 / 2
                    nominal = (CDec(nomEQ2)) 'PICO_db.Nominal '/ multiplier
                    HiLimit = nominal * ((CDec(PICO_db.Maximum) / 100) + 1) 'High Limit value
                    maximum = Math.Round(HiLimit, 3)

                    getLoLim = 100 - PICO_db.Minimum 'get the low limit less 100
                    getLolim2 = getLoLim / 100 'convert to %
                    LoLimit = nominal * getLolim2 ' Low Limit Value
                    minimum = Math.Round(LoLimit, 3)

                    txtRating.Text = PICO_db.Rating
                    txtPositive.Text = PICO_db.Maximum
                    txtNegative.Text = PICO_db.Minimum
                    txtNominal.Text = nominal
                    'txtHiLimit.Text = Math.Round(HiLimit, 3)
                    'txtLoLimit.Text = Math.Round(LoLimit, 3)

                    txtHiLimit.Text = UpperLimit
                    txtLoLimit.Text = LowerLimit

                    'Get_db() ' set Pcurrent and Pduration
                    'Hioki_Basis() 'Basis of Hiokis Upper and Lower Limit
                    'Hioki_set() 'Set Hiokis UCL and LCL

                    set_val() 'ser the progress bar parameter
                    ProgBar() 'progress bar value and maximum value

                    txtPartNo.ReadOnly = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    Private Sub txtLotNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles txtLotNumber.KeyUp

        If e.KeyCode = Keys.Enter Then
            If txtLotNumber.Text = "" Then
                MessageBox.Show("Please enter the correct Lot number!", "In-Line Pulse Test", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                btnStart.Focus()
                btnStart.Enabled = True
                btnClear2.Enabled = False
                btnChange.Enabled = True
                btnClear.Enabled = True
                txtLotNumber.ReadOnly = True
            End If
        End If
    End Sub

    Public Sub set_val()
        'maximum = (PICO_db.Maximum / multiplier)
        'minimum = (PICO_db.Minimum / multiplier)
        add = (CDec(txtHiLimit.Text) - nominal) / 9
        reduce = (nominal - CDec(txtLoLimit.Text)) / 9

        Label4.Text = Math.Round(nominal + (add * 10), 3)
        Label5.Text = Math.Round(nominal + (add * 9), 3)
        Label6.Text = Math.Round(nominal + (add * 8), 3)
        Label7.Text = Math.Round(nominal + (add * 7), 3)
        Label8.Text = Math.Round(nominal + (add * 6), 3)
        Label9.Text = Math.Round(nominal + (add * 5), 3)
        Label10.Text = Math.Round(nominal + (add * 4), 3)
        Label11.Text = Math.Round(nominal + (add * 3), 3)
        Label12.Text = Math.Round(nominal + (add * 2), 3)
        Label13.Text = Math.Round(nominal + (add * 1), 3)
        Label14.Text = nominal
        Label15.Text = Math.Round(nominal - (reduce * 1), 3)
        Label16.Text = Math.Round(nominal - (reduce * 2), 3)
        Label17.Text = Math.Round(nominal - (reduce * 3), 3)
        Label18.Text = Math.Round(nominal - (reduce * 4), 3)
        Label19.Text = Math.Round(nominal - (reduce * 5), 3)
        Label20.Text = Math.Round(nominal - (reduce * 6), 3)
        Label21.Text = Math.Round(nominal - (reduce * 7), 3)
        Label22.Text = Math.Round(nominal - (reduce * 8), 3)
        Label23.Text = Math.Round(nominal - (reduce * 9), 3)
        Label24.Text = Math.Round(nominal - (reduce * 10), 3)

        'Label4.Text = Math.Round(nominal + (add * 9), 3)
        'Label5.Text = Math.Round(nominal + (add * 8), 3)
        'Label6.Text = Math.Round(nominal + (add * 7), 3)
        'Label7.Text = Math.Round(nominal + (add * 6), 3)
        'Label8.Text = Math.Round(nominal + (add * 5), 3)
        'Label9.Text = Math.Round(nominal + (add * 4), 3)
        'Label10.Text = Math.Round(nominal + (add * 3), 3)
        'Label11.Text = Math.Round(nominal + (add * 2), 3)
        'Label12.Text = Math.Round(nominal + (add * 1), 3)
        'Label13.Text = Math.Round(nominal + (add * 0.5), 3)
        'Label14.Text = nominal
        'Label15.Text = Math.Round(nominal - (reduce * 0.5), 3)
        'Label16.Text = Math.Round(nominal - (reduce * 1), 3)
        'Label17.Text = Math.Round(nominal - (reduce * 2), 3)
        'Label18.Text = Math.Round(nominal - (reduce * 3), 3)
        'Label19.Text = Math.Round(nominal - (reduce * 4), 3)
        'Label20.Text = Math.Round(nominal - (reduce * 5), 3)
        'Label21.Text = Math.Round(nominal - (reduce * 6), 3)
        'Label22.Text = Math.Round(nominal - (reduce * 7), 3)
        'Label23.Text = Math.Round(nominal - (reduce * 8), 3)
        'Label24.Text = Math.Round(nominal - (reduce * 9), 3)
        splitval = Math.Round((CDec(Label13.Text) - CDec(Label15.Text)) / 3, 3)
        Console.WriteLine(splitval)
    End Sub

    Public Sub change_device()
        btnStart.Enabled = False

        labels()
        labelResult()
        ProgBar()
        txtHiokiValue.Text = ""
        txtResult.Text = ""
        txtResult.FillColor = Color.Black
        txtHiokiValue.FillColor = Color.Black

        txtPulseC.Text = ""
        txtPulseD.Text = ""
        txtPartNo.Text = ""
        txtLotNumber.Text = ""
        txtRating.Text = ""
        txtPositive.Text = ""
        txtNegative.Text = ""
        txtNominal.Text = ""
        txtHiLimit.Text = ""
        txtLoLimit.Text = ""
        txtTested.Text = "0"
        txtGood.Text = "0"
        txtHigh.Text = "0"
        txtOpen.Text = "0"
        txtLow.Text = "0"
        txtRej.Text = "0"
        txtStdev.Text = "0"
        txtCpk.Text = "0"
        txtYield.Text = ""

        txtLower.Text = ""
        txtUpper.Text = ""
        txtUnit.Text = ""

        txtPartNo.ReadOnly = False
        txtLotNumber.ReadOnly = False

        runcount = 0
        Array.Clear(store_res, 0, store_res.Length)
    End Sub

    Public Sub clear()

        ProgBar()
        labelResult()
        txtHiokiValue.Text = ""
        txtResult.Text = ""
        txtResult.FillColor = Color.Black
        txtHiokiValue.FillColor = Color.Black

        txtLotNumber.Text = ""

        txtTested.Text = "0"
        txtGood.Text = "0"
        txtHigh.Text = "0"
        txtOpen.Text = "0"
        txtLow.Text = "0"
        txtRej.Text = "0"
        txtStdev.Text = "0"
        txtCpk.Text = "0"
        txtYield.Text = ""

        txtLotNumber.ReadOnly = False

        runcount = 0
        Array.Clear(store_res, 0, store_res.Length)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you really want to exit?", "Exit In-Line Pulse Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.No Then
            e.Cancel = True
        Else
            'Application.ExitThread()
            End
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Try
        '    SerialPort1.WriteLine("*TRG")
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try

        test_result() 'GOOD/HIGH/LOW or OPEN 
        Prog_count() 'histogram counter
        statistics()

        'Savedata()
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If btnStart.Text = "Start" Then
            Timer1.Enabled = True
            btnStart.Text = "Stop"
            btnStart.FillColor = Color.Maroon
            btnStart.FillColor2 = Color.Red
            btnStart.ForeColor = Color.White
        Else
            Timer1.Enabled = False
            btnStart.Text = "Start"
            btnStart.FillColor = Color.LightGreen
            btnStart.FillColor2 = Color.Green
            btnStart.ForeColor = Color.Black
        End If
    End Sub

    Private Sub btnChange_Click(sender As Object, e As EventArgs) Handles btnChange.Click

        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you want continue?", "Change device and clear all data", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.Yes Then
            Savedata() 'Save data to csv file
            SaveRes()
            change_device()

            Timer1.Enabled = False
            btnClear2.Enabled = True
            btnChange.Enabled = False
            btnClear.Enabled = False
            btnStart.Text = "Start"
            btnStart.FillColor = Color.LightGreen
            btnStart.FillColor2 = Color.Green
            btnStart.ForeColor = Color.Black
        Else
            Return
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click


        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you want continue?", "Change Lot Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.Yes Then
            Savedata() 'Save data to csv file
            SaveRes()
            clear()

            Timer1.Enabled = False
            btnStart.Text = "Start"
            btnStart.FillColor = Color.LightGreen
            btnStart.FillColor2 = Color.Green
            btnStart.ForeColor = Color.Black
        Else
            Return
        End If

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you really want to exit?", "Exit In-Line Pulse Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.Yes Then
            'Application.ExitThread()
            End
        Else
            Return
        End If
    End Sub

    Public Sub Get_db()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\In-line Pulse Test_db.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfpulsetest")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From PulseTest_tb WHERE Part_number Like '%" & txtPartNo.Text & "%'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                txtPulseC.Text = Data.Rows(0).Item("_PulseC").ToString
                txtPulseD.Text = Data.Rows(0).Item("_PulseD").ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()

        End Try
    End Sub

    Public LowerLimit As Decimal
    Public UpperLimit As Decimal
    Public Sub Hioki_Basis()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\In-line Pulse Test_db.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfpulsetest")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From PulseTest_tb WHERE Part_number Like '%" & txtPartNo.Text & "%'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                txtLower.Text = Data.Rows(0).Item("_Hioki-LCL").ToString
                txtUpper.Text = Data.Rows(0).Item("_Hioki-UCL").ToString
                txtUnit.Text = Data.Rows(0).Item("_unit").ToString
                txtRange.Text = Data.Rows(0).Item("_range").ToString

                LowerLimit = txtLower.Text
                UpperLimit = txtUpper.Text
                Console.WriteLine(LowerLimit)
                Console.WriteLine(UpperLimit)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Unit As String
    Public Sub Hioki_set()

        If txtUnit.Text = "Ohms" Then
            Unit = "E+0"
            SerialPort1.WriteLine("RES:RANG " & txtRange.Text)
        ElseIf txtUnit.Text = "mOhms" Then
            Unit = "E-03"
            SerialPort1.WriteLine("RES:RANG " & txtRange.Text)
        End If

        Console.WriteLine(Unit)
        SerialPort1.WriteLine(":CALC:LIM:LOW " & txtLower.Text & Unit)
        SerialPort1.WriteLine(":CALC:LIM:UPP " & txtUpper.Text & Unit)
    End Sub

    Public holding As Boolean
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If SerialPort2.CDHolding = True Then
            If holding = False Then
                holding = True
                Try
                    SerialPort1.WriteLine("*TRG")
                Catch ex As Exception
                    MsgBox(ex.Message, vbCritical)
                End Try
            End If
        Else
            holding = False
        End If
    End Sub

    Public average As Decimal
    Public stdev As Decimal
    Public cpk_max As Decimal
    Public cpk_min As Decimal
    Public Sub statistics()
        average = 0
        For s As Integer = 0 To store_res.Length - 1
            If store_res(s) = "" Then
                If average > 0 Then
                    average = Math.Round(average / s, 2)
                    stdev = 0
                    For t As Integer = 0 To store_res.Length - 1
                        If store_res(t) = "" Then
                            If stdev > 0 Then
                                stdev = Math.Sqrt(CDec(stdev) / CDec(t))
                                cpk_max = (maximum - average) / (stdev * 3)
                                cpk_min = (average - minimum) / (stdev * 3)
                                txtCpk.Text = CStr(Math.Round(Math.Min(cpk_max, cpk_min), 2))
                                txtStdev.Text = CStr(Math.Round(stdev, 2))
                            End If
                            Exit For
                        Else
                            stdev = ((CDec(store_res(t)) - average) ^ 2) + stdev
                        End If
                    Next
                End If
                Exit For
            Else
                average = average + store_res(s)
            End If
        Next
    End Sub

    Public save As String
    Public get_FolderPath As String
    Public get_message As String

    Public get_FolderPath2 As String
    Public get_message2 As String
    Public DD As String

    Public Sub Savedata()
        save = txtLotNumber.Text
        DD = Format(Now, "MM-dd-hh:mm")
        'C:\DTR Line 1 In-line Pulse Test Database
        get_FolderPath = "C:\DTR Line 1 In-line Pulse Test Database\" & save & ".csv"

        get_message = """Yield,""" & "," & """Lot Number,""" & "," & """Total Good,""" & "," & """High Resistance,""" & "," & """Open Fuse,""" & "," & """Low Resistance,""" & "," & """CPK,""" & vbCrLf
        get_message = get_message & txtYield.Text & "," & txtLotNumber.Text & "," & txtGood.Text & "," & txtHigh.Text & "," & txtOpen.Text & "," & txtLow.Text & "," & txtCpk.Text & "," & vbCrLf

        My.Computer.FileSystem.WriteAllText(get_FolderPath, get_message, False)
        MessageBox.Show("The data is saved in " & get_FolderPath)


        ''---------------------------- Get Resistance Test ----------------------------
        'get_FolderPath = "C:\DTR Line 1 In-line Pulse Test Database\" & save & "_Resistance.csv"

        'get_message = """Lot Number,""" & vbCrLf
        'get_message = get_message & store_res(s) & vbCrLf

        'My.Computer.FileSystem.WriteAllText(get_FolderPath, get_message, False)
        'MessageBox.Show("The data is saved in " & get_FolderPath)

    End Sub

    Public Sub SaveRes()
        Dim CSVfile As String = String.Empty
        Dim Title As String
        'CSVfile = CSVfile & txtLotNumber.Text & vbCrLf
        CSVfile = CSVfile & "Fuse Resistance," & vbCrLf
        For row As Integer = 0 To store_res.Length - 1
            If store_res(row) > 0 Then
                CSVfile = CSVfile & store_res(row)
                CSVfile = CSVfile & vbCr & vbLf
            End If
        Next

        Title = "C:\DTR Line 1 In-line Pulse Test Database\" & txtLotNumber.Text & "_Resistance.csv"
        My.Computer.FileSystem.WriteAllText(Title, CSVfile, False)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Savedata()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        change_device()
    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles btnClear2.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you really want to clear all data?", "CLear All Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.Yes Then
            'Application.ExitThread()
            change_device()
        Else
            Return
        End If
    End Sub
End Class

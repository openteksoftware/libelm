Imports System.Threading

Public Class frmMain
    Private bConnected As Boolean
    Private lSensorPage As Integer
    Private WithEvents cELM As libelm.libelm.Core
    Private cLog As New libelm.cLogger
    Private SensorThread As Thread

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If bConnected Then
            ToolStripComboBox1.Enabled = True
            ToolStripComboBox2.Enabled = True

            bConnected = False
            ToolStripButton1.Text = "Connect"

            cELM.Disconnect()
        Else

            ToolStripComboBox1.Enabled = False
            ToolStripComboBox2.Enabled = False

            bConnected = True
            ToolStripButton1.Text = "Disconnect"

            cELM.COMPort = ToolStripComboBox1.Text
            cELM.BaudRate = ToolStripComboBox2.Text
            cELM.Connect()

            If cELM.isConnected() Then
                If Not SensorThread.IsAlive = False Then
                    'SensorThread.Start()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox1.Select(TextBox1.Text.Length, 0)
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            ' Send Data
            If TextBox2.Text.Length() = 0 Then
                e.Handled = True
                Return
            End If

            sendTerminalData(TextBox2.Text)
            TextBox2.Text = ""

            e.Handled = True
        End If
    End Sub

    Public Sub sendTerminalData(ByVal sText As String)
        sText = sText.Replace(vbCr, "")
        sText = sText.Replace(vbLf, "")
        cELM.sendRAWData(sText)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sendTerminalData(TextBox2.Text)
        TextBox2.Text = ""
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Length = 0 Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click

    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        Select Case TabControl1.SelectedIndex
            Case 0
                lSensorPage = 0
            Case 1 ' Digital Guages
                If Not SensorThread.IsAlive() Then
                    Try
                        SensorThread = New Thread(AddressOf SensorUpdateThread)
                        'SensorThread.IsBackground = True
                        SensorThread.Start()
                    Catch ex As Exception

                    End Try
                End If
                Select Case TabControl4.SelectedIndex
                    Case 0 ' Page 1
                        lSensorPage = 1
                    Case 1 ' Page 2
                        lSensorPage = 2
                    Case 2 ' Oxygen Sensors
                        lSensorPage = 3
                End Select
            Case 2
                lSensorPage = 0
            Case 3
                Select Case TabControl2.SelectedIndex
                    Case 0 ' General
                        lSensorPage = 4
                    Case 1 ' Error 
                        lSensorPage = 5
                    Case 2 ' Raw Communication
                        lSensorPage = 6
                End Select
        End Select
    End Sub

    Private Sub TabPage2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Not cELM.isConnected() Then
            Button1.Enabled = False
            TextBox2.Enabled = False
        Else
            Button1.Enabled = True
            TextBox2.Enabled = True
        End If

        Me.Text = "" & lSensorPage & " - " & SensorThread.IsAlive

        Select Case lSensorPage
            Case 0

            Case 1
            Case 2 ' Page 2
            Case 3 ' Oxygen Sensor
            Case 4 ' General Log Data
                If Not cELM.isConnected Then
                    Return
                End If

                If TextBox3.Text.Length() < cLog.getLogdata().Length() Then
                    TextBox3.Text = cLog.getLogdata()
                End If
            Case 5 ' Error Log Data
            Case 6 ' Raw Communication Log Data

        End Select
    End Sub

    Private Sub TabControl4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl4.Click
        Dim lOldPage As Integer = lSensorPage
        Select Case TabControl4.SelectedIndex
            Case 0 ' Page 1
                lSensorPage = 1
            Case 1 ' Page 2
                lSensorPage = 2
            Case 2 ' Oxygen Sensors
                lSensorPage = 3
        End Select

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        cELM = New libelm.libelm.Core
        cELM.Logger = cLog

        SensorThread = New Thread(AddressOf SensorUpdateThread)
        'SensorThread.IsBackground = True
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        TextBox3.Select(TextBox3.Text.Length(), 0)
        TextBox3.ScrollToCaret()
    End Sub

    Private Sub ChangeText(ByVal LabelObject As Label, ByVal Text As String)
        If LabelObject.InvokeRequired Then
            LabelObject.Invoke(Sub() LabelObject.Text = Text)
        End If
    End Sub

    Private Sub AppendTextBox(ByVal LabelObject As TextBox, ByVal Text As String)
        If LabelObject.InvokeRequired Then
            LabelObject.Invoke(Sub() LabelObject.Text = LabelObject.Text & Text & vbNewLine)
        End If
    End Sub

    Private Sub ChangeProgress(ByVal ProgressObject As ProgressBar, ByVal Progress As Integer)
        If ProgressObject.InvokeRequired Then
            ProgressObject.Invoke(Sub() ProgressObject.Value = Progress)
        End If
    End Sub

    Private Sub SensorUpdateThread()
        Dim dTemp As Double


        Select lSensorPage
            Case 1

                ' Vehicle MPH
                dTemp = cELM.getVehicleSpeed()
                ChangeProgress(ProgressBar1, Int(dTemp))
                ChangeText(Label1, FormatNumber(dTemp, 2) & "mph")

                Thread.Sleep(10)
                ' Engine Load
                dTemp = cELM.getEngineLoad()
                ChangeProgress(ProgressBar2, Int(dTemp))
                ChangeText(Label2, FormatNumber(dTemp, 0) & "%")

                ' Throttle Position
                'ProgressBar3.Value = Int(dTemp)
                'Label3.Text = FormatNumber(dTemp, 0) & "%"

                Thread.Sleep(10)
                ' Coolant Temperature
                dTemp = cELM.getCoolantTemperature()
                ChangeProgress(ProgressBar4, Int(dTemp))
                ChangeText(Label4, FormatNumber(dTemp, 2) & "F")

                ' MAF Flow Rate

                'ProgressBar5.Value = Int(dTemp)
                'Label5.Text = FormatNumber(dTemp, 2) & "grams/sec"

                ' Timing Advancment
                'ProgressBar6.Value = Int(dTemp)
                'Label6.Text = FormatNumber(dTemp, 2) & ""

                Thread.Sleep(10)
                ' Fuel Pressure
                dTemp = cELM.getFuelPressure()
                ChangeProgress(ProgressBar7, Int(dTemp))
                ChangeText(Label7, FormatNumber(dTemp, 2) & "kPa")

                ' Intake AIR Temperature
                'ProgressBar8.Value = Int(dTemp)
                'Label8.Text = FormatNumber(dTemp, 2) & "F"


                Thread.Sleep(10)
                ' Intake Manifold Pressure
                dTemp = cELM.getIntakeManifoldPressure()
                ChangeProgress(ProgressBar9, Int(dTemp))
                ChangeText(Label9, FormatNumber(dTemp, 2) & "kPa")

                ' Commanded EGR
                'ProgressBar10.Value = Int(dTemp)
                'Label10.Text = FormatNumber(dTemp, 2) & "%"

                ' EGR Error
                'ProgressBar11.Value = Int(dTemp)
                'Label11.Text = FormatNumber(dTemp, 2) & "%"

                ' Commanded EVAP Purge
                'ProgressBar12.Value = Int(dTemp)
                'Label12.Text = FormatNumber(dTemp, 2) & "%"

                ' Fuel Tank Level
                'ProgressBar13.Value = Int(dTemp)
                'Label13.Text = FormatNumber(dTemp, 2) & "%"

                ' EVAP System Pressure
                'ProgressBar14.Value = Int(dTemp)
                'Label14.Text = FormatNumber(dTemp, 2) & "Pa"

                ' Barometric Pressure
                'ProgressBar15.Value = Int(dTemp)
                'Label15.Text = FormatNumber(dTemp, 2) & "kPa"

                Thread.Sleep(10)
            Case 2

                ' Battery Voltage
                dTemp = cELM.getBatteryVoltage()
                ChangeProgress(ProgressBar16, Int(dTemp))
                ChangeText(Label17, FormatNumber(dTemp, 2) & "V")

                ' Control Module Voltage
                'ProgressBar17.Value = Int(dTemp)
                'Label17.Text = FormatNumber(dTemp, 2) & "V"

                ' Accelerator Pedal Position D
                'ProgressBar18.Value = Int(dTemp)
                'Label18.Text = FormatNumber(dTemp, 2) & "%"

                ' Accelerator Pedal Position E
                'ProgressBar19.Value = Int(dTemp)
                'Label19.Text = FormatNumber(dTemp, 2) & "%"

                ' Accelerator Pedal Position F
                'ProgressBar20.Value = Int(dTemp)
                'Label20.Text = FormatNumber(dTemp, 2) & "%"

                ' Commanded Equivalence Ratio
                'ProgressBar21.Value = Int(dTemp)
                'Label21.Text = FormatNumber(dTemp, 2) & "%"

                ' Commanded Throttle Actuator
                'ProgressBar22.Value = Int(dTemp)
                'Label22.Text = FormatNumber(dTemp, 2) & "%"

                ' Ethanol %
                'ProgressBar23.Value = Int(dTemp)
                'Label23.Text = FormatNumber(dTemp, 2) & "%"

                ' Engine Oil Temperature
                'ProgressBar24.Value = Int(dTemp)
                'Label24.Text = FormatNumber(dTemp, 2) & "%"

                Thread.Sleep(10)

                ' Bank 1 Short Term Fuel Trim
                dTemp = cELM.getShortTermFuelTrim(1)
                ChangeProgress(ProgressBar41, Int(dTemp))
                ChangeText(Label66, FormatNumber(dTemp, 2) & "%")

                Thread.Sleep(10)

                ' Bank 1 Long Term Fuel Trim
                dTemp = cELM.getLongTermFuelTrim(1)
                ChangeProgress(ProgressBar42, Int(dTemp))
                ChangeText(Label67, FormatNumber(dTemp, 2) & "%")

                Thread.Sleep(10)

                ' Bank 2 Short Term Fuel Trim
                dTemp = cELM.getShortTermFuelTrim(2)
                ChangeProgress(ProgressBar43, Int(dTemp))
                ChangeText(Label68, FormatNumber(dTemp, 2) & "%")

                Thread.Sleep(10)

                ' Bank 2 Long Term Fuel Trim
                dTemp = cELM.getLongTermFuelTrim(2)
                ChangeProgress(ProgressBar44, Int(dTemp))
                ChangeText(Label69, FormatNumber(dTemp, 2) & "%")

                Thread.Sleep(10)

            Case 3 ' Oxygen Sensors

                ' Bank 1 Sensor 1
                dTemp = cELM.getOxygenSensorVoltage(1, 1)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar25, Int(dTemp) + 10)
                End If
                ChangeText(Label27, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(1, 1)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar29, Int(dTemp) + 200)
                End If
                ChangeText(Label28, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label29, "Lean")
                Else
                    ' RICH
                    ChangeText(Label29, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 1 Sensor 2
                dTemp = cELM.getOxygenSensorVoltage(1, 2)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar27, Int(dTemp) + 10)
                End If
                ChangeText(Label32, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(1, 2)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar26, Int(dTemp) + 200)
                End If
                ChangeText(Label31, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label30, "Lean")
                Else
                    ' RICH
                    ChangeText(Label30, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 1 Sensor 3
                dTemp = cELM.getOxygenSensorVoltage(1, 3)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar30, Int(dTemp) + 10)
                End If
                ChangeText(Label38, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(1, 3)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar28, Int(dTemp) + 200)
                End If
                ChangeText(Label37, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label36, "Lean")
                Else
                    ' RICH
                    ChangeText(Label36, "Rich")
                End If


                Thread.Sleep(10)

                ' Bank 1 Sensor 4
                dTemp = cELM.getOxygenSensorVoltage(1, 4)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar32, Int(dTemp) + 10)
                End If
                ChangeText(Label43, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(1, 4)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar31, Int(dTemp) + 200)
                End If
                ChangeText(Label42, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label41, "Lean")
                Else
                    ' RICH
                    ChangeText(Label41, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 2 Sensor 1
                dTemp = cELM.getOxygenSensorVoltage(2, 1)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar40, Int(dTemp) + 10)
                End If
                ChangeText(Label63, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(2, 1)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar39, Int(dTemp) + 200)
                End If
                ChangeText(Label62, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label61, "Lean")
                Else
                    ' RICH
                    ChangeText(Label61, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 2 Sensor 2
                dTemp = cELM.getOxygenSensorVoltage(2, 2)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar38, Int(dTemp) + 10)
                End If
                ChangeText(Label58, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(2, 2)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar37, Int(dTemp) + 200)
                End If
                ChangeText(Label57, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label56, "Lean")
                Else
                    ' RICH
                    ChangeText(Label56, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 2 Sensor 3
                dTemp = cELM.getOxygenSensorVoltage(2, 3)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar36, Int(dTemp) + 10)
                End If
                ChangeText(Label53, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(2, 3)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar35, Int(dTemp) + 200)
                End If
                ChangeText(Label52, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label51, "Lean")
                Else
                    ' RICH
                    ChangeText(Label51, "Rich")
                End If

                Thread.Sleep(10)

                ' Bank 2 Sensor 4
                dTemp = cELM.getOxygenSensorVoltage(2, 4)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar34, Int(dTemp) + 10)
                End If
                ChangeText(Label48, FormatNumber(dTemp, 2) & "V")

                Thread.Sleep(5)

                ' Fuel Trim
                dTemp = cELM.getOxygenSensorFuelTrim(2, 4)
                If dTemp > 0 Then
                    ChangeProgress(ProgressBar33, Int(dTemp) + 200)
                End If
                ChangeText(Label47, FormatNumber(dTemp, 2) & "%")

                ' Is it rich or lean
                If dTemp < 0 Then
                    'LEAN
                    ChangeText(Label46, "Lean")
                Else
                    ' RICH
                    ChangeText(Label46, "Rich")
                End If
        End Select
    End Sub

    Private Sub cELM_RAWDataReceived(ByVal sData As String) Handles cELM.RAWDataReceived
        AppendTextBox(TextBox1, sData)
    End Sub

    Private Sub cELM_RAWDataSent(ByVal sData As String) Handles cELM.RAWDataSent
        AppendTextBox(TextBox1, sData)
    End Sub

    
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If Not SensorThread.IsAlive() Then
            Try
                SensorThread = New Thread(AddressOf SensorUpdateThread)
                SensorThread.IsBackground = True
                SensorThread.Start()
            Catch ex As Exception

            End Try
        End If
    End Sub
End Class '3045524733

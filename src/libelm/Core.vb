Imports System.IO.Ports

Namespace libelm

    Public Class cELM327_New
        Private bBuffer(1024) As Byte
        Private pPort As New System.IO.Ports.SerialPort
        Private sSerialPort As String = "COM4"
        Private lSerialBaud As Long = 38400
        Private This_Logger As cLogger
        Private bTransmissionECU As Boolean
        Private tSPIDs(0) As SupportedPID

        Private sELMModel As String = "Unknown"
        Private sProtocol As String = "Unknown"

        Private Structure OBDPacket
            Dim Valid As Boolean
            Dim Header() As Byte
            Dim PID() As Byte
            Dim Data() As Byte
        End Structure

        Private Structure SupportedPID
            Dim Mode As Byte
            Dim PID As Byte
            Dim Supported As Boolean
        End Structure


        Public Function getlibelmVersion() As String
            Return "libelm " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
        End Function
        ' *******************************************************************************************
        '       FUNCTION: createPacket()
        '      ARGUMENTS: {INTEGER}
        '        RETURNS: {OBDPACKET}
        '          NOTES: Creates a Empty OBDPacket
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Private Function createPacket(ByVal lPIDCount As Integer) As OBDPacket
            Dim tPacket As OBDPacket

            ReDim tPacket.PID(0 To lPIDCount - 1)

            ReDim tPacket.Data(0)
            ReDim tPacket.Header(0)

            Return tPacket
        End Function

        ' *******************************************************************************************
        '       FUNCTION: requestPacket()
        '      ARGUMENTS: {STRING}
        '        RETURNS: {OBDPACKET}
        '          NOTES: Querys the ECU for the PID
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Private Function requestPacket(ByVal sPID As String) As OBDPacket
            Dim tPacket As OBDPacket
            Dim bContinue As Boolean = True
            Dim sReturnVal As String = ""
            Dim lFileHandle As Integer
            Dim sSplit() As String
            lFileHandle = FreeFile()

            FileOpen(lFileHandle, "C:\Users\Ian\Desktop\OBDSim\OBDLog.txt", OpenMode.Append)

            sSplit = Split(sPID, " ")

            tPacket = createPacket(UBound(sSplit))
            tPacket.Valid = False
            ReDim tPacket.PID(2)
            For lIndex As Integer = 0 To 1
                tPacket.PID(lIndex) = ("&H" & sSplit(lIndex))
            Next lIndex

            Try
                pPort.WriteTimeout = 1000
                pPort.Write(sPID + vbCr)
                WriteLine(lFileHandle, sPID)
            Catch ex As TimeoutException
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return tPacket
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return tPacket
            End Try

            Dim count As Integer

            pPort.ReadTimeout = 5000
            Try
                While (bContinue = True)
                    count = pPort.Read(bBuffer, 0, 1024)
                    sReturnVal += System.Text.Encoding.Default.GetString(bBuffer, 0, count)
                    If (sReturnVal.Contains(">")) Then
                        bContinue = False
                    End If
                End While
                WriteLine(lFileHandle, sReturnVal)
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return tPacket
            End Try
            FileClose(lFileHandle)

            If sReturnVal.Contains("NO DATA") Then
                Return tPacket
            End If

            If sReturnVal.Contains("UNABLE TO CONNECT") Then
                Disconnect()
                Return tPacket
            End If

            If sReturnVal.Contains("SEARCHING") Then
                sReturnVal = sReturnVal.Replace("SEARCHING...", "")
            End If

            ' Remove unneeded stuff
            sReturnVal = sReturnVal.Replace(vbCr, "")
            sReturnVal = sReturnVal.Replace(vbLf, "")
            sReturnVal = sReturnVal.Replace(">", "")

            ' VIN Number Fix
            '"41 6B 10 014  0: 49 02 01 31 47 31  1: 41 4B 35 32 46 32 35  2: 37 35 36 32 32 32 37"
            If sReturnVal.Contains(":") Then
                '"41 6B 10 49 02 01 31 47 31 41 4B 35 32 46 32 35 37 35 36 32 32 32 37"
                sReturnVal = sReturnVal.Replace(" 0: ", "")
                sReturnVal = sReturnVal.Replace(" 1: ", "")
                sReturnVal = sReturnVal.Replace(" 2: ", "")
                sReturnVal = sReturnVal.Replace(" 3: ", "")
                sReturnVal = sReturnVal.Replace(" 4: ", "")
                ' Remove Lenght (20 Bytes)
                sReturnVal = sReturnVal.Replace("014 ", "")
            End If

            sReturnVal = sReturnVal.Trim()
            '"41 6B 10 41 00 BF 9F B9 90 05"

            sSplit = Split(sReturnVal, " ")
            Dim bResponse(0 To UBound(sSplit)) As Byte
            For lIndex As Integer = 0 To UBound(sSplit)
                bResponse(lIndex) = ("&H" & sSplit(lIndex))
            Next

            ' We need to fill the packet
            ReDim tPacket.Header(2)
            ReDim tPacket.Data(UBound(bResponse) - 5)
            tPacket.Header(0) = bResponse(0)
            tPacket.Header(1) = bResponse(1)
            tPacket.Header(2) = bResponse(2)

            For lIndex As Integer = 0 To (UBound(bResponse) - 5)
                tPacket.Data(lIndex) = bResponse(lIndex + 5)
            Next lIndex

            ' Check to see if Packet is Valid
            If bResponse(3) = &H40 + tPacket.PID(0) Then
                If bResponse(4) = tPacket.PID(1) Then
                    tPacket.Valid = True
                Else
                    This_Logger.Log("Invalid Packet!")
                End If
            Else
                This_Logger.Log("Invalid Packet!")
            End If

            Return tPacket
        End Function

        ' *******************************************************************************************
        '       FUNCTION: getELMProtocol()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {STRING}
        '          NOTES: Retrives the ELM Protocol currently in use
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function getELMProtocol() As String
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Try
                pPort.WriteTimeout = 1000
                pPort.Write("DPN" + vbCr)
                This_Logger.Log("DPN", cLogger.LogType.SERIAL)
            Catch ex As TimeoutException
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return False
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return False
            End Try

            Dim bContinue As Boolean = True
            Dim sReturnVal As String = ""
            Dim lCount As Integer

            pPort.ReadTimeout = 5000
            Try
                While (bContinue = True)
                    lCount = pPort.Read(bBuffer, 0, 1024)
                    sReturnVal += System.Text.Encoding.Default.GetString(bBuffer, 0, lCount)
                    If (sReturnVal.Contains(">")) Then
                        bContinue = False
                    End If
                End While
            Catch ex As Exception
                Return False
            End Try

            sReturnVal = sReturnVal.Replace(vbCr, "")
            sReturnVal = sReturnVal.Replace(vbLf, "")
            sReturnVal = sReturnVal.Replace(">", "")

            sReturnVal = sReturnVal.Trim()
            Select Case sReturnVal
                Case "0"
                    Return "Automatic"
                Case "1"
                    Return "SAE J1850 PWM (41.6 Kbaud)"
                Case "2"
                    Return "SAE J1850 VPM (10.4 Kbaud)"
                Case "3"
                    Return "ISO 9141-2 (5 baud init, 10.4 Kbaud)"
                Case "4"
                    Return "ISO 14230-4 KWP (5 baud init, 10.4 Kbaud)"
                Case "5"
                    Return "ISO 14230-4 KWP (fast init, 10.4 Kbaud)"
                Case "6"
                    Return "ISO 15765-4 CAN (11 bit ID, 500 Kbaud)"
                Case "7"
                    Return "ISO 15765-4 CAN (29 bit ID, 500 Kbaud)"
                Case "8"
                    Return "ISO 15765-4 CAN (11 bit ID, 250 Kbaud)"
                Case "9"
                    Return "ISO 15765-4 CAN (29 bit ID, 250 Kbaud)"
                Case "A"
                    Return "SAE J1939 CAN (29 bit ID, 250* Kbaud)"
                Case "B"
                    Return "USER1 CAN (11* bit ID, 125* Kbaud)"
                Case "C"
                    Return "USER2 CAN (11* bit ID, 50* kbaud)"
                Case Else
                    Return "Unknown: " & sReturnVal
            End Select
        End Function

        ' *******************************************************************************************
        '       PROPERTY: ELMModel (Read Only)
        '          NOTES: Retrives the ELM Model of the interface
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public ReadOnly Property ELMModel() As String
            Get
                If sELMModel <> "" Then
                    Return sELMModel
                Else
                    If isConnected() Then
                        Dim sResponse As String

                        sResponse = sendATCommand("AT I")

                        sResponse = sResponse.Replace(vbCr, " ")
                        sResponse = sResponse.Replace(vbLf, " ")
                        sResponse = sResponse.Replace(">", " ")
                        sResponse = sResponse.Replace(vbCr, " ")
                        sResponse.Trim()

                        sELMModel = sResponse
                        Return sELMModel
                    Else
                        ' We should never get this
                        Return "Unknown"
                    End If
                End If


                Return "Unknown"
            End Get
        End Property

        ' *******************************************************************************************
        '       PROPERTY: Logger (Write Only)
        '          NOTES: You set the logging class with this property
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public WriteOnly Property Logger() As cLogger
            Set(ByVal cThis As cLogger)
                This_Logger = cThis
            End Set
        End Property

        ' *******************************************************************************************
        '       FUNCTION: formatHEX()
        '      ARGUMENTS: {BYTE}
        '        RETURNS: {STRING}
        '          NOTES: Converts "F" to "0F"
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function formatHEX(ByVal bByte As Byte) As String
            Dim sTemp As String

            sTemp = "00" & Hex(bByte)

            sTemp = sTemp.Substring(Len(sTemp) - 2)

            Return sTemp
        End Function

        ' *******************************************************************************************
        '       FUNCTION: isPIDSupported()
        '      ARGUMENTS: {BYTE(2)}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Returns TRUE if the PID is Supported by the ECU
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Private Function isPIDSupported(ByVal bPID() As Byte) As Boolean
            Dim lIndex As Integer

            For lIndex = 0 To UBound(tSPIDs)
                If tSPIDs(lIndex).Mode = bPID(0) Then
                    If tSPIDs(lIndex).PID = bPID(1) Then
                        Return tSPIDs(lIndex).Supported
                    End If
                End If
            Next lIndex

            Return False
        End Function


        ' *******************************************************************************************
        '       FUNCTION: getBITArray()
        '      ARGUMENTS: {BYTE}
        '        RETURNS: {STRING}
        '          NOTES: Returns a String containing the Bits from the {BYTE}
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function getBITArray(ByVal bByte As Byte) As String
            Dim bArray As New BitArray({bByte})
            Dim sResult As String = ""

            If bByte <= &HF Then
                sResult = sResult & "0000"
            Else
                sResult = sResult & IIf(bArray.Get(7), "1", "0")
                sResult = sResult & IIf(bArray.Get(6), "1", "0")
                sResult = sResult & IIf(bArray.Get(5), "1", "0")
                sResult = sResult & IIf(bArray.Get(4), "1", "0")
            End If
            sResult = sResult & IIf(bArray.Get(3), "1", "0")
            sResult = sResult & IIf(bArray.Get(2), "1", "0")
            sResult = sResult & IIf(bArray.Get(1), "1", "0")
            sResult = sResult & IIf(bArray.Get(0), "1", "0")

            Return sResult
        End Function

        ' *******************************************************************************************
        '       FUNCTION: Connect()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Connects to the ECU
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Function Connect() As Boolean
            pPort.PortName = sSerialPort
            pPort.BaudRate = lSerialBaud
            pPort.DataBits = 8
            pPort.Parity = Parity.None
            pPort.StopBits = StopBits.One
            pPort.Handshake = Handshake.None

            Try
                pPort.WriteTimeout = 1000
                pPort.Open()
                This_Logger.Log("Opening Serial Port (Port: " & sSerialPort & ", " & lSerialBaud & " Baud)")
            Catch ex As TimeoutException
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return False
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return False
            End Try

            pPort.DiscardInBuffer()
            pPort.DiscardOutBuffer()
            If (pPort.BytesToRead > 0) Then

            End If

            If Not resetInterface() Then
                This_Logger.Log("Failed to reset obd interface", cLogger.LogType.ERRMSG)
                Return False
            End If

            ' Line Feed
            If Not SetOption("ATL0") Then
                This_Logger.Log("Failed to Set Line Feed", cLogger.LogType.ERRMSG)
                Return False
            End If

            ' ECHO
            If Not SetOption("ATE0") Then
                This_Logger.Log("Failed to set Echo", cLogger.LogType.ERRMSG)
                Return False
            End If

            ' SET PROTOCOL
            If Not SetOption("ATSP0") Then
                This_Logger.Log("Failed to set  Auto-Detect protocol", cLogger.LogType.ERRMSG)
                Return False
            End If

            ' Headers
            If Not SetOption("ATH1") Then
                This_Logger.Log("Failed to set Headers", cLogger.LogType.ERRMSG)
                Return False
            End If

            ' Set Baud Rate to 115200
            If increaseBaudRate() Then
                This_Logger.Log("Baud Rate is now 115200bps")
            End If

            pPort.DiscardInBuffer()
            pPort.DiscardOutBuffer()

            Dim tPacket As OBDPacket

            tPacket = requestPacket("01 00")
            tPacket = requestPacket("01 20")
            tPacket = requestPacket("01 40")
            tPacket = requestPacket("01 60")
            tPacket = requestPacket("01 80")
            tPacket = requestPacket("01 A0")
            tPacket = requestPacket("01 C0")

            Return True
        End Function

        ' *******************************************************************************************
        '       FUNCTION: resetInterface()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Resets the OBD Interface and preps the connection
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 08/07/2013
        ' *******************************************************************************************
        Public Function resetInterface() As Boolean
            ' Resets the Interface to begin communication
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Dim sResponse As String

            sResponse = sendATCommand("ATZ")

            If sResponse.Contains("ELM") Then
                Return True
            Else
                Return False
            End If
        End Function

        ' *******************************************************************************************
        '       FUNCTION: checkECUConnection()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Bascally Pings the ECU to see if we are still connected.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function checkECUConnection() As Boolean
            'AT PC - Resets Protocoll MUST ADD!!!!

            Return True
        End Function

        Private Function sendATCommand(ByVal sCommand As String) As String
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Try
                pPort.WriteTimeout = 1000
                pPort.Write(sCommand + vbCr)
                This_Logger.Log(sCommand, cLogger.LogType.SERIAL)
            Catch ex As TimeoutException
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return ""
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return ""
            End Try

            Dim bContinue As Boolean = True
            Dim sReturnVal As String = ""
            Dim count As Integer

            pPort.ReadTimeout = 5000
            Try
                While (bContinue = True)
                    count = pPort.Read(bBuffer, 0, 1024)
                    sReturnVal += System.Text.Encoding.Default.GetString(bBuffer, 0, count)
                    If (sReturnVal.Contains(">")) Then
                        bContinue = False
                    End If
                End While
            Catch ex As Exception
                This_Logger.Log(ex.Message, cLogger.LogType.ERRMSG)
                Return ""
            End Try

            Return sReturnVal
        End Function

        ' *******************************************************************************************
        '       FUNCTION: Disconnect()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {N/A}
        '          NOTES: Disconnects from the OBD Interface.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Sub Disconnect()
            Try
                If pPort.IsOpen = True Then



                    pPort.Close()
                    This_Logger.Log("Connection Closed!")
                End If
            Catch ex As Exception
                This_Logger.Log(ex.Message)
            End Try
        End Sub

        ' *******************************************************************************************
        '       FUNCTION: increaseBaudRate()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Attempts to set the fastest baud rate.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 08/07/2013
        ' *******************************************************************************************
        Public Function increaseBaudRate() As Boolean
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Dim sResponse As String

            sResponse = sendATCommand("AT BRD 23")

            sResponse = sResponse.Replace(vbCr, "")
            sResponse = sResponse.Replace(vbLf, "")
            sResponse = sResponse.Replace(">", "")

            If sResponse.Contains("?") Then
                ' Baud rate isnt supported
                pPort.BaudRate = 38400
                Return False
            End If

            ' Change the Port Baud Rate
            pPort.BaudRate = 115200


            ' Continue Reading 

            Dim bContinue As Boolean = True
            Dim lCount As Integer
            Dim sReturnVal As String = ""
            pPort.ReadTimeout = 5000
            Try
                While (bContinue = True)
                    lCount = pPort.Read(bBuffer, 0, 1024)
                    sReturnVal += System.Text.Encoding.Default.GetString(bBuffer, 0, lCount)
                    If (sReturnVal.Contains(">")) Then
                        bContinue = False
                    End If
                End While
            Catch ex As Exception
                Return False
            End Try

            If sReturnVal.Contains("ELM") Then
                pPort.Write(vbCr)
                Return True
            End If

            ' Baud Rate isnt supported
            pPort.BaudRate = 38400
            Return False
        End Function

        ' *******************************************************************************************
        '       FUNCTION: SetOption()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Sets a Option on a ELM Device
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Private Function SetOption(ByVal sCommand As String) As Boolean
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Dim sResponse As String

            sResponse = sendATCommand(sCommand)

            If sResponse.Contains("OK") Then
                Return True
            Else
                Return False
            End If
        End Function

        ' *******************************************************************************************
        '       FUNCTION: getBatteryVoltage()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {DOUBLE}
        '          NOTES: Gets the Vehicle battery voltage.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Function getBatteryVoltage() As Double
            ' AT RV
            ' 12.6V >
            pPort.DiscardOutBuffer()
            pPort.DiscardInBuffer()

            Dim sReturnVal As String

            sReturnVal = sendATCommand("AT RV")

            If sReturnVal <> "" Then
                sReturnVal = sReturnVal.Replace(">", "")
                sReturnVal = sReturnVal.Replace(vbLf, "")
                sReturnVal = sReturnVal.Replace(vbCr, "")
                sReturnVal = sReturnVal.Replace(" ", "")
                sReturnVal = sReturnVal.Replace("V", "")

                Return CDbl(sReturnVal)
            Else
                Return 0.0
            End If

            Return 0.0
        End Function

        ' *******************************************************************************************
        '       FUNCTION: setBatteryVoltageReference()
        '      ARGUMENTS: {DOUBLE}
        '        RETURNS: {NONE}
        '          NOTES: Sets the ELM Battery reference voltage
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Sub setBatteryVoltageReference(ByVal refVoltage As Double)
            ' AT CV 1247            ' 12.7v
            Dim sTemp As String

            sTemp = "0000" & CStr(refVoltage)
            sTemp = sTemp.Replace(".", "")
            ' 0000947
            sTemp = sTemp.Substring(Len(sTemp) - 4)
            ' 0947

            Dim sReturnVal As String

            sReturnVal = sendATCommand("AT CV " + sTemp)

            If (sReturnVal.Contains("OK")) Then
                This_Logger.Log("Battery reference Voltage has been set!")
                Return
            Else
                This_Logger.Log("Battery reference voltage failed to set!", cLogger.LogType.ERRMSG)
                Return
            End If

            Return
        End Sub

        ' *******************************************************************************************
        '       PROPERTY: COMPort
        '          NOTES: The COM Port for which the Serial Port, will use.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Property COMPort() As String
            Get
                Return sSerialPort
            End Get
            Set(ByVal sPort As String)
                sSerialPort = sPort
            End Set
        End Property

        ' *******************************************************************************************
        '       PROPERTY: BaudRate
        '          NOTES: The Baud Rate, which the Serial Port, will use.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/7/2013
        ' *******************************************************************************************
        Public Property BaudRate() As Long
            Get
                Return lSerialBaud
            End Get
            Set(ByVal lBaud As Long)
                lSerialBaud = lBaud
            End Set
        End Property

        ' *******************************************************************************************
        '       FUNCTION: isConnected()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Returns True if we are connected to a vehicle.
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function isConnected() As Boolean
            If pPort.IsOpen Then
                Return True
            End If

            Return False
        End Function

        ' *******************************************************************************************
        '       FUNCTION: CelsiusToFahrenheit()
        '      ARGUMENTS: {DOUBLE}
        '        RETURNS: {DOUBLE}
        '          NOTES: Converts Celsius to Fahrenheit
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function CelsiusToFahrenheit(ByVal dCelsius As Double) As Double
            Return ((dCelsius * 9) / 5) + 32
        End Function

        Public Sub checkForPIDSupport()
            Dim tPacket As OBDPacket

            This_Logger.Log("Checking for Mode 1 PID Support...")

            This_Logger.Log("Requesting 0x00...")
            tPacket = requestPacket("01 00")

            If tPacket.Valid = False Then
                This_Logger.Log("Internal Error Checking for PIDs")
                Return
            End If

            Dim sBITArray As String = ""
            For lIndex = 0 To UBound(tPacket.Data) - 1
                sBITArray = sBITArray + getBITArray(tPacket.Data(lIndex))
            Next lIndex

            '41 6B 10 41 00 BF 9F B9 90 05
            This_Logger.Log("Checking ECU for PID's")
            Dim bFirstRun As Boolean = True
            Dim bBase As Integer
            Dim lSupported As Integer
            For lIndex = 0 To 224
                If lIndex = 31 Or lIndex = 63 Or lIndex = 95 Or lIndex = 127 Or lIndex = 159 Or lIndex = 192 Then
                    If sBITArray.Substring(lIndex, 1) = "1" Then
                        bBase = bBase + 32
                        This_Logger.Log("Requesting 0x" & formatHEX(bBase) & "...")

                        ' Request the Next Group of PIDs
                        tPacket = requestPacket("01 " & formatHEX(bBase))

                        If tPacket.Valid = False Then
                            Exit For
                        End If

                        For lTemp As Integer = 0 To UBound(tPacket.Data) - 1
                            sBITArray = sBITArray + getBITArray(tPacket.Data(lTemp))
                        Next lTemp
                    Else
                        bBase = bBase + 32
                        Exit For
                    End If
                End If


                If sBITArray.Substring(lIndex, 1) = "1" Then
                    'This_Logger.Log("0x01 0x" & formatHEX(lIndex + 1))

                    'This means There is a Supported PID

                    lSupported = lSupported + 1
                End If

            Next lIndex
            This_Logger.Log("Finished!")
            This_Logger.Log("ECU Supported: " & FormatNumber(lSupported, 0) & " PID's")

        End Sub


        ' *******************************************************************************************
        ' *******************************************************************************************
        '                V E H I C L E   P I D   F U N C T I O N S
        ' *******************************************************************************************
        ' *******************************************************************************************

        ' *******************************************************************************************
        '       FUNCTION: transmissionECUPresent()
        '      ARGUMENTS: {NONE}
        '        RETURNS: {BOOLEAN}
        '          NOTES: Returns True if the vehicle has a Transmission ECU
        '  
        '         AUTHOR: Ian M. Thomas
        '  LAST MODIFIED: 8/9/2013
        ' *******************************************************************************************
        Public Function transmissionECUPresent()
            Return bTransmissionECU
        End Function

        Public Function getCoolantTemperature() As Double
            Dim tPacket As OBDPacket
            Dim dTemp As Double

            tPacket = requestPacket("01 05")

            If tPacket.Valid = False Then
                Return 0.0
            End If

            dTemp = tPacket.Data(0) - 40

            Return CelsiusToFahrenheit(dTemp)
        End Function

        Public Function getVehicleVIN() As String
            Dim tPacket As OBDPacket

            Dim sTemp As String = ""

            tPacket = requestPacket("09 02")
            If tPacket.Valid = False Then
                Return "N/A"
            End If
            ' We dont need to verify the VIN response it is rigged in processOBDResponse()

            For lIndex As Integer = 0 To UBound(tPacket.Data)
                sTemp = sTemp & Chr(tPacket.Data(lIndex))
            Next lIndex

            Return sTemp
        End Function

        Public Function getEngineRPM() As Integer
            Dim tPacket As OBDPacket

            Dim dTemp As Integer

            tPacket = requestPacket("01 0C")

            If tPacket.Valid = False Then
                Return 0
            End If

            dTemp = ((tPacket.Data(0) * 256) + tPacket.Data(1)) / 4

            Return dTemp
        End Function

        Public Function getVehicleSpeed() As Double
            Dim tPacket As OBDPacket

            Dim dTemp As Double

            tPacket = requestPacket("01 0D")

            If tPacket.Valid = False Then
                Return 0.0
            End If

            dTemp = tPacket.Data(0) * 0.621371

            Return dTemp
        End Function

        Public Function getODBStandard() As String
            Dim tPacket As OBDPacket

            tPacket = requestPacket("01 1C")

            If tPacket.Valid = False Then
                Return "Unknown"
            End If

            Select Case tPacket.Data(0)
                Case &H1
                    Return "OBD-II"
                Case &H2
                    Return "OBD (EPA)"
                Case &H3
                    Return "OBD & OBD-II"
                Case &H4
                    Return "OBD-I"
                Case &H5
                    Return "No OBD Standard"
                Case &H6
                    Return "EOBD (Europe)"
                Case &H7
                    Return "EOBD & OBD-II"
                Case &H8
                    Return "EOBD and OBD"
                Case &H9
                    Return "EOBD, OBD, & OBD-II"
                Case &HA
                    Return "JOBD (Japan)"
                Case &HB
                    Return "JOBD & OBD-II"
                Case &HC
                    Return "JOBD & EOBD"
                Case &HD
                    Return "JOBD, EOBD, & OBD-II"
            End Select

            Return "Unknown"
        End Function

        Public Function getEngineLoad() As Double
            Dim tPacket As OBDPacket
            Dim dTemp As Double

            tPacket = requestPacket("01 04")

            If tPacket.Valid = False Then
                Return 0.0
            End If

            dTemp = (tPacket.Data(0) / 255) * 100

            Return dTemp
        End Function

        Public Function getShortTermFuelTrim(ByVal lBank As Integer) As Double
            Dim tPacket As OBDPacket
            Dim dTemp As Double

            If lBank = 1 Then
                tPacket = requestPacket("01 06")
            Else
                tPacket = requestPacket("01 08")
            End If

            If tPacket.Valid = False Then
                Return 0.0
            End If

            dTemp = (tPacket.Data(0) - 128) * 100 / 128

            Return dTemp
        End Function

        Public Function getLongTermFuelTrim(ByVal lBank As Integer) As Double
            Dim tPacket As OBDPacket
            Dim dTemp As Double

            If lBank = 1 Then
                tPacket = requestPacket("01 07")
            Else
                tPacket = requestPacket("01 09")
            End If

            If tPacket.Valid = False Then
                Return 0.0
            End If

            dTemp = (tPacket.Data(0) - 128) * 100 / 128

            Return dTemp
        End Function
        
        Public Function getFuelPressure() As Double
            ' 01 0A

            Dim dTemp As Double
            Dim tPacket As OBDPacket

            tPacket = requestPacket("01 0A")

            If tPacket.Valid = False Then
                Return 0.0
            End If
            
            dTemp = (tPacket.Data(0) * 3)

            Return dTemp
        End Function

        Public Function getIntakeManifoldPressure() As Double
            ' 01 0B

            Dim bResponse() As Byte
            Dim tPacket As OBDPacket

            tPacket = requestPacket("01 0B")

            If tPacket.Valid = False Then
                Return 0.0
            End If

            Return tPacket.Data(0)
        End Function


    End Class
End Namespace


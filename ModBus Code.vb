Module ModBus_Code
    'Define ModBus Polling array
    'Columns are: 1) Address 2) Value
    Public VacuDrawArray(17, 1) As Single
    Public EuroThermStatus As String
    Public strINIFile As String
    Public strIPAdd As String

    Function FillVacuDrawArray() As Boolean
        FillVacuDrawArray = False
        Try
            VacuDrawArray(0, 0) = 41433 'TempSP Channel 1
            VacuDrawArray(1, 0) = 41436 'TempAct Channel 2
            VacuDrawArray(2, 0) = 41439 'TempOut Channel 3
            VacuDrawArray(3, 0) = 41442 'TempLoad Channel 4
            VacuDrawArray(4, 0) = 41445 'Vacuum Channel 5
            VacuDrawArray(5, 0) = 41448 'Pressure Channel 6
            VacuDrawArray(6, 0) = 41451 'HeaterLoadA Channel 7
            VacuDrawArray(7, 0) = 41454 'HeaterLoadB Channel 8
            VacuDrawArray(8, 0) = 41457 'TempWaterIn Channel 9
            VacuDrawArray(9, 0) = 41460 'TempWaterOut Channel 10
            VacuDrawArray(10, 0) = 41490 'CirculationFanOn Channel 20
            VacuDrawArray(11, 0) = 41493 'CoolingFanOn Channel 21
            VacuDrawArray(12, 0) = 41496 'RoughingPumpOn Channel 22
            VacuDrawArray(13, 0) = 41499 'BoosterOn Channel 23
            VacuDrawArray(14, 0) = 41502 'RecipeNumber Channel 24
            VacuDrawArray(15, 0) = 41505 'RecipeRunning Channel 25
            VacuDrawArray(16, 0) = 41508 'Active Alarm Channel 26
            VacuDrawArray(17, 0) = 41511 'Door Open Channel 27

            FillVacuDrawArray = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception in FillVacuDrawArray")
        End Try
    End Function

    Function Read_VacuDraw() As Boolean
        Dim n As Integer
        Dim booConn As Boolean
        Dim HR(1) As Integer
        Dim intHR As Integer
        Dim Machine = New EasyModbus.ModbusClient
        Read_VacuDraw = False

        booConn = False 'Used to determine if communcations with this machine was made
        Do
            n = n + 1
            Try
                If Machine.Connected = False Then
                    Machine.IPAddress = strIPAdd
                    Machine.Port = 502
                    Machine.Connect()
                Else
                    Threading.Thread.Sleep(100)
                End If
            Catch ex As Exception
                Machine.Disconnect()
            End Try
        Loop While n <= 3 And Machine.Connected = False
        If Machine.Connected = True Then
            booConn = True 'Connection is made
        Else
            WriteMessage("Error in Read_VacuDraw. Unable to connect to ModBus Server.")
            Call MakeBlankData()
            Exit Function
        End If

        Try
            'This gets the EuroTherm Device condirtion from Registers 22 and 23 (page 277 of USer Guide.
            'Register 22 is Instrument Status. See manual for results.
            'Register 23 is FormatNumber of config changes since Power Up.
            'This puts them as the first two digits (or more) in the status string.
            EuroThermStatus = ""
            HR = Machine.ReadHoldingRegisters(22, 2)
            EuroThermStatus = HR(0).ToString & HR(1).ToString & "-"

            'This iterates through the register array.
            For n = 0 To VacuDrawArray.GetUpperBound(0)
                HR = Machine.ReadHoldingRegisters(VacuDrawArray(n, 0), 2)
                intHR = CInt(HR(0))
                VacuDrawArray(n, 1) = intHR
                'This writes the status from the Eurotherm for each channel to a sequential string
                EuroThermStatus = EuroThermStatus & HR(1).ToString
            Next
            Machine.Disconnect()
            Read_VacuDraw = True
        Catch ex As Exception
            WriteMessage("Exception in Read_VacuDraw " & ex.Message)
            Call MakeBlankData()
        End Try

        'Debugging use
        'For n = 0 To VacuDrawArray.GetUpperBound(0)
        'Console.WriteLine(VacuDrawArray(n, 1).ToString)
        'Next n
    End Function

    Public Sub MakeBlankData()
        Dim n As Integer
        For n = 0 To VacuDrawArray.GetUpperBound(0)
            VacuDrawArray(n, 1) = 0
        Next
    End Sub
End Module

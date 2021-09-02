Imports System.IO
Imports System.Data.SqlClient

Module Accessories
    Public strServerConn As String
    Public strAppendQuery As String

    Function Read_INI(ByVal INI_File As String, ByVal INI_Section As String, ByVal INI_Key As String) As String
        Read_INI = vbLf & False.ToString
        Dim INILine, tmpKey, tmpKeyVal As String
        Dim InSection, IsSection, isKey As Boolean
        Dim n As Integer

        Try
            'Prep the passed strings
            If INI_File IsNot Nothing Then
                INI_File = Trim(UCase(INI_File))
            Else
                Throw New Exception("Passed INI_File parameter empty or null.")
            End If

            If INI_Section IsNot Nothing Then
                INI_Section = Trim(UCase(INI_Section))
            Else
                Throw New Exception("Passed INI_Section parameter empty or null.")
            End If

            If INI_Key IsNot Nothing Then
                INI_Key = Trim(UCase(INI_Key))
            Else

            End If

            'Checks for existence of INI_File string
            If Not File.Exists(INI_File) Then
                Throw New Exception("Passed INI_File filename does not exist.")
            End If

            'Processes open INI file to find specified Section and Key
            InSection = False
            IsSection = False
            tmpKey = ""
            tmpKeyVal = ""

            Dim INIReader As StreamReader = New StreamReader(INI_File)

            Do While Not INIReader.EndOfStream
                'Read one line from INI_File
                INILine = INIReader.ReadLine()
                INILine = Trim(INILine)
                'If line is neither blank nor a comment
                If Not INILine = "" And Not Mid(INILine, 1, 1) = "#" Then
                    'This determines if the line is a Section definition
                    If Mid(INILine, 1, 1) = "[" And Mid(INILine, Len(INILine), 1) = "]" Then
                        IsSection = True
                    Else
                        IsSection = False
                    End If
                    'If the line is a Section and the desired Section has not been found yet
                    'the flag InSection will be set to true and Keys can be looked for.
                    If IsSection = True And InSection = False Then
                        If UCase(INILine) = INI_Section Then
                            InSection = True
                        End If
                    ElseIf IsSection = True And InSection = True Then
                        'If the line is a Section and the desired Section has been found
                        'then the process can be exited.
                        Exit Do
                    End If
                    'If the desired Section has been read and the line is not the section header then
                    'the line can be evaluated for the Key
                    If InSection = True And IsSection = False Then
                        'Determines if the Line is a Key definition
                        'Position of the = sign will not be 0 if found
                        If Not InStr(1, INILine, "=") = 0 Then
                            isKey = True
                            n = InStr(1, INILine, "=")
                            'Evaluate Key based on position of = sign
                            If n = 1 Then
                                tmpKey = ""
                                tmpKeyVal = ""
                            Else
                                'Assign tmkKey and tmpKeyVal to results frm line parsing around '=' sign
                                tmpKey = UCase(Trim(Left(INILine, n - 1)))
                                tmpKeyVal = Trim(Right(INILine, Len(INILine) - n))
                                If tmpKey = INI_Key Then
                                    'Assigns value of tmpKeyVal to Function Return value and exits do loop
                                    Read_INI = tmpKeyVal
                                    Exit Do
                                End If
                            End If
                        Else
                            isKey = False
                        End If
                    End If
                End If
            Loop
            'All done with INIReader now
            INIReader.Close()

        Catch ex As Exception
            'May want to call a logwriter here if there is an exception thrown so the cause can be recorded.
        End Try
    End Function

    Function DataBase_Init() As Boolean
        DataBase_Init = False
        Dim strSQL As String
        Try
            'Testing connection 
            Dim DBConn As New SqlConnection(strServerConn)
            DBConn.Open()
            strSQL = "SELECT * FROM [ProcessData].[dbo].[VacuDraw]"
            Dim DBcmd1 As New SqlCommand(strSQL, DBConn)
            Dim DBReader1 As SqlDataReader = DBcmd1.ExecuteReader
            If DBReader1.HasRows = False Then
                DataBase_Init = False
                Throw New Exception("Cannot connect to DataBase.")
            Else
                DataBase_Init = True
            End If
            DBReader1.Close()
            DBcmd1.Dispose()
            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Sub Write_to_DB()
        Dim DBConn As New SqlConnection(strServerConn)
        Dim strAppendQuery As String
        Dim RecordCount As Integer
        Dim intTemp As Integer
        Dim sngTemp As Single
        Dim ParamString, Portion As String

        Portion = ""
        ParamString = ""
        Try
            Portion = "SQL Definitions"
            'Append aggregate values
            strAppendQuery = "INSERT INTO [VacuDraw] (TempSP, TempAct, TempOut, " &
                "TempLoad, Vacuum, Pressure, HeaterLoadA, HeaterLoadB, TempWaterIn, " &
                "TempWaterOut, CirculationFanOn, CoolingFanOn, RoughingPumpOn, " &
                "BoosterOn, RecipeNumber, RecipeRunning, AlarmCondition, DoorOpen, EurothermStatus )" &
                "VALUES (@P00, @P01, @P02, @P03, @P04, @P05, @P06, @P07, " &
                "@P08, @P09, @P10, @P11, @P12, @P13, @P14, @P15, @P16, @P17, @P18);"

            'Creates Parameters for database writing
            Dim SQLParams As New List(Of SqlParameter)
            Dim Param00 As New SqlParameter("@P00", 0) 'TempSP
            Dim Param01 As New SqlParameter("@P01", 0) 'TempAct
            Dim Param02 As New SqlParameter("@P02", 0) 'TempOut
            Dim Param03 As New SqlParameter("@P03", 0) 'TempLoad
            Dim Param04 As New SqlParameter("@P04", 0) 'Vacuum
            Dim Param05 As New SqlParameter("@P05", 0) 'Pressure
            Dim Param06 As New SqlParameter("@P06", 0) 'HeaterLoadA
            Dim Param07 As New SqlParameter("@P07", 0) 'HeaterLoadB
            Dim Param08 As New SqlParameter("@P08", 0) 'TempWaterIn
            Dim Param09 As New SqlParameter("@P09", 0) 'TempWaterOut
            Dim Param10 As New SqlParameter("@P10", 0) 'CirculationFanOn
            Dim Param11 As New SqlParameter("@P11", 0) 'CoolingFanOn
            Dim Param12 As New SqlParameter("@P12", 0) 'RoughingPumpOn
            Dim Param13 As New SqlParameter("@P13", 0) 'BoosterOn
            Dim Param14 As New SqlParameter("@P14", 0) 'RecipeNumber
            Dim Param15 As New SqlParameter("@P15", 0) 'RecipeRunning
            Dim Param16 As New SqlParameter("@P16", 0) 'RecipeRunning
            Dim Param17 As New SqlParameter("@P17", 0) 'RecipeRunning
            Dim Param18 As New SqlParameter("@P18", "X") 'Status String from EuroTherm transmitter
            'Add Params to List of Parameters
            SQLParams.Add(Param00) : SQLParams.Add(Param01) : SQLParams.Add(Param02)
            SQLParams.Add(Param03) : SQLParams.Add(Param04) : SQLParams.Add(Param05)
            SQLParams.Add(Param06) : SQLParams.Add(Param07) : SQLParams.Add(Param08)
            SQLParams.Add(Param09) : SQLParams.Add(Param10) : SQLParams.Add(Param11)
            SQLParams.Add(Param12) : SQLParams.Add(Param13) : SQLParams.Add(Param14)
            SQLParams.Add(Param15) : SQLParams.Add(Param16) : SQLParams.Add(Param17)
            SQLParams.Add(Param18)

            Portion = "Assignment of Values"
            'Assign values to Parameters
            'Temperature Set Point
            intTemp = CInt(VacuDrawArray(0, 1))
            If intTemp < 0 Or intTemp > 2000 Then intTemp = 0
            Param00.Value = intTemp

            'Actual Control Temperature
            intTemp = CInt(VacuDrawArray(1, 1))
            If intTemp < 0 Or intTemp > 2000 Then intTemp = 0
            Param01.Value = intTemp

            'Heating Output %
            sngTemp = VacuDrawArray(2, 1)
            If sngTemp < 0 Then sngTemp = 0
            Param02.Value = sngTemp

            'Work Load Temperature
            intTemp = CInt(VacuDrawArray(3, 1))
            If intTemp < 0 Or intTemp > 3000 Then intTemp = 0
            Param03.Value = intTemp

            'Vacuum level
            sngTemp = VacuDrawArray(4, 1) / 10
            If sngTemp < 0 Or sngTemp > 3200 Then sngTemp = 3200 'This sets it into a range that won't be charted.
            Param04.Value = sngTemp

            'Pressure Level
            sngTemp = VacuDrawArray(5, 1) / 10
            If sngTemp < -14.5 Then sngTemp = -14.5 'This sets it into a range that won't be charted.
            Param05.Value = sngTemp

            'Heater Load A
            sngTemp = VacuDrawArray(6, 1) / 100
            If sngTemp < 0 Then sngTemp = 0
            Param06.Value = sngTemp

            'Heater Load B
            sngTemp = VacuDrawArray(7, 1) / 100
            If sngTemp < 0 Then sngTemp = 0
            Param07.Value = sngTemp

            'Temperature Inlet Water
            intTemp = CInt(VacuDrawArray(8, 1))
            If intTemp < 0 Or intTemp > 1000 Then intTemp = 0
            Param08.Value = intTemp

            'Temperature Set Point
            intTemp = CInt(VacuDrawArray(9, 1))
            If intTemp < 0 Or intTemp > 1000 Then intTemp = 0
            Param09.Value = intTemp

            'Circulation Fan On
            If CInt(VacuDrawArray(10, 1)) = 1 Then
                Param10.Value = 1
            Else
                Param10.Value = 0
            End If

            'Cooling Fan On
            If CInt(VacuDrawArray(11, 1)) = 1 Then
                Param11.Value = 1
            Else
                Param11.Value = 0
            End If

            'Roughing Pump On
            If CInt(VacuDrawArray(12, 1)) = 1 Then
                Param12.Value = 1
            Else
                Param12.Value = 0
            End If

            'Booster On
            If CInt(VacuDrawArray(13, 1)) = 1 Then
                Param13.Value = 1
            Else
                Param13.Value = 0
            End If

            'Recipe Number
            intTemp = CInt(VacuDrawArray(14, 1))
            If intTemp < 0 Or intTemp > 200 Then
                Param14.Value = 0
            Else
                Param14.Value = intTemp
            End If

            'Recipe Running
            If CInt(VacuDrawArray(15, 1)) = 1 Then
                Param15.Value = 1
            Else
                Param15.Value = 0
            End If

            'Active Alarm
            If CInt(VacuDrawArray(16, 1)) = 1 Then
                Param16.Value = 1
            Else
                Param16.Value = 0
            End If

            'Door Open
            If CInt(VacuDrawArray(17, 1)) = 1 Then
                Param17.Value = 1
            Else
                Param17.Value = 0
            End If

            'EuroTherm Status
            Param18.Value = EuroThermStatus

            'This applies some logig to account for variances in the way data gets sent
            'to the Eurotherm data collector.
            If Param11.Value = 1 Then 'Cooling is on
                Param02.Value = 0 'No heater output is recorded
            End If

            'Debugging steps for Parameter values
            'Param00.Value = 0
            'Param01.Value = 0
            'Param02.Value = 0
            'Param03.Value = 0
            'Param04.Value = 0
            'Param05.Value = 0
            'Param06.Value = 0
            'Param07.Value = 0
            'Param08.Value = 0
            'Param09.Value = 0
            'Param10.Value = 0
            'Param11.Value = 0
            'Param12.Value = 0
            'Param13.Value = 0
            'Param14.Value = 0
            'Param15.Value = 0
            'Param15.Value = 0
            'Param17.Value = 0
            'Param18.Value = "XYZ"

            Portion = "Database Activities"
            'Execute Append query
            Dim DBcmd1 As New SqlCommand(strAppendQuery, DBConn)
            SQLParams.ForEach(Sub(p) DBcmd1.Parameters.Add(p))

            'Generate parameter list string, just in case
            ParamString = vbCrLf & "Parameter Values listed below:" & vbCrLf
            SQLParams.ForEach(Sub(p) ParamString = ParamString & p.ToString & ": " & p.SqlValue.ToString & vbCrLf)

            'Debugging steps for Parameter values
            'Console.WriteLine("Param00: " & Param00.Value.ToString)
            'Console.WriteLine("Param01: " & Param01.Value.ToString)
            'Console.WriteLine("Param02: " & Param02.Value.ToString)
            'Console.WriteLine("Param03: " & Param03.Value.ToString)
            'Console.WriteLine("Param04: " & Param04.Value.ToString)
            'Console.WriteLine("Param05: " & Param05.Value.ToString)
            'Console.WriteLine("Param06: " & Param06.Value.ToString)
            'Console.WriteLine("Param07: " & Param07.Value.ToString)
            'Console.WriteLine("Param08: " & Param08.Value.ToString)
            'Console.WriteLine("Param09: " & Param09.Value.ToString)
            'Console.WriteLine("Param10: " & Param10.Value.ToString)
            'Console.WriteLine("Param11: " & Param11.Value.ToString)
            'Console.WriteLine("Param12: " & Param12.Value.ToString)
            'Console.WriteLine("Param13: " & Param13.Value.ToString)
            'Console.WriteLine("Param14: " & Param14.Value.ToString)
            'Console.WriteLine("Param15: " & Param15.Value.ToString)
            'Console.WriteLine("Param16: " & Param15.Value.ToString)
            'Console.WriteLine("Param17: " & Param17.Value.ToString)
            'Console.WriteLine("Param18: " & Param18.Value.ToString)

            Dim dbDT As New DataTable
            Dim dbDA As New SqlDataAdapter(DBcmd1)
            RecordCount = dbDA.Fill(dbDT)

            MainForm.txtLastRead.Text = Now
            MainForm.txtLastRead.Refresh()

            DBcmd1.Dispose()
            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If
            'comment here.
        Catch ex As Exception
            WriteMessage(ex.Message & vbCrLf & Portion & ParamString)
        End Try
    End Sub

    Public Sub WriteMessage(txtMsg As String)
        Dim strMessage As String
        Try
            strMessage = vbCrLf & Now.ToString & vbTab & Trim(txtMsg)
            MainForm.txtMessages.AppendText(strMessage)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error on Message writing.")
        End Try
    End Sub
End Module

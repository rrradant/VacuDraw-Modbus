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
            If Not INI_File Is Nothing Then
                INI_File = Trim(UCase(INI_File))
            Else
                Throw New Exception("Passed INI_File parameter empty or null.")
            End If

            If Not INI_Section Is Nothing Then
                INI_Section = Trim(UCase(INI_Section))
            Else
                Throw New Exception("Passed INI_Section parameter empty or null.")
            End If

            If Not INI_Key Is Nothing Then
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
            strServerConn = Read_INI(strINIFile, "[Connection]", "ConnString")

            'Testing connection 
            Dim DBConn As New SqlConnection(strServerConn)
            DBConn.Open()
            strSQL = "SELECT * FROM [TemperatureSurvey].[dbo].[Equipment]"
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

        Try


            'Append aggregate values
            strAppendQuery = "INSERT INTO [SurveyTemps] (EquipID, TC01, TC02, TC03, TC04, TC05, TC06, " &
                        "TC07, TC08, TC09, TC10, TC11, TC12 )" &
                        "VALUES (@Equip, @R01, @R02, @R03, @R04, @R05, @R06, @R07, " &
                        "@R08, @R09, @R10, @R11, @R12);"

            'Creates Parameters for database writing
            Dim SQLParams As New List(Of SqlParameter)
            '"VALUES (@Stamp, @Comms, @TC1, @TC2, @Z1, @Z2, @Z3, @Z4, @Z5, @Z1A, @Z2A, @Z3A, @Z4A, @Z5A)
            Dim Param01 As New SqlParameter("@Equip", 0)
            Dim Param02 As New SqlParameter("@R01", 0)
            Dim Param03 As New SqlParameter("@R02", 0)
            Dim Param04 As New SqlParameter("@R03", 0)
            Dim Param05 As New SqlParameter("@R04", 0)
            Dim Param06 As New SqlParameter("@R05", 0)
            Dim Param07 As New SqlParameter("@R06", 0)
            Dim Param08 As New SqlParameter("@R07", 0)
            Dim Param09 As New SqlParameter("@R08", 0)
            Dim Param10 As New SqlParameter("@R09", 0)
            Dim Param11 As New SqlParameter("@R10", 0)
            Dim Param12 As New SqlParameter("@R11", 0)
            Dim Param13 As New SqlParameter("@R12", 0)

            SQLParams.Add(Param01) : SQLParams.Add(Param02) : SQLParams.Add(Param03)
            SQLParams.Add(Param04) : SQLParams.Add(Param05) : SQLParams.Add(Param06)
            SQLParams.Add(Param07) : SQLParams.Add(Param08) : SQLParams.Add(Param09)
            SQLParams.Add(Param10) : SQLParams.Add(Param11) : SQLParams.Add(Param12)
            SQLParams.Add(Param13)

            Param01.Value = MainForm.cmbSource.SelectedValue
            If Not TempArray(0, 3) = -1 Then Param02.Value = TempArray(0, 3)
            If Not TempArray(1, 3) = -1 Then Param03.Value = TempArray(1, 3)
            If Not TempArray(2, 3) = -1 Then Param04.Value = TempArray(2, 3)
            If Not TempArray(3, 3) = -1 Then Param05.Value = TempArray(3, 3)
            If Not TempArray(4, 3) = -1 Then Param06.Value = TempArray(4, 3)
            If Not TempArray(5, 3) = -1 Then Param07.Value = TempArray(5, 3)
            If Not TempArray(6, 3) = -1 Then Param08.Value = TempArray(6, 3)
            If Not TempArray(7, 3) = -1 Then Param09.Value = TempArray(7, 3)
            If Not TempArray(8, 3) = -1 Then Param10.Value = TempArray(8, 3)
            If Not TempArray(9, 3) = -1 Then Param11.Value = TempArray(9, 3)
            If Not TempArray(10, 3) = -1 Then Param12.Value = TempArray(10, 3)
            If Not TempArray(11, 3) = -1 Then Param13.Value = TempArray(11, 3)

            'Param03.Value = TempArray(1, 3)
            'Param04.Value = TempArray(2, 3)
            'Param05.Value = TempArray(3, 3)
            'Param06.Value = TempArray(4, 3)
            'Param07.Value = TempArray(5, 3)
            'Param08.Value = TempArray(6, 3)
            'Param09.Value = TempArray(7, 3)
            'Param10.Value = TempArray(8, 3)
            'Param11.Value = TempArray(9, 3)
            'Param12.Value = TempArray(10, 3)
            'Param13.Value = TempArray(11, 3)

            'Execute Append query
            Dim DBcmd1 As New SqlCommand(strAppendQuery, DBConn)
            SQLParams.ForEach(Sub(p) DBcmd1.Parameters.Add(p))
            Dim dbDT As New DataTable
            Dim dbDA As New SqlDataAdapter(DBcmd1)
            RecordCount = dbDA.Fill(dbDT)

            DBcmd1.Dispose()
            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Module

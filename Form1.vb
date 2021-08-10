Public Class MainForm


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strVer As String
        Try
            cmdStartPolling.Enabled = False
            cmdStopPolling.Enabled = False

            'Get Version and Deployment information
            strVer = "Program version: " & My.Application.Info.Version.ToString
            If Deployment.Application.ApplicationDeployment.IsNetworkDeployed = True Then
                strVer = strVer & vbNewLine
                strVer = strVer & "Deployment: " & Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            End If


            strINIFile = FileIO.FileSystem.CurrentDirectory.ToString & "\" & My.Application.Info.AssemblyName & ".ini"
            strIPAdd = Read_INI(strINIFile, "[Connection]", "IPAddress")
            strServerConn = Read_INI(strINIFile, "[Connection]", "ConnString")
            VersionText.Text = strVer 'System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            If IsNothing(strIPAdd) Then
                Throw New Exception("Cannot read IP Address from INI file.")
            End If
            If IsNothing(strServerConn) Then
                Throw New Exception("Cannot read Connection String from INI file.")
            End If
            If DataBase_Init() = False Then
                Throw New Exception("Cannot make connection to database.")
            End If
            cmdStartPolling.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Main Form Load")
        End Try


    End Sub

    Private Sub cmdStartPolling_Click(sender As Object, e As EventArgs) Handles cmdStartPolling.Click
        If FillVacuDrawArray() = True Then
            ModBus_Polling_Timer.Enabled = True
            cmdStopPolling.Enabled = True
            cmdStopPolling.Select()
            cmdStartPolling.Enabled = False
            If Read_VacuDraw() = True Then
                Call Write_to_DB()
            End If
        End If
    End Sub

    Private Sub ModBus_Polling_Timer_Tick(sender As Object, e As EventArgs) Handles ModBus_Polling_Timer.Tick
        If Read_VacuDraw() = True Then
            Call Write_to_DB()
        End If

    End Sub

    Private Sub cmdStopPolling_Click(sender As Object, e As EventArgs) Handles cmdStopPolling.Click
        ModBus_Polling_Timer.Enabled = False
        cmdStartPolling.Enabled = True
        cmdStartPolling.Select()
        cmdStopPolling.Enabled = False
    End Sub
End Class

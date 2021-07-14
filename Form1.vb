Public Class MainForm


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        strINIFile = FileIO.FileSystem.CurrentDirectory.ToString & "\" & My.Application.Info.AssemblyName & ".ini"
    End Sub

    Private Sub cmdStartPolling_Click(sender As Object, e As EventArgs) Handles cmdStartPolling.Click
        If FillVacuDrawArray() Then
            Call Read_VacuDraw()
        End If
    End Sub
End Class

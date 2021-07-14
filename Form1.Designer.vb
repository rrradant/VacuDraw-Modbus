<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ModBus_Polling_Timer = New System.Windows.Forms.Timer(Me.components)
        Me.cmdStartPolling = New System.Windows.Forms.Button()
        Me.cmdStopPolling = New System.Windows.Forms.Button()
        Me.txtLastRead = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMessages = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ModBus_Polling_Timer
        '
        Me.ModBus_Polling_Timer.Interval = 60000
        '
        'cmdStartPolling
        '
        Me.cmdStartPolling.Enabled = False
        Me.cmdStartPolling.Location = New System.Drawing.Point(26, 25)
        Me.cmdStartPolling.Name = "cmdStartPolling"
        Me.cmdStartPolling.Size = New System.Drawing.Size(92, 33)
        Me.cmdStartPolling.TabIndex = 0
        Me.cmdStartPolling.Text = "Start ModBus"
        Me.cmdStartPolling.UseVisualStyleBackColor = True
        '
        'cmdStopPolling
        '
        Me.cmdStopPolling.Enabled = False
        Me.cmdStopPolling.Location = New System.Drawing.Point(124, 25)
        Me.cmdStopPolling.Name = "cmdStopPolling"
        Me.cmdStopPolling.Size = New System.Drawing.Size(92, 33)
        Me.cmdStopPolling.TabIndex = 1
        Me.cmdStopPolling.Text = "Stop ModBus"
        Me.cmdStopPolling.UseVisualStyleBackColor = True
        '
        'txtLastRead
        '
        Me.txtLastRead.Location = New System.Drawing.Point(238, 38)
        Me.txtLastRead.Name = "txtLastRead"
        Me.txtLastRead.Size = New System.Drawing.Size(148, 20)
        Me.txtLastRead.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(235, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Last Read:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMessages
        '
        Me.txtMessages.Location = New System.Drawing.Point(26, 75)
        Me.txtMessages.Multiline = True
        Me.txtMessages.Name = "txtMessages"
        Me.txtMessages.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMessages.Size = New System.Drawing.Size(450, 130)
        Me.txtMessages.TabIndex = 4
        Me.txtMessages.WordWrap = False
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(505, 234)
        Me.Controls.Add(Me.txtMessages)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtLastRead)
        Me.Controls.Add(Me.cmdStopPolling)
        Me.Controls.Add(Me.cmdStartPolling)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.Text = "VacuDraw Data Acquisition"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ModBus_Polling_Timer As Timer
    Friend WithEvents cmdStartPolling As Button
    Friend WithEvents cmdStopPolling As Button
    Friend WithEvents txtLastRead As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtMessages As TextBox
End Class

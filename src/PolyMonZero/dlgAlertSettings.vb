''' <summary>
''' Global Matrix alert configuration (homeserver URL + token).
''' </summary>
Public Class dlgAlertSettings
    Inherits Form

    Private _Manager As MonitorManager
    Private txtHomeserver As New TextBox()
    Private txtToken As New TextBox()
    Private btnTest As New Button()
    Private btnOK As New Button()
    Private btnCancel As New Button()

    Public Sub New(ByVal manager As MonitorManager)
        _Manager = manager
        InitUI()
        txtHomeserver.Text = manager.MatrixHomeserver
        txtToken.Text = manager.MatrixToken
    End Sub

    Private Sub InitUI()
        Me.Text = "Alert Settings — Matrix"
        Me.Size = New Size(480, 220)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Font = New Font("Segoe UI", 9F)

        Dim tbl As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 4,
            .Padding = New Padding(12)
        }
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        tbl.Controls.Add(MakeLbl("Homeserver URL:"), 0, 0)
        txtHomeserver.Dock = DockStyle.Fill
        txtHomeserver.PlaceholderText = "https://matrix.example.com"
        tbl.Controls.Add(txtHomeserver, 1, 0)

        tbl.Controls.Add(MakeLbl("Access Token:"), 0, 1)
        txtToken.Dock = DockStyle.Fill
        txtToken.UseSystemPasswordChar = True
        tbl.Controls.Add(txtToken, 1, 1)

        ' Show/hide token + test button row
        Dim showToken As New CheckBox() With {.Text = "Show token", .Dock = DockStyle.Fill}
        AddHandler showToken.CheckedChanged, Sub(s, e) txtToken.UseSystemPasswordChar = Not showToken.Checked
        tbl.Controls.Add(showToken, 1, 2)

        Dim btnPanel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft}
        btnCancel.Text = "Cancel" : btnCancel.DialogResult = DialogResult.Cancel : btnCancel.Size = New Size(80, 28)
        btnOK.Text = "OK" : btnOK.Size = New Size(80, 28)
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        btnTest.Text = "Send Test" : btnTest.Size = New Size(88, 28)
        AddHandler btnTest.Click, AddressOf BtnTest_Click
        btnPanel.Controls.AddRange({btnCancel, btnOK, btnTest})
        tbl.SetColumnSpan(btnPanel, 2)
        tbl.Controls.Add(btnPanel, 0, 3)

        Me.Controls.Add(tbl)
        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        _Manager.MatrixHomeserver = txtHomeserver.Text.Trim()
        _Manager.MatrixToken = txtToken.Text.Trim()
        _Manager.MarkDirty()
        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub BtnTest_Click(sender As Object, e As EventArgs)
        Dim hs = txtHomeserver.Text.Trim()
        Dim tok = txtToken.Text.Trim()
        If String.IsNullOrWhiteSpace(hs) OrElse String.IsNullOrWhiteSpace(tok) Then
            MessageBox.Show("Enter homeserver URL and token first.", "Test Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim roomId = InputBox("Enter a Matrix room ID to send the test to:", "Test Alert", "!roomid:matrix.example.com")
        If String.IsNullOrWhiteSpace(roomId) Then Return
        Dim am As New AlertManager() With {.MatrixHomeserver = hs, .MatrixToken = tok}
        am.SendAsync(roomId, "[TEST] PolyMon Zero alert test — " & DateTime.Now.ToString("HH:mm:ss"))
        MessageBox.Show("Test message sent (check the room).", "Test Alert", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Shared Function MakeLbl(text As String) As Label
        Return New Label() With {.Text = text, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleRight}
    End Function

End Class

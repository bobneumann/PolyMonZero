Imports System.IO

''' <summary>
''' Add / Edit monitor dialog — General tab (name/group/interval/script) + Alerts tab.
''' </summary>
Public Class dlgMonitorProperties
    Inherits Form

    Private _Entry As MonitorEntry
    Private _Groups As List(Of MonitorGroup)

    ' General tab
    Private txtName As New TextBox()
    Private cmbGroup As New ComboBox()
    Private nudInterval As New NumericUpDown()
    Private txtScript As New TextBox()

    ' Alert tab
    Private chkAlertEnabled As New CheckBox()
    Private txtRoomId As New TextBox()
    Private radEveryN As New RadioButton()
    Private nudEveryN As New NumericUpDown()
    Private radSpecific As New RadioButton()
    Private chkNewFailure As New CheckBox()
    Private nudRepeatFailures As New NumericUpDown()
    Private chkFailToOK As New CheckBox()
    Private chkNewWarning As New CheckBox()
    Private nudRepeatWarnings As New NumericUpDown()
    Private chkWarnToOK As New CheckBox()

    Private btnOK As New Button()
    Private btnCancel As New Button()

    Public Sub New(ByVal entry As MonitorEntry, ByVal groups As List(Of MonitorGroup))
        _Entry = entry
        _Groups = groups
        InitUI()
        PopulateFromEntry()
    End Sub

    Private Sub InitUI()
        Me.Text = "Monitor Properties"
        Me.Size = New Size(580, 560)
        Me.MinimumSize = New Size(500, 500)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Font("Segoe UI", 9F)

        Dim tabs As New TabControl() With {.Dock = DockStyle.Fill}
        Dim tabGeneral As New TabPage("General")
        Dim tabAlerts As New TabPage("Alerts")
        tabs.TabPages.AddRange({tabGeneral, tabAlerts})

        BuildGeneralTab(tabGeneral)
        BuildAlertsTab(tabAlerts)

        ' OK / Cancel buttons below tabs
        Dim btnPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Bottom,
            .FlowDirection = FlowDirection.RightToLeft,
            .Height = 44,
            .Padding = New Padding(4)
        }
        btnOK.Text = "OK" : btnOK.Size = New Size(80, 28)
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        btnCancel.Text = "Cancel" : btnCancel.DialogResult = DialogResult.Cancel : btnCancel.Size = New Size(80, 28)
        btnPanel.Controls.AddRange({btnCancel, btnOK})

        Me.Controls.Add(tabs)
        Me.Controls.Add(btnPanel)
        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub

    Private Sub BuildGeneralTab(tab As TabPage)
        Dim tbl As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 6,
            .Padding = New Padding(10)
        }
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 26))
        tbl.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        tbl.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))

        tbl.Controls.Add(MakeLbl("Name:"), 0, 0)
        txtName.Dock = DockStyle.Fill
        tbl.Controls.Add(txtName, 1, 0)

        tbl.Controls.Add(MakeLbl("Group:"), 0, 1)
        cmbGroup.Dock = DockStyle.Fill : cmbGroup.DropDownStyle = ComboBoxStyle.DropDown
        tbl.Controls.Add(cmbGroup, 1, 1)

        tbl.Controls.Add(MakeLbl("Interval (sec):"), 0, 2)
        nudInterval.Dock = DockStyle.Fill : nudInterval.Minimum = 5 : nudInterval.Maximum = 86400 : nudInterval.Value = 60
        tbl.Controls.Add(nudInterval, 1, 2)

        tbl.Controls.Add(MakeLbl("PS Script:"), 0, 3)
        tbl.Controls.Add(New Label() With {.Text = "(paste or load from file)", .ForeColor = Color.Gray, .Dock = DockStyle.Fill}, 1, 3)

        txtScript.Multiline = True : txtScript.ScrollBars = ScrollBars.Both : txtScript.WordWrap = False
        txtScript.Dock = DockStyle.Fill : txtScript.Font = New Font("Consolas", 9F) : txtScript.AcceptsTab = True
        tbl.SetColumnSpan(txtScript, 2)
        tbl.Controls.Add(txtScript, 0, 4)

        Dim btnBrowse As New Button() With {.Text = "Load from .ps1 file...", .AutoSize = True}
        AddHandler btnBrowse.Click, AddressOf BrowseScript
        Dim browseRow As New FlowLayoutPanel() With {.Dock = DockStyle.Fill}
        browseRow.Controls.Add(btnBrowse)
        tbl.SetColumnSpan(browseRow, 2)
        tbl.Controls.Add(browseRow, 0, 5)

        tab.Controls.Add(tbl)
    End Sub

    Private Sub BuildAlertsTab(tab As TabPage)
        Dim pnl As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(10)}

        Dim y = 10

        ' Master enable
        chkAlertEnabled.Text = "Enable alerts for this monitor"
        chkAlertEnabled.Location = New Point(10, y) : chkAlertEnabled.AutoSize = True
        AddHandler chkAlertEnabled.CheckedChanged, AddressOf UpdateAlertControlStates
        pnl.Controls.Add(chkAlertEnabled)
        y += 30

        ' Room ID
        pnl.Controls.Add(New Label() With {.Text = "Matrix Room ID:", .Location = New Point(10, y + 3), .AutoSize = True})
        txtRoomId.Location = New Point(120, y) : txtRoomId.Width = 360 : txtRoomId.PlaceholderText = "!roomid:matrix.example.com"
        pnl.Controls.Add(txtRoomId)
        y += 34

        ' Separator
        Dim sep As New Label() With {.BorderStyle = BorderStyle.Fixed3D, .Height = 2, .Location = New Point(10, y), .Width = 500}
        pnl.Controls.Add(sep)
        y += 10

        ' Mode: Every N events
        radEveryN.Text = "Alert every" : radEveryN.Location = New Point(10, y) : radEveryN.AutoSize = True
        AddHandler radEveryN.CheckedChanged, AddressOf UpdateAlertControlStates
        pnl.Controls.Add(radEveryN)
        nudEveryN.Location = New Point(100, y - 2) : nudEveryN.Width = 60 : nudEveryN.Minimum = 1 : nudEveryN.Maximum = 9999 : nudEveryN.Value = 1
        pnl.Controls.Add(nudEveryN)
        pnl.Controls.Add(New Label() With {.Text = "poll cycles (regardless of status)", .Location = New Point(168, y + 3), .AutoSize = True})
        y += 30

        ' Mode: Specific conditions
        radSpecific.Text = "Alert on specific conditions:" : radSpecific.Location = New Point(10, y) : radSpecific.AutoSize = True
        AddHandler radSpecific.CheckedChanged, AddressOf UpdateAlertControlStates
        pnl.Controls.Add(radSpecific)
        y += 28

        Dim ix = 28  ' indent for condition controls

        ' New failure
        chkNewFailure.Text = "New failure (OK/Warn → Fail)" : chkNewFailure.Location = New Point(ix, y) : chkNewFailure.AutoSize = True
        pnl.Controls.Add(chkNewFailure)
        y += 26

        ' Repeat failures
        pnl.Controls.Add(New Label() With {.Text = "Repeat every", .Location = New Point(ix, y + 3), .AutoSize = True})
        nudRepeatFailures.Location = New Point(ix + 90, y) : nudRepeatFailures.Width = 60 : nudRepeatFailures.Minimum = 0 : nudRepeatFailures.Maximum = 9999 : nudRepeatFailures.Value = 0
        pnl.Controls.Add(nudRepeatFailures)
        pnl.Controls.Add(New Label() With {.Text = "consecutive failures (0 = no repeat)", .Location = New Point(ix + 158, y + 3), .AutoSize = True})
        y += 26

        ' Fail to OK
        chkFailToOK.Text = "Failure → OK recovery" : chkFailToOK.Location = New Point(ix, y) : chkFailToOK.AutoSize = True
        pnl.Controls.Add(chkFailToOK)
        y += 30

        ' Separator
        Dim sep2 As New Label() With {.BorderStyle = BorderStyle.Fixed3D, .Height = 2, .Location = New Point(ix, y), .Width = 460}
        pnl.Controls.Add(sep2)
        y += 10

        ' New warning
        chkNewWarning.Text = "New warning (OK → Warn)" : chkNewWarning.Location = New Point(ix, y) : chkNewWarning.AutoSize = True
        pnl.Controls.Add(chkNewWarning)
        y += 26

        ' Repeat warnings
        pnl.Controls.Add(New Label() With {.Text = "Repeat every", .Location = New Point(ix, y + 3), .AutoSize = True})
        nudRepeatWarnings.Location = New Point(ix + 90, y) : nudRepeatWarnings.Width = 60 : nudRepeatWarnings.Minimum = 0 : nudRepeatWarnings.Maximum = 9999 : nudRepeatWarnings.Value = 0
        pnl.Controls.Add(nudRepeatWarnings)
        pnl.Controls.Add(New Label() With {.Text = "consecutive warnings (0 = no repeat)", .Location = New Point(ix + 158, y + 3), .AutoSize = True})
        y += 26

        ' Warn to OK
        chkWarnToOK.Text = "Warning → OK recovery" : chkWarnToOK.Location = New Point(ix, y) : chkWarnToOK.AutoSize = True
        pnl.Controls.Add(chkWarnToOK)

        tab.Controls.Add(pnl)
    End Sub

    Private Sub PopulateFromEntry()
        txtName.Text = _Entry.MonitorName
        nudInterval.Value = Math.Max(nudInterval.Minimum, Math.Min(nudInterval.Maximum, _Entry.PollingIntervalSec))
        txtScript.Text = _Entry.Script

        cmbGroup.Items.Clear()
        For Each g In _Groups
            cmbGroup.Items.Add(g.Name)
        Next
        If Not cmbGroup.Items.Contains("Default") Then cmbGroup.Items.Add("Default")
        cmbGroup.Text = _Entry.GroupName

        ' Alerts tab
        chkAlertEnabled.Checked = _Entry.AlertEnabled
        txtRoomId.Text = _Entry.AlertRoomId
        If _Entry.AlertEveryNEvents Then
            radEveryN.Checked = True
            nudEveryN.Value = Math.Max(1, _Entry.AlertEveryNEventsCount)
        Else
            radSpecific.Checked = True
        End If
        chkNewFailure.Checked = _Entry.AlertOnNewFailure
        nudRepeatFailures.Value = Math.Max(0, _Entry.AlertRepeatEveryNFailures)
        chkFailToOK.Checked = _Entry.AlertOnFailToOK
        chkNewWarning.Checked = _Entry.AlertOnNewWarning
        nudRepeatWarnings.Value = Math.Max(0, _Entry.AlertRepeatEveryNWarnings)
        chkWarnToOK.Checked = _Entry.AlertOnWarnToOK

        UpdateAlertControlStates(Nothing, EventArgs.Empty)
    End Sub

    Private Sub UpdateAlertControlStates(sender As Object, e As EventArgs)
        Dim enabled = chkAlertEnabled.Checked
        txtRoomId.Enabled = enabled
        radEveryN.Enabled = enabled
        radSpecific.Enabled = enabled
        nudEveryN.Enabled = enabled AndAlso radEveryN.Checked
        Dim spec = enabled AndAlso radSpecific.Checked
        chkNewFailure.Enabled = spec
        nudRepeatFailures.Enabled = spec
        chkFailToOK.Enabled = spec
        chkNewWarning.Enabled = spec
        nudRepeatWarnings.Enabled = spec
        chkWarnToOK.Enabled = spec
    End Sub

    Private Sub BrowseScript(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
            If ofd.ShowDialog() = DialogResult.OK Then
                txtScript.Text = File.ReadAllText(ofd.FileName)
            End If
        End Using
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtName.Text) Then
            MessageBox.Show("Please enter a monitor name.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _Entry.MonitorName = txtName.Text.Trim()
        _Entry.GroupName = If(String.IsNullOrWhiteSpace(cmbGroup.Text), "Default", cmbGroup.Text.Trim())
        _Entry.PollingIntervalSec = CInt(nudInterval.Value)
        _Entry.Script = txtScript.Text

        _Entry.AlertEnabled = chkAlertEnabled.Checked
        _Entry.AlertRoomId = txtRoomId.Text.Trim()
        _Entry.AlertEveryNEvents = radEveryN.Checked
        _Entry.AlertEveryNEventsCount = CInt(nudEveryN.Value)
        _Entry.AlertOnNewFailure = chkNewFailure.Checked
        _Entry.AlertRepeatEveryNFailures = CInt(nudRepeatFailures.Value)
        _Entry.AlertOnFailToOK = chkFailToOK.Checked
        _Entry.AlertOnNewWarning = chkNewWarning.Checked
        _Entry.AlertRepeatEveryNWarnings = CInt(nudRepeatWarnings.Value)
        _Entry.AlertOnWarnToOK = chkWarnToOK.Checked

        Me.DialogResult = DialogResult.OK
    End Sub

    Private Shared Function MakeLbl(text As String) As Label
        Return New Label() With {.Text = text, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleRight}
    End Function

End Class

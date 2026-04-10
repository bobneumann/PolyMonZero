Imports System.IO
Imports System.Linq

Public Class frmMain
    Inherits Form

    Private _Manager As New MonitorManager()
    Private _AlertMgr As New AlertManager()
    Private _CardPanel As FlowLayoutPanel
    Private _GaugeDashboards As New Dictionary(Of MonitorEntry, frmGaugeDashboard)()
    Private _StatusBar As StatusStrip
    Private _lblCounts As ToolStripStatusLabel
    Private _lblFile As ToolStripStatusLabel
    Private _PauseAllBtn As ToolStripButton
    Private _RibbonPanel As pnlGroupedRibbon
    Private _btnViewTiles As ToolStripButton
    Private _btnViewGroups As ToolStripButton
    Private _RefreshTimer As New Timer() With {.Interval = 1000}

    Public Sub New()
        InitUI()
        AddHandler _Manager.MonitorUpdated, AddressOf OnMonitorUpdated
        AddHandler _Manager.DirtyChanged, AddressOf OnDirtyChanged
        AddHandler _Manager.AlertTriggered, AddressOf OnAlertTriggered
        AddHandler _RefreshTimer.Tick, AddressOf RefreshTick
        _RefreshTimer.Start()

        ' Auto-open last session
        Dim lastFile = MonitorManager.GetLastFilePath()
        If Not String.IsNullOrEmpty(lastFile) AndAlso File.Exists(lastFile) Then
            Try
                _Manager.LoadFromFile(lastFile)
                SyncAlertManager()
                RebuildCards()
            Catch
            End Try
        End If

        UpdateTitle()
        UpdateStatusBar()
    End Sub

#Region "UI Construction"
    Private Sub InitUI()
        Me.Text = "PolyMon Zero"
        Me.Size = New Size(900, 650)
        Me.MinimumSize = New Size(600, 400)
        Me.Font = New Font("Segoe UI", 9F)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Menu
        Dim menu As New MenuStrip() With {.Font = New Font("Segoe UI", 11F)}
        Dim fileMenu As New ToolStripMenuItem("&File")
        Dim monMenu As New ToolStripMenuItem("&Monitors")
        Dim settingsMenu As New ToolStripMenuItem("&Settings")
        Dim helpMenu As New ToolStripMenuItem("&Help")
        menu.Items.AddRange({fileMenu, monMenu, settingsMenu, helpMenu})

        AddMenuItem(fileMenu, "&New", Sub() FileNew())
        AddMenuItem(fileMenu, "&Open...", Sub() FileOpen())
        fileMenu.DropDownItems.Add(New ToolStripSeparator())
        AddMenuItem(fileMenu, "&Save", Sub() FileSave())
        AddMenuItem(fileMenu, "Save &As...", Sub() FileSaveAs())
        fileMenu.DropDownItems.Add(New ToolStripSeparator())
        AddMenuItem(fileMenu, "E&xit", Sub() Me.Close())

        AddMenuItem(monMenu, "&Add Monitor...", Sub() AddMonitor())
        AddMenuItem(monMenu, "Edit &Groups...", Sub() ShowManageGroups())
        monMenu.DropDownItems.Add(New ToolStripSeparator())
        AddMenuItem(monMenu, "&Pause All", Sub() PauseAll())
        AddMenuItem(monMenu, "&Resume All", Sub() ResumeAll())

        AddMenuItem(settingsMenu, "&Alert Settings (Matrix)...", Sub() ShowAlertSettings())

        AddMenuItem(helpMenu, "&About", Sub() ShowAbout())

        ' Add status bar and card panel first (bottom and fill),
        ' then toolbar and menu last — WinForms docks in reverse add order for Top/Bottom.
        Me.MainMenuStrip = menu

        ' Status bar (Bottom) — first so it docks to bottom before Fill is calculated
        _StatusBar = New StatusStrip()
        _lblFile = New ToolStripStatusLabel("No file") With {.Spring = True, .TextAlign = ContentAlignment.MiddleLeft}
        _lblCounts = New ToolStripStatusLabel("") With {.TextAlign = ContentAlignment.MiddleRight}
        _StatusBar.Items.AddRange({_lblFile, _lblCounts})
        Me.Controls.Add(_StatusBar)

        ' Content frame — dark border creates a sunken/embedded look for the content area
        Dim contentFrame As New Panel() With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(4),
            .BackColor = Color.FromArgb(110, 115, 128)
        }

        ' Card panel (Fill inside contentFrame)
        _CardPanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(8),
            .BackColor = Color.WhiteSmoke
        }
        contentFrame.Controls.Add(_CardPanel)

        ' Groups ribbon panel (Fill inside contentFrame, hidden until user switches view)
        _RibbonPanel = New pnlGroupedRibbon() With {.Dock = DockStyle.Fill, .Visible = False}
        AddHandler _RibbonPanel.EditRequested, AddressOf OnEditRequested
        AddHandler _RibbonPanel.DeleteRequested, AddressOf OnDeleteRequested
        AddHandler _RibbonPanel.RunNowRequested, AddressOf OnRunNowRequested
        AddHandler _RibbonPanel.GaugeViewRequested, AddressOf OnGaugeViewRequested
        contentFrame.Controls.Add(_RibbonPanel)

        Me.Controls.Add(contentFrame)

        ' Toolbar (Top) — added after Fill so it sits above it
        Dim toolbar As New ToolStrip() With {.GripStyle = ToolStripGripStyle.Hidden, .Dock = DockStyle.Top, .Font = New Font("Segoe UI", 11F)}
        Dim btnAdd As New ToolStripButton("+ Add Monitor") With {.DisplayStyle = ToolStripItemDisplayStyle.Text}
        AddHandler btnAdd.Click, Sub(s, e) AddMonitor()
        _PauseAllBtn = New ToolStripButton("Pause All") With {.DisplayStyle = ToolStripItemDisplayStyle.Text}
        AddHandler _PauseAllBtn.Click, Sub(s, e) TogglePauseAll()
        Dim btnGroups As New ToolStripButton("Groups...") With {.DisplayStyle = ToolStripItemDisplayStyle.Text}
        AddHandler btnGroups.Click, Sub(s, e) ShowManageGroups()
        _btnViewTiles = New ToolStripButton("Tiles") With {.DisplayStyle = ToolStripItemDisplayStyle.Text, .Checked = True, .CheckOnClick = False}
        AddHandler _btnViewTiles.Click, Sub(s, e) SwitchView("Tiles")
        _btnViewGroups = New ToolStripButton("Groups") With {.DisplayStyle = ToolStripItemDisplayStyle.Text, .Checked = False, .CheckOnClick = False}
        AddHandler _btnViewGroups.Click, Sub(s, e) SwitchView("Groups")
        toolbar.Items.AddRange({btnAdd, btnGroups, New ToolStripSeparator(), _PauseAllBtn,
            New ToolStripSeparator(), _btnViewTiles, _btnViewGroups})
        Me.Controls.Add(toolbar)

        ' Menu (Top) — added last so it docks to the very top
        Me.Controls.Add(menu)
    End Sub

    Private Shared Sub AddMenuItem(parent As ToolStripMenuItem, text As String, action As Action)
        Dim item As New ToolStripMenuItem(text)
        AddHandler item.Click, Sub(s, e) action()
        parent.DropDownItems.Add(item)
    End Sub
#End Region

#Region "Card Management"
    Private Sub RebuildCards()
        _CardPanel.SuspendLayout()
        ' Remove cards whose entry was deleted
        For Each card In _CardPanel.Controls.OfType(Of ucStatusCard)().ToList()
            If Not _Manager.Monitors.Contains(card.Entry) Then
                _CardPanel.Controls.Remove(card)
            End If
        Next
        ' Add cards for new entries
        For Each entry In _Manager.Monitors
            If Not _CardPanel.Controls.OfType(Of ucStatusCard)().Any(Function(c) c.Entry Is entry) Then
                Dim card As New ucStatusCard(entry)
                AddHandler card.EditRequested, AddressOf OnEditRequested
                AddHandler card.DeleteRequested, AddressOf OnDeleteRequested
                AddHandler card.RunNowRequested, AddressOf OnRunNowRequested
                AddHandler card.GaugeViewRequested, AddressOf OnGaugeViewRequested
                _CardPanel.Controls.Add(card)
            End If
        Next
        _CardPanel.ResumeLayout()
        If _RibbonPanel IsNot Nothing AndAlso _RibbonPanel.Visible Then
            _RibbonPanel.Rebuild(_Manager.Monitors, _Manager.Groups)
        End If
    End Sub

    Private Sub InvalidateAllCards()
        For Each card In _CardPanel.Controls.OfType(Of ucStatusCard)()
            card.Invalidate()
        Next
    End Sub
#End Region

#Region "Monitor Events"
    Private Sub OnMonitorUpdated(ByVal entry As MonitorEntry)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() OnMonitorUpdated(entry))
            Return
        End If
        ' Find the card and repaint it
        For Each card In _CardPanel.Controls.OfType(Of ucStatusCard)()
            If card.Entry Is entry Then
                card.Invalidate()
                Exit For
            End If
        Next
        If _RibbonPanel.Visible Then _RibbonPanel.InvalidateRow(entry)
        UpdateStatusBar()
    End Sub

    Private Sub OnDirtyChanged(ByVal isDirty As Boolean)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() OnDirtyChanged(isDirty))
            Return
        End If
        UpdateTitle()
    End Sub

    ' Tick every second to keep "Last: HH:mm:ss" fresh on all cards/rows
    Private Sub RefreshTick(sender As Object, e As EventArgs)
        If _CardPanel.Visible Then InvalidateAllCards()
        If _RibbonPanel.Visible Then _RibbonPanel.InvalidateAll()
        UpdateStatusBar()
    End Sub
#End Region

#Region "Card Handlers"
    Private Sub OnEditRequested(ByVal entry As MonitorEntry)
        entry.StopMonitor()
        Using dlg As New dlgMonitorProperties(entry, _Manager.Groups)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                _Manager.EnsureGroup(entry.GroupName)
                entry.StartMonitor()
                _Manager.MarkDirty()
            Else
                entry.StartMonitor()
            End If
        End Using
        UpdateTitle()
        UpdateStatusBar()
    End Sub

    Private Sub OnDeleteRequested(ByVal entry As MonitorEntry)
        Dim r = MessageBox.Show($"Delete monitor ""{entry.MonitorName}""?", "Confirm Delete",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If r = DialogResult.Yes Then
            _Manager.RemoveMonitor(entry)
            RebuildCards()
            UpdateStatusBar()
        End If
    End Sub

    Private Sub OnRunNowRequested(ByVal entry As MonitorEntry)
        entry.RunNow()
    End Sub

    Private Sub OnGaugeViewRequested(ByVal entry As MonitorEntry)
        ' Bring existing dashboard to front if already open; otherwise open a new one
        Dim existing As frmGaugeDashboard = Nothing
        If _GaugeDashboards.TryGetValue(entry, existing) AndAlso Not existing.IsDisposed Then
            existing.BringToFront()
            Return
        End If
        Dim dlg As New frmGaugeDashboard(entry)
        _GaugeDashboards(entry) = dlg
        AddHandler dlg.FormClosed, Sub(s, e) _GaugeDashboards.Remove(entry)
        dlg.Show(Me)
    End Sub
#End Region

#Region "Monitor Actions"
    Private Sub ShowManageGroups()
        Using dlg As New dlgManageGroups(_Manager)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                RebuildCards()
                UpdateTitle()
                UpdateStatusBar()
            End If
        End Using
    End Sub

    Private Sub AddMonitor()
        Dim entry = _Manager.AddMonitor()
        RebuildCards()
        Using dlg As New dlgMonitorProperties(entry, _Manager.Groups)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                _Manager.EnsureGroup(entry.GroupName)
                entry.StartMonitor()
            Else
                _Manager.RemoveMonitor(entry)
                RebuildCards()
            End If
        End Using
        UpdateTitle()
        UpdateStatusBar()
    End Sub

    Private _AllPaused As Boolean = False
    Private Sub TogglePauseAll()
        If _AllPaused Then
            ResumeAll()
        Else
            PauseAll()
        End If
    End Sub

    Private Sub PauseAll()
        _Manager.PauseAll()
        _AllPaused = True
        _PauseAllBtn.Text = "Resume All"
    End Sub

    Private Sub ResumeAll()
        _Manager.ResumeAll()
        _AllPaused = False
        _PauseAllBtn.Text = "Pause All"
    End Sub
#End Region

#Region "File Operations"
    Private Sub FileNew()
        If Not ConfirmDiscard() Then Return
        _Manager.NewFile()
        RebuildCards()
        UpdateTitle()
        UpdateStatusBar()
    End Sub

    Private Sub FileOpen()
        If Not ConfirmDiscard() Then Return
        Using ofd As New OpenFileDialog()
            ofd.Filter = "PolyMon Zero Files (*.pmz)|*.pmz|All Files (*.*)|*.*"
            ofd.Title = "Open Monitor File"
            If ofd.ShowDialog(Me) = DialogResult.OK Then
                Try
                    _Manager.LoadFromFile(ofd.FileName)
                    SyncAlertManager()
                    RebuildCards()
                    UpdateTitle()
                    UpdateStatusBar()
                Catch ex As Exception
                    MessageBox.Show("Failed to open file:" & Environment.NewLine & ex.Message,
                        "Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub FileSave()
        If String.IsNullOrEmpty(_Manager.CurrentFilePath) Then
            FileSaveAs()
        Else
            Try
                _Manager.SaveToFile(_Manager.CurrentFilePath)
                UpdateTitle()
            Catch ex As Exception
                MessageBox.Show("Save failed: " & ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub FileSaveAs()
        Using sfd As New SaveFileDialog()
            sfd.Filter = "PolyMon Zero Files (*.pmz)|*.pmz|All Files (*.*)|*.*"
            sfd.Title = "Save Monitor File"
            sfd.DefaultExt = "pmz"
            If Not String.IsNullOrEmpty(_Manager.CurrentFilePath) Then
                sfd.FileName = Path.GetFileName(_Manager.CurrentFilePath)
                sfd.InitialDirectory = Path.GetDirectoryName(_Manager.CurrentFilePath)
            End If
            If sfd.ShowDialog(Me) = DialogResult.OK Then
                Try
                    _Manager.SaveToFile(sfd.FileName)
                    UpdateTitle()
                Catch ex As Exception
                    MessageBox.Show("Save failed: " & ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Function ConfirmDiscard() As Boolean
        If Not _Manager.IsDirty Then Return True
        Dim r = MessageBox.Show("Save changes before continuing?", "Unsaved Changes",
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If r = DialogResult.Yes Then FileSave() : Return True
        If r = DialogResult.No Then Return True
        Return False
    End Function
#End Region

#Region "Alerts"
    Private Sub OnAlertTriggered(ByVal entry As MonitorEntry, ByVal message As String)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() OnAlertTriggered(entry, message))
            Return
        End If
        SyncAlertManager()
        _AlertMgr.SendAsync(entry.AlertRoomId, message)
    End Sub

    Private Sub SyncAlertManager()
        _AlertMgr.MatrixHomeserver = _Manager.MatrixHomeserver
        _AlertMgr.MatrixToken = _Manager.MatrixToken
    End Sub

    Private Sub ShowAlertSettings()
        Using dlg As New dlgAlertSettings(_Manager)
            dlg.ShowDialog(Me)
        End Using
        UpdateTitle()
    End Sub
#End Region

#Region "UI Helpers"
    Private Sub SwitchView(view As String)
        _CardPanel.Visible = (view = "Tiles")
        _RibbonPanel.Visible = (view = "Groups")
        _btnViewTiles.Checked = (view = "Tiles")
        _btnViewGroups.Checked = (view = "Groups")
        If view = "Groups" Then _RibbonPanel.Rebuild(_Manager.Monitors, _Manager.Groups)
    End Sub

    Private Sub UpdateTitle()
        Dim name = If(String.IsNullOrEmpty(_Manager.CurrentFilePath),
            "Untitled", Path.GetFileName(_Manager.CurrentFilePath))
        Me.Text = $"PolyMon Zero — {name}{If(_Manager.IsDirty, " *", "")}"
    End Sub

    Private Sub UpdateStatusBar()
        Dim total = _Manager.Monitors.Count
        Dim warns = _Manager.Monitors.Where(Function(m) m.LastResult IsNot Nothing AndAlso m.LastResult.Status = MonitorStatus.Warning).Count()
        Dim fails = _Manager.Monitors.Where(Function(m) m.LastResult IsNot Nothing AndAlso m.LastResult.Status = MonitorStatus.Fail).Count()
        _lblCounts.Text = $"Monitors: {total}   Warn: {warns}   Fail: {fails}"
        _lblFile.Text = If(String.IsNullOrEmpty(_Manager.CurrentFilePath), "No file",
            _Manager.CurrentFilePath)
    End Sub

    Private Sub ShowAbout()
        MessageBox.Show("PolyMon Zero" & Environment.NewLine &
            "Lightweight portable PowerShell monitor dashboard." & Environment.NewLine & Environment.NewLine &
            "Scripts inject: $Counters, $errlvl (0/1/2), $messages, $Counter",
            "About PolyMon Zero", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
#End Region

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If Not ConfirmDiscard() Then
            e.Cancel = True
            Return
        End If
        _RefreshTimer.Stop()
        MyBase.OnFormClosing(e)
    End Sub

End Class

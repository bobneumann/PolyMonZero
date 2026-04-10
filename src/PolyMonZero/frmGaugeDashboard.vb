''' <summary>
''' Modeless popup showing one skeuomorphic gauge (or numeric tile) per counter
''' for a single monitor entry. Subscribes to ResultReady and auto-refreshes.
''' Bottom toolbar: Run Now, interval control, and gauge-size slider.
''' Layout algorithm fills each row evenly and minimises empty slots in the last row.
''' </summary>
Public Class frmGaugeDashboard
    Inherits Form

    Private ReadOnly _Entry As MonitorEntry
    Private _GaugePanel As FlowLayoutPanel
    Private _lblStatus As Label
    Private _lblSizeVal As Label
    Private _nudInterval As NumericUpDown
    Private _trkSize As TrackBar
    Private _LastGaugeCount As Integer = -1
    Private _GaugeBaseSize As Integer = 175   ' controlled by size slider

    Public Sub New(ByVal entry As MonitorEntry)
        _Entry = entry
        InitUI()
        RefreshGauges()
        AddHandler _Entry.ResultReady, AddressOf OnResultReady
    End Sub

#Region "UI Construction"
    Private Sub InitUI()
        Me.Text = _Entry.MonitorName & " — Gauge View"
        Me.Size = New Size(980, 640)
        Me.MinimumSize = New Size(400, 360)
        Me.BackColor = Color.FromArgb(22, 22, 28)
        Me.Font = New Font("Segoe UI", 9F)
        Me.StartPosition = FormStartPosition.CenterParent

        ' ── Title strip (Top) ────────────────────────────────────────────────
        Dim titlePanel As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 38,
            .BackColor = Color.FromArgb(30, 32, 42)
        }

        _lblStatus = New Label() With {
            .AutoSize = True,
            .Font = New Font("Segoe UI", 16F),
            .ForeColor = Color.FromArgb(150, 150, 150),
            .Text = "●",
            .Location = New Point(8, 8)
        }

        Dim lblName As New Label() With {
            .AutoSize = False,
            .Font = New Font("Segoe UI", 12F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(210, 213, 225),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Text = _Entry.MonitorName,
            .Location = New Point(34, 0),
            .Size = New Size(500, 38)
        }

        titlePanel.Controls.Add(lblName)
        titlePanel.Controls.Add(_lblStatus)

        ' ── Toolbar (Bottom) ─────────────────────────────────────────────────
        Dim toolPanel As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 40,
            .BackColor = Color.FromArgb(26, 28, 36)
        }

        ' Run Now button
        Dim btnRun As New Button() With {
            .Text = "▶  Run Now",
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(180, 210, 255),
            .BackColor = Color.FromArgb(40, 80, 130),
            .Size = New Size(100, 26),
            .Location = New Point(8, 7)
        }
        btnRun.FlatAppearance.BorderColor = Color.FromArgb(60, 110, 180)
        AddHandler btnRun.Click, Sub(s, e) _Entry.RunNow()

        ' Separator
        Dim sep1 As New Label() With {
            .Text = "|", .ForeColor = Color.FromArgb(70, 73, 85),
            .AutoSize = True, .Location = New Point(116, 11)
        }

        ' Interval control
        Dim lblEvery As New Label() With {
            .Text = "Every:", .ForeColor = Color.FromArgb(155, 158, 170),
            .AutoSize = True, .Location = New Point(128, 13)
        }

        _nudInterval = New NumericUpDown() With {
            .Minimum = 5, .Maximum = 3600,
            .Value = Math.Max(5D, Math.Min(3600D, _Entry.PollingIntervalSec)),
            .Width = 70, .Location = New Point(178, 8),
            .BackColor = Color.FromArgb(36, 38, 46),
            .ForeColor = Color.FromArgb(200, 203, 215),
            .TextAlign = HorizontalAlignment.Center
        }
        AddHandler _nudInterval.ValueChanged, AddressOf OnIntervalChanged

        Dim lblSec As New Label() With {
            .Text = "sec", .ForeColor = Color.FromArgb(155, 158, 170),
            .AutoSize = True, .Location = New Point(253, 13)
        }

        ' Separator
        Dim sep2 As New Label() With {
            .Text = "|", .ForeColor = Color.FromArgb(70, 73, 85),
            .AutoSize = True, .Location = New Point(280, 11)
        }

        ' Gauge size slider
        Dim lblSize As New Label() With {
            .Text = "Gauge size:", .ForeColor = Color.FromArgb(155, 158, 170),
            .AutoSize = True, .Location = New Point(292, 13)
        }

        _trkSize = New TrackBar() With {
            .Minimum = 100, .Maximum = 350, .Value = _GaugeBaseSize,
            .SmallChange = 25, .LargeChange = 50,
            .TickStyle = TickStyle.None,
            .Width = 150, .Height = 26,
            .Location = New Point(375, 7)
        }
        AddHandler _trkSize.Scroll, AddressOf OnSizeSliderScroll

        _lblSizeVal = New Label() With {
            .Text = $"{_GaugeBaseSize}px",
            .ForeColor = Color.FromArgb(155, 158, 170),
            .AutoSize = True,
            .Location = New Point(530, 13)
        }

        toolPanel.Controls.AddRange(New Control() {
            btnRun, sep1, lblEvery, _nudInterval, lblSec,
            sep2, lblSize, _trkSize, _lblSizeVal
        })

        ' ── Gauge flow panel (Fill) ───────────────────────────────────────────
        ' Must be added to Me.Controls BEFORE the Top/Bottom docked panels so
        ' WinForms Fill docking works correctly.
        _GaugePanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(8),
            .BackColor = Color.FromArgb(22, 22, 28)
        }

        Me.Controls.Add(_GaugePanel)
        Me.Controls.Add(toolPanel)
        Me.Controls.Add(titlePanel)
    End Sub
#End Region

#Region "Refresh"
    Private Sub OnResultReady(ByVal entry As MonitorEntry)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() RefreshGauges())
            Return
        End If
        RefreshGauges()
    End Sub

    Public Sub RefreshGauges()
        Dim result = _Entry.LastResult

        ' Update status dot
        Dim status = If(result IsNot Nothing, result.Status, MonitorStatus.Unknown)
        _lblStatus.ForeColor = StatusColors.ForStatus(status)

        If result Is Nothing OrElse result.Counters.Count = 0 Then
            _GaugePanel.Controls.Clear()
            _GaugePanel.Controls.Add(New Label() With {
                .Text = "No counter data yet — waiting for first run.",
                .ForeColor = Color.FromArgb(140, 143, 155),
                .Font = New Font("Segoe UI", 10F),
                .AutoSize = True,
                .Margin = New Padding(16)
            })
            Return
        End If

        Dim counters = result.Counters

        ' Remove surplus gauges
        While _GaugePanel.Controls.Count > counters.Count
            _GaugePanel.Controls.RemoveAt(_GaugePanel.Controls.Count - 1)
        End While

        ' Add missing gauges
        While _GaugePanel.Controls.Count < counters.Count
            _GaugePanel.Controls.Add(New ucGauge())
        End While

        ' Update values
        For i = 0 To counters.Count - 1
            Dim gauge = DirectCast(_GaugePanel.Controls(i), ucGauge)
            Dim cv = counters(i)

            gauge.CounterName = cv.Name
            gauge.Value = cv.Value

            If cv.HasRange Then
                gauge.MinValue = cv.MinValue
                gauge.MaxValue = cv.MaxValue
                gauge.ShowAsGauge = True
            ElseIf ucGauge.IsGaugeType(cv.Name) Then
                gauge.MinValue = 0
                gauge.MaxValue = 100
                gauge.ShowAsGauge = True
            Else
                gauge.MinValue = 0
                gauge.MaxValue = 100
                gauge.ShowAsGauge = False
            End If

            gauge.Invalidate()
        Next

        ' Auto-size the form only the first time we know the gauge count
        If counters.Count <> _LastGaugeCount Then
            _LastGaugeCount = counters.Count
            AutoSizeToGauges(counters.Count)
        End If

        RelayoutGauges()
    End Sub

    ''' <summary>Set a sensible initial form size when the gauge count is first known.</summary>
    Private Sub AutoSizeToGauges(gaugeCount As Integer)
        Const MARGIN As Integer = 8
        Const PANEL_PAD As Integer = 16
        Const TITLE_H As Integer = 38
        Const TOOL_H As Integer = 40
        Const CHROME As Integer = 20

        Dim cellW = _GaugeBaseSize + MARGIN * 2
        Dim cellH = CInt(_GaugeBaseSize * 195.0 / 175.0) + MARGIN * 2
        Dim cols = Math.Max(1, Math.Min(gaugeCount, 5))
        Dim rows = CInt(Math.Ceiling(gaugeCount / cols))

        Dim idealW = cols * cellW + PANEL_PAD + CHROME
        Dim idealH = rows * cellH + PANEL_PAD + TITLE_H + TOOL_H + CHROME

        Dim screenArea = Screen.FromControl(Me).WorkingArea
        idealW = Math.Min(idealW, screenArea.Width - 40)
        idealH = Math.Min(idealH, screenArea.Height - 60)

        Me.Size = New Size(idealW, idealH)
    End Sub
#End Region

#Region "Layout"
    ''' <summary>
    ''' Choose the column count that best fills each row with minimal wasted slots.
    ''' Searches from maxCols downward; the first col-count that perfectly divides
    ''' gaugeCount wins. Otherwise the count with the fewest empty slots is used.
    ''' </summary>
    Private Function ComputeBestCols(gaugeCount As Integer, availableWidth As Integer) As Integer
        Dim cellW = _GaugeBaseSize + 16           ' gauge width + Margin(8)*2
        Dim maxCols = Math.Max(1, Math.Min(8, availableWidth \ cellW))
        Dim minCols = Math.Max(1, maxCols \ 2)    ' don't collapse below half

        Dim bestCols = maxCols
        Dim bestWaste = gaugeCount                 ' sentinel — any real waste beats this

        For cols = maxCols To minCols Step -1
            Dim waste = (cols - (gaugeCount Mod cols)) Mod cols
            If waste < bestWaste Then
                bestWaste = waste
                bestCols = cols
            End If
            If waste = 0 Then Exit For             ' perfect divisor found; largest possible
        Next
        Return bestCols
    End Function

    ''' <summary>
    ''' Size all gauge controls to best fill the available panel area.
    ''' • Maximized: auto-compute gauge size from window dimensions (ignores slider).
    ''' • Normal:    slider sets base size; even-row algo picks columns.
    ''' </summary>
    Private Sub RelayoutGauges()
        If _LastGaugeCount <= 0 Then Return

        Const PANEL_PAD As Integer = 16
        Const GAUGE_MARGIN As Integer = 16   ' Margin(8) × 2 sides

        ' Get usable panel dimensions (fall back to form if panel not yet laid out).
        Dim panelW = _GaugePanel.ClientSize.Width
        Dim panelH = _GaugePanel.ClientSize.Height
        If panelW < 80 Then panelW = Me.ClientSize.Width
        If panelH < 80 Then panelH = Me.ClientSize.Height - 78  ' approx title+toolbar

        Dim gaugeW As Integer
        Dim bestCols As Integer

        If Me.WindowState = FormWindowState.Maximized Then
            ' ── Maximized: fill the screen, leaving a comfortable margin on all 4 sides ──
            Const SCREEN_MARGIN As Integer = 24
            Dim fw = panelW - PANEL_PAD - SCREEN_MARGIN * 2
            Dim fh = panelH - PANEL_PAD - SCREEN_MARGIN * 2
            Dim best = ComputeAutoFill(_LastGaugeCount, fw, fh)
            gaugeW = best.GaugeW
            bestCols = best.Cols
        Else
            ' ── Normal window: slider sets preferred size, even-row algo picks cols ──
            Dim availW = panelW - PANEL_PAD
            bestCols = ComputeBestCols(_LastGaugeCount, availW)
            gaugeW = Math.Max(80, availW \ bestCols - GAUGE_MARGIN)
        End If

        Dim gaugeH = CInt(gaugeW * 195.0 / 175.0)
        _GaugePanel.SuspendLayout()
        For Each ctrl As Control In _GaugePanel.Controls
            If TypeOf ctrl Is ucGauge Then ctrl.Size = New Size(gaugeW, gaugeH)
        Next
        _GaugePanel.ResumeLayout()
    End Sub

    ''' <summary>
    ''' Find the column count (1–8) that yields the largest gauge size while
    ''' fitting entirely within availW × availH. A small waste penalty nudges the
    ''' result toward layouts where the last row is full (even rows).
    ''' </summary>
    Private Function ComputeAutoFill(n As Integer, availW As Integer, availH As Integer) As (GaugeW As Integer, Cols As Integer)
        Const GAUGE_MARGIN As Integer = 16
        Dim aspect = 195.0 / 175.0
        Dim bestScore As Double = -1
        Dim bestCols As Integer = 1
        Dim bestGW As Integer = 80

        For cols = 1 To Math.Min(n, 8)
            Dim rows = CInt(Math.Ceiling(n / cols))
            Dim gwW = availW / cols - GAUGE_MARGIN
            Dim gwH = (availH / rows - GAUGE_MARGIN) / aspect
            Dim gw = Math.Max(80.0, Math.Min(gwW, gwH))
            ' Slight penalty for wasted slots in the last row (prefer even rows).
            Dim waste = (cols - (n Mod cols)) Mod cols
            Dim score = gw - waste * 0.5
            If score > bestScore Then
                bestScore = score
                bestCols = cols
                bestGW = CInt(gw)
            End If
        Next

        Return (bestGW, bestCols)
    End Function

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        RelayoutGauges()
    End Sub
#End Region

#Region "Toolbar events"
    Private Sub OnIntervalChanged(sender As Object, e As EventArgs)
        _Entry.UpdateInterval(CInt(_nudInterval.Value))
    End Sub

    Private Sub OnSizeSliderScroll(sender As Object, e As EventArgs)
        _GaugeBaseSize = _trkSize.Value
        _lblSizeVal.Text = $"{_GaugeBaseSize}px"
        RelayoutGauges()
    End Sub
#End Region

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        RemoveHandler _Entry.ResultReady, AddressOf OnResultReady
        MyBase.OnFormClosed(e)
    End Sub

End Class

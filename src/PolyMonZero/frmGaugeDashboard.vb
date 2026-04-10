''' <summary>
''' Modeless popup showing one skeuomorphic gauge (or numeric tile) per counter
''' for a single monitor entry. Subscribes to ResultReady and auto-refreshes.
''' </summary>
Public Class frmGaugeDashboard
    Inherits Form

    Private ReadOnly _Entry As MonitorEntry
    Private _GaugePanel As FlowLayoutPanel
    Private _lblStatus As Label

    Public Sub New(ByVal entry As MonitorEntry)
        _Entry = entry
        InitUI()
        RefreshGauges()
        AddHandler _Entry.ResultReady, AddressOf OnResultReady
    End Sub

#Region "UI Construction"
    Private Sub InitUI()
        Me.Text = _Entry.MonitorName & " — Gauge View"
        Me.Size = New Size(580, 500)
        Me.MinimumSize = New Size(300, 280)
        Me.BackColor = Color.FromArgb(22, 22, 28)
        Me.Font = New Font("Segoe UI", 9F)
        Me.StartPosition = FormStartPosition.CenterParent

        ' Title strip
        Dim titlePanel As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 38,
            .BackColor = Color.FromArgb(30, 32, 42)
        }

        ' Colored status dot (● character)
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
            .Size = New Size(400, 38)
        }

        titlePanel.Controls.Add(lblName)
        titlePanel.Controls.Add(_lblStatus)
        Me.Controls.Add(titlePanel)

        ' Gauge flow panel
        _GaugePanel = New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(8),
            .BackColor = Color.FromArgb(22, 22, 28)
        }
        Me.Controls.Add(_GaugePanel)
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

        ' Update status dot color
        Dim status = If(result IsNot Nothing, result.Status, MonitorStatus.Unknown)
        _lblStatus.ForeColor = StatusColors.ForStatus(status)

        If result Is Nothing OrElse result.Counters.Count = 0 Then
            _GaugePanel.Controls.Clear()
            Dim lbl As New Label() With {
                .Text = "No counter data yet — waiting for first run.",
                .ForeColor = Color.FromArgb(140, 143, 155),
                .Font = New Font("Segoe UI", 10F),
                .AutoSize = True,
                .Margin = New Padding(16)
            }
            _GaugePanel.Controls.Add(lbl)
            Return
        End If

        Dim counters = result.Counters

        ' Remove surplus gauge controls
        While _GaugePanel.Controls.Count > counters.Count
            _GaugePanel.Controls.RemoveAt(_GaugePanel.Controls.Count - 1)
        End While

        ' Add missing gauge controls
        While _GaugePanel.Controls.Count < counters.Count
            _GaugePanel.Controls.Add(New ucGauge())
        End While

        ' Update each gauge from its counter
        For i = 0 To counters.Count - 1
            Dim gauge = DirectCast(_GaugePanel.Controls(i), ucGauge)
            Dim cv = counters(i)

            gauge.CounterName = cv.Name
            gauge.Value = cv.Value

            If cv.HasRange Then
                ' Script explicitly provided range
                gauge.MinValue = cv.MinValue
                gauge.MaxValue = cv.MaxValue
                gauge.ShowAsGauge = True
            ElseIf ucGauge.IsGaugeType(cv.Name) Then
                ' Auto-detected percentage gauge
                gauge.MinValue = 0
                gauge.MaxValue = 100
                gauge.ShowAsGauge = True
            Else
                ' Raw numeric — show as tile
                gauge.MinValue = 0
                gauge.MaxValue = 100
                gauge.ShowAsGauge = False
            End If

            gauge.Invalidate()
        Next
    End Sub
#End Region

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        RemoveHandler _Entry.ResultReady, AddressOf OnResultReady
        MyBase.OnFormClosed(e)
    End Sub

End Class

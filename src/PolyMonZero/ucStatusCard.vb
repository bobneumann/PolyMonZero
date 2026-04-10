Imports System.Linq

''' <summary>
''' A self-drawing status card for one monitor.
''' Shows: colored status bar (left edge), monitor name, status badge, last-run time, message, counters,
''' and a Run Now button below the badge.
''' Raises EditRequested, DeleteRequested, RunNowRequested.
''' </summary>
Public Class ucStatusCard
    Inherits Panel

    Public Event EditRequested(ByVal entry As MonitorEntry)
    Public Event DeleteRequested(ByVal entry As MonitorEntry)
    Public Event RunNowRequested(ByVal entry As MonitorEntry)
    Public Event GaugeViewRequested(ByVal entry As MonitorEntry)

    Private _Entry As MonitorEntry
    Private Const BAR_WIDTH As Integer = 8
    Private Const BADGE_W As Integer = 44
    Private Const BADGE_H As Integer = 22
    Private Const PAD As Integer = 8

    Private _RunNowRect As Rectangle   ' hit-test area for the Run Now button

    Public Sub New(ByVal entry As MonitorEntry)
        _Entry = entry
        Me.DoubleBuffered = True
        Me.Size = New Size(360, 112)
        Me.Cursor = Cursors.Default
        Me.Margin = New Padding(4)

        Dim ctx As New ContextMenuStrip()
        Dim mnuEdit As New ToolStripMenuItem("Edit...")
        Dim mnuRunNow As New ToolStripMenuItem("Run Now")
        Dim mnuGauge As New ToolStripMenuItem("Gauge View...")
        Dim mnuSep As New ToolStripSeparator()
        Dim mnuDelete As New ToolStripMenuItem("Delete")
        ctx.Items.AddRange({mnuEdit, mnuRunNow, mnuGauge, mnuSep, mnuDelete})
        AddHandler mnuEdit.Click, Sub(s, e) RaiseEvent EditRequested(_Entry)
        AddHandler mnuRunNow.Click, Sub(s, e) RaiseEvent RunNowRequested(_Entry)
        AddHandler mnuGauge.Click, Sub(s, e) RaiseEvent GaugeViewRequested(_Entry)
        AddHandler mnuDelete.Click, Sub(s, e) RaiseEvent DeleteRequested(_Entry)
        Me.ContextMenuStrip = ctx
    End Sub

    Public ReadOnly Property Entry As MonitorEntry
        Get
            Return _Entry
        End Get
    End Property

    Public Sub UpdateEntry(ByVal entry As MonitorEntry)
        _Entry = entry
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.Clear(Color.FromArgb(245, 245, 245))

        Dim status = If(_Entry.LastResult IsNot Nothing, _Entry.LastResult.Status, MonitorStatus.Unknown)
        Dim statusColor = StatusColors.ForStatus(status)
        Dim statusText = StatusColors.TextForStatus(status)

        ' Left status bar
        g.FillRectangle(New SolidBrush(statusColor), 0, 0, BAR_WIDTH, Me.Height)

        ' Card border
        g.DrawRectangle(Pens.LightGray, 0, 0, Me.Width - 1, Me.Height - 1)

        ' Layout constants
        Dim x = BAR_WIDTH + PAD
        Dim badgeX = Me.Width - BADGE_W - PAD
        Dim textW = badgeX - x - PAD

        ' Monitor name
        Using nameFont As New Font("Segoe UI", 9.5F, FontStyle.Bold)
            Dim nameRect = New Rectangle(x, PAD, textW, 20)
            g.DrawString(_Entry.MonitorName, nameFont, Brushes.Black, nameRect,
                New StringFormat() With {.Trimming = StringTrimming.EllipsisCharacter, .FormatFlags = StringFormatFlags.NoWrap})
        End Using

        ' Status badge (right side, upper half)
        Dim badgeY = PAD + 2
        Using badgeBrush As New SolidBrush(statusColor)
            g.FillRectangle(badgeBrush, badgeX, badgeY, BADGE_W, BADGE_H)
        End Using
        Using badgeFont As New Font("Segoe UI", 7.5F, FontStyle.Bold)
            Dim badgeSF As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString(statusText, badgeFont, Brushes.White, New Rectangle(badgeX, badgeY, BADGE_W, BADGE_H), badgeSF)
        End Using

        ' Run Now button — below the badge
        _RunNowRect = New Rectangle(badgeX, badgeY + BADGE_H + 4, BADGE_W, 18)
        g.FillRectangle(New SolidBrush(Color.FromArgb(228, 232, 240)), _RunNowRect)
        g.DrawRectangle(Pens.LightGray, _RunNowRect)
        Using btnFont As New Font("Segoe UI", 7.5F)
            Dim btnSF As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString("▶ Run", btnFont, Brushes.DimGray, _RunNowRect, btnSF)
        End Using

        ' Last run time
        Dim timeStr = If(_Entry.LastRunTime = DateTime.MinValue, "Never run", "Last: " & _Entry.LastRunTime.ToString("HH:mm:ss"))
        Using timeFont As New Font("Segoe UI", 7.5F)
            g.DrawString(timeStr, timeFont, Brushes.Gray, x, PAD + 22)
        End Using

        ' Counters + message
        Dim rowY = PAD + 42
        Dim fullW = badgeX - x - PAD
        Dim noWrap = New StringFormat() With {.Trimming = StringTrimming.EllipsisCharacter, .FormatFlags = StringFormatFlags.NoWrap}
        Using smallFont As New Font("Segoe UI", 8F)
            If _Entry.LastResult IsNot Nothing Then
                Dim counters = _Entry.LastResult.Counters
                Dim i = 0
                While i < counters.Count
                    Dim left = $"{counters(i).Name}: {counters(i).Value:G6}"
                    Dim right = If(i + 1 < counters.Count, $"  |  {counters(i + 1).Name}: {counters(i + 1).Value:G6}", "")
                    g.DrawString(left & right, smallFont, Brushes.DimGray,
                        New Rectangle(x, rowY, fullW, 15), noWrap)
                    rowY += 15
                    i += 2
                End While
                If Not String.IsNullOrEmpty(_Entry.LastResult.StatusMessage) Then
                    Dim firstLine = _Entry.LastResult.StatusMessage.Split({Environment.NewLine, vbLf}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                    g.DrawString(firstLine, smallFont, Brushes.DimGray,
                        New Rectangle(x, rowY, fullW, 15), noWrap)
                End If
            Else
                g.DrawString("Waiting for first run...", smallFont, Brushes.Gray, x, rowY)
            End If
        End Using
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso _RunNowRect.Contains(e.Location) Then
            RaiseEvent RunNowRequested(_Entry)
        End If
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)
        If Not _RunNowRect.Contains(e.Location) Then
            RaiseEvent EditRequested(_Entry)
        End If
    End Sub

End Class

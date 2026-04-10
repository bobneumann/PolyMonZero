Imports System.Linq

''' <summary>
''' Compact horizontal row for one monitor in the Groups view.
''' Option A sizing: 34px height, 10.5pt name, 9pt counters, 6px status bar, subtle group tint.
''' </summary>
Public Class ucRibbonRow
    Inherits Panel

    Public Event EditRequested(ByVal entry As MonitorEntry)
    Public Event DeleteRequested(ByVal entry As MonitorEntry)
    Public Event RunNowRequested(ByVal entry As MonitorEntry)
    Public Event GaugeViewRequested(ByVal entry As MonitorEntry)

    Private ReadOnly _Entry As MonitorEntry
    Private Const BAR As Integer = 6
    Private Const ROW_H As Integer = 42
    Private Const NAME_W As Integer = 210
    Private Const RUN_W As Integer = 28

    Private _GroupColor As Color = Color.White
    Private _RunNowRect As Rectangle

    Public Sub New(ByVal entry As MonitorEntry)
        _Entry = entry
        Me.DoubleBuffered = True
        Me.Height = ROW_H
        Me.BackColor = Color.White
        Me.Cursor = Cursors.Default

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

    ''' <summary>Set the group color so the row can apply a subtle matching tint.</summary>
    Public Property GroupColor As Color
        Get
            Return _GroupColor
        End Get
        Set(value As Color)
            _GroupColor = value
            Me.Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics

        ' Subtle group tint (6% group color, 94% white)
        Dim bg = BlendColor(Color.White, _GroupColor, 0.06)
        g.Clear(bg)

        Dim status = If(_Entry.LastResult IsNot Nothing, _Entry.LastResult.Status, MonitorStatus.Unknown)
        Dim statusColor = StatusColors.ForStatus(status)

        ' Left status bar
        g.FillRectangle(New SolidBrush(statusColor), 0, 0, BAR, Me.Height)

        ' Status dot
        Dim dotSize = 13
        Dim dotX = BAR + 7
        Dim dotY = (Me.Height - dotSize) \ 2
        g.FillEllipse(New SolidBrush(statusColor), dotX, dotY, dotSize, dotSize)

        ' Right-to-left: ▶ button | time | counters | name
        Dim rightEdge = Me.Width - 6

        ' ▶ Run button
        _RunNowRect = New Rectangle(rightEdge - RUN_W, 5, RUN_W, Me.Height - 10)
        g.FillRectangle(New SolidBrush(Color.FromArgb(215, 220, 232)), _RunNowRect)
        Using runFont As New Font("Segoe UI", 14F)
            Dim sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString("▶", runFont, Brushes.DimGray, _RunNowRect, sf)
        End Using
        rightEdge = _RunNowRect.Left - 6

        ' Last-run time
        Dim timeStr = If(_Entry.LastRunTime = DateTime.MinValue, "—", _Entry.LastRunTime.ToString("HH:mm:ss"))
        Using timeFont As New Font("Segoe UI", 12F)
            Dim tw = CInt(g.MeasureString(timeStr, timeFont).Width) + 2
            rightEdge -= tw
            g.DrawString(timeStr, timeFont, Brushes.Gray, rightEdge, (Me.Height - timeFont.Height) \ 2)
            rightEdge -= 8
        End Using

        ' Monitor name
        Dim x = dotX + dotSize + 6
        Using nameFont As New Font("Segoe UI", 13F, FontStyle.Bold)
            Dim sf = New StringFormat() With {
                .Trimming = StringTrimming.EllipsisCharacter,
                .FormatFlags = StringFormatFlags.NoWrap,
                .LineAlignment = StringAlignment.Center
            }
            g.DrawString(_Entry.MonitorName, nameFont, Brushes.Black,
                New RectangleF(x, 0, NAME_W, Me.Height), sf)
        End Using
        x += NAME_W + 10

        ' Counters / message
        Dim detail As String = ""
        If _Entry.LastResult IsNot Nothing Then
            If _Entry.LastResult.Counters.Count > 0 Then
                detail = String.Join("  |  ", _Entry.LastResult.Counters.Select(Function(cv) $"{cv.Name}: {cv.Value:G6}"))
            ElseIf Not String.IsNullOrEmpty(_Entry.LastResult.StatusMessage) Then
                detail = _Entry.LastResult.StatusMessage.Split({Environment.NewLine, vbLf},
                    StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            End If
        End If

        If Not String.IsNullOrEmpty(detail) AndAlso rightEdge > x + 10 Then
            Using counterFont As New Font("Segoe UI", 11F)
                Dim sf = New StringFormat() With {
                    .Trimming = StringTrimming.EllipsisCharacter,
                    .FormatFlags = StringFormatFlags.NoWrap,
                    .LineAlignment = StringAlignment.Center
                }
                g.DrawString(detail, counterFont, New SolidBrush(Color.FromArgb(20, 20, 30)),
                    New RectangleF(x, 0, rightEdge - x, Me.Height), sf)
            End Using
        End If

        ' Bottom separator
        g.DrawLine(New Pen(Color.FromArgb(220, 222, 228)), BAR, Me.Height - 1, Me.Width, Me.Height - 1)
    End Sub

    Private Shared Function BlendColor(base As Color, overlay As Color, alpha As Double) As Color
        Return Color.FromArgb(
            CInt(base.R * (1 - alpha) + overlay.R * alpha),
            CInt(base.G * (1 - alpha) + overlay.G * alpha),
            CInt(base.B * (1 - alpha) + overlay.B * alpha))
    End Function

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

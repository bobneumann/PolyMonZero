Imports System.Linq

''' <summary>
''' Groups view panel. Displays monitors as compact ribbon rows under collapsible group headers.
''' Groups are ordered by SortOrder; collapse state persists for the session.
''' </summary>
Public Class pnlGroupedRibbon
    Inherits Panel

    Public Event EditRequested(ByVal entry As MonitorEntry)
    Public Event DeleteRequested(ByVal entry As MonitorEntry)
    Public Event RunNowRequested(ByVal entry As MonitorEntry)
    Public Event GaugeViewRequested(ByVal entry As MonitorEntry)

    Private _Collapsed As New Dictionary(Of String, Boolean)()
    Private _LastMonitors As List(Of MonitorEntry)
    Private _LastGroups As List(Of MonitorGroup)

    Public Sub New()
        Me.AutoScroll = True
        Me.BackColor = Color.FromArgb(245, 245, 245)
        AddHandler Me.Resize, AddressOf RibbonResize
    End Sub

    Public Sub Rebuild(monitors As List(Of MonitorEntry), groups As List(Of MonitorGroup))
        _LastMonitors = monitors
        _LastGroups = groups
        RebuildInternal()
    End Sub

    Private Sub RebuildInternal()
        If _LastMonitors Is Nothing Then Return

        ' Detach toggle handlers before clearing
        For Each ctrl In Me.Controls.Cast(Of Control)().ToList()
            If TypeOf ctrl Is GroupHeader Then
                RemoveHandler DirectCast(ctrl, GroupHeader).ToggleCollapse, AddressOf OnToggleCollapse
            End If
        Next

        Me.SuspendLayout()
        Me.Controls.Clear()

        Dim y = 0
        Dim w = Me.ClientSize.Width

        ' Build ordered group name list
        Dim orderedGroups = _LastGroups.OrderBy(Function(g) g.SortOrder).Select(Function(g) g.Name).ToList()
        If Not orderedGroups.Contains("Default") AndAlso _LastMonitors.Any(Function(m) m.GroupName = "Default") Then
            orderedGroups.Insert(0, "Default")
        End If
        For Each m In _LastMonitors
            If Not orderedGroups.Contains(m.GroupName) Then orderedGroups.Add(m.GroupName)
        Next

        For Each groupName In orderedGroups
            Dim groupMonitors = _LastMonitors.Where(Function(m) m.GroupName = groupName).ToList()
            If groupMonitors.Count = 0 Then Continue For

            Dim group = _LastGroups.FirstOrDefault(Function(g) g.Name = groupName)
            Dim groupColor = If(group IsNot Nothing, group.Color, Color.SteelBlue)

            If Not _Collapsed.ContainsKey(groupName) Then _Collapsed(groupName) = False
            Dim isCollapsed = _Collapsed(groupName)

            Dim header As New GroupHeader(groupName, groupColor, groupMonitors, isCollapsed)
            header.Location = New Point(0, y)
            header.Width = w
            AddHandler header.ToggleCollapse, AddressOf OnToggleCollapse
            Me.Controls.Add(header)
            y += header.Height

            If Not isCollapsed Then
                For Each entry In groupMonitors
                    Dim row As New ucRibbonRow(entry) With {.GroupColor = groupColor}
                    row.Location = New Point(0, y)
                    row.Width = w
                    AddHandler row.EditRequested, Sub(en As MonitorEntry) RaiseEvent EditRequested(en)
                    AddHandler row.DeleteRequested, Sub(en As MonitorEntry) RaiseEvent DeleteRequested(en)
                    AddHandler row.RunNowRequested, Sub(en As MonitorEntry) RaiseEvent RunNowRequested(en)
                    AddHandler row.GaugeViewRequested, Sub(en As MonitorEntry) RaiseEvent GaugeViewRequested(en)
                    Me.Controls.Add(row)
                    y += row.Height
                Next
            End If
        Next

        Me.ResumeLayout()
    End Sub

    Private Sub OnToggleCollapse(groupName As String)
        _Collapsed(groupName) = Not _Collapsed.GetValueOrDefault(groupName, False)
        RebuildInternal()
    End Sub

    ''' <summary>Repaint the row for entry and all group headers (counts may have changed).</summary>
    Public Sub InvalidateRow(entry As MonitorEntry)
        For Each ctrl In Me.Controls.Cast(Of Control)().ToList()
            If TypeOf ctrl Is ucRibbonRow AndAlso DirectCast(ctrl, ucRibbonRow).Entry Is entry Then
                ctrl.Invalidate()
            ElseIf TypeOf ctrl Is GroupHeader Then
                ctrl.Invalidate()
            End If
        Next
    End Sub

    ''' <summary>Repaint everything (e.g. on the 1-second tick to refresh timestamps).</summary>
    Public Sub InvalidateAll()
        For Each ctrl In Me.Controls.Cast(Of Control)().ToList()
            ctrl.Invalidate()
        Next
    End Sub

    Private Sub RibbonResize(sender As Object, e As EventArgs)
        Dim w = Me.ClientSize.Width
        For Each ctrl In Me.Controls.Cast(Of Control)().ToList()
            ctrl.Width = w
        Next
    End Sub

    ' ── Group header (private) ─────────────────────────────────────────────
    Private Class GroupHeader
        Inherits Panel

        Public Event ToggleCollapse(groupName As String)

        Private ReadOnly _GroupName As String
        Private ReadOnly _Color As Color
        Private ReadOnly _Monitors As List(Of MonitorEntry)
        Private ReadOnly _Collapsed As Boolean
        Private Const HDR_H As Integer = 40
        Private Const BAR As Integer = 8

        Public Sub New(name As String, color As Color, monitors As List(Of MonitorEntry), collapsed As Boolean)
            _GroupName = name
            _Color = color
            _Monitors = monitors
            _Collapsed = collapsed
            Me.Height = HDR_H
            Me.DoubleBuffered = True
            Me.Cursor = Cursors.Hand
            ' Background = 15% group color blended into white
            Me.BackColor = BlendColor(Color.White, color, 0.15)
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            Dim g = e.Graphics
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(Me.BackColor)

            ' Group color bar (wider)
            g.FillRectangle(New SolidBrush(_Color), 0, 0, BAR, Me.Height)

            ' Collapse arrow
            Dim arrow = If(_Collapsed, "▶", "▼")
            Dim x = BAR + 10
            Using arrowFont As New Font("Segoe UI", 11F)
                g.DrawString(arrow, arrowFont, New SolidBrush(Color.FromArgb(100, 100, 110)),
                    x, (Me.Height - arrowFont.Height) \ 2)
                x += CInt(g.MeasureString(arrow, arrowFont).Width) + 5
            End Using

            ' Group name
            Using nameFont As New Font("Segoe UI", 13F, FontStyle.Bold)
                g.DrawString(_GroupName, nameFont, Brushes.Black, x, (Me.Height - nameFont.Height) \ 2)
            End Using

            ' Status pill badges (right side, drawn right-to-left)
            Dim okCount   = _Monitors.Where(Function(m) m.LastResult IsNot Nothing AndAlso m.LastResult.Status = MonitorStatus.OK).Count()
            Dim warnCount = _Monitors.Where(Function(m) m.LastResult IsNot Nothing AndAlso m.LastResult.Status = MonitorStatus.Warning).Count()
            Dim failCount = _Monitors.Where(Function(m) m.LastResult IsNot Nothing AndAlso m.LastResult.Status = MonitorStatus.Fail).Count()
            Dim unkCount  = _Monitors.Count - okCount - warnCount - failCount

            Dim pills As New List(Of (txt As String, col As Color))
            If unkCount > 0  Then pills.Add((txt:=$"{unkCount} ???", col:=Color.Gray))
            If okCount > 0   Then pills.Add((txt:=$"{okCount} OK",   col:=StatusColors.ForStatus(MonitorStatus.OK)))
            If warnCount > 0 Then pills.Add((txt:=$"{warnCount} WARN", col:=StatusColors.ForStatus(MonitorStatus.Warning)))
            If failCount > 0 Then pills.Add((txt:=$"{failCount} FAIL", col:=StatusColors.ForStatus(MonitorStatus.Fail)))

            Using pillFont As New Font("Segoe UI", 10F, FontStyle.Bold)
                Dim PILL_H = 22
                Dim PILL_PAD = 8
                Dim PILL_GAP = 4
                Dim rx = Me.Width - 10
                For Each pill In pills
                    Dim tw = CInt(g.MeasureString(pill.txt, pillFont).Width)
                    Dim pw = tw + PILL_PAD
                    rx -= pw
                    Dim pillRect = New Rectangle(rx, (Me.Height - PILL_H) \ 2, pw, PILL_H)
                    FillRoundedRect(g, New SolidBrush(pill.col), pillRect, 4)
                    Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(pill.txt, pillFont, Brushes.White, pillRect, sf)
                    rx -= PILL_GAP
                Next
            End Using

            ' Bottom border — slightly darker blend of group color
            g.DrawLine(New Pen(BlendColor(Color.Silver, _Color, 0.3)), 0, Me.Height - 1, Me.Width, Me.Height - 1)
        End Sub

        Private Shared Function BlendColor(base As Color, overlay As Color, alpha As Double) As Color
            Return Color.FromArgb(
                CInt(base.R * (1 - alpha) + overlay.R * alpha),
                CInt(base.G * (1 - alpha) + overlay.G * alpha),
                CInt(base.B * (1 - alpha) + overlay.B * alpha))
        End Function

        Private Shared Sub FillRoundedRect(g As Graphics, brush As Brush, rect As Rectangle, radius As Integer)
            Dim path As New Drawing2D.GraphicsPath()
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90)
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90)
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90)
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90)
            path.CloseFigure()
            g.FillPath(brush, path)
        End Sub

        Protected Overrides Sub OnClick(e As EventArgs)
            RaiseEvent ToggleCollapse(_GroupName)
        End Sub
    End Class

End Class

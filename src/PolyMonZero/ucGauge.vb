Imports System.Drawing.Drawing2D

''' <summary>
''' Skeuomorphic needle gauge or large numeric tile, drawn entirely in GDI+.
''' Set ShowAsGauge=True for percentage-type counters (0–100 arc + needle).
''' Set ShowAsGauge=False for raw numeric counters (large centered value, no arc).
''' </summary>
Public Class ucGauge
    Inherits Panel

    ' ── Data ──────────────────────────────────────────────────────────────
    Public Property CounterName As String = ""
    Public Property Value As Double = 0
    Public Property MinValue As Double = 0
    Public Property MaxValue As Double = 100
    Public Property T1Value As Double = 70
    Public Property T1Enabled As Boolean = False
    Public Property T2Value As Double = 90
    Public Property T2Enabled As Boolean = False
    Public Property ShowAsGauge As Boolean = True

    ' ── Geometry ──────────────────────────────────────────────────────────
    Private Const CX As Integer = 87
    Private Const CY As Integer = 83
    Private Const R_BEZEL As Integer = 80
    Private Const R_FACE As Integer = 70
    Private Const R_ARC As Integer = 58
    Private Const ARC_W As Integer = 13
    Private Const R_TICK_OUT As Integer = 53
    Private Const R_TICK_MAJ As Integer = 43
    Private Const R_TICK_MIN As Integer = 50
    Private Const R_LABEL As Integer = 34
    Private Const R_NEEDLE As Integer = 50
    Private Const R_HUB As Integer = 8
    Private Const START_DEG As Double = 150.0
    Private Const SWEEP_DEG As Double = 240.0

    Public Sub New()
        Me.DoubleBuffered = True
        Me.Size = New Size(175, 195)
        Me.BackColor = Color.FromArgb(22, 22, 28)
        Me.Margin = New Padding(8)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        g.Clear(Me.BackColor)

        If ShowAsGauge Then
            DrawGauge(g)
        Else
            DrawNumericTile(g)
        End If
    End Sub

    ' ── Gauge rendering ───────────────────────────────────────────────────
    Private Sub DrawGauge(g As Graphics)
        DrawBezel(g)

        ' Face
        g.FillEllipse(New SolidBrush(Color.FromArgb(18, 20, 26)),
            CX - R_FACE, CY - R_FACE, R_FACE * 2, R_FACE * 2)

        DrawZones(g)
        DrawTicks(g)
        DrawNeedle(g)

        ' Hub outer
        g.FillEllipse(New SolidBrush(Color.FromArgb(55, 57, 65)),
            CX - R_HUB, CY - R_HUB, R_HUB * 2, R_HUB * 2)
        ' Hub inner (highlight)
        g.FillEllipse(New SolidBrush(Color.FromArgb(140, 143, 155)),
            CX - R_HUB + 3, CY - R_HUB + 3, (R_HUB - 3) * 2, (R_HUB - 3) * 2)

        ' Value
        Dim valStr = $"{Value:G4}"
        If MaxValue = 100 AndAlso MinValue = 0 Then valStr = $"{Value:F1}%"
        Using vFont As New Font("Segoe UI", 12F, FontStyle.Bold)
            Dim vr = New Rectangle(CX - 52, CY + 22, 104, 22)
            Dim vsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString(valStr, vFont, New SolidBrush(ValueColor()), vr, vsf)
        End Using

        ' Name
        Using nFont As New Font("Segoe UI", 8.5F)
            Dim nr = New Rectangle(4, CY + 48, Me.Width - 8, 20)
            Dim nsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont, New SolidBrush(Color.FromArgb(155, 158, 170)), nr, nsf)
        End Using
    End Sub

    Private Sub DrawBezel(g As Graphics)
        ' Radial gradient — lighter center, darker rim
        Dim bRect = New Rectangle(CX - R_BEZEL, CY - R_BEZEL, R_BEZEL * 2, R_BEZEL * 2)
        Dim path As New GraphicsPath()
        path.AddEllipse(bRect)
        Using pgb As New PathGradientBrush(path)
            pgb.CenterPoint = New PointF(CX - 14, CY - 14)
            pgb.CenterColor = Color.FromArgb(88, 90, 100)
            pgb.SurroundColors = New Color() {Color.FromArgb(36, 38, 44)}
            g.FillPath(pgb, path)
        End Using
        ' Glare highlight (upper-left arc)
        Using glare As New Pen(Color.FromArgb(32, 255, 255, 255), 7)
            g.DrawArc(glare, CX - R_BEZEL + 5, CY - R_BEZEL + 5,
                (R_BEZEL - 5) * 2, (R_BEZEL - 5) * 2, 195, 115)
        End Using
        ' Inner rim shadow
        g.DrawEllipse(New Pen(Color.FromArgb(20, 22, 26), 2),
            CX - R_FACE - 3, CY - R_FACE - 3, (R_FACE + 3) * 2, (R_FACE + 3) * 2)
    End Sub

    Private Sub DrawZones(g As Graphics)
        Dim arcRect = New Rectangle(CX - R_ARC, CY - R_ARC, R_ARC * 2, R_ARC * 2)
        Dim range = MaxValue - MinValue
        Dim t1Pct = If(T1Enabled, Clamp01((T1Value - MinValue) / range), 0.7)
        Dim t2Pct = If(T2Enabled, Clamp01((T2Value - MinValue) / range), 0.9)
        t1Pct = Math.Max(0.05, Math.Min(0.93, t1Pct))
        t2Pct = Math.Max(t1Pct + 0.04, Math.Min(0.99, t2Pct))

        Dim gSweep = CSng(t1Pct * SWEEP_DEG)
        Dim ySweep = CSng((t2Pct - t1Pct) * SWEEP_DEG)
        Dim rSweep = CSng((1.0 - t2Pct) * SWEEP_DEG)

        Using p = New Pen(Color.FromArgb(45, 185, 100), ARC_W)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG), gSweep)
        End Using
        Using p = New Pen(Color.FromArgb(230, 150, 0), ARC_W)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG) + gSweep, ySweep)
        End Using
        Using p = New Pen(Color.FromArgb(210, 45, 60), ARC_W)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG) + gSweep + ySweep, rSweep)
        End Using
    End Sub

    Private Sub DrawTicks(g As Graphics)
        Dim totalDiv = 20  ' 4 minor between each of 5 major
        For i = 0 To totalDiv
            Dim pct = CDbl(i) / totalDiv
            Dim angleRad = (START_DEG + pct * SWEEP_DEG) * Math.PI / 180.0
            Dim isMajor = (i Mod 4 = 0)
            Dim innerR = If(isMajor, R_TICK_MAJ, R_TICK_MIN)

            Dim x1 = CSng(CX + R_TICK_OUT * Math.Cos(angleRad))
            Dim y1 = CSng(CY + R_TICK_OUT * Math.Sin(angleRad))
            Dim x2 = CSng(CX + innerR * Math.Cos(angleRad))
            Dim y2 = CSng(CY + innerR * Math.Sin(angleRad))

            If isMajor Then
                g.DrawLine(New Pen(Color.FromArgb(220, 222, 235), 2), x1, y1, x2, y2)
                ' Label
                Dim lv = MinValue + pct * (MaxValue - MinValue)
                Dim ls = If(MaxValue = 100, CInt(lv).ToString(), $"{lv:G3}")
                Dim lx = CSng(CX + R_LABEL * Math.Cos(angleRad))
                Dim ly = CSng(CY + R_LABEL * Math.Sin(angleRad))
                Using lf As New Font("Segoe UI", 7F)
                    Dim lsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(ls, lf, New SolidBrush(Color.FromArgb(165, 168, 180)), lx, ly)
                End Using
            Else
                g.DrawLine(New Pen(Color.FromArgb(100, 102, 112), 1), x1, y1, x2, y2)
            End If
        Next
    End Sub

    Private Sub DrawNeedle(g As Graphics)
        Dim pct = Clamp01(If(MaxValue > MinValue, (Value - MinValue) / (MaxValue - MinValue), 0))
        Dim angleRad = (START_DEG + pct * SWEEP_DEG) * Math.PI / 180.0
        Dim nx = CSng(CX + R_NEEDLE * Math.Cos(angleRad))
        Dim ny = CSng(CY + R_NEEDLE * Math.Sin(angleRad))

        ' Drop shadow
        g.DrawLine(New Pen(Color.FromArgb(120, 0, 0, 0), 3), CX + 2, CY + 2, nx + 2, ny + 2)
        ' Body
        g.DrawLine(New Pen(Color.FromArgb(210, 215, 230), 2), CX, CY, nx, ny)
        ' Bright tip highlight
        Dim tipX = CSng(CX + (R_NEEDLE - 8) * Math.Cos(angleRad))
        Dim tipY = CSng(CY + (R_NEEDLE - 8) * Math.Sin(angleRad))
        g.DrawLine(New Pen(Color.White, 1.5F), tipX, tipY, nx, ny)
    End Sub

    ' ── Numeric tile rendering ────────────────────────────────────────────
    Private Sub DrawNumericTile(g As Graphics)
        ' Subtle border
        g.DrawRectangle(New Pen(Color.FromArgb(48, 50, 58)), 0, 0, Me.Width - 1, Me.Height - 1)

        ' Name at top
        Using nFont As New Font("Segoe UI", 8.5F)
            Dim nr = New Rectangle(6, 10, Me.Width - 12, 18)
            Dim nsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont, New SolidBrush(Color.FromArgb(140, 143, 155)), nr, nsf)
        End Using

        ' Large value centered
        Dim valStr = FormatValue(Value)
        Using vFont As New Font("Segoe UI", 22F, FontStyle.Bold)
            Dim vr = New Rectangle(4, Me.Height \ 2 - 30, Me.Width - 8, 55)
            Dim vsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            ' Auto-shrink if too wide
            Dim sf2 = g.MeasureString(valStr, vFont)
            If sf2.Width > Me.Width - 12 Then
                Using vFont2 As New Font("Segoe UI", 16F, FontStyle.Bold)
                    g.DrawString(valStr, vFont2, New SolidBrush(ValueColor()), vr, vsf)
                End Using
            Else
                g.DrawString(valStr, vFont, New SolidBrush(ValueColor()), vr, vsf)
            End If
        End Using

        ' Name label again at bottom (for context when scrolled)
        Using nFont As New Font("Segoe UI", 8F)
            Dim br = New Rectangle(4, Me.Height - 26, Me.Width - 8, 18)
            Dim bsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont, New SolidBrush(Color.FromArgb(120, 123, 135)), br, bsf)
        End Using
    End Sub

    ' ── Helpers ───────────────────────────────────────────────────────────
    Private Function ValueColor() As Color
        If MaxValue <= MinValue Then Return Color.FromArgb(180, 183, 200)
        Dim pct = (Value - MinValue) / (MaxValue - MinValue)
        Dim t1Pct = If(T1Enabled, Clamp01((T1Value - MinValue) / (MaxValue - MinValue)), 0.7)
        Dim t2Pct = If(T2Enabled, Clamp01((T2Value - MinValue) / (MaxValue - MinValue)), 0.9)
        If pct >= t2Pct Then Return Color.FromArgb(210, 60, 75)
        If pct >= t1Pct Then Return Color.FromArgb(230, 160, 20)
        Return Color.FromArgb(60, 200, 120)
    End Function

    Private Shared Function FormatValue(v As Double) As String
        If Math.Abs(v) >= 1000 Then Return $"{v:N0}"
        If Math.Abs(v) >= 10 Then Return $"{v:F1}"
        Return $"{v:F2}"
    End Function

    Private Shared Function Clamp01(v As Double) As Double
        Return Math.Max(0.0, Math.Min(1.0, v))
    End Function

    ''' <summary>Returns True if this counter name suggests a 0–100 percentage gauge.</summary>
    Public Shared Function IsGaugeType(name As String) As Boolean
        If String.IsNullOrEmpty(name) Then Return False
        Dim n = name.ToLowerInvariant()
        If n.EndsWith("%") OrElse n.EndsWith("pct") OrElse n.EndsWith("percent") Then Return True
        If n.Contains("load") OrElse n.Contains("usage") OrElse n.Contains("utiliz") Then Return True
        ' Matches Core0, Core1, Core12, etc.
        If System.Text.RegularExpressions.Regex.IsMatch(name, "^Core\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then Return True
        Return False
    End Function

End Class

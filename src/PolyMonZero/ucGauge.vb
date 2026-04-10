Imports System.Drawing.Drawing2D

''' <summary>
''' Skeuomorphic needle gauge or large numeric tile, drawn entirely in GDI+.
''' All geometry scales dynamically with the control size so the dashboard is
''' user-resizable. Default size 175×195; set ShowAsGauge=True for 0–100 arc+needle.
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

    ' ── Design-time base size (geometry was authored at this size) ─────────
    Private Const BASE_W As Double = 175.0
    Private Const BASE_H As Double = 195.0

    ' ── Arc geometry (fixed) ───────────────────────────────────────────────
    Private Const START_DEG As Double = 150.0
    Private Const SWEEP_DEG As Double = 240.0

    ' ── Dynamic geometry (all scale with GaugeScale) ───────────────────────
    Private ReadOnly Property GaugeScale As Double
        Get
            Return Math.Min(Me.Width / BASE_W, Me.Height / BASE_H)
        End Get
    End Property

    Private ReadOnly Property GCX As Integer
        Get
            Return Me.Width \ 2
        End Get
    End Property

    Private ReadOnly Property GCY As Integer
        Get
            Return CInt(Me.Height * 83.0 / BASE_H)
        End Get
    End Property

    ''' <summary>Scale an integer pixel value by GaugeScale.</summary>
    Private Function Px(base As Integer) As Integer
        Return CInt(base * GaugeScale)
    End Function

    ''' <summary>Scale a font size by GaugeScale, with a 5.5pt floor.</summary>
    Private Function PxF(base As Single) As Single
        Return CSng(Math.Max(5.5, base * GaugeScale))
    End Function

    Public Sub New()
        Me.DoubleBuffered = True
        Me.Size = New Size(175, 195)
        Me.BackColor = Color.FromArgb(22, 22, 28)
        Me.Margin = New Padding(8)
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        g.Clear(Me.BackColor)
        If ShowAsGauge Then DrawGauge(g) Else DrawNumericTile(g)
    End Sub

    ' ── Gauge rendering ───────────────────────────────────────────────────
    Private Sub DrawGauge(g As Graphics)
        DrawBezel(g)

        ' Face
        Dim rFace = Px(70)
        g.FillEllipse(New SolidBrush(Color.FromArgb(18, 20, 26)),
            GCX - rFace, GCY - rFace, rFace * 2, rFace * 2)

        DrawZones(g)
        DrawTicks(g)
        DrawNeedle(g)

        ' Hub
        Dim rHub = Px(8)
        g.FillEllipse(New SolidBrush(Color.FromArgb(55, 57, 65)),
            GCX - rHub, GCY - rHub, rHub * 2, rHub * 2)
        g.FillEllipse(New SolidBrush(Color.FromArgb(140, 143, 155)),
            GCX - rHub + Px(3), GCY - rHub + Px(3), (rHub - Px(3)) * 2, (rHub - Px(3)) * 2)

        ' Value
        Dim valStr = $"{Value:G4}"
        If MaxValue = 100 AndAlso MinValue = 0 Then valStr = $"{Value:F1}%"
        Using vFont As New Font("Segoe UI", PxF(12F), FontStyle.Bold)
            Dim vr = New Rectangle(GCX - Px(52), GCY + Px(22), Px(104), Px(24))
            Dim vsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString(valStr, vFont, New SolidBrush(ValueColor()), vr, vsf)
        End Using

        ' Name
        Using nFont As New Font("Segoe UI", PxF(8.5F))
            Dim nr = New Rectangle(4, GCY + Px(50), Me.Width - 8, Px(22))
            Dim nsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont, New SolidBrush(Color.FromArgb(155, 158, 170)), nr, nsf)
        End Using
    End Sub

    Private Sub DrawBezel(g As Graphics)
        Dim rBezel = Px(80)
        Dim bRect = New Rectangle(GCX - rBezel, GCY - rBezel, rBezel * 2, rBezel * 2)
        Dim path As New GraphicsPath()
        path.AddEllipse(bRect)
        Using pgb As New PathGradientBrush(path)
            pgb.CenterPoint = New PointF(GCX - Px(14), GCY - Px(14))
            pgb.CenterColor = Color.FromArgb(88, 90, 100)
            pgb.SurroundColors = New Color() {Color.FromArgb(36, 38, 44)}
            g.FillPath(pgb, path)
        End Using
        Using glare As New Pen(Color.FromArgb(32, 255, 255, 255), CSng(Math.Max(2, 7 * GaugeScale)))
            g.DrawArc(glare, GCX - rBezel + Px(5), GCY - rBezel + Px(5),
                (rBezel - Px(5)) * 2, (rBezel - Px(5)) * 2, 195, 115)
        End Using
        Dim rFace = Px(70)
        g.DrawEllipse(New Pen(Color.FromArgb(20, 22, 26), CSng(Math.Max(1, 2 * GaugeScale))),
            GCX - rFace - Px(3), GCY - rFace - Px(3), (rFace + Px(3)) * 2, (rFace + Px(3)) * 2)
    End Sub

    Private Sub DrawZones(g As Graphics)
        Dim rArc = Px(58)
        Dim arcRect = New Rectangle(GCX - rArc, GCY - rArc, rArc * 2, rArc * 2)
        Dim arcW = CSng(Math.Max(3, 13 * GaugeScale))
        Dim range = MaxValue - MinValue
        Dim t1Pct = If(T1Enabled, Clamp01((T1Value - MinValue) / range), 0.7)
        Dim t2Pct = If(T2Enabled, Clamp01((T2Value - MinValue) / range), 0.9)
        t1Pct = Math.Max(0.05, Math.Min(0.93, t1Pct))
        t2Pct = Math.Max(t1Pct + 0.04, Math.Min(0.99, t2Pct))

        Dim gSweep = CSng(t1Pct * SWEEP_DEG)
        Dim ySweep = CSng((t2Pct - t1Pct) * SWEEP_DEG)
        Dim rSweep = CSng((1.0 - t2Pct) * SWEEP_DEG)

        Using p = New Pen(Color.FromArgb(45, 185, 100), arcW)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG), gSweep)
        End Using
        Using p = New Pen(Color.FromArgb(230, 150, 0), arcW)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG) + gSweep, ySweep)
        End Using
        Using p = New Pen(Color.FromArgb(210, 45, 60), arcW)
            p.EndCap = LineCap.Round : p.StartCap = LineCap.Round
            g.DrawArc(p, arcRect, CSng(START_DEG) + gSweep + ySweep, rSweep)
        End Using
    End Sub

    Private Sub DrawTicks(g As Graphics)
        Dim rTickOut = Px(53)
        Dim rTickMaj = Px(43)
        Dim rTickMin = Px(50)
        Dim rLabel   = Px(34)
        Dim totalDiv = 20
        For i = 0 To totalDiv
            Dim pct = CDbl(i) / totalDiv
            Dim angleRad = (START_DEG + pct * SWEEP_DEG) * Math.PI / 180.0
            Dim isMajor = (i Mod 4 = 0)
            Dim innerR = If(isMajor, rTickMaj, rTickMin)

            Dim x1 = CSng(GCX + rTickOut * Math.Cos(angleRad))
            Dim y1 = CSng(GCY + rTickOut * Math.Sin(angleRad))
            Dim x2 = CSng(GCX + innerR * Math.Cos(angleRad))
            Dim y2 = CSng(GCY + innerR * Math.Sin(angleRad))

            If isMajor Then
                g.DrawLine(New Pen(Color.FromArgb(220, 222, 235), CSng(Math.Max(1, 2 * GaugeScale))), x1, y1, x2, y2)
                Dim lv = MinValue + pct * (MaxValue - MinValue)
                Dim ls = If(MaxValue = 100, CInt(lv).ToString(), $"{lv:G3}")
                Dim lx = CSng(GCX + rLabel * Math.Cos(angleRad))
                Dim ly = CSng(GCY + rLabel * Math.Sin(angleRad))
                Using lf As New Font("Segoe UI", PxF(7F))
                    Dim lsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(ls, lf, New SolidBrush(Color.FromArgb(165, 168, 180)), lx, ly)
                End Using
            Else
                g.DrawLine(New Pen(Color.FromArgb(100, 102, 112), CSng(Math.Max(0.5F, GaugeScale))), x1, y1, x2, y2)
            End If
        Next
    End Sub

    Private Sub DrawNeedle(g As Graphics)
        ' Geometry — tip reaches into the coloured arc bands (arc centre is at Px(58))
        Dim rTip  As Integer = Px(63)   ' tip: inside the arc zone
        Dim rTail As Integer = Px(6)    ' counterweight tail (will be covered by the hub cap)
        Dim halfW As Single  = CSng(Math.Max(1.5, Px(5)))   ' half-width at base
        Const TIP_HALF As Single = 0.8F                      ' half-width at tip (near-point)

        Dim pct = Clamp01(If(MaxValue > MinValue, (Value - MinValue) / (MaxValue - MinValue), 0))
        Dim ang = (START_DEG + pct * SWEEP_DEG) * Math.PI / 180.0

        ' Unit vectors: along needle and perpendicular to it
        Dim dx = CSng(Math.Cos(ang))
        Dim dy = CSng(Math.Sin(ang))
        Dim perpX = -dy
        Dim perpY = dx

        Dim cx = CSng(GCX)
        Dim cy = CSng(GCY)

        ' Four corners: two at the wide base (behind pivot), two at the tapered tip
        Dim bL = New PointF(cx + halfW * perpX - rTail * dx,   cy + halfW * perpY - rTail * dy)
        Dim bR = New PointF(cx - halfW * perpX - rTail * dx,   cy - halfW * perpY - rTail * dy)
        Dim tL = New PointF(cx + TIP_HALF * perpX + rTip * dx, cy + TIP_HALF * perpY + rTip * dy)
        Dim tR = New PointF(cx - TIP_HALF * perpX + rTip * dx, cy - TIP_HALF * perpY + rTip * dy)
        Dim pts = New PointF() {bL, tL, tR, bR}

        ' Drop shadow — shifted polygon drawn underneath
        Dim sOff = CSng(Math.Max(1.2, 2.0 * GaugeScale))
        g.FillPolygon(New SolidBrush(Color.FromArgb(90, 0, 0, 0)),
            New PointF() {
                New PointF(bL.X + sOff, bL.Y + sOff),
                New PointF(tL.X + sOff, tL.Y + sOff),
                New PointF(tR.X + sOff, tR.Y + sOff),
                New PointF(bR.X + sOff, bR.Y + sOff)
            })

        ' Body: 3-D ridge effect via LinearGradientBrush perpendicular to needle.
        ' Gradient direction = across the needle width (left edge → right edge).
        ' 5-stop blend: dark orange edge → bevel → bright amber ridge → bevel → dark edge.
        Dim needlePath As New GraphicsPath()
        needlePath.AddPolygon(pts)

        If halfW >= 2 Then
            Dim gFrom = New PointF(cx + halfW * perpX, cy + halfW * perpY)
            Dim gTo   = New PointF(cx - halfW * perpX, cy - halfW * perpY)
            Using lgb As New LinearGradientBrush(gFrom, gTo, Color.Black, Color.Black)
                Dim cb As New ColorBlend(5)
                cb.Colors = New Color() {
                    Color.FromArgb(140,  55,  0),   ' dark outer edge
                    Color.FromArgb(210,  90,  5),   ' bevel shadow
                    Color.FromArgb(255, 175, 45),   ' bright amber ridge peak
                    Color.FromArgb(210,  90,  5),   ' bevel shadow
                    Color.FromArgb(140,  55,  0)    ' dark outer edge
                }
                cb.Positions = New Single() {0.0F, 0.2F, 0.5F, 0.8F, 1.0F}
                lgb.InterpolationColors = cb
                g.FillPath(lgb, needlePath)
            End Using
        Else
            g.FillPath(New SolidBrush(Color.FromArgb(220, 100, 0)), needlePath)
        End If

        needlePath.Dispose()

        ' Specular highlight line along the centre ridge (the peak of the sword cross-section)
        Dim hiliteStart = New PointF(cx + Px(7) * dx, cy + Px(7) * dy)      ' clear of the hub
        Dim hiliteEnd   = New PointF(cx + (rTip - Px(4)) * dx, cy + (rTip - Px(4)) * dy)
        Using hp As New Pen(Color.FromArgb(190, 255, 225, 115),
                            CSng(Math.Max(0.5F, 0.8F * GaugeScale)))
            g.DrawLine(hp, hiliteStart, hiliteEnd)
        End Using

        ' Subtle dark perimeter line to separate the needle from the face
        Using op As New Pen(Color.FromArgb(100, 80, 25, 0),
                            CSng(Math.Max(0.4F, 0.6F * GaugeScale)))
            g.DrawPolygon(op, pts)
        End Using
    End Sub

    ' ── Numeric tile rendering ────────────────────────────────────────────
    Private Sub DrawNumericTile(g As Graphics)
        g.DrawRectangle(New Pen(Color.FromArgb(48, 50, 58)), 0, 0, Me.Width - 1, Me.Height - 1)

        Using nFont As New Font("Segoe UI", PxF(8.5F))
            Dim nr = New Rectangle(6, 10, Me.Width - 12, Px(20))
            Dim nsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont, New SolidBrush(Color.FromArgb(140, 143, 155)), nr, nsf)
        End Using

        Dim valStr = FormatValue(Value)
        Dim vr = New Rectangle(4, Me.Height \ 2 - Px(30), Me.Width - 8, Px(55))
        Dim vsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
        Using vFont As New Font("Segoe UI", PxF(22F), FontStyle.Bold)
            If g.MeasureString(valStr, vFont).Width > Me.Width - 12 Then
                Using vFont2 As New Font("Segoe UI", PxF(16F), FontStyle.Bold)
                    g.DrawString(valStr, vFont2, New SolidBrush(ValueColor()), vr, vsf)
                End Using
            Else
                g.DrawString(valStr, vFont, New SolidBrush(ValueColor()), vr, vsf)
            End If
        End Using

        Using nFont2 As New Font("Segoe UI", PxF(8F))
            Dim br = New Rectangle(4, Me.Height - Px(26), Me.Width - 8, Px(20))
            Dim bsf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
            g.DrawString(CounterName, nFont2, New SolidBrush(Color.FromArgb(120, 123, 135)), br, bsf)
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

    Public Shared Function IsGaugeType(name As String) As Boolean
        If String.IsNullOrEmpty(name) Then Return False
        Dim n = name.ToLowerInvariant()
        If n.EndsWith("%") OrElse n.EndsWith("pct") OrElse n.EndsWith("percent") Then Return True
        If n.Contains("load") OrElse n.Contains("usage") OrElse n.Contains("utiliz") Then Return True
        If System.Text.RegularExpressions.Regex.IsMatch(name, "^Core\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then Return True
        Return False
    End Function

End Class

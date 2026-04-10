''' <summary>
''' Holds definition, runtime state, and history for one monitor.
''' Owns the polling timer, PowerShell execution, and alert state machine.
''' </summary>
Public Class MonitorEntry

    Public Event ResultReady(ByVal entry As MonitorEntry)
    Public Event AlertTriggered(ByVal entry As MonitorEntry, ByVal message As String)

#Region "Definition"
    Public Property MonitorID As Integer
    Public Property MonitorName As String = "New Monitor"
    Public Property GroupName As String = "Default"
    Public Property Script As String = ""
    Public Property PollingIntervalSec As Integer = 60
    Public Property RetentionPoints As Integer = 120
    Public Property IsPaused As Boolean = False
#End Region

#Region "Alert Settings"
    Public Property AlertEnabled As Boolean = False
    Public Property AlertRoomId As String = ""          ' Matrix room ID

    ''' <summary>When True, alert every N events regardless of status. When False, use specific conditions below.</summary>
    Public Property AlertEveryNEvents As Boolean = False
    Public Property AlertEveryNEventsCount As Integer = 1

    ' Failure alerts
    Public Property AlertOnNewFailure As Boolean = True
    Public Property AlertRepeatEveryNFailures As Integer = 0    ' 0 = no repeat
    Public Property AlertOnFailToOK As Boolean = True

    ' Warning alerts
    Public Property AlertOnNewWarning As Boolean = False
    Public Property AlertRepeatEveryNWarnings As Integer = 0   ' 0 = no repeat
    Public Property AlertOnWarnToOK As Boolean = False
#End Region

#Region "Runtime"
    Public Property LastResult As MonitorResult = Nothing
    Public Property LastRunTime As DateTime = DateTime.MinValue
    Public Property IsRunning As Boolean = False

    ''' <summary>Rolling history per counter: name -> list of (time, value)</summary>
    Public ReadOnly Property History As New Dictionary(Of String, List(Of (t As DateTime, v As Double)))
#End Region

#Region "Alert State (not persisted)"
    Private _PreviousStatus As MonitorStatus = MonitorStatus.Unknown
    Private _PreviousWasAlerted As Boolean = False  ' did the previous non-OK state produce an alert?
    Private _ConsecutiveFailures As Integer = 0
    Private _ConsecutiveWarnings As Integer = 0
    Private _EventCount As Integer = 0              ' for EveryNEvents mode
#End Region

    Private _PSMonitor As PowerShellMonitor
    Private WithEvents _Timer As New Timer()

#Region "Control"
    Public Sub StartMonitor()
        _Timer.Stop()
        _PSMonitor?.Dispose()
        _PSMonitor = Nothing
        If String.IsNullOrWhiteSpace(Script) Then Return
        _PSMonitor = New PowerShellMonitor(Script)
        IsRunning = True
        _Timer.Interval = PollingIntervalSec * 1000
        If Not IsPaused Then _Timer.Start()
    End Sub

    Public Sub StopMonitor()
        _Timer.Stop()
        IsRunning = False
    End Sub

    Public Sub SetPaused(ByVal paused As Boolean)
        IsPaused = paused
        If paused Then
            _Timer.Stop()
        ElseIf IsRunning Then
            _Timer.Start()
        End If
    End Sub

    Public Sub RunNow()
        _Timer.Stop()
        Timer_Tick(Nothing, EventArgs.Empty)
    End Sub
#End Region

#Region "Timer / Execution"
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles _Timer.Tick
        If _PSMonitor Is Nothing Then Return
        _Timer.Stop()
        Try
            Dim result = _PSMonitor.GetSample()
            Dim prev = _PreviousStatus

            LastResult = result
            LastRunTime = DateTime.Now

            ' Update counter history
            For Each cv In result.Counters
                If Not History.ContainsKey(cv.Name) Then History(cv.Name) = New List(Of (DateTime, Double))()
                History(cv.Name).Add((LastRunTime, cv.Value))
                While History(cv.Name).Count > RetentionPoints
                    History(cv.Name).RemoveAt(0)
                End While
            Next

            ' Alert state machine
            If AlertEnabled Then
                EvaluateAlerts(prev, result)
            End If

            _PreviousStatus = result.Status
            RaiseEvent ResultReady(Me)
        Catch
        Finally
            If IsRunning AndAlso Not IsPaused Then _Timer.Start()
        End Try
    End Sub

    Private Sub EvaluateAlerts(ByVal prev As MonitorStatus, ByVal result As MonitorResult)
        Dim curr = result.Status

        If AlertEveryNEvents Then
            ' Mode 1: alert every N poll cycles regardless of status
            _EventCount += 1
            If _EventCount >= AlertEveryNEventsCount Then
                _EventCount = 0
                FireAlert(FormatMessage(curr, Nothing), curr)
            End If
            Return
        End If

        ' Mode 2: specific conditions
        Select Case curr
            Case MonitorStatus.Fail
                If prev <> MonitorStatus.Fail Then
                    ' New failure
                    _ConsecutiveFailures = 1
                    _ConsecutiveWarnings = 0
                    If AlertOnNewFailure Then
                        FireAlert($"[FAIL] {MonitorName}" & FormatDetail(result), curr)
                        _PreviousWasAlerted = True
                    End If
                Else
                    ' Continuing failure
                    _ConsecutiveFailures += 1
                    If AlertRepeatEveryNFailures > 0 AndAlso _ConsecutiveFailures Mod AlertRepeatEveryNFailures = 0 Then
                        FireAlert($"[STILL FAILING x{_ConsecutiveFailures}] {MonitorName}" & FormatDetail(result), curr)
                    End If
                End If

            Case MonitorStatus.Warning
                If prev <> MonitorStatus.Warning Then
                    ' New warning
                    _ConsecutiveWarnings = 1
                    _ConsecutiveFailures = 0
                    If AlertOnNewWarning Then
                        FireAlert($"[WARN] {MonitorName}" & FormatDetail(result), curr)
                        _PreviousWasAlerted = True
                    End If
                Else
                    ' Continuing warning
                    _ConsecutiveWarnings += 1
                    If AlertRepeatEveryNWarnings > 0 AndAlso _ConsecutiveWarnings Mod AlertRepeatEveryNWarnings = 0 Then
                        FireAlert($"[STILL WARNING x{_ConsecutiveWarnings}] {MonitorName}" & FormatDetail(result), curr)
                    End If
                End If

            Case MonitorStatus.OK
                If prev = MonitorStatus.Fail Then
                    _ConsecutiveFailures = 0
                    If AlertOnFailToOK AndAlso _PreviousWasAlerted Then
                        FireAlert($"[RECOVERED] {MonitorName} (was FAIL)" & FormatDetail(result), curr)
                    End If
                    _PreviousWasAlerted = False
                ElseIf prev = MonitorStatus.Warning Then
                    _ConsecutiveWarnings = 0
                    If AlertOnWarnToOK AndAlso _PreviousWasAlerted Then
                        FireAlert($"[RECOVERED] {MonitorName} (was WARN)" & FormatDetail(result), curr)
                    End If
                    _PreviousWasAlerted = False
                End If
        End Select
    End Sub

    Private Sub FireAlert(ByVal message As String, ByVal status As MonitorStatus)
        RaiseEvent AlertTriggered(Me, message)
    End Sub

    Private Shared Function FormatDetail(ByVal result As MonitorResult) As String
        Dim sb As New System.Text.StringBuilder()
        If Not String.IsNullOrEmpty(result.StatusMessage) Then
            Dim firstLine = result.StatusMessage.Split({Environment.NewLine, vbLf}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            If Not String.IsNullOrEmpty(firstLine) Then sb.Append(vbLf & firstLine)
        End If
        If result.Counters.Count > 0 Then
            sb.Append(vbLf & String.Join(" | ", result.Counters.Select(Function(cv) $"{cv.Name}: {cv.Value:G6}")))
        End If
        Return sb.ToString()
    End Function

    Private Function FormatMessage(ByVal status As MonitorStatus, ByVal result As MonitorResult) As String
        Dim tag = $"[{StatusColors.TextForStatus(status)}] {MonitorName}"
        If result Is Nothing Then Return tag
        Return tag & FormatDetail(result)
    End Function
#End Region

    Public Sub Dispose()
        _Timer.Stop()
        _Timer.Dispose()
        _PSMonitor?.Dispose()
    End Sub

End Class

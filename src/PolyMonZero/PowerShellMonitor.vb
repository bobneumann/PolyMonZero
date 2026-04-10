Imports System.Management.Automation
Imports System.Management.Automation.Runspaces

''' <summary>
''' Executes a PowerShell script and returns a MonitorResult.
''' Script text is stored directly (not XML-wrapped).
''' Injects $Counters=$true, $errlvl=0, $messages=@(), $Counter=@()
''' </summary>
Public Class PowerShellMonitor
    Implements IDisposable

    Private _Script As String
    Private _RunSpace As Runspace
    Private _Disposed As Boolean

    Public Sub New(ByVal script As String)
        _Script = script
    End Sub

    Public Function GetSample() As MonitorResult
        Dim result As New MonitorResult()
        If String.IsNullOrWhiteSpace(_Script) Then Return result

        Try
            If _RunSpace Is Nothing Then
                _RunSpace = RunspaceFactory.CreateRunspace()
                _RunSpace.Open()
            End If

            _RunSpace.SessionStateProxy.SetVariable("Counters", True)
            _RunSpace.SessionStateProxy.SetVariable("errlvl", 0)
            _RunSpace.SessionStateProxy.SetVariable("messages", New String() {})
            _RunSpace.SessionStateProxy.SetVariable("Counter", New Object() {})

            Using pipe As Pipeline = _RunSpace.CreatePipeline(_Script)
                pipe.Invoke()
            End Using

            Dim errlvlRaw = _RunSpace.SessionStateProxy.GetVariable("errlvl")
            Dim messagesRaw = _RunSpace.SessionStateProxy.GetVariable("messages")

            Dim errlvl As Integer = 0
            If errlvlRaw IsNot Nothing AndAlso IsNumeric(errlvlRaw) Then errlvl = CInt(errlvlRaw)
            Select Case errlvl
                Case 0 : result.Status = MonitorStatus.OK
                Case 1 : result.Status = MonitorStatus.Warning
                Case Else : result.Status = MonitorStatus.Fail
            End Select

            If messagesRaw IsNot Nothing Then
                Dim msgs As New List(Of String)
                If TypeOf messagesRaw Is String() Then
                    msgs.AddRange(CType(messagesRaw, String()))
                ElseIf TypeOf messagesRaw Is System.Collections.IEnumerable Then
                    For Each o In CType(messagesRaw, System.Collections.IEnumerable)
                        If o IsNot Nothing Then msgs.Add(o.ToString())
                    Next
                Else
                    msgs.Add(messagesRaw.ToString())
                End If
                result.StatusMessage = String.Join(Environment.NewLine, msgs)
            End If

            ' Parse $Counter = @(,@("name", value), ...) or @(,@("name", value, min, max), ...)
            ' The second pipeline emits pscustomobjects with N, V, Mn, Mx, HR fields.
            Using cvPipe As Pipeline = _RunSpace.CreatePipeline(
                "if ($Counter) { foreach ($__p in $Counter) {" &
                " [pscustomobject]@{N=[string]$__p[0];V=[double]$__p[1];" &
                " Mn=if($__p.Count -ge 4){[double]$__p[2]}else{0.0};" &
                " Mx=if($__p.Count -ge 4){[double]$__p[3]}else{0.0};" &
                " HR=($__p.Count -ge 4)} } }")
                Dim cvResults = cvPipe.Invoke()
                For Each pso In cvResults
                    If pso Is Nothing Then Continue For
                    Dim n = pso.Properties("N")?.Value?.ToString()
                    Dim vProp = pso.Properties("V")?.Value
                    If Not String.IsNullOrEmpty(n) AndAlso vProp IsNot Nothing AndAlso IsNumeric(vProp) Then
                        Dim cv As New CounterValue() With {.Name = n, .Value = CDbl(vProp)}
                        Dim hrProp = pso.Properties("HR")?.Value
                        If hrProp IsNot Nothing AndAlso CBool(hrProp) Then
                            Dim mnProp = pso.Properties("Mn")?.Value
                            Dim mxProp = pso.Properties("Mx")?.Value
                            cv.MinValue = If(mnProp IsNot Nothing AndAlso IsNumeric(mnProp), CDbl(mnProp), 0)
                            cv.MaxValue = If(mxProp IsNot Nothing AndAlso IsNumeric(mxProp), CDbl(mxProp), 100)
                            cv.HasRange = True
                        End If
                        result.Counters.Add(cv)
                    End If
                Next
            End Using

        Catch ex As Exception
            result.Status = MonitorStatus.Fail
            result.StatusMessage = "Script error: " & ex.Message
        End Try

        Return result
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _Disposed Then Return
        _Disposed = True
        Try
            If _RunSpace IsNot Nothing Then
                If _RunSpace.RunspaceStateInfo.State <> RunspaceState.Closed Then _RunSpace.Close()
                _RunSpace.Dispose()
                _RunSpace = Nothing
            End If
        Catch
        End Try
    End Sub

End Class

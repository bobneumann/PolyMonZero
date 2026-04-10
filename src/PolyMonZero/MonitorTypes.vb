Public Enum MonitorStatus
    Unknown = 0
    OK = 1
    Warning = 2
    Fail = 3
End Enum

Public Class CounterValue
    Public Property Name As String = ""
    Public Property Value As Double = 0
    ''' <summary>Minimum of the gauge range. Only meaningful when HasRange=True.</summary>
    Public Property MinValue As Double = 0
    ''' <summary>Maximum of the gauge range. Only meaningful when HasRange=True.</summary>
    Public Property MaxValue As Double = 100
    ''' <summary>True when the script supplied a 4-element array: @("Name", value, min, max).</summary>
    Public Property HasRange As Boolean = False
End Class

Public Class MonitorResult
    Public Property Status As MonitorStatus = MonitorStatus.Unknown
    Public Property StatusMessage As String = ""
    Public Property Counters As New List(Of CounterValue)
End Class

Public Class MonitorGroup
    Public Property Name As String = "Default"
    Public Property Color As Color = Color.SteelBlue
    Public Property SortOrder As Integer = 0
End Class

Module StatusColors
    Public ReadOnly Property ForStatus(status As MonitorStatus) As Color
        Get
            Select Case status
                Case MonitorStatus.OK      : Return Color.FromArgb(60, 179, 113)   ' medium sea green
                Case MonitorStatus.Warning : Return Color.FromArgb(255, 165, 0)    ' orange
                Case MonitorStatus.Fail    : Return Color.FromArgb(220, 53, 69)    ' red
                Case Else                  : Return Color.FromArgb(150, 150, 150)  ' gray
            End Select
        End Get
    End Property

    Public ReadOnly Property TextForStatus(status As MonitorStatus) As String
        Get
            Select Case status
                Case MonitorStatus.OK      : Return "OK"
                Case MonitorStatus.Warning : Return "WARN"
                Case MonitorStatus.Fail    : Return "FAIL"
                Case Else                  : Return "???"
            End Select
        End Get
    End Property
End Module

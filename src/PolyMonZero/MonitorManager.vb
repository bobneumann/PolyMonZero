Imports System.IO
Imports System.Text.Json

''' <summary>
''' Owns the list of monitors and groups. Handles file I/O (.pmz JSON format).
''' </summary>
Public Class MonitorManager

    Public Event MonitorUpdated(ByVal entry As MonitorEntry)
    Public Event DirtyChanged(ByVal isDirty As Boolean)
    Public Event AlertTriggered(ByVal entry As MonitorEntry, ByVal message As String)

    Public ReadOnly Property Monitors As New List(Of MonitorEntry)
    Public ReadOnly Property Groups As New List(Of MonitorGroup)
    Public Property CurrentFilePath As String = ""

    ' Global Matrix alert config
    Public Property MatrixHomeserver As String = ""
    Public Property MatrixToken As String = ""

    Private _NextID As Integer = 1
    Private _IsDirty As Boolean = False

    Public Property IsDirty As Boolean
        Get
            Return _IsDirty
        End Get
        Private Set(value As Boolean)
            If _IsDirty <> value Then
                _IsDirty = value
                RaiseEvent DirtyChanged(value)
            End If
        End Set
    End Property

#Region "Monitor Management"
    Public Function AddMonitor() As MonitorEntry
        Dim entry As New MonitorEntry() With {
            .MonitorID = _NextID,
            .GroupName = If(Groups.Count > 0, Groups(0).Name, "Default")
        }
        _NextID += 1
        Attach(entry)
        IsDirty = True
        Return entry
    End Function

    Public Sub RemoveMonitor(ByVal entry As MonitorEntry)
        entry.StopMonitor()
        Detach(entry)
        Monitors.Remove(entry)
        entry.Dispose()
        IsDirty = True
    End Sub

    Private Sub Attach(ByVal entry As MonitorEntry)
        AddHandler entry.ResultReady, AddressOf OnResultReady
        AddHandler entry.AlertTriggered, AddressOf OnAlertTriggered
        Monitors.Add(entry)
    End Sub

    Private Sub Detach(ByVal entry As MonitorEntry)
        RemoveHandler entry.ResultReady, AddressOf OnResultReady
        RemoveHandler entry.AlertTriggered, AddressOf OnAlertTriggered
    End Sub

    Private Sub OnResultReady(ByVal entry As MonitorEntry)
        IsDirty = True
        RaiseEvent MonitorUpdated(entry)
    End Sub

    Private Sub OnAlertTriggered(ByVal entry As MonitorEntry, ByVal message As String)
        RaiseEvent AlertTriggered(entry, message)
    End Sub

    Public Sub MarkDirty()
        IsDirty = True
    End Sub
#End Region

#Region "Group Management"
    Public Function EnsureGroup(ByVal name As String) As MonitorGroup
        Dim g = Groups.FirstOrDefault(Function(x) x.Name = name)
        If g Is Nothing Then
            g = New MonitorGroup() With {.Name = name}
            Groups.Add(g)
        End If
        Return g
    End Function

    Public Sub AddGroup(ByVal name As String)
        If Not Groups.Any(Function(x) x.Name = name) Then
            Groups.Add(New MonitorGroup() With {.Name = name})
            IsDirty = True
        End If
    End Sub

    Public Sub RemoveGroup(ByVal name As String)
        Dim g = Groups.FirstOrDefault(Function(x) x.Name = name)
        If g Is Nothing Then Return
        For Each m In Monitors.Where(Function(x) x.GroupName = name)
            m.GroupName = "Default"
        Next
        Groups.Remove(g)
        IsDirty = True
    End Sub
#End Region

#Region "Pause / Resume"
    Public Sub PauseAll()
        For Each m In Monitors
            m.SetPaused(True)
        Next
    End Sub

    Public Sub ResumeAll()
        For Each m In Monitors
            m.SetPaused(False)
        Next
    End Sub
#End Region

#Region "File I/O — .pmz JSON"

    Private Class JFile
        Public Property version As Integer = 2
        Public Property matrixHomeserver As String = ""
        Public Property matrixToken As String = ""
        Public Property groups As New List(Of JGroup)
        Public Property monitors As New List(Of JMonitor)
    End Class

    Private Class JGroup
        Public Property name As String = ""
        Public Property color As String = "#4682B4"
        Public Property sortOrder As Integer = 0
    End Class

    Private Class JMonitor
        Public Property id As Integer
        Public Property name As String = "Monitor"
        Public Property group As String = "Default"
        Public Property pollSec As Integer = 60
        Public Property retention As Integer = 120
        Public Property isPaused As Boolean = False
        Public Property script As String = ""
        ' Alert settings
        Public Property alertEnabled As Boolean = False
        Public Property alertRoomId As String = ""
        Public Property alertEveryNEvents As Boolean = False
        Public Property alertEveryNEventsCount As Integer = 1
        Public Property alertOnNewFailure As Boolean = True
        Public Property alertRepeatEveryNFailures As Integer = 0
        Public Property alertOnFailToOK As Boolean = True
        Public Property alertOnNewWarning As Boolean = False
        Public Property alertRepeatEveryNWarnings As Integer = 0
        Public Property alertOnWarnToOK As Boolean = False
    End Class

    ' ── Last-file persistence ─────────────────────────────────────────────
    Private Shared ReadOnly _LastFilePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PolyMonZero", "lastfile.txt")

    Public Shared Function GetLastFilePath() As String
        Try
            If File.Exists(_LastFilePath) Then Return File.ReadAllText(_LastFilePath, System.Text.Encoding.UTF8).Trim()
        Catch
        End Try
        Return ""
    End Function

    Private Shared Sub SaveLastFilePath(path As String)
        Try
            Dim dir = System.IO.Path.GetDirectoryName(_LastFilePath)
            If dir IsNot Nothing Then Directory.CreateDirectory(dir)
            File.WriteAllText(_LastFilePath, path, System.Text.Encoding.UTF8)
        Catch
        End Try
    End Sub

    Public Sub NewFile()
        StopAll()
        Monitors.Clear()
        Groups.Clear()
        _NextID = 1
        CurrentFilePath = ""
        MatrixHomeserver = ""
        MatrixToken = ""
        IsDirty = False
    End Sub

    Public Sub LoadFromFile(ByVal path As String)
        StopAll()
        Monitors.Clear()
        Groups.Clear()
        _NextID = 1

        Dim json = File.ReadAllText(path, System.Text.Encoding.UTF8)
        Dim opts As New JsonSerializerOptions() With {.PropertyNameCaseInsensitive = True}
        Dim jf = JsonSerializer.Deserialize(Of JFile)(json, opts)
        If jf Is Nothing Then Return

        MatrixHomeserver = If(jf.matrixHomeserver, "")
        MatrixToken = If(jf.matrixToken, "")

        For Each jg In jf.groups
            If Not Groups.Any(Function(x) x.Name = jg.name) Then
                Groups.Add(New MonitorGroup() With {
                    .Name = jg.name,
                    .Color = ParseColor(jg.color, Color.SteelBlue),
                    .SortOrder = jg.sortOrder
                })
            End If
        Next

        For Each jm In jf.monitors
            Dim entry As New MonitorEntry() With {
                .MonitorID = jm.id,
                .MonitorName = jm.name,
                .GroupName = If(String.IsNullOrEmpty(jm.group), "Default", jm.group),
                .PollingIntervalSec = jm.pollSec,
                .RetentionPoints = jm.retention,
                .IsPaused = jm.isPaused,
                .Script = jm.script,
                .AlertEnabled = jm.alertEnabled,
                .AlertRoomId = If(jm.alertRoomId, ""),
                .AlertEveryNEvents = jm.alertEveryNEvents,
                .AlertEveryNEventsCount = If(jm.alertEveryNEventsCount > 0, jm.alertEveryNEventsCount, 1),
                .AlertOnNewFailure = jm.alertOnNewFailure,
                .AlertRepeatEveryNFailures = jm.alertRepeatEveryNFailures,
                .AlertOnFailToOK = jm.alertOnFailToOK,
                .AlertOnNewWarning = jm.alertOnNewWarning,
                .AlertRepeatEveryNWarnings = jm.alertRepeatEveryNWarnings,
                .AlertOnWarnToOK = jm.alertOnWarnToOK
            }
            If entry.MonitorID >= _NextID Then _NextID = entry.MonitorID + 1
            EnsureGroup(entry.GroupName)
            Attach(entry)
            entry.StartMonitor()
        Next

        CurrentFilePath = path
        SaveLastFilePath(path)
        IsDirty = False
    End Sub

    Public Sub SaveToFile(ByVal path As String)
        Dim jf As New JFile() With {
            .matrixHomeserver = MatrixHomeserver,
            .matrixToken = MatrixToken
        }
        For Each g In Groups
            jf.groups.Add(New JGroup() With {
                .name = g.Name,
                .color = ColorTranslator.ToHtml(g.Color),
                .sortOrder = g.SortOrder
            })
        Next
        For Each m In Monitors
            jf.monitors.Add(New JMonitor() With {
                .id = m.MonitorID,
                .name = m.MonitorName,
                .group = m.GroupName,
                .pollSec = m.PollingIntervalSec,
                .retention = m.RetentionPoints,
                .isPaused = m.IsPaused,
                .script = m.Script,
                .alertEnabled = m.AlertEnabled,
                .alertRoomId = m.AlertRoomId,
                .alertEveryNEvents = m.AlertEveryNEvents,
                .alertEveryNEventsCount = m.AlertEveryNEventsCount,
                .alertOnNewFailure = m.AlertOnNewFailure,
                .alertRepeatEveryNFailures = m.AlertRepeatEveryNFailures,
                .alertOnFailToOK = m.AlertOnFailToOK,
                .alertOnNewWarning = m.AlertOnNewWarning,
                .alertRepeatEveryNWarnings = m.AlertRepeatEveryNWarnings,
                .alertOnWarnToOK = m.AlertOnWarnToOK
            })
        Next
        Dim opts As New JsonSerializerOptions() With {.WriteIndented = True}
        File.WriteAllText(path, JsonSerializer.Serialize(jf, opts), System.Text.Encoding.UTF8)
        CurrentFilePath = path
        SaveLastFilePath(path)
        IsDirty = False
    End Sub
#End Region

#Region "Helpers"
    Private Sub StopAll()
        For Each m In Monitors
            Detach(m)
            m.StopMonitor()
            m.Dispose()
        Next
    End Sub

    Private Shared Function ParseColor(s As String, def As Color) As Color
        If String.IsNullOrEmpty(s) Then Return def
        Try
            Return ColorTranslator.FromHtml(s)
        Catch
            Return def
        End Try
    End Function
#End Region

End Class

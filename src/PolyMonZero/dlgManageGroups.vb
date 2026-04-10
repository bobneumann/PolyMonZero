Imports System.Linq

''' <summary>
''' Add, rename, reorder, delete groups, and pick their color.
''' Works on a local copy; changes applied to MonitorManager on OK (including monitor renames).
''' </summary>
Public Class dlgManageGroups
    Inherits Form

    Private ReadOnly _Manager As MonitorManager
    Private ReadOnly _Working As New List(Of MonitorGroup)()
    ''' <summary>Maps working copy -> original group name (Nothing if the group is new).</summary>
    Private ReadOnly _OriginalName As New Dictionary(Of MonitorGroup, String)()

    Private _ListBox As ListBox

    Public Sub New(manager As MonitorManager)
        _Manager = manager
        For Each g In manager.Groups.OrderBy(Function(x) x.SortOrder)
            Dim copy As New MonitorGroup() With {.Name = g.Name, .Color = g.Color, .SortOrder = g.SortOrder}
            _Working.Add(copy)
            _OriginalName(copy) = g.Name
        Next
        InitUI()
        RefreshList(-1)
    End Sub

#Region "UI Construction"
    Private Sub InitUI()
        Me.Text = "Manage Groups"
        Me.Size = New Size(400, 320)
        Me.MinimumSize = New Size(340, 260)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Font("Segoe UI", 9F)

        ' OK / Cancel (Bottom — add first so Fill resolves correctly)
        Dim okPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Bottom,
            .FlowDirection = FlowDirection.RightToLeft,
            .Height = 44,
            .Padding = New Padding(4)
        }
        Dim btnOK As New Button() With {.Text = "OK", .Size = New Size(80, 28)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .Size = New Size(80, 28), .DialogResult = DialogResult.Cancel}
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        okPanel.Controls.AddRange({btnCancel, btnOK})
        Me.Controls.Add(okPanel)
        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel

        ' Main area: list + side buttons
        Dim main As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .Padding = New Padding(8)
        }
        main.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        main.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 108))

        ' ListBox (owner-draw for color swatch)
        _ListBox = New ListBox() With {
            .Dock = DockStyle.Fill,
            .DrawMode = DrawMode.OwnerDrawFixed,
            .ItemHeight = 24
        }
        AddHandler _ListBox.DrawItem, AddressOf ListBox_DrawItem
        main.Controls.Add(_ListBox, 0, 0)

        ' Side buttons
        Dim side As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown,
            .Padding = New Padding(4, 0, 0, 0)
        }
        Dim bAdd    As New Button() With {.Text = "Add",      .Size = New Size(96, 26)}
        Dim bRename As New Button() With {.Text = "Rename",   .Size = New Size(96, 26)}
        Dim bDelete As New Button() With {.Text = "Delete",   .Size = New Size(96, 26)}
        Dim bUp     As New Button() With {.Text = "▲  Up",    .Size = New Size(96, 26)}
        Dim bDown   As New Button() With {.Text = "▼  Down",  .Size = New Size(96, 26)}
        Dim bColor  As New Button() With {.Text = "Color...", .Size = New Size(96, 26)}
        ' Add a little spacing before Up/Down and Color
        bUp.Margin = New Padding(0, 6, 0, 0)
        bColor.Margin = New Padding(0, 6, 0, 0)
        AddHandler bAdd.Click,    AddressOf BtnAdd_Click
        AddHandler bRename.Click, AddressOf BtnRename_Click
        AddHandler bDelete.Click, AddressOf BtnDelete_Click
        AddHandler bUp.Click,     AddressOf BtnUp_Click
        AddHandler bDown.Click,   AddressOf BtnDown_Click
        AddHandler bColor.Click,  AddressOf BtnColor_Click
        side.Controls.AddRange({bAdd, bRename, bDelete, bUp, bDown, bColor})
        main.Controls.Add(side, 1, 0)

        Me.Controls.Add(main)
    End Sub
#End Region

#Region "List management"
    ''' <summary>Rebuild the ListBox. Pass selectedGroup to re-select it, or -1 to keep current index.</summary>
    Private Overloads Sub RefreshList(keepSelectedGroup As MonitorGroup)  ' select by object ref
        Dim sorted = _Working.OrderBy(Function(g) g.SortOrder).ToList()
        _ListBox.Items.Clear()
        For Each g In sorted
            _ListBox.Items.Add(g)
        Next
        If keepSelectedGroup IsNot Nothing Then
            _ListBox.SelectedItem = keepSelectedGroup
        ElseIf _ListBox.Items.Count > 0 Then
            _ListBox.SelectedIndex = 0
        End If
    End Sub

    Private Overloads Sub RefreshList(selectIndex As Integer)  ' select by index
        Dim sorted = _Working.OrderBy(Function(g) g.SortOrder).ToList()
        _ListBox.Items.Clear()
        For Each g In sorted
            _ListBox.Items.Add(g)
        Next
        If _ListBox.Items.Count > 0 Then
            _ListBox.SelectedIndex = Math.Max(0, Math.Min(selectIndex, _ListBox.Items.Count - 1))
        End If
    End Sub

    Private Sub ListBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 Then Return
        Dim g = DirectCast(_ListBox.Items(e.Index), MonitorGroup)
        e.DrawBackground()
        ' Color swatch
        Dim swatch = New Rectangle(e.Bounds.X + 4, e.Bounds.Y + 5, 14, e.Bounds.Height - 10)
        e.Graphics.FillRectangle(New SolidBrush(g.Color), swatch)
        e.Graphics.DrawRectangle(Pens.Gray, swatch)
        ' Name
        Dim brush = If((e.State And DrawItemState.Selected) <> 0, SystemBrushes.HighlightText, Brushes.Black)
        e.Graphics.DrawString(g.Name, e.Font, brush,
            swatch.Right + 6, e.Bounds.Y + (e.Bounds.Height - e.Font.Height) \ 2)
        e.DrawFocusRectangle()
    End Sub

    Private Function SelectedGroup() As MonitorGroup
        Return If(TypeOf _ListBox.SelectedItem Is MonitorGroup,
            DirectCast(_ListBox.SelectedItem, MonitorGroup), Nothing)
    End Function
#End Region

#Region "Button handlers"
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        Dim name = InputBox("Group name:", "Add Group", "")
        If String.IsNullOrWhiteSpace(name) Then Return
        name = name.Trim()
        If _Working.Any(Function(g) g.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) Then
            MessageBox.Show("A group with that name already exists.", "Add Group",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim nextOrder = If(_Working.Count > 0, _Working.Max(Function(g) g.SortOrder) + 1, 0)
        Dim newGroup As New MonitorGroup() With {.Name = name, .Color = Color.SteelBlue, .SortOrder = nextOrder}
        _Working.Add(newGroup)
        ' No entry in _OriginalName means it's new
        RefreshList(newGroup)
    End Sub

    Private Sub BtnRename_Click(sender As Object, e As EventArgs)
        Dim g = SelectedGroup()
        If g Is Nothing Then Return
        Dim name = InputBox("New name:", "Rename Group", g.Name)
        If String.IsNullOrWhiteSpace(name) OrElse name.Trim() = g.Name Then Return
        name = name.Trim()
        If _Working.Any(Function(x) x IsNot g AndAlso x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) Then
            MessageBox.Show("A group with that name already exists.", "Rename Group",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        g.Name = name
        RefreshList(g)
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs)
        Dim g = SelectedGroup()
        If g Is Nothing Then Return
        If _Working.Count = 1 Then
            MessageBox.Show("Cannot delete the last group.", "Delete Group",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim r = MessageBox.Show($"Delete ""{g.Name}""? Monitors in this group will move to the first remaining group.",
            "Delete Group", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If r <> DialogResult.Yes Then Return
        Dim idx = _ListBox.SelectedIndex
        _Working.Remove(g)
        _OriginalName.Remove(g)
        RefreshList(idx)
    End Sub

    Private Sub BtnUp_Click(sender As Object, e As EventArgs)
        Dim g = SelectedGroup()
        If g Is Nothing Then Return
        Dim sorted = _Working.OrderBy(Function(x) x.SortOrder).ToList()
        Dim i = sorted.IndexOf(g)
        If i <= 0 Then Return
        Dim tmp = sorted(i - 1).SortOrder
        sorted(i - 1).SortOrder = sorted(i).SortOrder
        sorted(i).SortOrder = tmp
        RefreshList(g)
    End Sub

    Private Sub BtnDown_Click(sender As Object, e As EventArgs)
        Dim g = SelectedGroup()
        If g Is Nothing Then Return
        Dim sorted = _Working.OrderBy(Function(x) x.SortOrder).ToList()
        Dim i = sorted.IndexOf(g)
        If i >= sorted.Count - 1 Then Return
        Dim tmp = sorted(i + 1).SortOrder
        sorted(i + 1).SortOrder = sorted(i).SortOrder
        sorted(i).SortOrder = tmp
        RefreshList(g)
    End Sub

    Private Sub BtnColor_Click(sender As Object, e As EventArgs)
        Dim g = SelectedGroup()
        If g Is Nothing Then Return
        Using dlg As New ColorDialog() With {.Color = g.Color, .FullOpen = True}
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                g.Color = dlg.Color
                _ListBox.Invalidate()
            End If
        End Using
    End Sub
#End Region

#Region "Apply"
    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        ApplyToManager()
        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub ApplyToManager()
        ' 1. Handle renames: match working copy to original by _OriginalName key
        For Each kvp In _OriginalName
            Dim copy = kvp.Key
            Dim origName = kvp.Value
            If copy.Name = origName Then Continue For  ' no rename
            Dim mgGroup = _Manager.Groups.FirstOrDefault(Function(g) g.Name = origName)
            If mgGroup IsNot Nothing Then
                ' Rename monitors in this group
                For Each m In _Manager.Monitors.Where(Function(x) x.GroupName = origName)
                    m.GroupName = copy.Name
                Next
                mgGroup.Name = copy.Name
            End If
        Next

        ' 2. Delete removed groups (manager.RemoveGroup reassigns monitors to Default)
        For Each origName In _OriginalName.Values.ToList()
            If Not _Working.Any(Function(g) _OriginalName.ContainsKey(g) AndAlso _OriginalName(g) = origName) Then
                _Manager.RemoveGroup(origName)
            End If
        Next

        ' 3. Add new groups (those without an entry in _OriginalName)
        For Each g In _Working.Where(Function(x) Not _OriginalName.ContainsKey(x))
            _Manager.AddGroup(g.Name)
        Next

        ' 4. Sync color + sort order for all remaining groups
        For Each g In _Working
            Dim mgGroup = _Manager.Groups.FirstOrDefault(Function(x) x.Name = g.Name)
            If mgGroup IsNot Nothing Then
                mgGroup.Color = g.Color
                mgGroup.SortOrder = g.SortOrder
            End If
        Next

        _Manager.MarkDirty()
    End Sub
#End Region

End Class

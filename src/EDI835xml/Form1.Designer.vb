<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar2 = New System.Windows.Forms.ToolStripProgressBar()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AccessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.XMLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExportToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcelToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AccessToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.XMLToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotepadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenDirectoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.ListView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(570, 263)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "File"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Path"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Size"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Type"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Last Updated"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 241)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(570, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(121, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(121, 17)
        Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
        '
        'ToolStripProgressBar2
        '
        Me.ToolStripProgressBar2.Name = "ToolStripProgressBar2"
        Me.ToolStripProgressBar2.Size = New System.Drawing.Size(100, 16)
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.ExportToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(570, 24)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.OpenToolStripMenuItem.Text = "&Open"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcelToolStripMenuItem, Me.AccessToolStripMenuItem, Me.XMLToolStripMenuItem})
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.ExportToolStripMenuItem.Text = "&Export"
        '
        'ExcelToolStripMenuItem
        '
        Me.ExcelToolStripMenuItem.Name = "ExcelToolStripMenuItem"
        Me.ExcelToolStripMenuItem.Size = New System.Drawing.Size(110, 22)
        Me.ExcelToolStripMenuItem.Text = "E&xcel"
        '
        'AccessToolStripMenuItem
        '
        Me.AccessToolStripMenuItem.Name = "AccessToolStripMenuItem"
        Me.AccessToolStripMenuItem.Size = New System.Drawing.Size(110, 22)
        Me.AccessToolStripMenuItem.Text = "&Access"
        '
        'XMLToolStripMenuItem
        '
        Me.XMLToolStripMenuItem.Name = "XMLToolStripMenuItem"
        Me.XMLToolStripMenuItem.Size = New System.Drawing.Size(110, 22)
        Me.XMLToolStripMenuItem.Text = "&XML"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionsToolStripMenuItem.Text = "O&ptions"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExportToolStripMenuItem1, Me.OpenDirectoryToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(155, 48)
        '
        'ExportToolStripMenuItem1
        '
        Me.ExportToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcelToolStripMenuItem1, Me.AccessToolStripMenuItem1, Me.XMLToolStripMenuItem1, Me.NotepadToolStripMenuItem})
        Me.ExportToolStripMenuItem1.Name = "ExportToolStripMenuItem1"
        Me.ExportToolStripMenuItem1.Size = New System.Drawing.Size(154, 22)
        Me.ExportToolStripMenuItem1.Text = "Export"
        '
        'ExcelToolStripMenuItem1
        '
        Me.ExcelToolStripMenuItem1.Name = "ExcelToolStripMenuItem1"
        Me.ExcelToolStripMenuItem1.Size = New System.Drawing.Size(120, 22)
        Me.ExcelToolStripMenuItem1.Text = "Excel"
        '
        'AccessToolStripMenuItem1
        '
        Me.AccessToolStripMenuItem1.Name = "AccessToolStripMenuItem1"
        Me.AccessToolStripMenuItem1.Size = New System.Drawing.Size(120, 22)
        Me.AccessToolStripMenuItem1.Text = "Access"
        '
        'XMLToolStripMenuItem1
        '
        Me.XMLToolStripMenuItem1.Name = "XMLToolStripMenuItem1"
        Me.XMLToolStripMenuItem1.Size = New System.Drawing.Size(120, 22)
        Me.XMLToolStripMenuItem1.Text = "XML"
        '
        'NotepadToolStripMenuItem
        '
        Me.NotepadToolStripMenuItem.Name = "NotepadToolStripMenuItem"
        Me.NotepadToolStripMenuItem.Size = New System.Drawing.Size(120, 22)
        Me.NotepadToolStripMenuItem.Text = "Notepad"
        '
        'OpenDirectoryToolStripMenuItem
        '
        Me.OpenDirectoryToolStripMenuItem.Name = "OpenDirectoryToolStripMenuItem"
        Me.OpenDirectoryToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.OpenDirectoryToolStripMenuItem.Text = "Open Directory"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(570, 263)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.ListView1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "EDI 835 -> XML"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar2 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExcelToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AccessToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents XMLToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ExportToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExcelToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AccessToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents XMLToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotepadToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenDirectoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick
        Me.ListView1.ListViewItemSorter = New ListViewItemComparer(e.Column)
        ListView1.Sort()

    End Sub

    Private Sub ListView1_ItemActivate(sender As Object, e As EventArgs) Handles ListView1.ItemActivate
        On Error Resume Next
        ' TODO
        ' Might be best to use Ref to ListView1 instead of casting sender
        ' Dim item = ListView1.SelectedItems(0).Text
        ' Want the file name found in Column 1
        Dim item = CType(sender, ListView).SelectedItems(0)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = item.Text

    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        Call LoadModelForm()
    End Sub
End Class

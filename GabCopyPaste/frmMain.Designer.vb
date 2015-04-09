<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.notification = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStripMain = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ShowMainWindowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tooltip = New System.Windows.Forms.ToolTip(Me.components)
        Me.rtfPreview = New System.Windows.Forms.RichTextBox()
        Me.SplitContainerBottom = New System.Windows.Forms.SplitContainer()
        Me.webPreview = New System.Windows.Forms.WebBrowser()
        Me.picPreview = New System.Windows.Forms.PictureBox()
        Me.SplitContainerMain = New System.Windows.Forms.SplitContainer()
        Me.SplitContainerTop = New System.Windows.Forms.SplitContainer()
        Me.tableSlot = New System.Windows.Forms.TableLayoutPanel()
        Me.tableQueue = New System.Windows.Forms.TableLayoutPanel()
        Me.ContextMenuStripMain.SuspendLayout()
        Me.SplitContainerBottom.Panel1.SuspendLayout()
        Me.SplitContainerBottom.Panel2.SuspendLayout()
        Me.SplitContainerBottom.SuspendLayout()
        CType(Me.picPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerMain.Panel1.SuspendLayout()
        Me.SplitContainerMain.Panel2.SuspendLayout()
        Me.SplitContainerMain.SuspendLayout()
        Me.SplitContainerTop.Panel1.SuspendLayout()
        Me.SplitContainerTop.Panel2.SuspendLayout()
        Me.SplitContainerTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtLog
        '
        Me.txtLog.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLog.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtLog.Location = New System.Drawing.Point(0, 0)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ReadOnly = True
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(179, 232)
        Me.txtLog.TabIndex = 1
        '
        'notification
        '
        Me.notification.BalloonTipTitle = "== GabCopyPaste =="
        Me.notification.ContextMenuStrip = Me.ContextMenuStripMain
        Me.notification.Icon = CType(resources.GetObject("notification.Icon"), System.Drawing.Icon)
        Me.notification.Text = "GabCopyPaste"
        Me.notification.Visible = True
        '
        'ContextMenuStripMain
        '
        Me.ContextMenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptionsToolStripMenuItem, Me.ToolStripSeparator2, Me.ShowMainWindowToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitToolStripMenuItem})
        Me.ContextMenuStripMain.Name = "ContextMenuStripMain"
        Me.ContextMenuStripMain.Size = New System.Drawing.Size(179, 82)
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(175, 6)
        '
        'ShowMainWindowToolStripMenuItem
        '
        Me.ShowMainWindowToolStripMenuItem.Name = "ShowMainWindowToolStripMenuItem"
        Me.ShowMainWindowToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ShowMainWindowToolStripMenuItem.Text = "&Show main window"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(175, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ExitToolStripMenuItem.Text = "&Exit"
        '
        'rtfPreview
        '
        Me.rtfPreview.AcceptsTab = True
        Me.rtfPreview.BackColor = System.Drawing.Color.White
        Me.rtfPreview.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.rtfPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtfPreview.Location = New System.Drawing.Point(0, 0)
        Me.rtfPreview.Name = "rtfPreview"
        Me.rtfPreview.ReadOnly = True
        Me.rtfPreview.Size = New System.Drawing.Size(313, 232)
        Me.rtfPreview.TabIndex = 4
        Me.rtfPreview.Text = ""
        Me.rtfPreview.WordWrap = False
        '
        'SplitContainerBottom
        '
        Me.SplitContainerBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerBottom.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerBottom.Name = "SplitContainerBottom"
        '
        'SplitContainerBottom.Panel1
        '
        Me.SplitContainerBottom.Panel1.Controls.Add(Me.rtfPreview)
        Me.SplitContainerBottom.Panel1.Controls.Add(Me.webPreview)
        Me.SplitContainerBottom.Panel1.Controls.Add(Me.picPreview)
        '
        'SplitContainerBottom.Panel2
        '
        Me.SplitContainerBottom.Panel2.Controls.Add(Me.txtLog)
        Me.SplitContainerBottom.Size = New System.Drawing.Size(504, 234)
        Me.SplitContainerBottom.SplitterDistance = 315
        Me.SplitContainerBottom.SplitterWidth = 8
        Me.SplitContainerBottom.TabIndex = 5
        '
        'webPreview
        '
        Me.webPreview.AllowNavigation = False
        Me.webPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.webPreview.IsWebBrowserContextMenuEnabled = False
        Me.webPreview.Location = New System.Drawing.Point(0, 0)
        Me.webPreview.MinimumSize = New System.Drawing.Size(20, 20)
        Me.webPreview.Name = "webPreview"
        Me.webPreview.ScriptErrorsSuppressed = True
        Me.webPreview.Size = New System.Drawing.Size(313, 232)
        Me.webPreview.TabIndex = 0
        Me.webPreview.Visible = False
        '
        'picPreview
        '
        Me.picPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.picPreview.Location = New System.Drawing.Point(0, 0)
        Me.picPreview.Name = "picPreview"
        Me.picPreview.Size = New System.Drawing.Size(313, 232)
        Me.picPreview.TabIndex = 5
        Me.picPreview.TabStop = False
        Me.picPreview.Visible = False
        '
        'SplitContainerMain
        '
        Me.SplitContainerMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainerMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.SplitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerMain.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerMain.Name = "SplitContainerMain"
        Me.SplitContainerMain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainerMain.Panel1
        '
        Me.SplitContainerMain.Panel1.Controls.Add(Me.SplitContainerTop)
        '
        'SplitContainerMain.Panel2
        '
        Me.SplitContainerMain.Panel2.Controls.Add(Me.SplitContainerBottom)
        Me.SplitContainerMain.Size = New System.Drawing.Size(504, 662)
        Me.SplitContainerMain.SplitterDistance = 420
        Me.SplitContainerMain.SplitterWidth = 8
        Me.SplitContainerMain.TabIndex = 6
        '
        'SplitContainerTop
        '
        Me.SplitContainerTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainerTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerTop.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerTop.Name = "SplitContainerTop"
        '
        'SplitContainerTop.Panel1
        '
        Me.SplitContainerTop.Panel1.Controls.Add(Me.tableSlot)
        '
        'SplitContainerTop.Panel2
        '
        Me.SplitContainerTop.Panel2.Controls.Add(Me.tableQueue)
        Me.SplitContainerTop.Size = New System.Drawing.Size(504, 420)
        Me.SplitContainerTop.SplitterDistance = 245
        Me.SplitContainerTop.SplitterWidth = 8
        Me.SplitContainerTop.TabIndex = 4
        '
        'tableSlot
        '
        Me.tableSlot.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.tableSlot.ColumnCount = 1
        Me.tableSlot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableSlot.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tableSlot.Location = New System.Drawing.Point(0, 0)
        Me.tableSlot.Name = "tableSlot"
        Me.tableSlot.RowCount = 10
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableSlot.Size = New System.Drawing.Size(243, 418)
        Me.tableSlot.TabIndex = 3
        '
        'tableQueue
        '
        Me.tableQueue.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.tableQueue.ColumnCount = 1
        Me.tableQueue.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableQueue.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tableQueue.Location = New System.Drawing.Point(0, 0)
        Me.tableQueue.Name = "tableQueue"
        Me.tableQueue.RowCount = 10
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.tableQueue.Size = New System.Drawing.Size(249, 418)
        Me.tableQueue.TabIndex = 4
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(504, 662)
        Me.Controls.Add(Me.SplitContainerMain)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "GabCopyPaste"
        Me.ContextMenuStripMain.ResumeLayout(False)
        Me.SplitContainerBottom.Panel1.ResumeLayout(False)
        Me.SplitContainerBottom.Panel2.ResumeLayout(False)
        Me.SplitContainerBottom.Panel2.PerformLayout()
        Me.SplitContainerBottom.ResumeLayout(False)
        CType(Me.picPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerMain.Panel1.ResumeLayout(False)
        Me.SplitContainerMain.Panel2.ResumeLayout(False)
        Me.SplitContainerMain.ResumeLayout(False)
        Me.SplitContainerTop.Panel1.ResumeLayout(False)
        Me.SplitContainerTop.Panel2.ResumeLayout(False)
        Me.SplitContainerTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents notification As System.Windows.Forms.NotifyIcon
    Friend WithEvents tooltip As System.Windows.Forms.ToolTip
    Friend WithEvents rtfPreview As System.Windows.Forms.RichTextBox
    Friend WithEvents SplitContainerBottom As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainerMain As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainerTop As System.Windows.Forms.SplitContainer
    Friend WithEvents tableSlot As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents tableQueue As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents webPreview As System.Windows.Forms.WebBrowser
    Friend WithEvents picPreview As System.Windows.Forms.PictureBox
    Friend WithEvents ContextMenuStripMain As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowMainWindowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator

End Class

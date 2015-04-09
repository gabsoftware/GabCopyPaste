Imports GabSoftware.Utils
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging
Imports System.IO

Public Class frmMain


    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessage", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As UInteger, ByVal wParam As Integer, ByVal lParam As StringBuilder) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function SendMessageTimeout(ByVal windowHandle As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal flags As SendMessageTimeoutFlags, ByVal timeout As UInteger, ByRef result As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)> _
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, <[In](), Out()> ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, EntryPoint:="GetGUIThreadInfo", SetLastError:=True)> _
    Public Shared Function GetGUIThreadInfo(ByVal idThread As Integer, <[In](), Out()> ByRef lpgui As GUITHREADINFO) As Boolean
    End Function

    '<DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    'Public Shared Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    'End Function

    <DllImport("kernel32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Sub SetLastError(ByVal errorCode As Integer)
    End Sub

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function AttachThreadInput(ByVal idAttach As System.UInt32, ByVal idAttachTo As System.UInt32, ByVal fAttach As Boolean) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetFocus() As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function AddClipboardFormatListener(ByVal hwnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function RemoveClipboardFormatListener(ByVal hwnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetClipboardViewer(ByVal hWndNewViewer As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function ChangeClipboardChain(ByVal hWndRemove As IntPtr, ByVal hWndNewNext As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenClipboard(ByVal hWndNewOwner As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function EmptyClipboard() As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CloseClipboard() As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetOpenClipboardWindow() As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetClipboardOwner() As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetClipboardData(ByVal format As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CountClipboardFormats() As Integer
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function EnumClipboardFormats(ByVal format As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetClipboardFormatName(ByVal format As UInteger, <[In](), Out()> ByRef lpString As StringBuilder, ByVal cchMax As Integer) As Integer
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetPriorityClipboardFormat(ByVal paFormatPriorityList() As UInteger, ByVal cFormats As Integer) As Integer
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetClipboardData(ByVal format As UInteger, ByVal hMem As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Friend Shared Function GlobalLock(ByVal handle As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Friend Shared Function GlobalUnlock(ByVal handle As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)> _
    Friend Shared Function GlobalSize(ByVal hMem As IntPtr) As UIntPtr 'UIntPtr remplace le type SIZE_T (unsigned int) (32bits/64bits pas pareils)
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True, SetLastError:=True)> _
    Public Shared Function GlobalAlloc(ByVal uFlags As GlobalAllocFlags, ByVal dwBytes As UIntPtr) As IntPtr 'UIntPtr remplace le type SIZE_T (unsigned int) (32bits/64bits pas pareils)
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True, SetLastError:=True)> _
    Public Shared Function GlobalFree(ByVal handle As IntPtr) As IntPtr
    End Function

    Const WM_COPY = &H301
    Const WM_PASTE = &H302
    Const WM_GETTEXT = &HD
    Const WM_CLIPBOARDUPDATE = &H31D
    Const WM_DRAWCLIPBOARD = &H308
    Const WM_CHANGECBCHAIN = &H30D

    <StructLayout(LayoutKind.Sequential)>
    Public Structure GUITHREADINFO
        Public cbSize As UInteger
        Public flags As GuiThreadInfoFlags
        Public hwndActive As IntPtr
        Public hwndFocus As IntPtr
        Public hwndCapture As IntPtr
        Public hwndMenuOwner As IntPtr
        Public hwndMoveSize As IntPtr
        Public hwndCaret As IntPtr
        Public rcCaret As RECT
    End Structure

    Public Enum GuiThreadInfoFlags As UInteger
        GUI_CARETBLINKING = &H1
        GUI_INMENUMODE = &H4
        GUI_INMOVESIZE = &H2
        GUI_POPUPMENUMODE = &H10
        GUI_SYSTEMMENUMODE = &H8
        GUI_16BITTASK = &H20 'winver >= 5.01
    End Enum

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Public Enum PredefinedClipboardFormats As UInteger

        CF_TEXT = &H1           'Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
        CF_BITMAP = &H2         'A handle to a bitmap (HBITMAP).
        CF_METAFILEPICT = &H3   'Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
        CF_SYLK = &H4           'Microsoft Symbolic Link (SYLK) format.
        CF_DIF = &H5            'Software Arts' Data Interchange Format.
        CF_TIFF = &H6           'Tagged-image file format.
        CF_OEMTEXT = &H7        'Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
        CF_DIB = &H8            'A memory object containing a BITMAPINFO structure followed by the bitmap bits.
        CF_PALETTE = &H9        'Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.
        'If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the SelectPalette and RealizePalette functions to realize (compare) any other data in the clipboard against that logical palette.
        'When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format.
        CF_PENDATA = &HA        'Data for the pen extensions to the Microsoft Windows for Pen Computing.
        CF_RIFF = &HB           'Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
        CF_WAVE = &HC           'Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
        CF_UNICODETEXT = &HD    'Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
        CF_ENHMETAFILE = &HE    'A handle to an enhanced metafile (HENHMETAFILE).
        CF_HDROP = &HF          'A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
        CF_LOCALE = &H10        'The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text. 
        'An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text.
        'Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead.
        'The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion.
        CF_DIBV5 = &H11 'A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
        CF_MAX = &H11 'idem CF_DIBV5
        CF_OWNERDISPLAY = &H80  'Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
        CF_DSPTEXT = &H81       'Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
        CF_DSPBITMAP = &H82     'Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
        CF_DSPMETAFILEPICT = &H83   'Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
        CF_DSPENHMETAFILE = &H8E    'Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
        CF_PRIVATEFIRST = &H200  'Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
        CF_PRIVATELAST = &H2FF   'See CF_PRIVATEFIRST.
        CF_GDIOBJFIRST = &H300   'Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.
        'Handles associated with clipboard formats in this range are not automatically deleted using the GlobalFree function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag.
        CF_GDIOBJLAST = &H3FF 'See CF_GDIOBJFIRST.


        'not standards but quite common
        CF_FILENAME = &HC006
        CF_FILENAMEW = &HC007
        CF_DATAOBJECT = &HC009
        CF_OLEPRIVATEDATA = &HC013
        CF_SHELLIDLIST = &HC048
        CF_HIDDENTEXTBANNERFORMAT = &HC070
        CF_RICHTEXTFORMAT = &HC087
        CF_SHELLOBJECTOFFSETS = &HC099
        CF_PREFERREDDROPEFFECT = &HC0A2
        CF_RICHTEXTFORMAT2 = &HC0C2


    End Enum

    Public Enum SendMessageTimeoutFlags As UInteger
        SMTO_NORMAL = &H0  'The calling thread is not prevented from processing other requests while waiting for the function to return.
        SMTO_BLOCK = &H1 'Prevents the calling thread from processing any other requests until the function returns.
        SMTO_ABORTIFHUNG = &H2 'The function returns without waiting for the time-out period to elapse if the receiving thread appears to not respond or "hangs."
        SMTO_NOTIMEOUTIFNOTHUNG = &H8  'The function does not enforce the time-out period as long as the receiving thread is processing messages.
        SMTO_ERRORONEXIT = &H20 'The function should return 0 if the receiving window is destroyed or its owning thread dies while the message is being processed.
    End Enum

    Public Enum GlobalAllocFlags As UInteger
        GMEM_FIXED = &H0   'Allocates fixed memory. The return value is a pointer.
        GMEM_MOVEABLE = &H2    'Allocates movable memory. Memory blocks are never moved in physical memory, but they can be moved within the default heap. 
        'The return value is a handle to the memory object. To translate the handle into a pointer, use the GlobalLock function.
        'This value cannot be combined with GMEM_FIXED.

        GMEM_ZEROINIT = &H40    'Initializes memory contents to zero.
        GPTR = &H40 'Combines GMEM_FIXED and GMEM_ZEROINIT.
        GHND = &H42 'Combines GMEM_MOVEABLE and GMEM_ZEROINIT.
    End Enum


    ''' <summary>
    ''' objet qui capture les touches du clavier
    ''' </summary>
    ''' <remarks></remarks>
    Friend WithEvents keyHook As GabKeyboardHook
    Private Shared slots As New List(Of List(Of ClipData))(10)
    Private Shared queue As New Queue(Of List(Of ClipData))(10)
    Private Shared nextClipboardViewer As IntPtr
    Private Shared supportsNewAPI As Boolean = False
    Private Shared listens As Boolean = False
    Private Shared ignore As Boolean = False
    Private Shared loaded As Boolean = False
    Private Shared autoconvert As Boolean = False
    Private Shared pasting As Boolean = False
    Private Shared forceClose As Boolean = False
    Private Shared dlgOpts As dlgOptions = Nothing
    Private Shared activatedFirstTime As Boolean = False

    Private Class ClipData
        Private _data As Object
        Private _format As String

        Public Property Data As Object
            Get
                Return _data
            End Get
            Set(ByVal value As Object)
                _data = value
            End Set
        End Property

        Public Property Format As String
            Get
                Return _format
            End Get
            Set(ByVal value As String)
                _format = value
            End Set
        End Property

        Public Sub New(ByVal thisData As Object, ByVal thisFormat As String)
            _data = thisData
            _format = thisFormat
        End Sub

    End Class


    'Private Class BinaryClipData
    '    Private _data As Byte()
    '    Private _format As UInteger
    '    Private _formatName As String

    '    Public Property Data As Byte()
    '        Get
    '            Return _data
    '        End Get
    '        Set(ByVal value As Byte())
    '            _data = value
    '        End Set
    '    End Property

    '    Public Property Format As UInteger
    '        Get
    '            Return _format
    '        End Get
    '        Set(ByVal value As UInteger)
    '            _format = value
    '        End Set
    '    End Property

    '    Public Property FormatName As String
    '        Get
    '            Return _formatName
    '        End Get
    '        Set(ByVal value As String)
    '            _formatName = value
    '        End Set
    '    End Property

    '    Public Sub New(ByVal thisData As Object, ByVal thisFormat As UInteger, ByVal thisFormatName As String)
    '        _data = thisData
    '        _format = thisFormat
    '        _formatName = thisFormatName
    '    End Sub

    'End Class

    Protected Overrides Sub WndProc(ByRef m As Message)
        'Dim kargs As KeyEventArgs

        If loaded Then
            If listens Then
                If supportsNewAPI Then



                    Select Case m.Msg
                        Case WM_CLIPBOARDUPDATE

                            ' IO.File.AppendAllText(Application.StartupPath & "\prout.txt", "WM_CLIPBOARDUPDATE message: " & m.Msg & vbNewLine & _
                            '"LParam: " & m.LParam.ToString & vbNewLine & _
                            '"WParam: " & m.WParam.ToString & vbNewLine & _
                            '"Result: " & m.Result.ToString & vbNewLine & vbNewLine)

                            ClipboardDataChanged()

                        Case Else

                            ' IO.File.AppendAllText(Application.StartupPath & "\prout.txt", "Message: " & m.Msg & vbNewLine & _
                            '"LParam: " & m.LParam.ToString & vbNewLine & _
                            '"WParam: " & m.WParam.ToString & vbNewLine & _
                            '"Result: " & m.Result.ToString & vbNewLine & vbNewLine)

                            MyBase.WndProc(m)
                    End Select
                Else
                    Select Case m.Msg
                        Case WM_DRAWCLIPBOARD


                            ClipboardDataChanged()


                            SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam)
                            SetLastError(0)
                        Case WM_CHANGECBCHAIN
                            If m.WParam = nextClipboardViewer Then
                                nextClipboardViewer = m.LParam
                            Else
                                SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam)
                            End If
                        Case Else
                            MyBase.WndProc(m)
                    End Select
                End If
            Else
                MyBase.WndProc(m)
            End If
        Else
            MyBase.WndProc(m)
        End If
        'MyBase.WndProc(m)
    End Sub

    Private Sub listenToClipboard()
        Try
            'ajoute la fenêtre comme moniteur du presse papier
            supportsNewAPI = True
            If Not AddClipboardFormatListener(Me.Handle) Then
                'AddClipboardFormatListener failed...
                txtLog.Text = "Could not set the application as a Clipboard format listener!" & vbNewLine & txtLog.Text
                listens = False
            Else
                listens = True
            End If
        Catch ex As Exception
            supportsNewAPI = False
            SetLastError(0)
            nextClipboardViewer = SetClipboardViewer(Me.Handle)
            If nextClipboardViewer = IntPtr.Zero And Marshal.GetLastWin32Error <> 0 Then
                'SetClipboardViewer failed...
                txtLog.Text = "Could not set the application as a Clipboard viewer!" & vbNewLine & txtLog.Text
                listens = False
            Else
                listens = True
            End If
        End Try
    End Sub

    Private Sub stopListenToClipboard()
        'enlève la fenêtre de la liste des moniteurs du presse papier
        If listens Then
            If supportsNewAPI Then
                RemoveClipboardFormatListener(Me.Handle)
            Else
                ChangeClipboardChain(Me.Handle, nextClipboardViewer)
            End If
        End If
    End Sub





    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'on désinstalle le hook du clavier
        Me.keyHook.Stop(True, False)

        For Each L As List(Of ClipData) In slots
            If L IsNot Nothing Then
                L.Clear()
            End If
        Next

        slots.Clear()
        queue.Clear()

        stopListenToClipboard()

    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not forceClose Then
            MinimizeToTray()
            e.Cancel = True
        End If
    End Sub

    Private Sub PostInitializeComponents()
        Dim p As Panel
        'Dim btnSave As Button
        'Dim btnPreview As Button
        Dim btnDelete As Button
        Dim btnCopy As Button
        Dim btnTransfer As Button
        Dim l As Label

        'Ajoute les contrôles
        Me.SuspendLayout()
        Me.tableSlot.SuspendLayout()
        Me.tableQueue.SuspendLayout()

        'controles pour les slots
        For i As Integer = 0 To 9
            p = New Panel()
            p.SuspendLayout()
            'btnSave = New Button()
            'btnPreview = New Button()
            btnDelete = New Button()
            btnCopy = New Button()
            l = New Label()

            '
            'tableSlot
            '
            Me.tableSlot.Controls.Add(p, 0, i)

            '
            'panel
            '
            'p.Controls.Add(btnSave)
            'p.Controls.Add(btnPreview)
            p.Controls.Add(btnDelete)
            p.Controls.Add(btnCopy)
            p.Controls.Add(l)
            p.Dock = System.Windows.Forms.DockStyle.Fill
            p.Name = "panSlot_" & i
            If i = 0 Then
                p.BackColor = Color.FromArgb(&HB2, &HCC, &HFF)
            End If
            AddHandler p.MouseEnter, AddressOf panSlot_MouseEnter


            ''
            ''btnSlotSave
            ''
            'btnSave.Dock = System.Windows.Forms.DockStyle.Right
            'btnSave.Image = Global.GabCopyPaste.My.Resources.Resources.saveHS
            'btnSave.Location = New System.Drawing.Point(44, 0)
            'btnSave.Name = "btnSlotSave_" & i
            'btnSave.Size = New System.Drawing.Size(40, 45)
            'Me.tooltip.SetToolTip(btnSave, "Save to file")
            'btnSave.UseVisualStyleBackColor = True
            'btnSave.Visible = False
            'AddHandler btnSave.MouseEnter, AddressOf ctrlSlot_MouseEnter

            ''
            ''btnSlotPreview
            ''
            'btnPreview.Dock = System.Windows.Forms.DockStyle.Right
            'btnPreview.Image = Global.GabCopyPaste.My.Resources.Resources.ZoomHS
            'btnPreview.Location = New System.Drawing.Point(84, 0)
            'btnPreview.Name = "btnSlotPreview_" & i
            'btnPreview.Size = New System.Drawing.Size(40, 45)
            'Me.tooltip.SetToolTip(btnPreview, "Preview/edit")
            'btnPreview.UseVisualStyleBackColor = True
            'btnPreview.Visible = False
            'AddHandler btnPreview.MouseEnter, AddressOf ctrlSlot_MouseEnter

            '
            'btnSlotDelete
            '
            btnDelete.Dock = System.Windows.Forms.DockStyle.Right
            btnDelete.Image = Global.GabCopyPaste.My.Resources.Resources.DeleteHS
            btnDelete.Location = New System.Drawing.Point(124, 0)
            btnDelete.Name = "btnSlotDelete_" & i
            btnDelete.Size = New System.Drawing.Size(80, 45)
            Me.tooltip.SetToolTip(btnDelete, "Delete data in Slot")
            btnDelete.UseVisualStyleBackColor = True
            btnDelete.Visible = False
            AddHandler btnDelete.MouseEnter, AddressOf ctrlSlot_MouseEnter
            AddHandler btnDelete.Click, AddressOf btnSlotDelete_Click

            '
            'btnSlotCopy
            '
            btnCopy.Dock = System.Windows.Forms.DockStyle.Right
            btnCopy.Image = Global.GabCopyPaste.My.Resources.Resources.FillDownHS
            btnCopy.Location = New System.Drawing.Point(164, 0)
            btnCopy.Name = "btnSlotCopy_" & i
            btnCopy.Size = New System.Drawing.Size(80, 45)
            Me.tooltip.SetToolTip(btnCopy, "Place the data from this Slot in the Clipboard" & vbNewLine & "CTRL+SHIFT + " & i & " or CTRL+ALT+Numpad " & i)
            btnCopy.UseVisualStyleBackColor = True
            btnCopy.Visible = False
            AddHandler btnCopy.MouseEnter, AddressOf ctrlSlot_MouseEnter
            AddHandler btnCopy.Click, AddressOf btnSlotCopy_Click

            '
            'lblSlot
            '
            l.Dock = System.Windows.Forms.DockStyle.Left
            l.Location = New System.Drawing.Point(0, 0)
            l.Name = "lblSlot_" & i
            l.Size = New System.Drawing.Size(55, 45)
            l.TabIndex = 1
            l.Text = "Slot #" & i & vbNewLine & "(empty)"
            AddHandler l.MouseEnter, AddressOf ctrlSlot_MouseEnter

            p.ResumeLayout(False)

        Next
        Me.tableSlot.ResumeLayout(False)

        'controles pour la queue
        For i As Integer = 0 To 9
            p = New Panel()
            p.SuspendLayout()
            'btnSave = New Button()
            btnTransfer = New Button()
            'btnPreview = New Button()
            btnCopy = New Button()
            l = New Label()

            '
            'tableQueue
            '
            Me.tableQueue.Controls.Add(p, 0, i)

            '
            'panel
            '
            'p.Controls.Add(btnSave)
            'p.Controls.Add(btnPreview)
            p.Controls.Add(btnTransfer)
            p.Controls.Add(btnCopy)
            p.Controls.Add(l)
            p.Dock = System.Windows.Forms.DockStyle.Fill
            p.Name = "panQueue_" & i
            AddHandler p.MouseEnter, AddressOf panQueue_MouseEnter

            ''
            ''btnQueueSave
            ''
            'btnSave.Dock = System.Windows.Forms.DockStyle.Right
            'btnSave.Image = Global.GabCopyPaste.My.Resources.Resources.saveHS
            'btnSave.Location = New System.Drawing.Point(44, 0)
            'btnSave.Name = "btnQueueSave_" & i
            'btnSave.Size = New System.Drawing.Size(40, 45)
            'Me.tooltip.SetToolTip(btnSave, "Save to file")
            'btnSave.UseVisualStyleBackColor = True
            'btnSave.Visible = False
            'AddHandler btnSave.MouseEnter, AddressOf ctrlQueue_MouseEnter

            ''
            ''btnQueuePreview
            ''
            'btnPreview.Dock = System.Windows.Forms.DockStyle.Right
            'btnPreview.Image = Global.GabCopyPaste.My.Resources.Resources.ZoomHS
            'btnPreview.Location = New System.Drawing.Point(84, 0)
            'btnPreview.Name = "btnQueuePreview_" & i
            'btnPreview.Size = New System.Drawing.Size(40, 45)
            'Me.tooltip.SetToolTip(btnPreview, "Preview/edit")
            'btnPreview.UseVisualStyleBackColor = True
            'btnPreview.Visible = False
            'AddHandler btnPreview.MouseEnter, AddressOf ctrlQueue_MouseEnter

            '
            'btnQueueTransfer
            '
            btnTransfer.Dock = System.Windows.Forms.DockStyle.Right
            btnTransfer.Image = Global.GabCopyPaste.My.Resources.Resources.TransferHS
            btnTransfer.Location = New System.Drawing.Point(124, 0)
            btnTransfer.Name = "btnQueueTransfer_" & i
            btnTransfer.Size = New System.Drawing.Size(80, 45)
            Me.tooltip.SetToolTip(btnTransfer, "Transfer data to a Slot")
            btnTransfer.UseVisualStyleBackColor = True
            btnTransfer.Visible = False
            AddHandler btnTransfer.MouseEnter, AddressOf ctrlQueue_MouseEnter
            AddHandler btnTransfer.Click, AddressOf btnQueueTransfer_Click

            '
            'btnQueueCopy
            '
            btnCopy.Dock = System.Windows.Forms.DockStyle.Right
            btnCopy.Image = Global.GabCopyPaste.My.Resources.Resources.FillDownHS
            btnCopy.Location = New System.Drawing.Point(164, 0)
            btnCopy.Name = "btnQueueCopy_" & i
            btnCopy.Size = New System.Drawing.Size(80, 45)
            Me.tooltip.SetToolTip(btnCopy, "Place this data in the Clipboard")
            btnCopy.UseVisualStyleBackColor = True
            btnCopy.Visible = False
            AddHandler btnCopy.MouseEnter, AddressOf ctrlQueue_MouseEnter
            AddHandler btnCopy.Click, AddressOf btnQueueCopy_Click

            '
            'lblQueue
            '
            l.Dock = System.Windows.Forms.DockStyle.Left
            l.Location = New System.Drawing.Point(0, 0)
            l.Name = "lblQueue_" & i
            l.Size = New System.Drawing.Size(80, 45)
            l.TabIndex = 1
            l.Text = "History item #" & i & vbNewLine & "(empty)"
            AddHandler l.MouseEnter, AddressOf ctrlQueue_MouseEnter

            p.ResumeLayout(False)

        Next
        Me.tableQueue.ResumeLayout(False)

        Me.ResumeLayout(False)
        Me.PerformLayout()

        'On capture toutes les touches même si on a pas le focus
        Me.keyHook = New GabKeyboardHook(True)

        'initialise les slots
        For i As Integer = 0 To slots.Capacity - 1
            slots.Add(Nothing)
            queue.Enqueue(Nothing)
        Next

        listenToClipboard()

        loaded = True

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'MinimizeToTray()
        'loaded = True
        'MsgBox("loaded")

    End Sub

    Private Sub keyHook_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles keyHook.KeyDown

        Me.ProcessKeys(e, Me.ContainsFocus)

    End Sub

    Private Sub ProcessKeys(ByVal touche As System.Windows.Forms.KeyEventArgs, ByVal hasFocus As Boolean)

        'txtLog.Text += vbNewLine & touche.KeyCode.ToString & " - " & touche.KeyData.ToString & " - " & touche.KeyValue.ToString
        'MsgBox(touche.KeyCode.ToString & " - " & touche.KeyData.ToString & " - " & touche.KeyValue.ToString & " - " & touche.Control.ToString & " - " & touche.Alt.ToString & " - " & touche.Shift.ToString)
        ''touche.Handled = True
        'Return

        'If touche.Control And touche.Alt And _
        '    (touche.KeyCode = Keys.D1 Or _
        '     touche.KeyCode = Keys.D2 Or _
        '     touche.KeyCode = Keys.D3 Or _
        '     touche.KeyCode = Keys.D4 Or _
        '     touche.KeyCode = Keys.D5 Or _
        '     touche.KeyCode = Keys.D6 Or _
        '     touche.KeyCode = Keys.D7 Or _
        '     touche.KeyCode = Keys.D8 Or _
        '     touche.KeyCode = Keys.D9 Or _
        '     touche.KeyCode = Keys.D0) Then 'évite de réagir à AltGr+Dx
        '    Return
        'End If


        Select Case True

            Case touche.Control And touche.Alt And _
            (touche.KeyCode = Keys.D1 Or _
             touche.KeyCode = Keys.D2 Or _
             touche.KeyCode = Keys.D3 Or _
             touche.KeyCode = Keys.D4 Or _
             touche.KeyCode = Keys.D5 Or _
             touche.KeyCode = Keys.D6 Or _
             touche.KeyCode = Keys.D7 Or _
             touche.KeyCode = Keys.D8 Or _
             touche.KeyCode = Keys.D9 Or _
             touche.KeyCode = Keys.D0) 'évite de réagir à AltGr+Dx
                Return
                Exit Select

            Case touche.Control And touche.Shift And touche.KeyCode = Keys.V
                touche.Handled = True
                ShowMainWindow()
                Exit Select

            Case touche.KeyCode = Keys.V And touche.Control
                'touche.Handled = True
                pasting = True
                Exit Select

                '
                'PASTE : CTRL + SHIFT + (Fx | Dx) | CTRL + ALT + Numpadx
                '
            Case ((touche.KeyCode = Keys.F1 Or touche.KeyCode = Keys.D1) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad1 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(1)
                Exit Select

            Case ((touche.KeyCode = Keys.F2 Or touche.KeyCode = Keys.D2) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad2 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(2)
                Exit Select

            Case ((touche.KeyCode = Keys.F3 Or touche.KeyCode = Keys.D3) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad3 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(3)
                Exit Select

            Case ((touche.KeyCode = Keys.F4 Or touche.KeyCode = Keys.D4) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad4 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(4)
                Exit Select

            Case ((touche.KeyCode = Keys.F5 Or touche.KeyCode = Keys.D5) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad5 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(5)
                Exit Select

            Case ((touche.KeyCode = Keys.F6 Or touche.KeyCode = Keys.D6) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad6 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(6)
                Exit Select

            Case ((touche.KeyCode = Keys.F7 Or touche.KeyCode = Keys.D7) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad7 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(7)
                Exit Select

            Case ((touche.KeyCode = Keys.F8 Or touche.KeyCode = Keys.D8) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad8 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(8)
                Exit Select

            Case ((touche.KeyCode = Keys.F9 Or touche.KeyCode = Keys.D9) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad9 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(9)
                Exit Select

            Case ((touche.KeyCode = Keys.F10 Or touche.KeyCode = Keys.Oem7 Or touche.KeyCode = Keys.D0) And touche.Control And touche.Shift) Or (touche.KeyCode = Keys.NumPad0 And touche.Control And touche.Alt)
                touche.Handled = True
                PasteFromSlot(0)
                Exit Select



                '
                ' COPY : CTRL + (Numpad | Dx)
                '
            Case (touche.KeyCode = Keys.NumPad1 Or touche.KeyCode = Keys.D1) And touche.Control
                touche.Handled = True
                CopyToSlot(1)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad2 Or touche.KeyCode = Keys.D2) And touche.Control
                touche.Handled = True
                CopyToSlot(2)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad3 Or touche.KeyCode = Keys.D3) And touche.Control
                touche.Handled = True
                CopyToSlot(3)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad4 Or touche.KeyCode = Keys.D4) And touche.Control
                touche.Handled = True
                CopyToSlot(4)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad5 Or touche.KeyCode = Keys.D5) And touche.Control
                touche.Handled = True
                CopyToSlot(5)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad6 Or touche.KeyCode = Keys.D6) And touche.Control
                touche.Handled = True
                CopyToSlot(6)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad7 Or touche.KeyCode = Keys.D7) And touche.Control
                touche.Handled = True
                CopyToSlot(7)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad8 Or touche.KeyCode = Keys.D8) And touche.Control
                touche.Handled = True
                CopyToSlot(8)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad9 Or touche.KeyCode = Keys.D9) And touche.Control
                touche.Handled = True
                CopyToSlot(9)
                Exit Select

            Case (touche.KeyCode = Keys.NumPad0 Or touche.KeyCode = Keys.Oem7 Or touche.KeyCode = Keys.D0) And touche.Control
                touche.Handled = True
                CopyToSlot(0)
                Exit Select

        End Select


    End Sub

    Private Sub ClipboardDataChanged()
        AddToQueue()
        UpdateQueueUI()
    End Sub

    Private Sub AddToQueue()
        Dim io As IDataObject = Nothing
        Dim newdata As New List(Of ClipData)
        Dim nextdata As New List(Of ClipData)
        Dim olddata As List(Of ClipData)
        Dim identical As Boolean = True
        Dim bnew() As Byte
        Dim bnext() As Byte

        If pasting Then
            pasting = False
            txtLog.Text = "AddToQueue avoided" & vbNewLine & txtLog.Text
            Return
        End If

        Dim noformat As Boolean = True
        txtLog.Text = "Data added to History" & vbNewLine & txtLog.Text

        stopListenToClipboard()

        io = Clipboard.GetDataObject()

        If io Is Nothing Then
            txtLog.Text = "Clipboard was empty! Aborted." & vbNewLine & txtLog.Text
            Return
        End If

        For Each f As String In io.GetFormats(autoconvert)
            noformat = False
            If io.GetDataPresent(f, autoconvert) Then

                If Not "EnhancedMetafile;MetaFilePict;Link Source".Contains(f) Then
                    newdata.Add(New ClipData(io.GetData(f, autoconvert), f))
                    txtLog.Text = "Added format: " & f & vbNewLine & txtLog.Text
                Else
                    txtLog.Text = "Format '" & f & "' has been ignored" & vbNewLine & txtLog.Text
                End If

            End If
        Next
        If noformat Then
            txtLog.Text = "No format! Aborted." & vbNewLine & txtLog.Text
            listenToClipboard()
            Return
        End If

        'try to compare the previous entry to the new one, and do not add a new entry if identical
        nextdata = queue.ToArray()(9)
        Try
            If nextdata IsNot Nothing And newdata IsNot Nothing Then
                If nextdata.Count = newdata.Count Then
                    For i As Integer = 0 To nextdata.Count - 1
                        If identical = False Then
                            Exit For
                        End If
                        If nextdata(i).Format = newdata(i).Format Then

                            If Not "[Embed Source];[Rich Text Format];[HTML Format];[ObjectLink];[Shell Object Offsets]".Contains("[" & nextdata(i).Format & "]") Then 'unreliable formats for comparison
                                If TypeOf nextdata(i).Data Is System.String Then

                                    'System.IO.File.WriteAllText(Application.StartupPath & "\bnext." & nextdata(i).Format.Replace(" ", "") & ".txt", DirectCast(nextdata(i).Data, String))
                                    'System.IO.File.WriteAllText(Application.StartupPath & "\bnew." & newdata(i).Format.Replace(" ", "") & ".txt", DirectCast(newdata(i).Data, String))

                                    If DirectCast(nextdata(i).Data, String).Length = DirectCast(newdata(i).Data, String).Length Then
                                        If String.Compare(DirectCast(nextdata(i).Data, String), DirectCast(newdata(i).Data, String)) <> 0 Then
                                            identical = False
                                            Exit For
                                        End If
                                    Else
                                        identical = False
                                        Exit For
                                    End If

                                ElseIf TypeOf nextdata(i).Data Is MemoryStream Then
                                    If DirectCast(nextdata(i).Data, MemoryStream).Length = DirectCast(newdata(i).Data, MemoryStream).Length Then
                                        bnext = DirectCast(nextdata(i).Data, MemoryStream).ToArray()
                                        bnew = DirectCast(newdata(i).Data, MemoryStream).ToArray()

                                        'System.IO.File.WriteAllBytes(Application.StartupPath & "\bnext." & nextdata(i).Format.Replace(" ", "") & ".bin", bnext)
                                        'System.IO.File.WriteAllBytes(Application.StartupPath & "\bnew." & newdata(i).Format.Replace(" ", "") & ".bin", bnew)

                                        For j As Integer = 0 To bnext.Length - 1
                                            If bnext(j).CompareTo(bnew(j)) <> 0 Then
                                                identical = False
                                                Exit For
                                            End If
                                        Next
                                    Else
                                        identical = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Else
                            identical = False
                            Exit For
                        End If
                    Next
                Else
                    identical = False
                End If
            Else
                identical = False
            End If
        Finally

        End Try
        If identical Then
            listenToClipboard()
            Return 'we do not need to add again something we already just added
        End If

        olddata = queue.Dequeue()
        If olddata IsNot Nothing Then
            olddata.Clear()
        End If

        queue.Enqueue(newdata)

        listenToClipboard()

        'notification.BalloonTipText = "Data added to History!"
        'notification.ShowBalloonTip(5000)

    End Sub

    Private Sub PasteFromQueue(ByVal position As Integer, Optional ByVal dontPaste As Boolean = False)

        Dim data As List(Of ClipData) = Nothing
        Dim dobj As DataObject = Nothing

        txtLog.Text = "History item #" & position & " selected" & vbNewLine & txtLog.Text

        data = queue.ToArray()(9 - position)

        If data IsNot Nothing AndAlso data.Count > 0 Then

            txtLog.Text = "Processing..." & vbNewLine & txtLog.Text

            stopListenToClipboard()

            dobj = New DataObject()
            Clipboard.Clear()

            For Each cd As ClipData In data
                dobj.SetData(cd.Format, autoconvert, cd.Data)
                txtLog.Text = "Added data to DataObject under format: " & cd.Format & vbNewLine & txtLog.Text
            Next

            Clipboard.SetDataObject(dobj, True, 10, 100)

            listenToClipboard()

            txtLog.Text = "DataObject added to clipboard" & vbNewLine & txtLog.Text
            notification.BalloonTipText = "Pasted from History item #" & position & vbNewLine & "Now use CTRL+V to paste!"
            notification.ShowBalloonTip(5000)
        Else
            txtLog.Text = "History item #" & position & " was empty!" & vbNewLine & txtLog.Text
            notification.BalloonTipText = "History item #" & position & " was empty!"
            notification.ShowBalloonTip(5000)
        End If

    End Sub

    Private Sub UpdateQueueUI()
        Dim p As Panel
        Dim l As Label
        Dim datalist() As List(Of ClipData)
        Dim realindex As Integer

        datalist = queue.ToArray()

        For i As Integer = 0 To datalist.Length - 1
            realindex = 9 - i
            p = DirectCast(Me.tableQueue.Controls("panQueue_" & i), Panel)
            l = DirectCast(p.Controls("lblQueue_" & i), Label)
            If datalist(realindex) Is Nothing Then
                l.Text = "History item #" & i & vbNewLine & "(empty)"
            ElseIf datalist(realindex).Count = 0 Then
                l.Text = "History item #" & i & vbNewLine & "(empty)"
            Else
                l.Text = "History item #" & i & vbNewLine & "IN USE"
            End If

        Next
    End Sub

    Private Sub CopyToSlot(ByVal NumSlot As Integer, Optional ByVal FromSlot As Integer = 0, Optional ByVal fromQueue As Integer = -1)

        Dim io As IDataObject = Nothing
        Dim data As New List(Of ClipData)
        Dim fromData As List(Of ClipData)
        Dim p As Panel
        Dim l As Label
        Dim noformat As Boolean = True

        'Dim tryAlternate As Boolean = False
        'Dim nb As Integer
        'Dim iFormat As UInteger
        'Dim formatName As StringBuilder = Nothing
        'Dim formatNameMaxSize As Integer = 800
        'Dim clipboardDataPointer As IntPtr
        'Dim buffer() As Byte
        'Dim is64 As Boolean = UIntPtr.Size = 8
        'Dim length As UIntPtr
        'Dim iLength64 As UInt64
        'Dim iLength32 As UInt32
        'Dim gLock As IntPtr
        'Dim clipdatabackup As List(Of BinaryClipData) = Nothing
        'Dim clipdatanew As List(Of BinaryClipData) = Nothing
        'Dim gui As GUITHREADINFO
        'Dim dwResult As IntPtr
        'Dim hWndTarget As IntPtr
        'Dim thisThreadId As Integer
        'Dim targetThreadId As Integer
        'Dim lpdwtargetPid As Integer
        'Dim hwndFocus As IntPtr
        'Dim res As Boolean
        'Dim lasterr As Integer


        If NumSlot >= 0 And NumSlot <= 9 Then

            stopListenToClipboard()

            If FromSlot = 0 Then
                txtLog.Text = "Copy to slot #" & NumSlot & vbNewLine & txtLog.Text


                ''saves clipboard
                'If OpenClipboard(IntPtr.Zero) Then
                '    nb = CountClipboardFormats()
                '    'if at least one format is present in the Clipboard
                '    If nb > 0 Then
                '        iFormat = EnumClipboardFormats(0)
                '        clipdatabackup = New List(Of BinaryClipData)
                '        While iFormat > 0
                '            'get the name of the format
                '            If formatName IsNot Nothing AndAlso formatName.Length > 0 Then
                '                formatName.Remove(0, formatName.Length)
                '            Else
                '                formatName = New StringBuilder(formatNameMaxSize)
                '            End If

                '            If [Enum].IsDefined(GetType(PredefinedClipboardFormats), iFormat) Then
                '                'this is a predefined format
                '                formatName.Append([Enum].GetName(GetType(PredefinedClipboardFormats), iFormat))
                '            Else
                '                'this is a custom format
                '                formatName.Append("(unknown format: " & iFormat & ")")
                '            End If

                '            'Get pointer to clipboard data in the selected format
                '            clipboardDataPointer = GetClipboardData(iFormat)
                '            If clipboardDataPointer <> IntPtr.Zero Then

                '                'Do a bunch of crap necessary to copy the data from the memory
                '                'the above pointer points at to a place we can access it.
                '                length = GlobalSize(clipboardDataPointer)
                '                If length <> UIntPtr.Zero Then
                '                    If is64 Then
                '                        iLength64 = length.ToUInt64
                '                    Else
                '                        iLength32 = length.ToUInt32
                '                    End If

                '                    gLock = GlobalLock(clipboardDataPointer)
                '                    If gLock <> IntPtr.Zero Then
                '                        'Init a buffer which will contain the clipboard data
                '                        buffer = New Byte(IIf(is64, iLength64 - 1L, iLength32 - 1)) {}

                '                        'Copy clipboard data to buffer
                '                        Marshal.Copy(gLock, buffer, 0, IIf(is64, iLength64, iLength32))

                '                        'for debugging purpose only, display the string contained by the byte array
                '                        'If iFormat = PredefinedClipboardFormats.CF_UNICODETEXT Then
                '                        '    MsgBox(Encoding.Unicode.GetString(buffer), , formatName.ToString())
                '                        'Else
                '                        '    MsgBox(Encoding.ASCII.GetString(buffer), , formatName.ToString())
                '                        'End If

                '                        'save the data
                '                        clipdatabackup.Add(New BinaryClipData(buffer, iFormat, formatName.ToString()))

                '                        'unlock
                '                        'GlobalUnlock is success if it returns zero AND GetLastError = NO_ERROR
                '                        res = GlobalUnlock(clipboardDataPointer)
                '                        lasterr = Marshal.GetLastWin32Error()
                '                        If res = False And lasterr = 0 Then
                '                            txtLog.Text = "Could not unlock the global lock!" & vbNewLine & txtLog.Text
                '                        End If
                '                    Else
                '                        txtLog.Text = "Global lock failed !" & vbNewLine & txtLog.Text
                '                    End If
                '                Else
                '                    txtLog.Text = "Data length is null !" & vbNewLine & txtLog.Text
                '                End If
                '            Else
                '                txtLog.Text = "Couldn't get clipboard data for format #" & iFormat & vbNewLine & txtLog.Text
                '            End If

                '            'get next format
                '            iFormat = EnumClipboardFormats(iFormat)
                '        End While
                '    Else
                '        txtLog.Text = "Clipboard seems to be empty" & vbNewLine & txtLog.Text
                '    End If

                '    If Not CloseClipboard() Then
                '        txtLog.Text = "Couldn't close clipboard !" & vbNewLine & txtLog.Text
                '    End If

                'Else
                '    txtLog.Text = "Couldn't open clipboard !" & vbNewLine & txtLog.Text
                'End If


                ''tries to copy !
                'thisThreadId = GetWindowThreadProcessId(Me.Handle, IntPtr.Zero)
                'targetThreadId = GetWindowThreadProcessId(hWndTarget, lpdwtargetPid)
                'gui.cbSize = Marshal.SizeOf(gui)
                'If GetGUIThreadInfo(targetThreadId, gui) Then
                '    If SendMessageTimeout(gui.hwndFocus, WM_COPY, 0, 0, SendMessageTimeoutFlags.SMTO_BLOCK, 500, dwResult) = 0 Then
                '        txtLog.Text = "SendMessageTimeout failed!" & vbNewLine & txtLog.Text
                '        tryAlternate = True
                '    Else
                '        Threading.Thread.Sleep(500) 'TODO: use the clipboard viewer API to get notified of the change !
                '    End If
                'Else
                '    hWndTarget = GetForegroundWindow()
                '    If hWndTarget <> IntPtr.Zero Then
                '        thisThreadId = GetWindowThreadProcessId(Me.Handle, IntPtr.Zero)
                '        targetThreadId = GetWindowThreadProcessId(hWndTarget, lpdwtargetPid)
                '        If Not thisThreadId <> targetThreadId And thisThreadId > 0 And targetThreadId > 0 Then
                '            If AttachThreadInput(thisThreadId, targetThreadId, True) Then
                '                hwndFocus = GetFocus()
                '                If hwndFocus <> IntPtr.Zero Then
                '                    If SendMessageTimeout(hwndFocus, WM_COPY, 0, 0, SendMessageTimeoutFlags.SMTO_BLOCK, 500, dwResult) = 0 Then
                '                        txtLog.Text = "SendMessageTimeout failed!" & vbNewLine & txtLog.Text
                '                        tryAlternate = True
                '                    Else
                '                        Threading.Thread.Sleep(500) 'TODO: use the clipboard viewer API to get notified of the change !
                '                    End If
                '                    If Not AttachThreadInput(thisThreadId, targetThreadId, False) Then
                '                        txtLog.Text = "Error in AttachThreadInput! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
                '                        tryAlternate = True
                '                    End If
                '                Else
                '                    txtLog.Text = "Error in GetFocus! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
                '                    AttachThreadInput(thisThreadId, targetThreadId, False)
                '                    tryAlternate = True
                '                End If
                '            Else
                '                txtLog.Text = "Error in AttachThreadInput! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
                '                tryAlternate = True
                '            End If
                '        Else
                '            txtLog.Text = "Error, thisThreadId is equal to the targetThreadId!" & vbNewLine & txtLog.Text
                '            tryAlternate = True
                '        End If
                '    Else
                '        tryAlternate = True
                '        txtLog.Text = "Error in GetForegroundWindow! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
                '    End If
                'End If


                'If tryAlternate Then
                '    txtLog.Text = "Using alternate way" & vbNewLine & txtLog.Text
                '    SendKeys.SendWait("^c") 'envoie CTRL+C à l'application qui a le focus, 5 fois, parce que parfois le copier ne se fait pas.
                '    SendKeys.SendWait("^c")
                '    SendKeys.SendWait("^c")
                '    SendKeys.SendWait("^c")
                '    SendKeys.SendWait("^c")
                'End If

            End If


            'saves the clipboard data in the specified slot


            'If OpenClipboard(IntPtr.Zero) Then
            '    nb = CountClipboardFormats()
            '    'if at least one format is present in the Clipboard
            '    If nb > 0 Then
            '        iFormat = EnumClipboardFormats(0)
            '        clipdatanew = New List(Of BinaryClipData)
            '        While iFormat > 0
            '            'get the name of the format
            '            If formatName IsNot Nothing AndAlso formatName.Length > 0 Then
            '                formatName.Remove(0, formatName.Length)
            '            Else
            '                formatName = New StringBuilder(formatNameMaxSize)
            '            End If

            '            If [Enum].IsDefined(GetType(PredefinedClipboardFormats), iFormat) Then
            '                'this is a predefined format
            '                formatName.Append([Enum].GetName(GetType(PredefinedClipboardFormats), iFormat))
            '            Else
            '                'this is a custom format
            '                formatName.Append("(unknown format: " & iFormat & ")")
            '            End If

            '            'Get pointer to clipboard data in the selected format
            '            clipboardDataPointer = GetClipboardData(iFormat)
            '            If clipboardDataPointer <> IntPtr.Zero Then

            '                'Do a bunch of crap necessary to copy the data from the memory
            '                'the above pointer points at to a place we can access it.
            '                length = GlobalSize(clipboardDataPointer)
            '                If length <> UIntPtr.Zero Then
            '                    If is64 Then
            '                        iLength64 = length.ToUInt64
            '                    Else
            '                        iLength32 = length.ToUInt32
            '                    End If

            '                    gLock = GlobalLock(clipboardDataPointer)
            '                    If gLock <> IntPtr.Zero Then
            '                        'Init a buffer which will contain the clipboard data
            '                        buffer = New Byte(IIf(is64, iLength64 - 1L, iLength32 - 1)) {}

            '                        'Copy clipboard data to buffer
            '                        Marshal.Copy(gLock, buffer, 0, IIf(is64, iLength64, iLength32))

            '                        'for debugging purpose only, display the string contained by the byte array
            '                        If iFormat = PredefinedClipboardFormats.CF_UNICODETEXT Then
            '                            MsgBox(Encoding.Unicode.GetString(buffer), , formatName.ToString())
            '                        Else
            '                            MsgBox(Encoding.ASCII.GetString(buffer), , formatName.ToString())
            '                        End If

            '                        'save the data
            '                        clipdatanew.Add(New BinaryClipData(buffer, iFormat, formatName.ToString()))

            '                        'unlock
            '                        'GlobalUnlock is success if it returns zero AND GetLastError = NO_ERROR
            '                        res = GlobalUnlock(clipboardDataPointer)
            '                        lasterr = Marshal.GetLastWin32Error()
            '                        If res = False And lasterr = 0 Then
            '                            MsgBox("Could not unlock the global lock!")
            '                        End If
            '                    Else
            '                        MsgBox("Global lock failed !")
            '                    End If
            '                Else
            '                    MsgBox("Data length is null !")
            '                End If
            '            Else
            '                MsgBox("Couldn't get clipboard data for format #" & iFormat)
            '            End If

            '            'get next format
            '            iFormat = EnumClipboardFormats(iFormat)
            '        End While
            '    Else
            '        MsgBox("Clipboard seems to be empty")
            '    End If

            '    If Not CloseClipboard() Then
            '        MsgBox("Couldn't close clipboard !")
            '    End If

            'Else
            '    MsgBox("Couldn't open clipboard !")
            'End If



            If fromQueue > -1 And fromQueue < 10 Then
                fromData = queue.ToArray()(9 - fromQueue)
                If fromData IsNot Nothing Then
                    io = New DataObject()
                    For Each cd As ClipData In fromData
                        io.SetData(cd.Format, cd.Data)
                    Next
                End If
            Else
                io = Clipboard.GetDataObject()
            End If



            If io Is Nothing Then
                txtLog.Text = "Clipboard was empty! Aborted." & vbNewLine & txtLog.Text
                Return
            End If

            For Each f As String In io.GetFormats(autoconvert)
                noformat = False
                If io.GetDataPresent(f, autoconvert) Then
                    If Not "EnhancedMetafile;MetaFilePict;Link Source".Contains(f) Then
                        data.Add(New ClipData(io.GetData(f, autoconvert), f))
                        txtLog.Text = "Added format: " & f & vbNewLine & txtLog.Text
                    Else
                        txtLog.Text = "Format '" & f & "' has been ignored" & vbNewLine & txtLog.Text
                    End If
                End If
            Next
            If noformat Then
                txtLog.Text = "No format! Aborted." & vbNewLine & txtLog.Text
                Return
            End If

            If slots(NumSlot) IsNot Nothing And NumSlot > 0 Then
                slots(NumSlot).Clear()
            End If
            slots(NumSlot) = data

            If NumSlot > 0 Then
                slots(0) = data
                p = DirectCast(Me.tableSlot.Controls("panSlot_0"), Panel)
                l = DirectCast(p.Controls("lblSlot_0"), Label)
                l.Text = "Slot #0" & vbNewLine & "IN USE"
            End If

            p = DirectCast(Me.tableSlot.Controls("panSlot_" & NumSlot), Panel)
            l = DirectCast(p.Controls("lblSlot_" & NumSlot), Label)
            l.Text = "Slot #" & NumSlot & vbNewLine & "IN USE"



            If FromSlot = 0 Then
                notification.BalloonTipText = "Data copied to Slot #" & NumSlot
                notification.ShowBalloonTip(5000)
            End If



            ''restores clipboard
            'If OpenClipboard(IntPtr.Zero) Then
            '    If EmptyClipboard() Then
            '        For Each bcd As BinaryClipData In clipdatabackup
            '            If is64 Then
            '                length = New UIntPtr(Convert.ToUInt64(bcd.Data.LongLength))
            '            Else
            '                length = New UIntPtr(Convert.ToUInt32(bcd.Data.Length))
            '            End If
            '            clipboardDataPointer = GlobalAlloc(GlobalAllocFlags.GMEM_MOVEABLE, length)
            '            If clipboardDataPointer <> IntPtr.Zero Then
            '                gLock = GlobalLock(clipboardDataPointer)
            '                If gLock <> IntPtr.Zero Then

            '                    Marshal.Copy(bcd.Data, 0, gLock, length)

            '                    'GlobalUnlock is success if it returns zero AND GetLastError = NO_ERROR
            '                    res = GlobalUnlock(clipboardDataPointer)
            '                    lasterr = Marshal.GetLastWin32Error()
            '                    If res = False And lasterr = 0 Then
            '                        If SetClipboardData(bcd.Format, clipboardDataPointer) = IntPtr.Zero Then
            '                            txtLog.Text = "SetClipboardData failed!" & vbNewLine & txtLog.Text
            '                            GlobalFree(clipboardDataPointer)
            '                        End If
            '                    Else
            '                        txtLog.Text = "Could not unlock the global lock!" & vbNewLine & txtLog.Text
            '                    End If

            '                Else
            '                    txtLog.Text = "GlobalLock failed!" & vbNewLine & txtLog.Text
            '                End If
            '            Else
            '                txtLog.Text = "GlobalAlloc failed!" & vbNewLine & txtLog.Text
            '            End If
            '        Next
            '    Else
            '        txtLog.Text = "Could not emptying the clipboard!" & vbNewLine & txtLog.Text
            '    End If
            '    If Not CloseClipboard() Then
            '        txtLog.Text = "Could not close the clipboard !" & vbNewLine & txtLog.Text
            '    End If
            'Else
            '    txtLog.Text = "Couldn't open clipboard !" & vbNewLine & txtLog.Text
            'End If

            listenToClipboard()

        Else
            txtLog.Text = "Bad slot number! Aborted." & vbNewLine & txtLog.Text
        End If


    End Sub

    Private Sub PasteFromSlot(ByVal NumSlot As Integer, Optional ByVal dontPaste As Boolean = False)

        Dim data As List(Of ClipData) = Nothing
        Dim dobj As DataObject = Nothing


        'Dim tryAlternate As Boolean = False
        'Dim gui As GUITHREADINFO
        'Dim dwResult As IntPtr
        'Dim hWndTarget As IntPtr
        'Dim thisThreadId As Integer
        'Dim targetThreadId As Integer
        'Dim lpdwtargetPid As Integer
        'Dim hwndFocus As IntPtr




        txtLog.Text = "Slot #" & NumSlot & " selected" & vbNewLine & txtLog.Text

        data = slots(NumSlot)

        If data IsNot Nothing AndAlso data.Count > 0 Then

            txtLog.Text = "Processing..." & vbNewLine & txtLog.Text

            stopListenToClipboard()

            dobj = New DataObject()
            Clipboard.Clear()

            For Each cd As ClipData In data
                dobj.SetData(cd.Format, autoconvert, cd.Data)
                txtLog.Text = "Added data to DataObject under format: " & cd.Format & vbNewLine & txtLog.Text
            Next

            Clipboard.SetDataObject(dobj, True)



            ''tries to paste !
            'thisThreadId = GetWindowThreadProcessId(Me.Handle, IntPtr.Zero)
            'targetThreadId = GetWindowThreadProcessId(hWndTarget, lpdwtargetPid)
            'gui.cbSize = Marshal.SizeOf(gui)
            'If GetGUIThreadInfo(targetThreadId, gui) Then
            '    If SendMessageTimeout(gui.hwndFocus, WM_PASTE, 0, 0, SendMessageTimeoutFlags.SMTO_BLOCK, 500, dwResult) = 0 Then
            '        txtLog.Text = "SendMessageTimeout failed!" & vbNewLine & txtLog.Text
            '        tryAlternate = True
            '    Else
            '        Threading.Thread.Sleep(500) 'TODO: use the clipboard viewer API to get notified of the change !
            '    End If
            'Else
            '    hWndTarget = GetForegroundWindow()
            '    If hWndTarget <> IntPtr.Zero Then
            '        thisThreadId = GetWindowThreadProcessId(Me.Handle, IntPtr.Zero)
            '        targetThreadId = GetWindowThreadProcessId(hWndTarget, lpdwtargetPid)
            '        If Not thisThreadId <> targetThreadId And thisThreadId > 0 And targetThreadId > 0 Then
            '            If AttachThreadInput(thisThreadId, targetThreadId, True) Then
            '                hwndFocus = GetFocus()
            '                If hwndFocus <> IntPtr.Zero Then
            '                    If SendMessageTimeout(hwndFocus, WM_PASTE, 0, 0, SendMessageTimeoutFlags.SMTO_BLOCK, 500, dwResult) = 0 Then
            '                        txtLog.Text = "SendMessageTimeout failed!" & vbNewLine & txtLog.Text
            '                        tryAlternate = True
            '                    Else
            '                        Threading.Thread.Sleep(500) 'TODO: use the clipboard viewer API to get notified of the change !
            '                    End If
            '                    If Not AttachThreadInput(thisThreadId, targetThreadId, False) Then
            '                        txtLog.Text = "Error in AttachThreadInput! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
            '                        tryAlternate = True
            '                    End If
            '                Else
            '                    txtLog.Text = "Error in GetFocus! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
            '                    AttachThreadInput(thisThreadId, targetThreadId, False)
            '                    tryAlternate = True
            '                End If
            '            Else
            '                txtLog.Text = "Error in AttachThreadInput! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
            '                tryAlternate = True
            '            End If
            '        Else
            '            txtLog.Text = "Error, thisThreadId is equal to the targetThreadId!" & vbNewLine & txtLog.Text
            '            tryAlternate = True
            '        End If
            '    Else
            '        tryAlternate = True
            '        txtLog.Text = "Error in GetForegroundWindow! Code:" & Marshal.GetLastWin32Error() & vbNewLine & txtLog.Text
            '    End If
            'End If


            'If tryAlternate Then
            '    txtLog.Text = "Using alternate way" & vbNewLine & txtLog.Text
            '    SendKeys.SendWait("^v") 'envoie CTRL+C à l'application qui a le focus, 5 fois, parce que parfois le copier ne se fait pas.
            '    SendKeys.SendWait("^v")
            '    SendKeys.SendWait("^v")
            '    SendKeys.SendWait("^v")
            '    SendKeys.SendWait("^v")
            'End If



            'If Not dontPaste Then
            '    SendKeys.SendWait("^v") 'cela ne change rien de coller plusieurs fois dans les applications qui ne veulent pas coller.
            'End If

            listenToClipboard()

            txtLog.Text = "DataObject added to clipboard" & vbNewLine & txtLog.Text
            notification.BalloonTipText = "Pasted from Slot #" & NumSlot & vbNewLine & "Now use CTRL+V to paste!"
            notification.ShowBalloonTip(5000)
        Else
            txtLog.Text = "Slot #" & NumSlot & " was empty!" & vbNewLine & txtLog.Text
            notification.BalloonTipText = "Slot #" & NumSlot & " was empty!"
            notification.ShowBalloonTip(5000)
        End If



        'Dim hWndTarget As IntPtr = GetForegroundWindow()
        'Dim title As StringBuilder = New StringBuilder(256)
        'GetWindowText(hWndTarget, title, 256)
        'Dim a As String = title.ToString()

        'Dim thisPid As IntPtr = GetWindowThreadProcessId(Me.Handle, IntPtr.Zero)

        'Dim targetPid As IntPtr = GetWindowThreadProcessId(hWndTarget, IntPtr.Zero)

        'If Not thisPid.Equals(targetPid) Then

        '    AttachThreadInput(thisPid, targetPid, True)
        '    Dim lngResult As IntPtr = SendMessage(GetFocus(), WM_PASTE, IntPtr.Zero, IntPtr.Zero)
        '    AttachThreadInput(thisPid, targetPid, False)

        'End If




    End Sub

    Private Sub panSlot_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim panSlot As Panel = DirectCast(sender, Panel)
        Dim name As String = panSlot.Name
        Dim numSlot As Integer = Val(name.Split("_")(1))

        showcontrols(numSlot, True)
        preview(numSlot, True)

    End Sub

    Private Sub ctrlSlot_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ctrlSlot As Control = DirectCast(sender, Control)
        Dim name As String = ctrlSlot.Name
        Dim numSlot As Integer = Val(name.Split("_")(1))

        showcontrols(numSlot, True)
        preview(numSlot, True)
    End Sub

    Private Sub panQueue_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim panQueue As Panel = DirectCast(sender, Panel)
        Dim name As String = panQueue.Name
        Dim numItem As Integer = Val(name.Split("_")(1))

        showcontrols(numItem, False)
        preview(numItem, False)

    End Sub

    Private Sub ctrlQueue_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ctrlQueue As Control = DirectCast(sender, Control)
        Dim name As String = ctrlQueue.Name
        Dim numItem As Integer = Val(name.Split("_")(1))

        showcontrols(numItem, False)
        preview(numItem, False)
    End Sub

    Private Sub preview(ByVal position As Integer, Optional ByVal fromSlot As Boolean = True)
        Dim data As List(Of ClipData) = Nothing
        Dim dobj As DataObject = Nothing
        Dim tmpHTML As String
        Dim beginHTML As Integer
        Dim endHTML As Integer
        Dim file As String
        Dim filelist As StringBuilder
        Dim txtHTML As String

        If fromSlot Then
            data = slots(position)
        Else
            data = queue.ToArray()(9 - position)
        End If

        If data IsNot Nothing AndAlso data.Count > 0 Then
            dobj = New DataObject()
            For Each cd As ClipData In data
                dobj.SetData(cd.Format, autoconvert, cd.Data)
            Next

            If dobj.ContainsImage() Then 'display image
                webPreview.Visible = False
                rtfPreview.Visible = False
                picPreview.Visible = True
                picPreview.Image = dobj.GetImage()

            ElseIf dobj.ContainsText() Then
                If dobj.ContainsText(TextDataFormat.Html) Then 'display copied html fragment
                    txtHTML = dobj.GetText(TextDataFormat.Html)
                    beginHTML = txtHTML.IndexOf("StartHTML:")
                    endHTML = txtHTML.IndexOf("EndHTML:")
                    If beginHTML > 0 And endHTML > 0 Then

                        tmpHTML = txtHTML.Substring(beginHTML, 32)
                        tmpHTML = tmpHTML.Split(vbNewLine)(0).Split(":")(1)
                        beginHTML = Val(tmpHTML)

                        tmpHTML = txtHTML.Substring(endHTML, 32)
                        tmpHTML = tmpHTML.Split(vbNewLine)(0).Split(":")(1)
                        endHTML = Val(tmpHTML)

                        tmpHTML = txtHTML.Substring(beginHTML, endHTML - beginHTML)

                        rtfPreview.Visible = False
                        picPreview.Visible = False
                        webPreview.Visible = True

                        If webPreview.Document Is Nothing Then
                            webPreview.DocumentText = tmpHTML
                        Else
                            webPreview.Document.OpenNew(False)
                            webPreview.Document.Write(tmpHTML)
                        End If

                    ElseIf dobj.ContainsText(TextDataFormat.Rtf) Then 'html errors, try rtf
                        webPreview.Visible = False
                        picPreview.Visible = False
                        rtfPreview.Visible = True
                        rtfPreview.ResetText()
                        rtfPreview.Rtf = dobj.GetText(TextDataFormat.Rtf)

                    ElseIf dobj.ContainsText(TextDataFormat.UnicodeText) Then 'no rtf, try unicode text
                        webPreview.Visible = False
                        picPreview.Visible = False
                        rtfPreview.Visible = True
                        rtfPreview.Rtf = ""
                        rtfPreview.Text = dobj.GetText(TextDataFormat.UnicodeText)

                    ElseIf dobj.ContainsText(TextDataFormat.Text) Then 'no unicode, display simple text
                        webPreview.Visible = False
                        picPreview.Visible = False
                        rtfPreview.Visible = True
                        rtfPreview.Rtf = ""
                        rtfPreview.Text = dobj.GetText(TextDataFormat.Text)

                    Else 'display raw data
                        webPreview.Visible = False
                        picPreview.Visible = False
                        rtfPreview.Visible = True
                        rtfPreview.Rtf = ""
                        rtfPreview.Text = dobj.GetText()
                    End If

                ElseIf dobj.ContainsText(TextDataFormat.Rtf) Then 'display rtf data
                    webPreview.Visible = False
                    picPreview.Visible = False
                    rtfPreview.Visible = True
                    rtfPreview.ResetText()
                    rtfPreview.Rtf = dobj.GetText(TextDataFormat.Rtf)

                ElseIf dobj.ContainsText(TextDataFormat.CommaSeparatedValue) Then 'display csv data
                    webPreview.Visible = False
                    picPreview.Visible = False
                    rtfPreview.Visible = True
                    rtfPreview.Rtf = ""
                    rtfPreview.Text = dobj.GetText(TextDataFormat.CommaSeparatedValue)

                ElseIf dobj.ContainsText(TextDataFormat.UnicodeText) Then 'display unicode text data
                    webPreview.Visible = False
                    picPreview.Visible = False
                    rtfPreview.Visible = True
                    rtfPreview.Rtf = ""
                    rtfPreview.Text = dobj.GetText(TextDataFormat.UnicodeText)

                ElseIf dobj.ContainsText(TextDataFormat.Text) Then 'display text data
                    webPreview.Visible = False
                    picPreview.Visible = False
                    rtfPreview.Visible = True
                    rtfPreview.Rtf = ""
                    rtfPreview.Text = dobj.GetText(TextDataFormat.Text)

                Else 'display raw data
                    webPreview.Visible = False
                    picPreview.Visible = False
                    rtfPreview.Visible = True
                    rtfPreview.Rtf = ""
                    rtfPreview.Text = dobj.GetText()
                End If

            ElseIf dobj.ContainsFileDropList() Then 'display file list
                webPreview.Visible = False
                picPreview.Visible = False
                rtfPreview.Visible = True
                rtfPreview.Rtf = ""
                filelist = New StringBuilder()
                filelist.AppendLine("File list:")
                For Each file In dobj.GetFileDropList()
                    filelist.AppendLine(file)
                Next
                rtfPreview.Text = filelist.ToString()
                filelist.Remove(0, filelist.Length)

            ElseIf dobj.ContainsAudio() Then 'show that we have audio data
                webPreview.Visible = False
                picPreview.Visible = False
                rtfPreview.Visible = True
                rtfPreview.Rtf = ""
                rtfPreview.Text = "(Audio data)"
            Else
                webPreview.Visible = False
                picPreview.Visible = False
                rtfPreview.Visible = True
                rtfPreview.Rtf = ""
                rtfPreview.Text = "(Preview not available yet for this data type)"
            End If

        Else
            webPreview.Visible = False
            picPreview.Visible = False
            rtfPreview.Visible = True
            'rtfPreview.Rtf = ""
            'rtfPreview.Text = "---- EMPTY ----"
            rtfPreview.ResetText()
            If fromSlot Then
                'rtfPreview.AppendText(vbNewLine & _
                '                      vbNewLine & _
                '                      "CTRL + " & position & vbNewLine & _
                '                      "or" & vbNewLine & _
                '                      "CTRL + Numpad " & position & vbNewLine & _
                '                      IIf(position = 0, "or" & vbNewLine & "CTRL + [ ` ² ]" & vbNewLine, "") & _
                '                      "to copy Clipboard content into this Slot" & vbNewLine & _
                '                      vbNewLine & _
                '                      "CTRL + SHIFT + " & position & vbNewLine & _
                '                      "or" & vbNewLine & _
                '                      "CTRL + ALT + Numpad " & position & vbNewLine & _
                '                       IIf(position = 0, "or" & vbNewLine & "CTRL + SHIFT + [ ` ² ]" & vbNewLine, "") & _
                '                      "to copy this slot into the Clipboard")
                rtfPreview.Rtf = _
"{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}{\f1\fnil\fcharset0 System;}{\f2\fnil\fcharset0 Calibri;}}" & vbNewLine & _
"{\colortbl ;\red192\green192\blue192;\red255\green0\blue0;\red79\green129\blue189;}" & vbNewLine & _
"{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\qc\highlight1\b\f0\fs20 ---------- EMPTY ----------\highlight0\b0\fs17\par" & vbNewLine & _
"\pard\par" & vbNewLine & _
"\cf2\b\f1 CTRL \cf0 +\cf2  " & position & "\cf0\b0\f0\par" & vbNewLine & _
"or\par" & vbNewLine & _
"\cf2\b\f1 CTRL \cf0 +\cf2  Numpad " & position & "\cf0\b0\f0\par" & vbNewLine & _
IIf(position = 0, "or\par" & vbNewLine & _
"\cf2\b\f1 CTRL \cf0 + [\cf2  ` \'b2 \cf0 ]\b0\f0\par" & vbNewLine, "") & _
"to copy Clipboard content into this Slot\par" & vbNewLine & _
"\par" & vbNewLine & _
"\cf3\b\f1 CTRL \cf0 +\cf3  SHIFT \cf0 +\cf3  " & position & "\cf0\b0\f0\par" & vbNewLine & _
"or\par" & vbNewLine & _
"\cf3\b\f1 CTRL \cf0 +\cf3  ALT \cf0 +\cf3  Numpad " & position & "\cf0\b0\f0\par" & vbNewLine & _
IIf(position = 0, "or\par" & vbNewLine & _
"\cf3\b\f1 CTRL \cf0 +\cf3  SHIFT \cf0 + [\cf3  ` \'b2 \cf0 ]\b0\f0\par" & vbNewLine, "") & _
"\pard\sa200\sl276\slmult1 to copy this slot into the Clipboard\lang12\f2\fs22\par" & vbNewLine & _
"}" & vbNullChar

            Else
                rtfPreview.Rtf = _
"{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}{\f1\fnil\fcharset0 Calibri;}}" & vbNewLine & _
"{\colortbl ;\red192\green192\blue192;}" & vbNewLine & _
"{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\qc\highlight1\b\f0\fs20 ---------- EMPTY ----------\highlight0\lang12\b0\f1\fs22\par" & vbNewLine & _
"}" & vbNullChar

            End If
        End If



    End Sub

    Private Sub btnSlotDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim p As Panel
        Dim l As Label
        Dim btn As Control = DirectCast(sender, Button)
        Dim name As String = btn.Name
        Dim numSlot As Integer = Val(name.Split("_")(1))

        If slots(numSlot) IsNot Nothing Then
            slots(numSlot).Clear()
        End If

        p = DirectCast(Me.tableSlot.Controls("panSlot_" & numSlot), Panel)
        l = DirectCast(p.Controls("lblSlot_" & numSlot), Label)
        l.Text = "Slot #" & numSlot & vbNewLine & "(empty)"

    End Sub

    Private Sub btnSlotCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Control = DirectCast(sender, Button)
        Dim name As String = btn.Name
        Dim numSlot As Integer = Val(name.Split("_")(1))

        PasteFromSlot(numSlot, True)

    End Sub

    Private Sub btnQueueCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Control = DirectCast(sender, Button)
        Dim name As String = btn.Name
        Dim position As Integer = Val(name.Split("_")(1))

        PasteFromQueue(position, True)

    End Sub

    Private Sub btnQueueTransfer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Control = DirectCast(sender, Button)
        Dim name As String = btn.Name
        Dim position As Integer = Val(name.Split("_")(1))
        Dim ianswer As Integer
        Dim answer As String

        answer = InputBox("Type the slot number (0-9):", "Where?", "1")
        While Not Int32.TryParse(answer, ianswer)
            answer = InputBox("Error! Type the slot number (0-9):", "Where?", "1")
        End While

        CopyToSlot(ianswer, 0, position)

    End Sub

    Private Sub showcontrols(ByVal position As Integer, Optional ByVal fromSlot As Boolean = True)

        Dim p As Panel = Nothing
        'Dim btnSave As Button = Nothing
        Dim btnCopy As Button = Nothing
        Dim btnDelete As Button = Nothing
        'Dim btnPreview As Button = Nothing
        Dim btnTransfer As Button = Nothing

        If fromSlot Then
            p = tableSlot.Controls("panSlot_" & position)
            'btnSave = p.Controls("btnSlotSave_" & position)
            btnCopy = p.Controls("btnSlotCopy_" & position)
            btnDelete = p.Controls("btnSlotDelete_" & position)
            'btnPreview = p.Controls("btnSlotPreview_" & position)
        Else
            p = tableQueue.Controls("panQueue_" & position)
            'btnSave = p.Controls("btnQueueSave_" & position)
            btnCopy = p.Controls("btnQueueCopy_" & position)
            'btnPreview = p.Controls("btnQueuePreview_" & position)
            btnTransfer = p.Controls("btnQueueTransfer_" & position)
        End If

        'btnSave.Show()
        btnCopy.Show()
        'btnPreview.Show()
        If fromSlot Then
            btnDelete.Show()
        Else
            btnTransfer.Show()
        End If

        For i As Integer = 0 To 9
            If i <> position Then

                If fromSlot Then
                    p = tableSlot.Controls("panSlot_" & i)
                    'btnSave = p.Controls("btnSlotSave_" & i)
                    btnCopy = p.Controls("btnSlotCopy_" & i)
                    btnDelete = p.Controls("btnSlotDelete_" & i)
                    'btnPreview = p.Controls("btnSlotPreview_" & i)
                Else
                    p = tableQueue.Controls("panQueue_" & i)
                    'btnSave = p.Controls("btnQueueSave_" & i)
                    btnCopy = p.Controls("btnQueueCopy_" & i)
                    'btnPreview = p.Controls("btnQueuePreview_" & i)
                    btnTransfer = p.Controls("btnQueueTransfer_" & i)
                End If

                'btnSave.Hide()
                btnCopy.Hide()
                'btnPreview.Hide()
                If fromSlot Then
                    btnDelete.Hide()
                Else
                    btnTransfer.Hide()
                End If
            End If
        Next

    End Sub

    Private Sub ShowMainWindow()
        'Me.WindowState = FormWindowState.Normal
        'Me.ShowInTaskbar = True

        Me.Hide()
        Me.Show()
        Me.BringToFront()
        Me.Activate()

    End Sub

    Private Sub MinimizeToTray()
        'Me.ShowInTaskbar = False
        'Me.SendToBack()
        'Me.WindowState = FormWindowState.Minimized
        Me.Hide()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        forceClose = True
        Application.Exit()
    End Sub

    Private Sub notification_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles notification.MouseDoubleClick
        ShowMainWindow()
    End Sub

    Private Sub ShowMainWindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowMainWindowToolStripMenuItem.Click
        ShowMainWindow()
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        If dlgOpts Is Nothing OrElse dlgOpts.Created = False Then
            dlgOpts = New dlgOptions()
        End If
        dlgOpts.ShowDialog()

    End Sub


    Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'If Not activatedFirstTime Then
        '    activatedFirstTime = True
        '    MinimizeToTray()
        'End If
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'If Not activatedFirstTime Then
        '    activatedFirstTime = True
        '    MinimizeToTray()
        'End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        PostInitializeComponents()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub




End Class

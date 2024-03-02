Public Class Form2


    Const DWM_TNP_VISIBLE = 8
    Const DWM_TNP_RECTDESTINATION = 1
    Const DWM_TNP_RECTSOURCE = 2
    Const DWM_TNP_OPACITY = 4
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As IntPtr, ByVal bRevert As Int32) As Int32
    Declare Function DrawMenuBar Lib "user32" Alias "DrawMenuBar" (ByVal hwnd As IntPtr) As Int32
    Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Int32, ByVal nPosition As Int32, ByVal wFlags As Int32, ByVal wIDNewItem As Int32, ByVal lpNewItem As String) As Int32
    Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Int32, ByVal nPosition As Int32, ByVal wFlags As Int32, ByVal wIDNewItem As Int32, ByVal lpNewItem As String) As Int32

    Private Const MF_BYCOMMAND = &H0&
    Private Const MF_BYPOSITION = &H400&
    Private Const MF_STRING = &H0&
    Private Const MF_SEPARATOR = &H800&
    Private Const WM_SYSCOMMAND = &H112
    Private Const MF_CHECKED = &H8&
    Private Const SW_SHOWMAXIMIZED = &H3&

    Const WS_CHILD = &H40000000

    Const SWP_SHOWWINDOW = &H40

    Structure Rect
        Sub New(ByVal _left As Integer, ByVal _top As Integer, ByVal _right As Integer, ByVal _bottom As Integer)
            Left = _left
            Right = _right
            Top = _top
            Bottom = _bottom
        End Sub
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure


    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As IntPtr, ByVal nlndex As Integer, ByVal wNewLong As IntPtr) As IntPtr
    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As IntPtr) As IntPtr

    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Int32
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal nCmdShow As IntPtr) As Int32
    Private Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As IntPtr, ByVal y As IntPtr, ByVal cx As IntPtr, ByVal cy As IntPtr, ByVal wFlags As IntPtr) As Int32

    Private Declare Function DwmUpdateThumbnailProperties Lib "dwmapi.dll" (ByVal hThumb As IntPtr, ByRef dwmp As Form1.DWM_PROPERTIES) As Integer
    Declare Function GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As IntPtr, ByRef lpRect As Rect) As Long
    Private Declare Function ChildWindowFromPoint Lib "user32" Alias "ChildWindowFromPoint" (ByVal hWndParent As IntPtr, ByVal x As IntPtr, ByVal y As IntPtr) As IntPtr
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Const SC_MOVE = &HF010&
    Const HTCAPTION = 2
    Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean
    Declare Function SetCurrentProcessExplicitAppUserModelID Lib "Shell32.dll" (ByVal AppID As String) As Integer

    Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As IntPtr, ByVal nIndex As IntPtr) As IntPtr
    Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As IntPtr, ByVal nIndex As IntPtr, ByVal dwNewLong As IntPtr) As IntPtr

    Declare Function IsZoomed Lib "user32" Alias "IsZoomed" (ByVal hwnd As IntPtr) As IntPtr

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal dwTime As IntPtr, ByVal dwFlags As IntPtr) As Boolean

    Declare Function UpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hwnd As IntPtr) As IntPtr

    Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As IntPtr, ByVal wFlags As IntPtr, ByVal x As IntPtr, ByVal y As IntPtr, ByVal HWnd As IntPtr, ByVal lptpm As IntPtr) As IntPtr

    Private Declare Function DwmUnregisterThumbnail Lib "dwmapi.dll" (ByVal hThumbnailId As IntPtr) As IntPtr

    Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As IntPtr, ByVal attr As IntPtr, ByRef attrValue As IntPtr, ByVal attrSize As IntPtr) As IntPtr

    Private Const WM_PAINT = &HF
    Private Const WM_PRINT = &H317

    Const MF_APPEND = &H100&
    Const TPM_LEFTALIGN = &H0&
    Const MF_DISABLED = &H2&
    Const MF_GRAYED = &H1&
    Const TPM_RETURNCMD = &H100&
    Const TPM_RIGHTBUTTON = &H2&

    Public Enum DWMWINDOWATTRIBUTE
        DWMWA_ALLOW_NCPAINT = 4
        DWMWA_CAPTION_BUTTON_BOUNDS = 5
        DWMWA_FLIP3D_POLICY = 8
        DWMWA_FORCE_ICONIC_REPRESENTATION = 7
        DWMWA_LAST = 9
        DWMWA_NCRENDERING_ENABLED = 1
        DWMWA_NCRENDERING_POLICY = 2
        DWMWA_NONCLIENT_RTL_LAYOUT = 6
        DWMWA_TRANSITIONS_FORCEDISABLED = 3
    End Enum


    Public hwnd As IntPtr
    Private stat As Boolean
    Public x As Integer
    Public y As Integer
    Public winstyle As IntPtr
    Private lv As Boolean

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim hrect As Rect

        If m.Msg = WM_SYSCOMMAND Then
            If m.WParam.ToInt32 = 2002 Then
                If Not stat Then
                    ModifyMenu(GetSystemMenu(Me.Handle, False), 0, MF_BYPOSITION Or MF_STRING Or MF_CHECKED, 2002, "接管窗口")
                    Panel1.Visible = True
                    Form1.dwmp.fVisible = False
                    DwmUpdateThumbnailProperties(Form1.tid, Form1.dwmp)
                    GetWindowRect(hwnd, hrect)

                    Panel1.Left = -x ' Me.Left
                    Panel1.Top = -y 'Me.Top
                    Panel1.Width = My.Computer.Screen.Bounds.Width
                    Panel1.Height = My.Computer.Screen.Bounds.Height
                    Dim max As Boolean = IsZoomed(hwnd)
                    SetParent(hwnd, Panel1.Handle)
                    If max Then ShowWindow(hwnd, SW_SHOWMAXIMIZED)
                Else
                    ModifyMenu(GetSystemMenu(Me.Handle, False), 0, MF_BYPOSITION Or MF_STRING, 2002, "接管窗口")
                    Panel1.Visible = False
                    Dim max As Boolean = IsZoomed(hwnd)
                    SetParent(hwnd, 0)
                    Me.TopMost = True

                    If max Then ShowWindow(hwnd, SW_SHOWMAXIMIZED)
                    Form1.dwmp.fVisible = True
                    DwmUpdateThumbnailProperties(Form1.tid, Form1.dwmp)

                End If
                DrawMenuBar(GetSystemMenu(Me.Handle, False))
                stat = Not stat
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form2_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Me.Opacity = 0
        AnimateWindow(Me.hwnd, 0, 0)
        End
    End Sub


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InsertMenu(GetSystemMenu(Me.Handle, False), 0, MF_BYPOSITION Or MF_SEPARATOR, 2001, "")
        InsertMenu(GetSystemMenu(Me.Handle, False), 0, MF_BYPOSITION Or MF_STRING, 2002, "接管窗口")
        Dim r As New Random
        Dim rwl As IntPtr
        rwl = FindWindow("Shell_TrayWnd", 0)
        SendMessage(rwl, WM_PAINT, 0, 0)
        SendMessage(rwl, WM_PRINT, 0, 0)
        UpdateWindow(rwl)
        SetCurrentProcessExplicitAppUserModelID("Window.sliced.id" + r.Next().ToString())

    End Sub

    Private Sub Form2_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If stat Then
            Dim max As Boolean = IsZoomed(hwnd)
            SetParent(hwnd, 0)
            If max Then ShowWindow(hwnd, SW_SHOWMAXIMIZED)
        End If
        Me.Opacity = 0
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Panel1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        Dim dx As Double
        Dim dy As Double
        If Me.WindowState <> FormWindowState.Normal Then
            Exit Sub
        End If
        dx = 1.0
        dy = (Me.Height - SystemInformation.CaptionHeight - SystemInformation.FixedFrameBorderSize.Height - SystemInformation.BorderSize.Height) * 1.0 / (Me.Width - SystemInformation.FixedFrameBorderSize.Width - SystemInformation.BorderSize.Width)

        If Me.Height > 200 Or e.Delta > 0 Then
            Me.ResizeRedraw = False
            Me.Height += dy * e.Delta / 5
            Me.Width += dx * e.Delta / 5
            Me.Top -= dy * e.Delta / 10
            Me.Left -= dx * e.Delta / 10

            Form1.dwmp.rcDest.Right += dx * e.Delta / 5   '= Me.Width - SystemInformation.FixedFrameBorderSize.Width
            Form1.dwmp.rcDest.Bottom += dy * e.Delta / 5    '= Me.Height - SystemInformation.CaptionHeight - SystemInformation.FixedFrameBorderSize.Height
            'Debug.Print(dy)
            'Form1.dwmp.rcDest.Right = Me.Width - SystemInformation.FixedFrameBorderSize.Width - SystemInformation.BorderSize.Width
            'Form1.dwmp.rcDest.Bottom = Me.Height - SystemInformation.CaptionHeight - SystemInformation.FixedFrameBorderSize.Height - SystemInformation.BorderSize.Height
            DwmUpdateThumbnailProperties(Form1.tid, Form1.dwmp)
            Me.ResizeRedraw = True
        End If
    End Sub

    Private Sub Form2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
        End If
        If e.Button = Windows.Forms.MouseButtons.Left And e.Clicks = 2 Then
            If Me.WindowState = FormWindowState.Normal Then
                Me.WindowState = FormWindowState.Maximized
                Me.BackColor = Color.Black
            ElseIf Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
                Me.BackColor = SystemColors.Control
            End If
            ' Me.Refresh()
            ' Application.DoEvents()
            Dim r As Double
            r = Form1.dwmp.rcDest.Right / Form1.dwmp.rcDest.Bottom
            Dim oldr As Integer, oldb As Integer
            oldr = Form1.dwmp.rcDest.Right
            oldb = Form1.dwmp.rcDest.Bottom
            Form1.dwmp.rcDest.Right = Me.Width - SystemInformation.FixedFrameBorderSize.Width - SystemInformation.BorderSize.Width
            Form1.dwmp.rcDest.Bottom = Me.Height - SystemInformation.CaptionHeight - SystemInformation.FixedFrameBorderSize.Height - SystemInformation.BorderSize.Height
            If Me.WindowState = FormWindowState.Maximized Then
                If (Form1.dwmp.rcDest.Right / oldr) > (Form1.dwmp.rcDest.Bottom / oldb) Then
                    Form1.dwmp.rcDest.Right = Form1.dwmp.rcDest.Bottom * r
                    Dim d As Integer
                    d = Me.Width - (Form1.dwmp.rcDest.Right - Form1.dwmp.rcDest.Left)
                    Form1.dwmp.rcDest.Left += d / 2
                    Form1.dwmp.rcDest.Right += d / 2
                Else
                    Form1.dwmp.rcDest.Bottom = Form1.dwmp.rcDest.Right / r
                    Dim d As Integer
                    d = Me.Height - (Form1.dwmp.rcDest.Bottom - Form1.dwmp.rcDest.Top)
                    Form1.dwmp.rcDest.Top += d / 2
                    Form1.dwmp.rcDest.Bottom += d / 2
                End If
            Else
                Form1.dwmp.rcDest.Left = 0
                Form1.dwmp.rcDest.Top = 0
            End If
            DwmUpdateThumbnailProperties(Form1.tid, Form1.dwmp)
        End If
    End Sub

    Private Sub Form2_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Dim x As IntPtr
        x = GetClassLong(hwnd, -14)
        If x = 0 Then x = GetClassLong(hwnd, -34)
        If x <> 0 Then Me.Icon = Icon.FromHandle(x)
    End Sub

    Public Sub New()
        ' 此调用是设计器所必需的。
        InitializeComponent()
        'SetStyle(ControlStyles.SupportsTransparentBackColor, False)
        ' 在 InitializeComponent() 调用之后添加任何初始化。
    End Sub
End Class
Public Class Form1

    Const DWM_TNP_VISIBLE = 8
    Const DWM_TNP_RECTDESTINATION = 1
    Const DWM_TNP_RECTSOURCE = 2
    Const DWM_TNP_OPACITY = 4

    Private Const DWM_BB_ENABLE = &H1&
    Private Const DWM_BB_BLURREGION = &H2&
    Private Const DWM_BB_TRANSITIONONMAXIMIZED = &H4

    Private Const WM_SYSCOMMAND = &H112

    Private Structure POINTAPI
        Dim x As Long
        Dim y As Long
    End Structure

    Private Structure DWM_BLURBEHIND
        Dim dwFlags As Long
        Dim fEnable As Long
        Dim hRgnBlur As Long
        Dim fTransitionOnMaximized As Long
    End Structure

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

    Structure DWM_PROPERTIES
        Dim dwFlags As Integer
        Dim rcDest As Rect
        Dim rcSrc As Rect
        Dim opacity As Byte
        Dim fVisible As Boolean
        Dim fSourceClientAreaOnly As Boolean
    End Structure

    Public dwmp As New DWM_PROPERTIES
    Public tid As IntPtr

    Declare Function GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As IntPtr, ByRef lpRect As Rect) As Long
    Declare Function GetParent Lib "user32" Alias "GetParent" (ByVal hwnd As IntPtr) As IntPtr

    Private Declare Function ScreenToClient Lib "user32" (ByRef hwnd As IntPtr, ByRef lpPoint As POINTAPI) As Integer

    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As String, ByVal cch As Integer) As Integer
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As IntPtr) As Integer
    Private Declare Function DwmEnableBlurBehindWindow Lib "dwmapi" (ByVal hWnd As IntPtr, ByRef pBlurBehind As DWM_BLURBEHIND) As Long

    Private Declare Function DwmRegisterThumbnail Lib "dwmapi.dll" (ByVal hWndDest As IntPtr, ByVal hWndSrc As IntPtr, ByRef dwThumb As IntPtr) As Integer
    Private Declare Function DwmUpdateThumbnailProperties Lib "dwmapi.dll" (ByVal hThumb As IntPtr, ByRef dwmp As DWM_PROPERTIES) As Integer
    Private Declare Function WindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Integer, ByVal yPoint As Integer) As Integer

    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal dwTime As IntPtr, ByVal dwFlags As IntPtr) As Boolean

    Dim st As Boolean
    Dim start_x As Integer
    Dim start_y As Integer

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Click

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        st = False
    End Sub

    Private Sub Form1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        Panel1.Visible = True
        Panel1.Top = e.Y
        Panel1.Left = e.X
        start_x = e.X
        start_y = e.Y
        st = True
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If st Then
            Dim x As Integer
            Dim y As Integer
            Dim w As Integer
            Dim h As Integer
            If e.Y > start_y Then
                h = e.Y - Panel1.Top
                y = start_y
            Else
                y = e.Y
                h = start_y - e.Y
            End If

            If e.X > start_x Then
                w = e.X - Panel1.Left
                x = start_x
            Else
                x = e.X
                w = start_x - e.X
            End If
            Panel1.SetBounds(x, y, w, h)
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        st = False



        Form2.Top = 65535
        Form2.Visible = False
        Form2.Show()
        Form2.Visible = False
        Form2.Top = Panel1.Top - SystemInformation.CaptionHeight - SystemInformation.FixedFrameBorderSize.Height '- SystemInformation.BorderSize.Height
        Form2.Left = Panel1.Left - SystemInformation.FixedFrameBorderSize.Width '- SystemInformation.BorderSize.Width
        Form2.Height = Panel1.Height + SystemInformation.CaptionHeight + SystemInformation.FixedFrameBorderSize.Height + SystemInformation.BorderSize.Height
        Form2.Width = Panel1.Width + SystemInformation.FixedFrameBorderSize.Width + SystemInformation.BorderSize.Width

        Dim hwnd As Integer
        hwnd = WindowFromPoint(Panel1.Left + Panel1.Width / 2, Panel1.Top + Panel1.Height / 2)
        While GetParent(hwnd).ToInt32 > 0
            hwnd = GetParent(hwnd)
        End While
        Form2.hwnd = hwnd
        Dim hrect As Rect
        GetWindowRect(hwnd, hrect)
        Dim rst As Integer
        rst = DwmRegisterThumbnail(Form2.Handle, hwnd, tid)

        dwmp.opacity = 255
        dwmp.fVisible = True

        dwmp.rcSrc = New Rect(Panel1.Left - hrect.Left, Panel1.Top - hrect.Top, Panel1.Right - hrect.Left, Panel1.Bottom - hrect.Top)
        Form2.x = Panel1.Left
        Form2.y = Panel1.Top
        dwmp.rcDest = New Rect(Left, Top, Panel1.Width, Panel1.Height)
        dwmp.dwFlags = DWM_TNP_VISIBLE + DWM_TNP_RECTDESTINATION + DWM_TNP_RECTSOURCE

        DwmUpdateThumbnailProperties(tid, dwmp)

        Dim bb As DWM_BLURBEHIND
        bb.dwFlags = DWM_BB_ENABLE
        bb.fEnable = 0
        bb.hRgnBlur = 0
        bb.fTransitionOnMaximized = 1

        DwmEnableBlurBehindWindow(tid, bb)

        Dim strTitle As String = Space(GetWindowTextLength(hwnd) + 1) '构造窗口标题字符串缓冲区
        GetWindowText(hwnd, strTitle, strTitle.Length) '获取窗口文字
        Form2.Text = "窗口切片 - " + strTitle

        Me.ShowInTaskbar = False
        While (Me.Opacity > 0.2)
            Me.Opacity = Me.Opacity * 0.9
            Application.DoEvents()
            Threading.Thread.Sleep(20)
        End While

        Me.Hide()


        Form2.Visible = True
        If e.Button.ToString = "Right" Then
            'Threading.Thread.Sleep(200)
            PostMessage(Form2.Handle, WM_SYSCOMMAND, 2002, 0)
        End If
    End Sub

    Private Sub Form1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.DoubleClick
        End
    End Sub

    Private Sub Form1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        End
    End Sub

    Private Sub Form1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Leave
        End
    End Sub

    Private Sub Form1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

    End Sub
End Class

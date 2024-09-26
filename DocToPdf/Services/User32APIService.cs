using System.Runtime.InteropServices;
using System.Text;

namespace DocToPdf.Services
{
    [Flags]
    public enum MonitorFromWindowFlags
    {
        DefaultToNull    = 0,
        DefaultToPrimary = 1,
        DefaultToNearest = 2,
    }
    [Flags]
    public enum SetWindowPosFlags
    {
        NOSIZE         = 0x0001,
        NOMOVE         = 0x0002,
        NOZORDER       = 0x0004,
        NOREDRAW       = 0x0008,
        NOACTIVATE     = 0x0010,
        DRAWFRAME      = 0x0020,
        FRAMECHANGED   = 0x0020,
        SHOWWINDOW     = 0x0040,
        HIDEWINDOW     = 0x0080,
        NOCOPYBITS     = 0x0100,
        NOOWNERZORDER  = 0x0200,
        NOREPOSITION   = 0x0200,
        NOSENDCHANGING = 0x0400,
        DEFERERASE     = 0x2000,
        ASYNCWINDOWPOS = 0x4000,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;

        public RECT(int left, int top, int right, int bottom)
        {
            Left   = left;
            Top    = top;
            Right  = right;
            Bottom = bottom;
        }

        public int Height
        {
            get { return Bottom - Top; }
        }

        public int Width
        {
            get { return Right - Left; }
        }
    }
     
    public class User32APIService
    {
        [DllImport("user32.dll", EntryPoint = "SetParent")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);       
        
        [DllImport("user32.dll", EntryPoint = "DeferWindowPos")]
        public static extern IntPtr DeferWindowPos(IntPtr hWinPosInfo, IntPtr hWnd,
        IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", EntryPoint = "BeginDeferWindowPos")]
        public static extern IntPtr BeginDeferWindowPos(int nNumWindows);

        [DllImport("user32.dll", EntryPoint = "EndDeferWindowPos")]
        public static extern bool EndDeferWindowPos(IntPtr hWinPosInfo);

        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32", EntryPoint = "GetForegroundWindow")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", EntryPoint = "MoveWindow", SetLastError = true)]
        public static extern bool MoveWindow(IntPtr hwnd, int x, int y, int cx, int cy, bool repaint);

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll", EntryPoint = "ShowWindowAsync", SetLastError = true)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow", SetLastError = false)]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", EntryPoint = "GetWindowRect", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int width, int height, SetWindowPosFlags uFlags);
  
        [DllImport("user32.dll", EntryPoint = "MonitorFromWindow")]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, MonitorFromWindowFlags dwFlags);

        [DllImport("user32.dll", EntryPoint = "GetActiveWindow")]
        public static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll", EntryPoint = "SetActiveWindow")]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        public static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "PostMessage")]
        public static extern int PostMessage(int hwnd, int wMsg, int wParam, int lParam);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "SendMessage")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);
        [DllImport("user32.dll", EntryPoint = "GetWindow")]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
    }

    /// <summary>
    /// User32 관련 랩퍼 멤버 함수는 FDUser32APIWrapper 클래스에 추가 하십시오.
    /// </summary>
    public class User32APIWrapper : User32APIService
    {
        public User32APIWrapper()
        {

        }
        public static IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam)
        {
            return SendMessage(hWnd, (uint)msg, (IntPtr)wParam, null);
        }
        public static bool IsWindowForeground()
        {
            uint GW_HWNDPREV = 3;

            IntPtr foregroundWindow = GetForegroundWindow();               // 현재 활성화된 윈도우 핸들을 가져옴
            IntPtr thisWindow       = GetWindow(IntPtr.Zero, GW_HWNDPREV); // 이전 윈도우 핸들을 가져옴

            return (foregroundWindow == thisWindow);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;


namespace DuteIT
{
    public class 窗口
    {
        #region WinodwsAPI

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string IpClassName, string IpWindowName);

       
        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll ", EntryPoint = "SendMessageA")]
        public static extern int SendMessageA(IntPtr hwnd, uint wMsg, int wParam, string lParam);

        [DllImport("user32.dll ", EntryPoint = "SendMessage")]
        public static extern int SendMessage(IntPtr hwnd, uint wMsg, int wParam, int lParam);

        [DllImport("user32.dll", EntryPoint = "GetParent")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]
        public static extern bool GetCursorPos(out Point pt);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr WindowFromPoint(Point pt);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, [Out, MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowRect(IntPtr hwnd, ref Rectangle rc);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetClientRect(IntPtr hwnd, ref Rectangle rc);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool MoveWindow(IntPtr hwnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int ScreenToClient(IntPtr hWnd, ref Rectangle rect);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hWndParent, ChildWindowsProc lpEnumFunc, int lParam);

        #endregion
        private delegate bool WNDENUMPROC(IntPtr hWnd, int lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();
        public static void 窗口_取屏幕句柄()
        {
            GetDesktopWindow();
        }

        /// <summary>
        /// 根据标题查找窗体句柄
        /// </summary>
        /// <param name="title">标题内容</param>
        /// <returns></returns>
        public static IntPtr 窗口_标题找句柄(string title)
        {
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.MainWindowTitle.IndexOf(title) != -1)
                {
                    return p.MainWindowHandle;
                }
            }
            return IntPtr.Zero;
        }

        /// </summary>
        /// 根据标题和类名找句柄
        /// <summary>
        /// <param name="IpClassName">窗口类名 如果为"" 则只根据标题查找</param>
        /// <param name="IpClassName">窗口标题 如果为"" 则只根据类名查找</param>
        /// <returns>找不到则返回0</returns>
        public static int 窗口_找句柄(string IpClassName, string IpTitleName)
        {
            if (IpTitleName == "" && IpClassName != "")
            {
                return (int)FindWindow(IpClassName, null);
            }
            else if (IpClassName == "" && IpTitleName != "")
            {
                return (int)FindWindow(null, IpTitleName);
            }
            else if (IpClassName != "" && IpTitleName != "")
            {
                return (int)FindWindow(IpClassName, IpTitleName);
            }
            return 0;
        }

        [DllImport("user32.dll", EntryPoint = "FindWindowA")]
        private static extern IntPtr FindWindowA(string IpClassName, string IpWindowName);
        /// <summary>
        /// 窗口_找句柄
        /// </summary>
        /// <param name="IpClassName">类名</param>
        ///  <param name="IpWindowName">标题</param>
        /// <returns></returns>
        public static IntPtr 窗口_找精确句柄(string IpClassName, string IpWindowName)
        {
          return  FindWindowA(IpClassName, IpWindowName);
        }
        

        /// <summary>
        /// 查找句柄
        /// </summary>
        /// <param name="IpClassName">类名</param>
        /// <returns></returns>
        public static IntPtr 窗口_类名找句柄(string IpClassName)
        {
            return FindWindow(IpClassName, null);
        }

        /// <summary>
        /// 找到句柄
        /// </summary>
        /// <param name="p">坐标</param>
        /// <returns></returns>
        public static IntPtr GetHandle(Point p)
        {
            return WindowFromPoint(p);
        }
        
        [DllImport("user32.dll")]

        private static extern int ShowWindow(IntPtr hwnd, int nCmdShow);
        /// <summary>
        /// 窗口_置状态
        /// </summary>
        /// <param name="hwnd">句柄</param>
        /// <param name="nCmdShow">0 隐藏取消激活 1 还原激活 2 最小化激活 3 最大化激活 4 还原 6 最小化取消激活 7 最小化 9 还原激活</param>
        public static int 窗口_置状态(IntPtr hwnd, int nCmdShow)
        {
            return ShowWindow(hwnd, nCmdShow);
        }

        [DllImport("user32.dll")]

        private static extern int CloseWindow(IntPtr hwndw);
        /// <summary>
        /// 窗口_最小化
        /// </summary>
        /// <param name="hwndw">句柄</param>
        public static bool 窗口_最小化(IntPtr hwndw)
        { int 是否成功 = CloseWindow(hwndw);
            if (是否成功 ==0)
            {
                return false;
            }
            else return true;
           
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        public static bool 窗口_是否可见(IntPtr hWnd)
        {
          return  IsWindowVisible( hWnd);
        }
    

        [DllImport("user32.dll")]       
        private static extern int UpdateWindow(IntPtr hWnd);
        /// <summary>
        /// 窗口_刷新
        /// </summary>
        /// <param name="hwndw">句柄</param>
        /// <returns>0为刷新失败,反之成功</returns>
        public static int 窗口_刷新(IntPtr hwndw)
        { 
            return UpdateWindow(hwndw);
        }
       
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        ///  ///<summary>
        /// 该函数设置由不同线程产生的窗口的显示状态
        /// </summary>
        /// <param name="hWnd">窗口句柄</param>
        /// <param name="fAltTab">窗口模式</param>
        /// <returns>0 隐藏取消激活 1 还原激活 2 最小化激活 3 最大化激活 4 还原 6 最小化取消激活 7 最小化 9 还原激活</returns> 
        public static void 窗口_显示隐藏(IntPtr hWnd, int fAltTab)
        {
            ShowWindowAsync(hWnd, fAltTab);
        }
        /// <summary>
        /// 激活窗口API
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="fAltTab"></param>
        [DllImport("user32.dll")]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
        /// <summary>
        /// 激活窗口 
        /// </summary>
        /// <param name="hwnd"></param>
        public static void 窗口_激活(IntPtr hWnd, bool fAltTab)
        {
            SwitchToThisWindow(hWnd, fAltTab);
        }
        /// <summary>
        /// 激活窗口 并显示到最前面
        /// </summary>
        /// <param name="hwnd"></param>
        public static void 窗口_激活显示(IntPtr hwnd)
        {
            ShowWindowAsync(hwnd, 1);//显示
            SetForegroundWindow(hwnd);//当到最前端
        }
        public static bool 窗口_置焦点(IntPtr hWnd)
        {
         return   SetForegroundWindow(hWnd);//当到最前端
        }
        /// <summary>
        ///  该函数将创建指定窗口的线程设置到前台，并且激活该窗口。键盘输入转向该窗口，并为用户改各种可视的记号。
        ///  系统给创建前台窗口的线程分配的权限稍高于其他线程。 
        /// </summary>
        /// <param name="hWnd">将被激活并被调入前台的窗口句柄</param>
        /// <returns>如果窗口设入了前台，返回值为非零；如果窗口未被设入前台，返回值为零</returns>
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        
       
       

      

        /// <summary>
        /// 窗口置顶
        /// </summary>
        /// <param name="hWnd">窗口句柄</param>
        /// <param name="is_activ">是否置顶 为false 取消置顶</param>
        /// <returns></returns>
        public static int 窗口_置顶(IntPtr hWnd, bool is_activ)
        {
            int is_top ;
            if (is_activ == true)
            {
                is_top = -1;
            }
            else is_top = -2;

            return SetWindowPos(hWnd, is_top, 0, 0, 0, 0, 1 | 2);
        }

        public delegate bool ChildWindowsProc(IntPtr hwnd, int lParam);

       


        /// <summary>
        /// 窗口置顶 或设置大小
        /// </summary>
        /// <param name="窗口句柄"></param>
        /// <param name="最前-1,最后-2"></param>
        /// <param name="X坐标"></param>
        /// <param name="Y坐标"></param>
        /// <param name="改变程序大小"></param>
        /// <param name="改变程序大小"></param>
        /// <param name="选项"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);




        [DllImport("user32.dll")]
        private static extern bool EnumWindows(WNDENUMPROC lpEnumFunc, int lParam);
        [DllImport("user32.dll")]
        private static extern int GetWindowTextW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder lpString, int nMaxCount);

    
        public enum ShowEnum
        {
            SW_Close = 0,
            SW_NORMAL = 1,
            SW_MINIMIZE = 2,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_RESTORE = 9,//还原
            SW_SHOWDEFAULT = 10
        }
        /// <summary>
        /// 查找子窗口句柄
        /// </summary>
        /// <param name="hwndParent">  要查找子窗口的父窗口句柄,是 0, 则函数以桌面窗口为父窗口</param>
        /// <param name="hwndChildAfter">前一个同目录级同名窗口句柄</param>
        /// <param name="lpszClass">类名</param>
        /// <param name="lpszWindow">标题</param>
        /// <returns></returns>
        public static IntPtr 窗口_取子窗口精确句柄(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow)
        {
            return FindWindowEx(hwndParent, hwndChildAfter, lpszClass, lpszWindow);
        }

        /// <summary>
        /// 查找子窗口句柄
        /// </summary>
        /// <param name="hwndParent">  要查找子窗口的父窗口句柄,是 0, 则函数以桌面窗口为父窗口</param>
        /// <param name="hwndChildAfter">前一个同目录级同名窗口句柄</param>
        /// <param name="lpszClass">类名</param>
        /// <returns></returns>
        public static IntPtr 窗口_取子窗口句柄(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass)
        {
            return FindWindowEx(hwndParent, hwndChildAfter, lpszClass, null);
        }

        /// <summary>
        /// 查找全部子窗口句柄
        /// </summary>
        /// <param name="hwndParent">父窗口句柄</param>
        /// <param name="className">类名</param>
        /// <returns></returns>
        public static List<IntPtr> 窗口_取全部子窗口句柄(IntPtr hwndParent, string className)
        {
            List<IntPtr> resultList = new List<IntPtr>();
            for (IntPtr hwndClient = 窗口_取子窗口句柄(hwndParent, IntPtr.Zero, className); hwndClient != IntPtr.Zero; hwndClient = 窗口_取子窗口句柄(hwndParent, hwndClient, className))
            {
                resultList.Add(hwndClient);
            }

            return resultList;
        }

        /// <summary>
        /// 给窗口发送文本内容
        /// </summary>
        /// <param name="hWnd">句柄</param>
        /// <param name="lParam">要发送的内容</param>
        public static void 窗口_发送文本(IntPtr hWnd, string lParam)
        {

            SendMessageA(hWnd, WindowsMessage.WM_SETTEXT, 0, lParam);
        }

        /// <summary>
        /// 获得窗口内容或标题
        /// </summary>
        /// <param name="hWnd">句柄</param>
        /// <returns></returns>
        public static string 窗口_取窗口标题(IntPtr hWnd)
        {
            StringBuilder result = new StringBuilder(128);
            GetWindowText(hWnd, result, result.Capacity);
            return result.ToString();
        }

        /// <summary>
        /// 窗口在屏幕位置
        /// </summary>
        /// <param name="hWnd">句柄</param>
        /// <returns></returns>
        public static Rectangle 窗口_在屏幕位置(IntPtr hWnd)
        {
            Rectangle result = default(Rectangle);
            GetWindowRect(hWnd, ref result);
            result.Width = result.Width - result.X;
            result.Height = result.Height - result.Y;

            return result;
        }

        /// <summary>
        /// 窗口相对屏幕位置转换成父窗口位置
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="rect"></param>
        /// <returns></returns>
        public static Rectangle 窗口_相对屏幕位置转换成父窗口位置(IntPtr hWnd, Rectangle rect)
        {
            Rectangle result = rect;
            ScreenToClient(hWnd, ref result);
            return result;
        }

        /// <summary>
        /// 窗口大小
        /// </summary>
        /// <param name="hWnd"></param>
        /// <returns></returns>
        public static Rectangle 窗口_窗口大小(IntPtr hWnd)
        {
            Rectangle result = default(Rectangle);
            GetClientRect(hWnd, ref result);
            return result;
        }

        ///=====================================================================  
        //判断窗口是否存在
        [DllImport("user32", EntryPoint = "IsWindow")]
        private static extern bool IsWindow(IntPtr hWnd);
        


        /// <summary>
        /// 判断窗口是否存在
        /// </summary>
        /// <param name="Hwnd">窗口句柄</param>
        /// <returns>存在返回 true 不存在返回 false</returns>
        public static bool 窗口_窗口是否存在(int Hwnd)
        {
            if (IsWindow((IntPtr)Hwnd))
            {
                return true;
            }
            return false;
        }
        //修改指定窗口标题
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowText(IntPtr hWnd, string text);

        /// <summary>
        /// 设置窗口标题
        /// </summary>
        /// <param name="Hwnd">窗口句柄</param>
        /// <param name="newtext">新标题</param>
        public static void 窗口_置标题(int Hwnd, string newtext)
        {
            SetWindowText((IntPtr)Hwnd, newtext);
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsZoomed(IntPtr hWnd);

        //获得顶层窗口
        [DllImport("user32", EntryPoint = "GetForegroundWindow")]
        private static extern IntPtr GetForegroundwindow();

        /// <summary>
        /// 获得顶层窗口
        /// </summary>
        /// <returns>返回 窗口句柄</returns>
        public static int 窗口_获取顶层窗口()
        {
            return (int)GetForegroundwindow();
        }

        /// <summary>
        /// 得到窗口状态
        /// </summary>
        /// <param name="Hwnd">窗口句柄</param>
        /// <param name="flag">
        /// 操作方式
        /// 1：判断窗口是否最小化
        /// 2：判断窗口是否最大化
        /// 3：判断窗口是否激活
        /// </param>
        /// <returns>满足条件返回 true</returns>
        public static bool 窗口_判断当前窗口状态(int Hwnd, int flag)
        {
            switch (flag)
            {
                case 1:
                    return IsIconic((IntPtr)Hwnd);
                case 2:
                    return IsZoomed((IntPtr)Hwnd);
                case 3:
                    if (Hwnd != 窗口_获取顶层窗口())
                    {
                        return false;
                    }
                    break;
            }
            return true;

        }

        /// <summary>
        /// 得到窗口上一级窗口的句柄
        /// </summary>
        /// <param name="ChildHwnd">子窗口句柄</param>
        /// <returns> 返回 窗口句柄 找不到返回 0</returns>
        public static int 窗口_取上一级句柄(int ChildHwnd)
        {
            return (int)GetParent((IntPtr)ChildHwnd);
        }

        /// </summary>
        /// 得到指定窗口类名
        /// <summary>
        /// <param name="hWnd">句柄</param>
        /// <returns>找不到返回""</returns>
        public static string 窗口_取类名(int hWnd)
        {
            return GetClassName((IntPtr)hWnd);
        }
        /// </summary>
        /// 得到指定窗口类名
        /// <summary>
        /// <param name="hWnd">句柄</param>
        /// <returns>找不到返回""</returns>
        public static string GetClassName(IntPtr hWnd)
        {
            StringBuilder lpClassName = new StringBuilder(128);
            if (GetClassName(hWnd, lpClassName, lpClassName.Capacity) == 0)
            {
                return "";
            }
            return lpClassName.ToString();
        }

        /// <summary>
        /// 得到指定坐标的窗口句柄
        /// </summary>
        /// <param name="x">X坐标</param>
        /// <param name="y">Y坐标</param>
        /// <returns>找不到返回 0</returns>
        public static int 窗口_取指定坐标句柄(int x, int y)
        {
            Point p = new Point(x, y);
            IntPtr formHandle = WindowFromPoint(p);//得到窗口句柄
            return (int)formHandle;
        }


        /// <summary>
        /// 窗口_取句柄_模糊
        /// </summary>
        /// <param name="title">窗口标题</param>
        /// <returns>返回 窗口句柄 找不到返回 0</returns>
        public static int 窗口_取句柄_模糊(string title)
        {
            //按照窗口标题来寻找窗口句柄
            Process[] ps = Process.GetProcesses();
            string WindowHwnd ;
            foreach (Process p in ps)
            {
                if (p.MainWindowTitle.IndexOf(title) != -1)
                {
                    WindowHwnd = p.MainWindowHandle.ToString();
                    return int.Parse(WindowHwnd);
                }
            }
            return 0;
        }



        /// <summary>
        /// 根据窗口标题模糊查找符合条件的所有窗口句柄
        /// </summary>
        /// <param name="title">窗口标题关键字</param>
        /// <returns>返回 窗口句柄 多个句柄以"|" 隔开，找不到返回""</returns>
        public static string 窗口_取句柄_模糊取全部句柄(string title)
        {
            //按照窗口标题来寻找窗口句柄
            Process[] ps = Process.GetProcesses();
            string WindowHwnd = "";
            foreach (Process p in ps)
            {
                if (p.MainWindowTitle.IndexOf(title) != -1)
                {
                    if (WindowHwnd == "")
                    {
                        WindowHwnd = p.MainWindowHandle.ToString();
                    }
                    else
                    {
                        WindowHwnd = WindowHwnd + "|" + p.MainWindowHandle.ToString();
                    }
                }
            }
            if (WindowHwnd == "")
            {
                return "";
            }
            return WindowHwnd;
        }

        /// <summary>
        /// 不改变尺寸移动窗口到指定位置
        /// </summary>
        /// <param name="Hwnd">窗口句柄</param>
        /// <param name="X">目的地左上角X</param>
        /// <param name="Y">目的地左上角Y</param>
        /// <returns>移动成功返回 true</returns>
        public static bool 窗口_移动(int Hwnd, int X, int Y)
        {
            Rectangle rect = new Rectangle();
            GetWindowRect((IntPtr)Hwnd, ref rect);
            return MoveWindow((IntPtr)Hwnd, rect.Left, rect.Top, rect.Right, rect.Bottom, true);
        }

        /// <summary>
        /// 改变尺寸移动窗口到指定位置
        /// </summary>
        /// <param name="Hwnd">窗口句柄</param>
        /// <param name="X">目的地左上角X</param>
        /// <param name="Y">目的地左上角Y</param>
        /// <param name="Width">新宽度</param>
        /// <param name="Height">新高度</param>
        /// <returns>移动成功返回 true</returns>
        public static bool 窗口_移动并改变尺寸(int Hwnd, int X, int Y, int Width, int Height)
        {
            return MoveWindow((IntPtr)Hwnd, X, Y, Width, Height, true);
        }

        [DllImport("user32.dll", EntryPoint = "SetParent")]
        public static extern int SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        /// <summary>
        /// 窗口_置父
        /// </summary>
        /// <param name="hWndChild">父窗口句柄</param>
        /// <param name="hWndNewParent">子窗口句柄</param>
        /// <returns></returns>
        public static int 窗口_置父 (IntPtr hWndChild, IntPtr hWndNewParent)
        {
            return SetParent( hWndChild,  hWndNewParent);
        }
        [DllImport("user32.dll")]
        public static extern void PostMessage(IntPtr hwnd, int msg, int wParam, int lParam);
        public static void 窗口_关闭(IntPtr hwnd, int msg, int wParam, int lParam)
        {
            PostMessage( hwnd, 16, 0, 0);
        }
    }
}

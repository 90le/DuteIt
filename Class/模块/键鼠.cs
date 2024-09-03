using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace DuteIT
{
    public class 键鼠

    {
        //[DllImport("user32.dll", EntryPoint = "ClientToScreen")]
        //private static extern int ClientToScreen(IntPtr hWnd, Point p);
        ///// <summary>
        ///// 鼠标移动后台
        ///// </summary>
        ///// <param name="hWnd">句柄</param>
        ///// <param name="p">坐标</param>
        //public static void 鼠标_移动后台(IntPtr hWnd, Point p)
        //{
        //    ClientToScreen(hWnd, p);

        //}

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern int SetCursorPos(int x, int y);
        /// <summary>
        ///鼠标_移动
        /// </summary>
        /// <param name="x">X坐标</param>
        /// <param name="y">Y坐标</param>
        public static int 鼠标_移动(int x, int y)
        {
          return  SetCursorPos(x, y);

        }


        enum MouseEventFlag : uint
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800,
            VirtualDesk = 0x4000,
            Absolute = 0x8000
        }

        [DllImport("user32.dll")]
        static extern void mouse_event(MouseEventFlag flags, int dx, int dy, int data, int extraInfo);
        /// <summary>
        /// 鼠标_单机
        /// </summary>
        /// <param name="mes">1 = 鼠标左键单击；2 = 鼠标右键单击；</param>
        public static void 鼠标_单机(int mes)
        {
            if (mes==1)
            {
                mouse_event(MouseEventFlag.LeftDown, 0, 0, 0, 0);
                Thread.Sleep(20);
                mouse_event(MouseEventFlag.LeftUp, 0, 0, 0, 0);

            }
            if (mes == 2)
            {
                mouse_event(MouseEventFlag.RightDown, 0, 0, 0, 0);
                Thread.Sleep(20);
                mouse_event(MouseEventFlag.RightUp, 0, 0, 0, 0);
            }
       
        }

        /// <summary>
        /// 鼠标_消息
        /// </summary>
        /// <param name="HWD">句柄</param>
        /// <param name="键"> 1 #左键   2 #右键   3 #中键  </param>
        /// <param name="控制"> 1 #单击   2 #双击   3 #按下  4 #放开</param>
        public static void 鼠标_消息(IntPtr HWD, int 键, int 控制)
        {
            if (键 == 1)
            {
                if (控制 == 1)
                {
                    PostMessageA(HWD, 513, 1, 0);//左键按下
                    PostMessageA(HWD, 514, 0, 0);//左键放开
                }
                else if (控制 == 2)
                {
                    PostMessageA(HWD, 513, 1, 0);
                    PostMessageA(HWD, 514, 0, 0);
                    PostMessageA(HWD, 515, 0, 0);

                }
                else if (控制 == 3)
                {
                    PostMessageA(HWD, 513, 1, 0);
                }
                else if (控制 == 4)
                {
                    PostMessageA(HWD, 514, 0, 0);
                }

            }
            if (键 == 2)
            {
                if (控制 == 1)
                {
                    PostMessageA(HWD, 516, 2, 0);
                    PostMessageA(HWD, 517, 2, 0);
                }
                else if (控制 == 2)
                {
                    PostMessageA(HWD, 516, 2, 0);
                    PostMessageA(HWD, 517, 2, 0);
                    PostMessageA(HWD, 518, 0, 0);
                }
                else if (控制 == 3)
                {
                    PostMessageA(HWD, 516, 2, 0);
                }
                else if (控制 == 4)
                {
                    PostMessageA(HWD, 517, 2, 0);
                }

            }
            if (键 == 3)
            {
                if (控制 == 1)
                {
                    PostMessageA(HWD, 519, 16, 0);
                    PostMessageA(HWD, 520, 0, 0);
                }
                else if (控制 == 2)
                {
                    PostMessageA(HWD, 519, 16, 0);
                       PostMessageA(HWD, 520, 0, 0);
                        PostMessageA(HWD, 521, 0, 0);

                }
                else if (控制 == 3)
                {
                    PostMessageA(HWD, 519, 16, 0);
                }
                else if (控制 == 4)
                {
                    PostMessageA(HWD, 520, 0, 0);
                }


            }
        }     

        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]
        private static extern bool GetCursorPos(out Point pt);
        /// <summary>
        /// 获取鼠标位置的坐标
        /// </summary>
        /// <returns></returns>
        public static Point 鼠标_取当前位置坐标()
        {
            Point p = new Point();
            if (GetCursorPos(out p))
            {
                return p;
            }
            return default(Point);
        }
       

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr WindowFromPoint(Point pt);
        /// <summary>
        /// 找到句柄
        /// </summary>
        /// <param name="p">坐标</param>
        /// <returns></returns>
        public static IntPtr GetHandle(Point p)
        {
            return WindowFromPoint(p);
        }

        /// <summary>
        /// 得到鼠标指向的HWD
        /// </summary>
        /// <returns>找不到则返回-1</returns>
        public static int 鼠标_取当前HWD()
        {
            try
            {
                return (int)GetHandle(鼠标_取当前位置坐标());
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return -1;
        }

        //消息发送API
        [DllImport("User32.dll", EntryPoint = "PostMessageA")]
        private static extern int PostMessageA(
            IntPtr hWnd,        // 信息发往的窗口的句柄
            int Msg,            // 消息ID
            int wParam,         // 参数1
            int lParam            // 参数2
        );
        
        public static void 键盘_消息(IntPtr hWnd, int Msg, byte keyy)//句柄,按键功能,键盘按键
        {
     
            if (Msg == 1 )                        //输入字符(大写)
            {
                PostMessageA(hWnd, 258, keyy, 0);
            }
            else if (Msg == 2)                     //输入字符(小写)
            {
                PostMessageA(hWnd, 260, keyy, 0);
            }
            else if (Msg == 3)                      
            {
                PostMessageA(hWnd, 260, keyy, 0);//3=按下
            }
            else if (Msg == 4)
            {
                PostMessageA(hWnd, 261, keyy, 0);//4=放开
            }
            else if (Msg == 5)
            {
                PostMessageA(hWnd, 260, keyy, 0);//5=单击
                PostMessageA(hWnd, 261, keyy, 0);

            }
        }

        [DllImport("user32.dll", EntryPoint = "keybd_event")]

        private static extern void keybd_event(
       byte bVk, //虚拟键值
       byte bScan,// 一般为0
       int dwFlags, //这里是整数类型 0 为按下，2为释放
       int dwExtraInfo //这里是整数类型 一般情况下设成为0
        );
        public static void 键盘_单机(byte keyy, int dwFlags)//键代码,0 为按下，2为释放
        {
           
           keybd_event(keyy, 0, dwFlags, 0);
            
        }

  

    }
}

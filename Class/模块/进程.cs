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
    public  class 进程
    {
        /// <summary>
        /// 根据进程名获得窗口句柄 - 不需要带上进程后缀
        /// </summary>
        /// <param name="ProssName">进程名</param>
        /// <returns>窗口句柄 多个用"|"隔开 找不到返回 ""</returns>
        public static string 进程_进程名取多个句柄(string ProssName)
        {
            string Hwnd = "";
            Process[] pp = Process.GetProcessesByName(ProssName);
            for (int i = 0; i < pp.Length; i++)
            {
                if (pp[i].ProcessName == ProssName)
                {
                    if (Hwnd == "")
                    {
                        Hwnd = pp[i].MainWindowHandle.ToString();
                    }
                    else
                    {
                        Hwnd = Hwnd + "|" + pp[i].MainWindowHandle.ToString();
                    }
                }
            }
            return Hwnd;
        }


        /// <summary>
        /// 根据进程名获得窗口句柄 - 不需要带上进程后缀
        /// </summary>
        /// <param name="ProssName">进程名</param>
        /// <returns>窗口句柄 找不到返回 0</returns>
        public static int 进程_进程名取句柄(string ProssName)
        {
            Process[] pp = Process.GetProcessesByName(ProssName);
            for (int i = 0; i < pp.Length; i++)
            {
                if (pp[i].ProcessName == ProssName)
                {
                    return (int)pp[i].MainWindowHandle;
                }
            }
            return 0;
        }



      


        /// <summary>
        /// 根据进程名结束进程 不需要后缀 多个 相同进程名 会被一起结束
        /// </summary>
        /// <param name="ProssName"></param>
        public static void 进程_取进程名结束 (string ProssName)
        {
            string newName = ProssName.Replace(".exe", "");
            try
            {

                Process[] pp = Process.GetProcessesByName(newName);

                for (int i = 0; i < pp.Length; i++)

                {
                    if (pp[i].ProcessName == newName)
                    {
                        pp[i].Kill();

                    }

                }

            }

            catch (System.Exception ex)

            {
                Console.WriteLine(ex.Message);

            }

        }
        /// <summary>
        /// 根据进程名和进程PID，关闭指定进程 - 进程名不需要带后缀
        /// </summary>
        /// <param name="ProssName">进程名</param>
        /// <param name="ClosePid">需要关闭的进程PID</param>
        public static void 进程_取名称和PID结束(string ProssName, int ClosePid)
        {
            Process[] pp = Process.GetProcessesByName(ProssName);
            for (int i = 0; i < pp.Length; i++)
            {
                if (pp[i].Id == ClosePid)
                {
                    pp[i].Kill();
                }
            }
        }
        /// <summary>
        /// 根据进程名获得进程Process对象的集合 - 不需要带上进程后缀
        /// </summary>
        /// <param name="ProssName">进程名</param>
        /// <param name="Pro">进程Process 对象集合</param>
        /// <returns>找不到返回 false</returns>
        public static bool 进程_进程名取Process集合(string ProssName, ref List<Process> Pro)
        {
            bool finded = false;
            Process[] pp = Process.GetProcessesByName(ProssName);
            for (int i = 0; i < pp.Length; i++)
            {
                if (pp[i].ProcessName == ProssName)
                {
                    finded = true;
                    Pro.Add(pp[i]);
                }
            }
            if (finded)
            {
                return true;
            }
            return false;
        }

        [DllImport("user32", EntryPoint = "GetWindowThreadProcessId")]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);

        /// <summary>
        /// 根据窗口句柄获得进程PID和线程PID
        /// </summary>
        /// <param name="hwnd">句柄</param>
        /// <param name="pid">返回 进程PID</param>
        /// <returns>方法的返回值，线程PID，进程PID和线程PID这两个概念不同</returns>
        public static int 进程_窗口取进程PID(int hwnd, out int pid)
        {
            pid = 0;
            return GetWindowThreadProcessId((IntPtr)hwnd, out pid);
        }
        /// <summary>
        /// 根据窗口标题获得进程Process对象
        /// </summary>
        /// <param name="Title">窗口标题</param>
        /// <param name="Pro">进程Process 对象</param>
        /// <returns>找不到返回 false</returns>
        public static bool 进程_窗口标题取进程Process(string Title, out Process Pro)
        {
            Pro = null;
            Process[] arrayProcess = Process.GetProcesses();
            foreach (Process p in arrayProcess)
            {
                if (p.MainWindowTitle.IndexOf(Title) != -1)
                {
                    Pro = p;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 根据窗口标题查找窗口进程PID-返回List
        /// </summary>
        /// <param name="windowTitle">窗口标题</param>
        /// <returns>List</returns>
        public static List<int> 进程_窗口标题取进程PED(string Title)
        {
            List<int> list = new List<int>();
            Process[] arrayProcess = Process.GetProcesses();
            foreach (Process p in arrayProcess)
            {
                if (p.MainWindowTitle.IndexOf(Title) != -1)
                {
                    list.Add(p.Id);

                }
            }
            return list;
        }
        /// <summary>
        /// 根据进程名获得进程PID - 不需要带上进程后缀
        /// </summary>
        /// <param name="ProssName">进城名</param>
        /// <returns>进城PID 找不到返回 0</returns>
        public static int 进程_名取PID(string ProssName)
        {
            Process[] pp = Process.GetProcessesByName(ProssName);
            for (int i = 0; i < pp.Length; i++)
            {
                if (pp[i].ProcessName == ProssName)
                {
                    return pp[i].Id;
                }
            }
            return 0;
        }

        /// <summary>
        /// 通过句柄获得进程路径
        /// </summary>
        /// <param name="hwnd">句柄</param>
        /// <returns>返回 进程路径 找不到返回""</returns>
        public static string 进程_句柄取进程路径(int hwnd)
        {
            string path = "";
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if ((int)p.MainWindowHandle == hwnd)
                {
                    path = p.MainModule.FileName.ToString();
                }
            }
            return path;
        }
        /// <summary>
        /// 通过进程名获得进程路径 不需要后缀
        /// </summary>
        /// <param name="hwnd">句柄</param>
        /// <returns>返回 进程路径 找不到返回""</returns>
        public static string 进程_名取路径(string prossName)
        {
            string path = "";
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName == prossName)
                {
                    path = p.MainModule.FileName.ToString();
                    return path;
                }
            }


            return "";
        }
        /// <summary>
        /// 取进程ID
        /// </summary>
        /// <param name="hWndParent">窗口句柄</param>
        /// <param name="intPtr"> 进程id</param>
        /// <returns>拥有窗口的线程的标识符</returns>
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWndParent, ref IntPtr lpdwProcessId);
        /// <summary>
        /// 进程_取窗口进程ID
        /// </summary>
        /// <param name="hWndParent">窗口句柄</param>
        /// <param name="lpdwProcessId">进程id</param>
        /// <returns>拥有窗口的线程的标识符</returns>
        public static int 进程_取窗口进程ID(IntPtr hWndParent, ref IntPtr lpdwProcessId)
        {
            return GetWindowThreadProcessId(hWndParent, ref lpdwProcessId);
        }
    }
}

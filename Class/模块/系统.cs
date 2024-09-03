using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace DuteIT
{
    public class 系统
    {     /// <summary>
          /// 物理内存
          /// </summary>
          /// <returns></returns>
        public static string 系统_物理内存()
        {
            var st = string.Empty;
            var mc = new ManagementClass("Win32_ComputerSystem");
            var moc = mc.GetInstances();
            foreach (var o in moc)
            {
                var mo = (ManagementObject)o;
                st = mo["TotalPhysicalMemory"].ToString();
            }
            return st;
        }
        /// <summary>
        /// 获得CPU编号
        /// </summary>
        /// <returns></returns>
        public static string 系统_获得CPU编号()
        {
            var cpuid = string.Empty;
            var mc = new ManagementClass("Win32_Processor");
            var moc = mc.GetInstances();
            foreach (var o in moc)
            {
                var mo = (ManagementObject)o;
                cpuid = mo.Properties["ProcessorId"].Value.ToString();
            }
            return cpuid;
        }
        /// <summary>
        /// 获取硬盘序列号
        /// </summary>
        /// <returns></returns>
        public static string 系统_获取硬盘序列号()
        {
            //这种模式在插入一个U盘后可能会有不同的结果，如插入我的手机时
            //这名话解决有多个物理盘时产生的问题，只取第一个物理硬盘

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");
                string strHardDiskID = null;
                foreach (ManagementObject mo in searcher.Get())
                {
                    strHardDiskID = mo["SerialNumber"].ToString().Trim();

                    break;
                }
                return strHardDiskID;
            }
            catch
            {
                return "";
            }

        }
        /// <summary>
        /// 获取硬盘的大小
        /// </summary>
        /// <returns></returns>
        public static string 系统_获取硬盘的大小()
        {


            ManagementClass mc = new ManagementClass("Win32_DiskDrive");
            ManagementObjectCollection moj = mc.GetInstances();
            foreach (ManagementObject m in moj)
            {
                long.TryParse(m.Properties["Size"].Value.ToString(), out long size);

                if (size > 0)
                {
                    size = size / 1024 / 1024 / 1024;
                    return size.ToString() + " G";
                }

                return "-1";
            }
            return "-1";



        }
        /// <summary>
        /// 获取网卡硬件地址
        /// </summary>
        /// <returns></returns> 
        public static string 系统_获取网卡硬件地址()
        {
            var mac = "";
            var mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            var moc = mc.GetInstances();
            foreach (var o in moc)
            {
                var mo = (ManagementObject)o;
                if (!(bool)mo["IPEnabled"]) continue;
                mac = mo["MacAddress"].ToString();
                break;
            }
            return mac;
        }
        /// <summary>
        /// 获取IP地址
        /// </summary>
        /// <returns></returns>
        public static string 系统_获取IP地址()
        {
            var st = string.Empty;
            var mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            var moc = mc.GetInstances();
            foreach (var o in moc)
            {
                var mo = (ManagementObject)o;
                if (!(bool)mo["IPEnabled"]) continue;
                var ar = (Array)(mo.Properties["IpAddress"].Value);
                st = ar.GetValue(0).ToString();
                break;
            }
            return st;
        }
        /// <summary>
        /// 获取计算机名
        /// </summary>
        /// <returns></returns>
        public static string 系统_获取计算机名()
        {
            return Environment.MachineName;
        }
        /// <summary>
        /// 操作系统类型
        /// </summary>
        /// <returns></returns> 
        public static string 系统_取操作系统类型()
        {
            var st = string.Empty;
            var mc = new ManagementClass("Win32_ComputerSystem");
            var moc = mc.GetInstances();
            foreach (var o in moc)
            {
                var mo = (ManagementObject)o;
                st = mo["SystemType"].ToString();
            }
            return st;
        }

        /// <summary>
        /// 显卡PNPDeviceID
        /// </summary>
        /// <returns></returns>
        public static string 系统_显卡PNPDeviceID()
        {
            var st = "";
            var mos = new ManagementObjectSearcher("Select * from Win32_VideoController");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["PNPDeviceID"].ToString();
            }
            return st;
        }

        /// <summary>
        /// 声卡PNPDeviceID
        /// </summary>
        /// <returns></returns>
        public static string 系统_声卡PNPDeviceID()
        {
            var st = string.Empty;
            var mos = new ManagementObjectSearcher("Select * from Win32_SoundDevice");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["PNPDeviceID"].ToString();
            }
            return st;
        }

        /// <summary>
        /// CPU版本信息
        /// </summary>
        /// <returns></returns>
        public static string 系统_CPU版本信息()
        {
            var st = string.Empty;
            var mos = new ManagementObjectSearcher("Select * from Win32_Processor");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["Version"].ToString();
            }
            return st;
        }
        /// <summary>
        /// CPU名称信息
        /// </summary>
        /// <returns></returns>
        public static string 系统_CPU名称信息()
        {
            var st = string.Empty;
            var driveId = new ManagementObjectSearcher("Select * from Win32_Processor");
            foreach (var o in driveId.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["Name"].ToString();
            }
            return st;
        }
        /// <summary>
        /// CPU制造厂商
        /// </summary>
        /// <returns></returns>
        public static string 系统_CPU制造商()
        {
            var st = string.Empty;
            var mos = new ManagementObjectSearcher("Select * from Win32_Processor");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["Manufacturer"].ToString();
            }
            return st;
        }
        /// <summary>
        /// 主板制造厂商
        /// </summary>
        /// <returns></returns>
        public static string 系统_主板制造商()
        {
            var query = new SelectQuery("Select * from Win32_BaseBoard");
            var mos = new ManagementObjectSearcher(query);
            var data = mos.Get().GetEnumerator();
            data.MoveNext();
            var board = data.Current;
            return board.GetPropertyValue("Manufacturer").ToString();
        }
        /// <summary>
        /// 主板编号
        /// </summary>
        /// <returns></returns>
        public static string 系统_主板编号()
        {
            var st = string.Empty;
            var mos = new ManagementObjectSearcher("Select * from Win32_BaseBoard");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["SerialNumber"].ToString();
            }
            return st;
        }
        /// <summary>
        /// 主板型号
        /// </summary>
        /// <returns></returns>
        public static string 系统_主板型号()
        {
            var st = string.Empty;
            var mos = new ManagementObjectSearcher("Select * from Win32_BaseBoard");
            foreach (var o in mos.Get())
            {
                var mo = (ManagementObject)o;
                st = mo["Product"].ToString();
            }
            return st;
        }

        /// <summary>
        /// 判断操作系统是否为Windows98
        /// </summary>
        public static bool 系统_是否Windows98
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32Windows) && (Environment.OSVersion.Version.Minor == 10) && (Environment.OSVersion.Version.Revision.ToString() != "2222A");
            }
        }
        /// <summary>
        /// 判断操作系统是否为Windows98第二版
        /// </summary>
        public static bool 系统_是否Windows98第二版
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32Windows) && (Environment.OSVersion.Version.Minor == 10) && (Environment.OSVersion.Version.Revision.ToString() == "2222A");
            }
        }

        /// <summary>
        /// 判断操作系统是否为Windows2000
        /// </summary>
        public static bool 系统_是否Windows2000
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major == 5) && (Environment.OSVersion.Version.Minor == 0);
            }
        }
        /// <summary>
        /// 判断操作系统是否为WindowsXP
        /// </summary>
        public static bool 系统_是否WindowsXP
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major == 5) && (Environment.OSVersion.Version.Minor == 1);
            }
        }

        /// <summary>
        /// 判断操作系统是否为Windows2003
        /// </summary>
        public static bool 系统_是否Windows2003
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major == 5) && (Environment.OSVersion.Version.Minor == 2);
            }
        }
        /// <summary>
        /// 判断操作系统是否为WindowsVista
        /// </summary>
        public static bool 系统_是否WindowsVista
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major == 6) && (Environment.OSVersion.Version.Minor == 0);
            }
        }
        /// <summary>
        /// 判断操作系统是否为Windows7
        /// </summary>
        public static bool 系统_是否Windows7
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major == 6) && (Environment.OSVersion.Version.Minor == 1);
            }
        }

        /// <summary>
        /// 判断操作系统是否为Unix
        /// </summary>
        public static bool 系统_是否Unix
        {
            get
            {
                return Environment.OSVersion.Platform == PlatformID.Unix;
            }
        }

        /// <summary>
        /// 是否是64位 false 为32位
        /// </summary>
        /// <returns></returns>
        public static bool 系统_是否64位()
        {
            bool type;
            type = Environment.Is64BitOperatingSystem;
            return type;
        }

        /// <summary>
        /// 获取计算机名
        /// </summary>
        /// <returns></returns>
        public static string 系统_取计算机名()
        {
            return Environment.MachineName;
        }
        /// <summary>
        /// 操作系统的登录用户名
        /// </summary>
        /// <returns></returns> 
        public static string 系统_取登录用户名()
        {
            return Environment.UserName;
        }

        //判断本机是否联网
        [DllImport("wininet.dll", EntryPoint = "InternetGetConnectedState")]
        //判断网络状况的方法,返回值true为连接，false为未连接
        private extern static bool InternetGetConnectedState(out int conState, int reder);
        #region 关机 重启 注销
        [DllImport("user32.dll", EntryPoint = "ExitWindowsEx", CharSet = CharSet.Ansi)]
        private static extern int ExitWindowsEx(int uFlags, int dwReserved);

        /// <summary>
        /// 判断是否联网
        /// </summary>
        /// <returns></returns>
        public static bool 系统_是否联网()
        {
            int n = 0;
            return InternetGetConnectedState(out n, 0);

        }

        /// <summary>
        /// 注销
        /// </summary>
        public static void 系统_注销()
        {
            //注销计算机
            ExitWindowsEx(0, 0);
        }

        /// <summary>
        /// 关机
        /// </summary>
        public static void 系统_关机()
        {
            //关机
            System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
            myProcess.StartInfo.FileName = "cmd.exe";//启动cmd命令
            myProcess.StartInfo.UseShellExecute = false;//是否使用系统外壳程序启动进程
            myProcess.StartInfo.RedirectStandardInput = true;//是否从流中读取
            myProcess.StartInfo.RedirectStandardOutput = true;//是否写入流
            myProcess.StartInfo.RedirectStandardError = true;//是否将错误信息写入流
            myProcess.StartInfo.CreateNoWindow = true;//是否在新窗口中启动进程
            myProcess.Start();//启动进程
            myProcess.StandardInput.WriteLine("shutdown -s -t 0");//执行关机命令
        }

        /// <summary>
        /// 重启电脑
        /// </summary>
        public static void 系统_重启()
        {
            //重启
            System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
            myProcess.StartInfo.FileName = "cmd.exe";//启动cmd命令
            myProcess.StartInfo.UseShellExecute = false;//是否使用系统外壳程序启动进程
            myProcess.StartInfo.RedirectStandardInput = true;//是否从流中读取
            myProcess.StartInfo.RedirectStandardOutput = true;//是否写入流
            myProcess.StartInfo.RedirectStandardError = true;//是否将错误信息写入流
            myProcess.StartInfo.CreateNoWindow = true;//是否在新窗口中启动进程
            myProcess.Start();//启动进程
            myProcess.StandardInput.WriteLine("shutdown -r -t 0");//执行重启计算机命令
        }


        #endregion
        #region 更改分辨率
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int ChangeDisplaySettings([In] ref DEVMODE lpDevMode, int dwFlags);

        public enum DMDO
        {
            DEFAULT = 0,
            D90 = 1,
            D180 = 2,
            D270 = 3
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        struct DEVMODE
        {
            public const int DM_DISPLAYFREQUENCY = 0x400000;
            public const int DM_PELSWIDTH = 0x80000;
            public const int DM_PELSHEIGHT = 0x100000;
            public const int DM_BITSPERPEL = 262144;
            private const int CCHDEVICENAME = 32;
            private const int CCHFORMNAME = 32;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public int dmPositionX;
            public int dmPositionY;
            public DMDO dmDisplayOrientation;
            public int dmDisplayFixedOutput;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }


        /// <summary>
        /// 更改分辨率
        /// </summary>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="displayFrequency">刷新频率 60，75，85，100</param>
        public static void 系统_更改分辨率(int width, int height, int displayFrequency = 60)
        {
            long RetVal = 0;
            DEVMODE dm = new DEVMODE();
            dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
            dm.dmPelsWidth = width;//宽
            dm.dmPelsHeight = height;//高
            dm.dmDisplayFrequency = displayFrequency;//刷新率
            dm.dmFields = DEVMODE.DM_PELSWIDTH | DEVMODE.DM_PELSHEIGHT | DEVMODE.DM_DISPLAYFREQUENCY | DEVMODE.DM_BITSPERPEL;
            RetVal = ChangeDisplaySettings(ref dm, 0);
        }

        #endregion
        /// <summary>
        /// 判断指定端口号是否被占用 占用返回true
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        internal static bool 系统_端口是否被占用(Int32 port)
        {
            bool result = false;
            try
            {
                System.Net.NetworkInformation.IPGlobalProperties iproperties = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties();
                System.Net.IPEndPoint[] ipEndPoints = iproperties.GetActiveTcpListeners();
                foreach (var item in ipEndPoints)
                {
                    if (item.Port == port)
                    {
                        result = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            return result;
        }
        /// <summary>
        /// 系统_宽带拨号连接
        /// </summary>
        /// <param name="UserS">宽带账户</param>
        /// <param name="PwdS">宽带密码</param>
        /// <returns></returns>
        public static string 系统_宽带拨号连接(string UserS, string PwdS)
        {
            string arg = @"rasdial.exe 宽带连接" + " " + UserS + " " + PwdS;
            return InvokeCmd(arg);
        }

        private static string InvokeCmd(string cmdArgs)
        {
            string Tstr = "";
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();


            p.StandardInput.WriteLine(cmdArgs);
            p.StandardInput.WriteLine("exit");
            Tstr = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            p.Close();
            return Tstr;
        }

     

    }
}

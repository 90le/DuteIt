using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DuteIT
{
    public class 文件
    {       /// <summary>
            /// 运行一个指定文件或者程序
            /// </summary>
            /// <param name="Path">文件路径</param>
            /// <returns>失败返回false</returns>
        public static bool 文件_运行(string Path)
        {
            try
            {
                Process pro = new Process();
                pro.StartInfo.FileName = @Path;
                pro.Start();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }
        /// <summary>
        /// 运行一个指定文件或者程序可以带上参数
        /// </summary>
        /// <param name="Path">文件路径</param>
        /// <param name="Flag">附带参数</param>
        /// <returns>失败返回false</returns>
        public static bool 文件_运行带参数(string Path, string Flag)
        {
            try
            {
                Process pro = new Process();
                pro.StartInfo.FileName = @Path;
                pro.StartInfo.Arguments = Flag;
                pro.Start();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }
        [DllImport("Kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool CloseHandle(IntPtr handle);
        /// <summary>
        /// 文件_关闭
        /// </summary>
        /// <param name="hwd">窗口句柄</param>
        public static void 文件_关闭(IntPtr hwd)
        {
            if (hwd != null)
            {
                CloseHandle(hwd);
            }
        }


    /// <summary>
    /// 移动文件
    /// </summary>
    /// <param name="dirPath">文件原始路径</param>
    /// <param name="tarPath">文件目标路径</param>
    /// <param name="name">文件名</param>
    public static void 文件_移动(string dirPath, string tarPath, string name)
        {
            bool flag = false;
            foreach (string d in Directory.GetFileSystemEntries(dirPath))
            {
                if (File.Exists(dirPath + @"\" + name))
                {
                    flag = true;
                }
            }//end of for

            if (!flag)
            {
                Console.WriteLine("目标文件 " + name + " 不存在");
                return;
            }

            File.Move(dirPath + @"\" + name, tarPath + @"\" + name);
        }
        /// <summary>
        /// 复制文件
        /// </summary>
        /// <param name="dirPath">文件原始路径</param>
        /// <param name="tarPath">文件目标路径</param>
        /// <param name="name">文件名</param>
        public static void 文件_复制(string dirPath, string tarPath, string name)
        {
            bool flag = false;
            foreach (string d in Directory.GetFileSystemEntries(dirPath))
            {
                if (File.Exists(dirPath + @"\" + name))
                {
                    flag = true;
                }
            }//end of for

            if (!flag)
            {
                Console.WriteLine("目标文件 " + name + " 不存在");
                return;
            }

            File.Copy(dirPath + @"\" + name, tarPath + @"\" + name);

        }

        /// <summary>
        /// 文件_删除
        /// </summary>
        /// <param name="filepath">路径</param>
        public static void 文件_删除(string filepath)
        {
            DirectoryInfo dir = new DirectoryInfo(filepath);
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
            foreach (FileSystemInfo i in fileinfo)
            {
                if (i is DirectoryInfo)            //判断是否文件夹
                {
                    DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                    subdir.Delete(true);          //删除子目录和文件
                }
                else
                {
                    File.Delete(i.FullName);      //删除指定文件
                }
            }
        }



    }
}

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace DuteIT
{

    public static class Weibo_Class
    {
        public static Excel.Application App = Globals.ThisAddIn.Application;
        public static Excel.Workbook wbS = App.ActiveWorkbook;
        public static Excel.Worksheet ShK = wbS.ActiveSheet;
        public static Stopwatch stopwatch = new Stopwatch();


        /// <summary> 
        /// 从一条新浪微博地址中获取这条微博的id
        /// </summary>  
        /// <param name="url">一条微博地址</param>  
        /// <returns>微博id</returns>  
        public static string GetIdFromUrl(string url)
        {
            string mid = GetMidFromUrl(url);

            if (string.IsNullOrEmpty(mid))
            {
                return string.Empty;
            }
            else
            {
                return Mid2Id(mid);
            }
        }

        /// <summary> 
        /// 从一条新浪微博地址中获取这条微博的mid
        /// </summary>  
        /// <param name="url">一条微博地址</param>  
        /// <returns>微博mid</returns>  
        public static string GetMidFromUrl(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                return "";
            }
            if (url.IndexOf('?') != -1)
            {
                url = url.Substring(0, url.IndexOf('?'));
            }
            Regex reg = new Regex(@"^http://(e\.)?weibo\.com/[0-9a-zA-Z]+/(?<id>[0-9a-zA-Z]+)$", RegexOptions.IgnoreCase);
            Match match = reg.Match(url);
            if (match.Success)
            {
                return match.Result("${id}");
            }
            return "";
        }

        /// <summary>
        /// 将新浪微博mid转换成id
        /// </summary>
        /// <param name="mid">微博mid</param>
        /// <returns>返回微博的id</returns>
        public static string Mid2Id(string mid)
        {
            string id = "";
            for (int i = mid.Length - 4; i > -4; i = i - 4) //从最后往前以4字节为一组读取URL字符
            {
                int offset1 = i < 0 ? 0 : i;
                int len = i < 0 ? mid.Length % 4 : 4;
                var str = mid.Substring(offset1, len);

                str = Str62toInt(str);
                if (offset1 > 0) //若不是第一组，则不足7位补0
                {
                    while (str.Length < 7)
                    {
                        str = "0" + str;
                    }
                }
                id = str + id;
            }
            return id;
        }
        /// <summary>
        /// 新浪微博id转换为mid
        /// </summary>
        /// <param name="id">微博id</param>
        /// <returns>返回微博的mid</returns>
        public static string Id2Mid(string id)
        {
            string mid = "", strTemp;
            int startIdex, len;

            for (var i = id.Length - 7; i > -7; i = i - 7) //从最后往前以7字节为一组读取mid
            {
                startIdex = i < 0 ? 0 : i;
                len = i < 0 ? id.Length % 7 : 7;
                strTemp = id.Substring(startIdex, len);
                mid = IntToStr62(Convert.ToInt32(strTemp)) + mid;
            }
            return mid;
        }

        /// <summary>
        /// 微博爬虫
        /// </summary>
        /// <param name="mode"></param>
        public static void WeiBoCrawlers(string mode)
        {
            //获取SUB组合成Cookie共后续api请求使用
            string tid = HttpUtility.UrlEncode(JObject.Parse(textSub(PostData("https://passport.weibo.com/visitor/genvisitor", "gen_callback")))["data"]["tid"].ToString());
            JObject resString = JObject.Parse(textSub(Get("https://passport.weibo.com/visitor/visitor?a=incarnate&t=" + tid + "&w=3&c=100&cb=cross_domain&from=weibo")));
            string sub = "SUB=" + resString["data"]["sub"].ToString() + ";";
            string subp = "SUBP=" + resString["data"]["subp"].ToString() + ";";

            //执行api操作
            userApiData(sub, mode);
            
        }


        public static void userApiData(string cookie,string mode)
        {
            Excel.Range Rng = App.Selection;//获取当前选中单元格

            object[,] strArray = Rng.Value2;
            //获取配置设置
            int RowNum = 1;//每10条提现一次运行状态
            Random TispTime = new Random(); //用于生成随机数，让线程休眠
            
            int TotalSize = strArray.GetLength(0);
            int Size = TotalSize <= 2 ? TotalSize : TotalSize / 2;
            int arrSize = TotalSize % Size == 0 ? TotalSize / Size : TotalSize / Size + 1;

            Task.Run(() =>
            {
                List<Task> taskList = new List<Task>();
                int countkey = 1;
                for (int t = 0; t < arrSize; t++)
                {
                    int s = t;
                    //新建一个线程
                    taskList.Add(Task.Run(() =>
                    {
                        //打标签
                        List<string> listTag = new List<string>();
                        for (int j = (s * Size) + 1; j <= Size * (s + 1); j++)
                        {
                            if (j <= TotalSize)
                            {
                                if (countkey % RowNum == 0)
                                {
                                    Thread.Sleep(TispTime.Next(500, 2000));
                                    try { App.StatusBar = "正在处理第" + countkey.ToString() + "行内容..."; } catch (Exception) { }
                                }

                                for (int i = 1; i <= strArray.GetLength(1); i++)
                                {
                                    try
                                    {
                                        var dd = Get("https://weibo.com/ajax/profile/info?custom=" + strArray[j, i].ToString(), cookie);
                                        switch (mode)
                                        {
                                            case "userfollowers":
                                                strArray[j, i] = (int)JObject.Parse(dd)["data"]["user"]["followers_count"];
                                                break;
                                            case "userfriends":
                                                strArray[j, i] = (int)JObject.Parse(dd)["data"]["user"]["friends_count"];
                                                break;
                                            case "userstatuses":
                                                strArray[j, i] = (int)JObject.Parse(dd)["data"]["user"]["statuses_count"];
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        strArray[j, i] = ex.Message.ToString();
                                    }

                                }

                                countkey++;
                            }
                        }
                    }));
                }
                //等待所有线程结束
                Task.WaitAll(taskList.ToArray());

                wbS.Activate();
                ShK.Activate();
                Rng.Value2 = strArray;

                //停止计时
                stopwatch.Stop();
                App.StatusBar = "标签已处理完成！共计耗时：" + stopwatch.Elapsed.ToString();
                
            });
        }

        #region 进制转换

        //62进制转成10进制
        public static string Str62toInt(string str62)
            {
                Int64 i64 = 0;
                for (int i = 0; i < str62.Length; i++)
                {
                    Int64 Vi = (Int64)Math.Pow(62, (str62.Length - i - 1));
                    char t = str62[i];
                    i64 += Vi * GetInt10(t.ToString());
                }
                return i64.ToString();
            }
            //10进制转成62进制
            public static string IntToStr62(int int10)
            {
                string s62 = "";
                int r = 0;
                while (int10 != 0)
                {
                    r = int10 % 62;
                    s62 = Get62key(r) + s62;
                    int10 = int10 / 62;
                }
                return s62;
            }
            // 62进制字典
            private static string str62keys = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVTXYZ";
            //获取key对应的62进制整数  
            private static Int64 GetInt10(string key)
            {
                return str62keys.IndexOf(key);
            }
            //获取62进制整数对应的key
            private static string Get62key(int int10)
            {
                if (int10 < 0 || int10 > 61)
                    return "";
                return str62keys.Substring(int10, 1);
            }

        #endregion

        #region 其他类

        static string textSub(string str)
        {
            int startIndex = str.IndexOf("(") + 1;
            int endIndex = str.IndexOf(")") - startIndex;
            return str.Substring(startIndex, endIndex);
        }


        public static string GetData(string Url, string cookie)
        {
            string body = "";
            var clientHandler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
            };
            var client = new HttpClient(clientHandler);
            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri("https://weibo.com/ajax/profile/info?custom=3203137375"),
                Headers =
                {
                    { "accept", "text/html" },
                    { "user-agent", "Mozilla/5.0 (Windows NT 10; Win64; x64; rv:64.0) Gecko/20100101 Firefox/64.0" },
                    { "Cookie",cookie },
                },
            };
            try
            {

                HttpResponseMessage response = client.SendAsync(request).Result;
                body = response.Content.ReadAsStringAsync().Result.ToString();
            }
            catch (Exception ex)
            {

                return ex.ToString();
            }

            return body;
        }

        public static string PostData(string Url, string inputData)
        {
            string str = "";
            try
            {
                HttpClient client = new HttpClient();
                var postContent = new MultipartFormDataContent();
                postContent.Add(new StringContent(inputData), "cb");
                HttpResponseMessage response = client.PostAsync(Url, postContent).Result;
                str = response.Content.ReadAsStringAsync().Result.ToString();

            }
            catch (Exception ex)
            {

                return ex.ToString();
            }

            return str;
        }

        /// <summary>
        /// Post请求
        /// </summary>
        public static string Post(string Url, string postDataStr, string cookies = null)
        {
            byte[] byteData = Encoding.UTF8.GetBytes(postDataStr);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "POST";
            if (cookies != null)
                request.Headers.Add("Cookie", cookies);
            request.ContentType = "application/json;charset=UTF-8";
            request.UserAgent = RandomBrowserUa();
            request.Accept = "application/json, text/plain, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
            request.ContentLength = byteData.Length;
            //发送数据
            using (Stream resquestStream = request.GetRequestStream())
            {
                resquestStream.Write(byteData, 0, byteData.Length);
                resquestStream.Close();
            }
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string encoding = response.ContentEncoding;
            if (encoding == null || encoding.Length < 1)
            {
                encoding = "UTF-8"; //默认编码  
            }
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));
            string retString = reader.ReadToEnd();
            return retString;
        }

        /// <summary>
        /// Get请求
        /// </summary>
        public static string Get(string Url, string cookies = null)
        {
            string retString = "";
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                //GET请求
                request.Method = "GET";
                request.ReadWriteTimeout = 5000;
                request.ContentType = "application/json;charset=UTF-8";

                if (cookies != null)
                    request.Headers.Add("Cookie", cookies);
                request.UserAgent = RandomBrowserUa();
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//执行get请求
                Stream myResponseStream = response.GetResponseStream();
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));

                //返回内容JSON
                retString = myStreamReader.ReadToEnd();
            }
            catch (Exception ex)
            {

                return ex.ToString();
            }

            return retString;
        }

        /// <summary>
        /// cookie字符转CookieContainer
        /// </summary>
        /// <param name="cooikes">cookie字符串</param>
        /// <param name="domain">域名</param>
        /// <returns></returns>
        public static CookieContainer Cookiestr2CookieContainer(string cooikes, string domain)
        {
            CookieCollection collection = new CookieCollection();
            Regex _cookieRegex = new Regex("(\\S*?)=(.*?)(?:;|$)", RegexOptions.Compiled);  //正则表达式
            MatchCollection matchCollection = _cookieRegex.Matches(cooikes);    // 根据 Cookies 规则匹配键值  GetCookies (domain)取cookie字符串
            foreach (Match item in matchCollection)
            {
                //string temp = item.Value.Replace(";","");
                //temp = item.Value.Replace(" ", "");
                //string[] ckstr = temp.Split('=');
                //collection.Add(new Cookie(ckstr[0].ToString(), ckstr[1].ToString()));
                collection.Add(new Cookie(item.Groups[1].Value, item.Groups[2].Value, "", domain));   // 转换为 Cookie 对象
            }
            CookieContainer cc = new CookieContainer();
            cc.Add(collection);
            return cc;
        }

        /// <summary>
        /// 随机生成ua
        /// </summary>
        /// <returns></returns>
        public static string RandomBrowserUa()
        {

            string[] ua = new string[] {
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
                "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
                "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
                "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1",
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5",
                "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5",
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
                "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
                "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
                "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
                "Mozilla/5.0 (Macintosh; U; Mac OS X Mach-O; en-US; rv:2.0a) Gecko/20040614 Firefox/3.0.0 ",
                "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10.5; en-US; rv:1.9.0.3) Gecko/2008092414 Firefox/3.0.3",
                "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10.5; en-US; rv:1.9.1) Gecko/20090624 Firefox/3.5",
                "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10.6; en-US; rv:1.9.2.14) Gecko/20110218 AlexaToolbar/alxf-2.0 Firefox/3.6.14",
                "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10.5; en-US; rv:1.9.2.15) Gecko/20110303 Firefox/3.6.15",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
                "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
                "Opera/9.80 (Android 2.3.4; Linux; Opera mobi/adr-1107051709; U; zh-cn) Presto/2.8.149 Version/11.10",
                "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US) AppleWebKit/531.21.8 (KHTML, like Gecko) Version/4.0.4 Safari/531.21.10",
                "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US) AppleWebKit/533.17.8 (KHTML, like Gecko) Version/5.0.1 Safari/533.17.8",
                "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5",
                "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0",
                "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
                "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
                "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
            };
            Random rd = new Random();
            int index = rd.Next(0, ua.Length);
            return ua[index];
        }
        #endregion
    }
}

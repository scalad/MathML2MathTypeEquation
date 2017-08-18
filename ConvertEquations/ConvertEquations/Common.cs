using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ConvertEquations
{
    /// <summary>
    /// MathML处理
    /// </summary>
    class MathML
    {
        static string[] riginal = {"∆"};

        static string[] replace = { "&#x0394;"};
        //预处理mathml
        public static string preproccessMathml(string mathml)
        {
            int rilength = riginal.Length;
            int relength = replace.Length;
            for (int i = 0; i < rilength; i++)
            {
                mathml.Replace(riginal[i], replace[i]);
            }
            return mathml;
        }
    }

    /// <summary>
    /// Word操作工具类
    /// </summary>
    class WordUtils
    {
        public static object nothing = System.Reflection.Missing.Value;

        /// <summary>
        /// 向左移动moveCount个光标
        /// </summary>
        /// <param name="document"></param>
        /// <param name="moveCount"></param>
        public static void moveLeft(Word.Document document, int moveCount)
        {
            if (moveCount <= 0) return;
            object moveUnit = Microsoft.Office.Interop.Word.WdUnits.wdWord;
            object moveExtend = Microsoft.Office.Interop.Word.WdMovementType.wdExtend;
            document.Application.Selection.MoveLeft(ref moveUnit, moveCount, ref nothing);
        }
    }

    /// <summary>
    /// Global and static variables
    /// </summary>
    class Common
    {
        //Filter the src in the html img tag
        public static string HTML_IMG = @"<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*(?<imgUrl>[^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>";
    }

    /// <summary>
    /// HTML translator util
    /// </summary>
    class HTMLUtils
    {
        // 将HTML代码复制到Windows剪贴板，并保证中
        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();
        [DllImport("user32.dll")]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("user32.dll")]
        static extern bool CloseClipboard();
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint RegisterClipboardFormatA(string lpszFormat);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint GlobalSize(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalUnlock(IntPtr hMem);

        /// <summary>
        /// copy the html into clipboard
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        static public bool CopyHTMLToClipboard(string html)
        {
            uint CF_HTML = RegisterClipboardFormatA("HTML Format");
            bool bResult = false;
            if (OpenClipboard(IntPtr.Zero))
            {
                if (EmptyClipboard())
                {
                    byte[] bs = System.Text.Encoding.UTF8.GetBytes(html);

                    int size = Marshal.SizeOf(typeof(byte)) * bs.Length;

                    IntPtr ptr = Marshal.AllocHGlobal(size);
                    Marshal.Copy(bs, 0, ptr, bs.Length);

                    IntPtr hRes = SetClipboardData(CF_HTML, ptr);
                    CloseClipboard();
                }
            }
            return bResult;
        }

        //将HTML代码按照Windows剪贴板格进行格式化
        public static string HtmlClipboardData(string html)
        {
            StringBuilder sb = new StringBuilder();
            Encoding encoding = Encoding.UTF8; //Encoding.GetEncoding(936);
            string Header = @"Version: 1.0
                            StartHTML: {0:000000}
                            EndHTML: {1:000000}
                            StartFragment: {2:000000}
                            EndFragment: {3:000000}
                            ";
            string HtmlPrefix = @"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">
                            <html>
                            <head>
                            <meta http-equiv=Content-Type content=""text/html; charset={0}"">
                            </head>
                            <body>
                            <!--StartFragment-->
                            ";
            HtmlPrefix = string.Format(HtmlPrefix, "gb2312");

            string HtmlSuffix = @"<!--EndFragment--></body></html>";

            // Get lengths of chunks
            int HeaderLength = encoding.GetByteCount(Header);
            HeaderLength -= 16; // extra formatting characters {0:000000}
            int PrefixLength = encoding.GetByteCount(HtmlPrefix);
            int HtmlLength = encoding.GetByteCount(html);
            int SuffixLength = encoding.GetByteCount(HtmlSuffix);

            // Determine locations of chunks
            int StartHtml = HeaderLength;
            int StartFragment = StartHtml + PrefixLength;
            int EndFragment = StartFragment + HtmlLength;
            int EndHtml = EndFragment + SuffixLength;

            // Build the data
            sb.AppendFormat(Header, StartHtml, EndHtml, StartFragment, EndFragment);
            sb.Append(HtmlPrefix);
            sb.Append(html);
            sb.Append(HtmlSuffix);

            //Console.WriteLine(sb.ToString());
            return sb.ToString();
        }
    }

    /// <summary>
    /// System Code
    /// </summary>
    sealed class ResultCode
    {
        public static string SUCCESS = "转换成功";
        public static string WORD_ERROR = "Word错误";
        public static string WORD_SAVE_ERROR = "Word存储错误";
        public static string EXCEL_ERROR = "Excel错误";
        public static string EXCEL_READ_ERROR = "Excel读取错误";
        public static string FILE_PATH_ERROR = "文件位置错误";
    }

    /// <summary>
    /// System Utils
    /// </summary>
    class Utils
    {
        /// <summary>
        /// delete download file
        /// </summary>
        /// <param name="names">file full path collections</param>
        public static void deleteFile(List<string> names)
        {
            if (names == null || names.Count <= 0)
            {
                return;
            }
            else
            {
                foreach (string name in names)
                {
                    if (File.Exists(name))
                    {
                        FileInfo info = new FileInfo(name);
                        if (info.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            info.Attributes = FileAttributes.Normal;
                        File.Delete(name);
                    }
                }
            }
        }

        /// <summary>
        /// kill a process
        /// </summary>
        /// <param name="processName"></param>
        public static void killAllProcess(string processName)
        {
            System.Diagnostics.Process[] myPs;
            myPs = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process p in myPs)
            {
                if (p.Id != 0)
                {
                    try
                    {
                        if (p.Modules != null)
                            if (p.Modules.Count > 0)
                            {
                                System.Diagnostics.ProcessModule pm = p.Modules[0];
                                if (pm.ModuleName.ToLower() == processName)
                                    p.Kill();
                            }
                    }
                    catch
                    { }
                }
            }
        }

        //filter html tag
        public static string NoHTML(string Htmlstring)
        {
            //delete the script
            Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"–>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<!–.*", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);
            Htmlstring.Replace("<", "");
            Htmlstring.Replace(">", "");
            Htmlstring.Replace("\r\n", "");
            return Htmlstring;
        }

        public static string GetInputFolder(string strFile)
        {
            FileInfo fi = new FileInfo(Application.ExecutablePath);
            string strRet = System.IO.Path.Combine(fi.Directory.Parent.Parent.FullName, "Data");
            return System.IO.Path.Combine(strRet, strFile);
        }
    }

}

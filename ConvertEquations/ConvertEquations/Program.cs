using System;
using System.Data;
using System.IO;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MTSDKDN;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using System.Collections.Generic;
using EnvDTE;
using EnvDTE80;

namespace ConvertEquations
{
    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }

    public class MessageFilter : IOleMessageFilter
    {
        //
        // Class containing the IOleMessageFilter
        // thread error-handling functions.

        // Start the filter.
        public static void Register()
        {
            Console.WriteLine("Register method");
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, out oldFilter);
        }

        // Done with the filter, close it.
        public static void Revoke()
        {
            Console.WriteLine("revoke method...");
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        //
        // IOleMessageFilter functions.
        // Handle incoming thread requests.
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
          lpInterfaceInfo)
        {
            Console.WriteLine("HandleInComingCall method");
            return 0;
        }

        // Thread call was rejected, so try again.
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr
          hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {
                // Retry the thread call immediately if return >=0 & 
                // <100.
                return 99;
            }
            // Too busy; cancel call.
            return -1;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
          int dwTickCount, int dwPendingType)
        {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int
          CoRegisterMessageFilter(IOleMessageFilter newFilter, out 
          IOleMessageFilter oldFilter);
    }

    class Program
    {
        //用于作为函数的默认参数
        public static object nothing = System.Reflection.Missing.Value;

        //微软提供的可调用的API入口
        [DllImport("shell32.dll ")]
        public static extern int ShellExecute(IntPtr hwnd, String lpszOp, String lpszFile, String lpszParams, String lpszDir, int FsShowCmd);

        //主程序入口，必须以单线程方式启动
        [STAThread]
        static void Main(string[] args)
        {
            string ConString = System.Configuration.ConfigurationManager.AppSettings["filepath"];
            string filename = System.Configuration.ConfigurationManager.AppSettings["filename"];
            Program program = new Program();
            program.MathML2MathTypeWord(program, new ConvertEquation(), filename);
        }

        public string MathML2MathTypeWord(Program p, ConvertEquation ce, string filename)
        {
            Utils.killAllProcess("winword.exe");
            Utils.killAllProcess("mathtype.exe");
            Utils.killAllProcess("excel.exe");
            object name = "c:\\yb3.docx";

            //create document
            Word.Application newapp = new Word.Application();
            //create a word document
            Word.Document newdoc = newapp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            //是否显示word程序界面
            newapp.Visible = false;
          

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            
            string path = Utils.GetInputFolder(filename);
           
            Excel.Application excel = new Excel.Application();//lauch excel application  
            if (excel == null)
            {
                return ResultCode.EXCEL_READ_ERROR;
            }
            excel.Visible = false; 
            excel.UserControl = true;
            // 以只读的形式打开EXCEL文件  
            workbook = excel.Application.Workbooks.Open(path, nothing, true, nothing, nothing, nothing,
             nothing, nothing, nothing, true, nothing, nothing, nothing, nothing, nothing);
            //取得第一个工作薄
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            //取得总记录行数   (包括标题列)
            int iRowCount = worksheet.UsedRange.Rows.Count;
            int iColCount = worksheet.UsedRange.Columns.Count;

            //生成列头
            List<string> titles = new List<String>();
            for (int i = 0; i < iColCount; i++)
            {
                var txt = ((Excel.Range)worksheet.Cells[1, i + 1]).Text.ToString();
                titles.Add(txt.ToString()+ ": ");

            }
            //生成行数据
            Excel.Range range;
            //从第二行开始
            int rowIdx = 2;
            int count = 0;
            object anchor = null;
            Console.ReadLine();
            for (int iRow = rowIdx; iRow <= iRowCount; iRow++)
            {
                for (int iCol = 1; iCol <= iColCount; iCol++)
                {
                    //插入列头
                    newapp.Selection.Font.Color = Word.WdColor.wdColorBlue;
                    newapp.Selection.TypeText(titles[iCol - 1]);
                    //得到单元格内容
                    range = (Excel.Range)worksheet.Cells[iRow, iCol];
                    string d = range.Text.ToString();
                    string[] oneLevelData = d.Split(new string[] { "<math", "</math>" }, StringSplitOptions.None);

                    try
                    {
                        foreach (string datas in oneLevelData)
                        {
                            if (datas.StartsWith(" xmlns="))
                            {
                                // MML in a text file to clipboard text
                                ce.Convert(new EquationInputFileText("<math" + datas + "</math>", ClipboardFormats.cfMML), new EquationOutputClipboardText());
                                count++;
                                newapp.Selection.Paste();
                                Console.WriteLine("插入公式完成");
                                if (count == 9)
                                {
                                    Utils.killAllProcess("mathtype.exe");
                                    count = 0;
                                }
                            }
                            else
                            {
                                string[] tags = datas.Split(new string[] { "<img" }, StringSplitOptions.None);
                                foreach(string tag in tags)
                                {
                                    Console.WriteLine(tag);
                                    string matchString = Regex.Match("<img " + tag, "<img.+?src=[\"'](.+?)[\"'].*?>", RegexOptions.IgnoreCase).Groups[1].Value;
                                    if (matchString != null && !"".Equals(matchString))
                                    {
                                        object SaveWithDocument = true;
                                        anchor = newdoc.Application.Selection.Range;
                                        newapp.Selection.Move();
                                        //插入图片
                                        newdoc.Application.ActiveDocument.InlineShapes.AddPicture(matchString, ref nothing, ref SaveWithDocument, ref anchor);
                                        newapp.Selection.Move();
                                        Console.WriteLine("插入图片完成");
                                    }
                                    newapp.Selection.Font.Color = Word.WdColor.wdColorBlack;
                                    string text = Utils.NoHTML("<span><img " + tag + "</span>");
                                    if (text != null && !"".Equals(text))
                                    {
                                        //去除空格、插入文本b
                                        newapp.Selection.TypeText(text.Trim());
                                        newapp.Selection.Move();
                                        Console.WriteLine("插入文本完成 >>> " + text);
                                    }
                                    else
                                    {
                                        int location = text.Trim().IndexOf("</");
                                        Console.WriteLine(location);
                                        newapp.Selection.TypeText(tag.Trim());
                                        newapp.Selection.Move();
                                        Console.WriteLine("插入文本完成 >>> " + tag);
                                    }
                                }
                            }
                        }
                        newapp.Selection.TypeParagraph();
                    }
                    catch (Exception et)
                    {
                        Console.WriteLine(et);
                    }
                }
                newapp.Selection.TypeParagraph();
                //清空粘贴板，否则会将前一次粘贴记录保留。
                Clipboard.SetDataObject("", true);
            }

            try
            {
                newdoc.SaveAs(ref name, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing);
            }
            catch (Exception ex)
            {
                newdoc.Close(true, ref nothing, ref nothing);
                Console.WriteLine(ex);
            }
            excel.Quit();
            excel = null;
            newdoc = null;
            newapp = null;
            Console.WriteLine("Transaction finish");
            Console.ReadLine();
            return ResultCode.SUCCESS;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Drawing.Imaging;
using MTSDKDN;
using Word = Microsoft.Office.Interop.Word;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using System.Text.RegularExpressions;
using System.Linq.Expressions;

namespace ConvertEquations
{
	/// SDK
	#region MTSDK class
	class MTSDK
	{
		// c-tor
		public MTSDK() { }

		// vars
		protected bool m_bDidInit = false;

		// init
		public bool Init()
		{
			if (!m_bDidInit)
			{
				Int32 result = MathTypeSDK.Instance.MTAPIConnectMgn(MTApiStartValues.mtinitLAUNCH_AS_NEEDED, 30);
				if (result == MathTypeReturnValue.mtOK)
				{
					m_bDidInit = true;
					return true;
				}
				else
					return false;
			}
			return true;
		}

		// de-init
		public bool DeInit()
		{
			if (m_bDidInit)
			{
				m_bDidInit = false;
				MathTypeSDK.Instance.MTAPIDisconnectMgn();
			}
			return true;
		}

	}
	#endregion

	/// Output Equation Classes
	#region EquationOutput Classes
	abstract class EquationOutput
	{
		// c-tor
		public EquationOutput(string strOutTrans) 
		{
			if (!string.IsNullOrEmpty(strOutTrans))
				this.strOutTrans = strOutTrans;
			else
				this.strOutTrans = string.Empty;
		}

		protected EquationOutput() { }

		// properties
		protected short m_iType;
		public short iType 
		{ 
			get { return m_iType; }
			protected set { m_iType = value; }
		}

		protected short m_iFormat;
		public short iFormat
		{
			get { return m_iFormat; }
			protected set { m_iFormat = value; }
		}

		private string m_strFileName;
		public string strFileName
		{
			get { return m_strFileName; }
			set { m_strFileName = value; }
		}

		private string m_strEquation;
		public string strEquation
		{
			get { return m_strEquation; }
			set { m_strEquation = value; }
		}

		// output translator
		protected string m_strOutTrans;
		public string strOutTrans
		{
			get { return m_strOutTrans; }
			set { m_strOutTrans = value; }
		}

		// save equation to its destination
		abstract public bool Put();
	}

	abstract class EquationOutputClipboard : EquationOutput
	{
		public EquationOutputClipboard(string strOutTrans)
			: base(strOutTrans)
		{
			strFileName = string.Empty;
			iType = MTXFormEqn.mtxfmCLIPBOARD;
		}

		public EquationOutputClipboard()
			: base()
		{
			strFileName = string.Empty;
			iType = MTXFormEqn.mtxfmCLIPBOARD;
		}

		public override bool Put() { return true; }
	}

	class EquationOutputClipboardText : EquationOutputClipboard
	{
		public EquationOutputClipboardText(string strOutTrans)
			: base(strOutTrans)
		{
			iFormat = MTXFormEqn.mtxfmTEXT;
		}

		public EquationOutputClipboardText()
			: base()
		{
			iFormat = MTXFormEqn.mtxfmTEXT;
		}

		public override string ToString() { return "Clipboard Text";  }
	}
	
	#endregion

	/// Input Equation Classes
	#region EquationInput Classes
	abstract class EquationInput
	{
		// c-tor
		public EquationInput(string strInTrans) 
		{
			if (!string.IsNullOrEmpty(strInTrans))
				this.strInTrans = strInTrans;
			else
				this.strInTrans = string.Empty;
		}

		protected short m_iType;
		public short iType
		{
			get { return m_iType; }
			protected set { m_iType = value; }
		}

		protected short m_iFormat;
		public short iFormat
		{
			get { return m_iFormat; }
			protected set { m_iFormat = value; }
		}

		// the equation as a string
		protected string m_strEquation;
		public string strEquation
		{
			get { return m_strEquation; }
			set { m_strEquation = value; }
		}

		// the equation as a byte array
		protected byte[] m_bEquation;
		public byte[] bEquation
		{
			get { return m_bEquation; }
			set { m_bEquation = value; }
		}

		// MTEF byte array
		protected byte[] m_bMTEF;
		public byte[] bMTEF
		{
			get { return m_bMTEF; }
			set { m_bMTEF = value; }
		}

		// MTEF byte array length
		protected int m_iMTEF_Length;
		public int iMTEF_Length
		{
			get { return m_iMTEF_Length; }
			set { m_iMTEF_Length = value; }
		}

		// MTEF string
		protected string m_strMTEF;
		public string strMTEF
		{
			get { return m_strMTEF; }
			set { m_strMTEF = value; }
		}

		// input translator
		protected string m_strInTrans;
		public string strInTrans
		{
			get { return m_strInTrans; }
			set { m_strInTrans = value; }
		}

		// the source equation file
		protected string m_strFileName;
		public string strFileName
		{
			get { return m_strFileName; }
			set { m_strFileName = value; }
		}

		protected MTSDK sdk = new MTSDK();

		// get the equation from the source
		abstract public bool Get();

		// get binary MTEF
		abstract public bool GetMTEF();
	}

	abstract class EquationInputFile : EquationInput
	{
		public EquationInputFile(string strFileName, string strInTrans)
			: base(strInTrans)
		{
			this.strFileName = strFileName;
			iType = MTXFormEqn.mtxfmLOCAL;
		}
	}

	class EquationInputFileText : EquationInputFile
	{
		public EquationInputFileText(string strFileName, string strInTrans)
			: base(strFileName, strInTrans)
		{
			iFormat = MTXFormEqn.mtxfmMTEF;
		}

		public override string ToString() { return "Text file"; }

		override public bool Get()
		{
			try
			{
				strEquation = strFileName;
				return true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				return false;
			}
		}

		override public bool GetMTEF()
		{
			bool bReturn = false;

			if (!sdk.Init())
				return bReturn;

			IDataObject dataObject = MathTypeSDK.getIDataObject();

			if (dataObject == null)
			{
				sdk.DeInit();
				return bReturn;
			}

			FORMATETC formatEtc = new FORMATETC();
			STGMEDIUM stgMedium = new STGMEDIUM();

			try
			{
				// Setup the formatting information to use for the conversion.
				formatEtc.cfFormat = (Int16)DataFormats.GetFormat(strInTrans).Id;
				formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
				formatEtc.lindex = -1;
				formatEtc.ptd = (IntPtr)0;
				formatEtc.tymed = TYMED.TYMED_HGLOBAL;

				// Setup the MathML content to convert
				stgMedium.unionmember = Marshal.StringToHGlobalAuto(strEquation);
				stgMedium.tymed = TYMED.TYMED_HGLOBAL;
				stgMedium.pUnkForRelease = 0;

				// Perform the conversion
				dataObject.SetData(ref formatEtc, ref stgMedium, false);

				// Set the format for the output
				formatEtc.cfFormat = (Int16)DataFormats.GetFormat("MathType EF").Id;
				formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
				formatEtc.lindex = -1;
				formatEtc.ptd = (IntPtr)0;
				formatEtc.tymed = TYMED.TYMED_ISTORAGE;

				// Create a blank data structure to hold the converted result.
				stgMedium = new STGMEDIUM();
				stgMedium.tymed = TYMED.TYMED_NULL;
				stgMedium.pUnkForRelease = 0;

				// Get the conversion result in MTEF format
				dataObject.GetData(ref formatEtc, out stgMedium);
			}
			catch (COMException e)
			{
				Console.WriteLine("MathML conversion to MathType threw an exception: " + Environment.NewLine + e.ToString());
				sdk.DeInit();
				return bReturn;
			}

			// The pointer now becomes a Handle reference.
			HandleRef handleRef = new HandleRef(null, stgMedium.unionmember);

			try
			{
				// Lock in the handle to get the pointer to the data
				IntPtr ptrToHandle = MathTypeSDK.GlobalLock(handleRef);

				// Get the size of the memory block
				m_iMTEF_Length = MathTypeSDK.GlobalSize(handleRef);

				// New an array of bytes and Marshal the data across.
				m_bMTEF = new byte[m_iMTEF_Length];
				Marshal.Copy(ptrToHandle, m_bMTEF, 0, m_iMTEF_Length);
				m_strMTEF = System.Text.ASCIIEncoding.ASCII.GetString(m_bMTEF);
				bReturn = true;
			}
			catch (Exception e)
			{
				Console.WriteLine("Generation of image from MathType failed: " + Environment.NewLine + e.ToString());
			}
			finally
			{
				MathTypeSDK.GlobalUnlock(handleRef);
			}

			sdk.DeInit();
			return bReturn;
		}
	}

	#endregion

	/// ConvertEquation Classes
	class ConvertEquation
	{
		protected EquationInput m_ei;
		protected EquationOutput m_eo;
		protected MTSDK m_sdk = new MTSDK();

		// c-tor
		public ConvertEquation() { }

		// convert
		virtual public bool Convert(EquationInput ei, EquationOutput eo)
		{
			m_ei = ei;
			m_eo = eo;
			return Convert();
		}

		virtual protected bool Convert()
		{
			bool bReturn = false;

			Console.WriteLine("Converting {0} to {1}", m_ei.ToString(), m_eo.ToString());

			Console.WriteLine("Get equation: {0}", m_ei.strFileName);
			if (m_ei.Get())
			{
				Console.WriteLine("Get MTEF");
				if (m_ei.GetMTEF())
				{
					Console.WriteLine("Convert Equation");
					if (ConvertToOutput())
					{
						Console.WriteLine("Write equation: {0}", m_eo.strFileName);
						if (m_eo.Put())
							bReturn = true;
					}
				}
			}

			Console.WriteLine("Convert success: {0}\r\n", bReturn.ToString());
			return bReturn;
		}

		protected bool SetTranslator()
		{
			if (string.IsNullOrEmpty(m_eo.strOutTrans))
				return true;

			Int32 stat = MathTypeSDK.Instance.MTXFormSetTranslatorMgn(
				MTXFormSetTranslator.mtxfmTRANSL_INC_NAME + MTXFormSetTranslator.mtxfmTRANSL_INC_DATA, 
				m_eo.strOutTrans);
			return stat == MathTypeReturnValue.mtOK;
		}

		protected bool ConvertToOutput()
		{
			bool bResult = false;
			try
			{
				if (!m_sdk.Init())
					return false;

				if (MathTypeSDK.Instance.MTXFormResetMgn() == MathTypeReturnValue.mtOK && 
					SetTranslator())
				{
					Int32 stat = 0;
					Int32 iBufferLength = 5000;
					StringBuilder strDest = new StringBuilder(iBufferLength);
					MTAPI_DIMS dims = new MTAPI_DIMS();

					// convert
					stat = MathTypeSDK.Instance.MTXFormEqnMgn(
						m_ei.iType,
						m_ei.iFormat,
						m_ei.bMTEF,
						m_ei.iMTEF_Length,
						m_eo.iType,
						m_eo.iFormat,
						strDest,
						iBufferLength,
						m_eo.strFileName,
						ref dims);

					// save equation
					if (stat == MathTypeReturnValue.mtOK)
					{
						m_eo.strEquation = strDest.ToString();
						bResult = true;
					}
				}

				m_sdk.DeInit();
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
			return bResult;
		}
	}

	/// Program Class
	#region Program Class
	class Program
	{
		public string GetInputFolder(string strFile)
		{
			FileInfo fi = new FileInfo(Application.ExecutablePath);
			string strRet = System.IO.Path.Combine(fi.Directory.Parent.Parent.FullName, "Data");
			return System.IO.Path.Combine(strRet, strFile);
		}

		protected int iFileNum = 0;
		public string GetOutputFile(string strExt)
		{
			string strRet = Path.GetTempPath();
			string strFileName;
			strFileName = string.Format("Output{0}.{1}", iFileNum++, strExt);
			return System.IO.Path.Combine(strRet, strFileName);
		}

		public void MessagePause(string strMessage)
		{
			Console.WriteLine(strMessage);
			//Console.ReadKey(true);
		}

        [DllImport("shell32.dll ")]
        public static extern int ShellExecute(IntPtr hwnd, String lpszOp, String lpszFile, String lpszParams, String lpszDir, int FsShowCmd); 

		public void LocalToClipboard(Program p, ConvertEquation ce)
		{
            killAllProcess("winword.exe");
            killAllProcess("mathtype.exe");
            object name = "e:\\yb3.doc";
            string path = p.GetInputFolder("MathML8.txt");
            string d = System.IO.File.ReadAllText(path);

            //create document
            Word.Application newapp = new Word.Application();
            //用于作为函数的默认参数
            object nothing = System.Reflection.Missing.Value;
            //create a word document
            Word.Document newdoc = newapp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            //是否显示word程序界面
            newapp.Visible = false;

            string[] oneLevelData = d.Split(new string[] { "<math", "</math>" },StringSplitOptions.None);

            foreach (string datas in oneLevelData)
            {
                Word.Paragraph data = newdoc.Content.Paragraphs.Add(nothing);

                if (datas.StartsWith(" xmlns="))
                {
                    // MML in a text file to clipboard text
                    ce.Convert(new EquationInputFileText("<math" + datas + "</math>", ClipboardFormats.cfMML), new EquationOutputClipboardText());
                    try
                    {
                        Object objUnit = Word.WdUnits.wdStory;
                        newapp.Selection.EndKey(ref objUnit);
                        newdoc.Select();
                        data.Range.Paste();
                        //移动焦点并换行  
                        object count = 14;
                        newapp.Selection.MoveDown(ref nothing, ref count, ref nothing);//移动焦点
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                    finally
                    {
                        killAllProcess("mathtype.exe");
                    }
                }
                else
                {
                    Console.WriteLine(datas);
                    Regex regex = new Regex(@"<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*(?<imgUrl>[^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>", RegexOptions.IgnoreCase);
                    MatchCollection matches = regex.Matches(datas.ToString());
                    int i = 0;
                    string[] sUrlList = new string[matches.Count];
                    foreach (Match match in matches)
                    {
                        sUrlList[i++] = match.Groups["imgUrl"].Value;
                        Console.WriteLine(sUrlList[0]);
                        object LinkToFile = false;  
                        object SaveWithDocument = true;  
                        object Anchor = newdoc.Application.Selection.Range;
                        newdoc.Application.ActiveDocument.InlineShapes.AddPicture(sUrlList[0], ref LinkToFile, ref SaveWithDocument, ref Anchor);
                    }

                    string para = NoHTML(datas);
                    if (para != null && para != "")
                    {
                        newapp.Selection.TypeText(para);
                    }
                }
                //保存文档
                newdoc.SaveAs(ref name, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing);
                //清空粘贴板，否则会将前一次粘贴记录保留。
                Clipboard.SetDataObject("", true);
            }
            
            //关闭文档
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            newdoc.Close(ref nothing, ref nothing, ref nothing);
            newapp.Application.Quit(ref saveOption, ref nothing, ref nothing);
            newdoc = null;
            newapp = null;
            ShellExecute(IntPtr.Zero, "open", name.ToString(), "", "", 3);
            p.MessagePause("Inspect the clipboard, then press any key");
		}

        // 杀掉所有winword.exe进程
        protected void killAllProcess(string processName) 
        {
            System.Diagnostics.Process[] myPs;
            myPs = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process p in myPs)
            {
                if (p.Id != 0)
                {
                    string myS = "WINWORD.EXE" + p.ProcessName + "  ID:" + p.Id.ToString();
                    try
                    {
                        if (p.Modules != null)
                            if (p.Modules.Count > 0)
                            {
                                System.Diagnostics.ProcessModule pm = p.Modules[0];
                                myS += "\n Modules[0].FileName:" + pm.FileName;
                                myS += "\n Modules[0].ModuleName:" + pm.ModuleName;
                                myS += "\n Modules[0].FileVersionInfo:\n" + pm.FileVersionInfo.ToString();
                                if (pm.ModuleName.ToLower() == processName)
                                    p.Kill();
                            }
                    }
                    catch
                    { }
                }
            }
        }

        public static string NoHTML(string Htmlstring)
        {
            //删除脚本
            Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "",
            RegexOptions.IgnoreCase);
            //删除HTML 
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"–>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<!–.*", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">",
            RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ",
            RegexOptions.IgnoreCase);
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

        [STAThread] 
		static void Main(string[] args)
		{
			Program p = new Program();
			ConvertEquation ce = new ConvertEquation();

			p.LocalToClipboard(p, ce);
        
            Console.ReadLine();
		}
	}
	#endregion
}

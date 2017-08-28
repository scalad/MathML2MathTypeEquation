using System;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Drawing.Imaging;
using MTSDKDN;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace ConvertEquations
{
    /// SDK
    #region MTSDK class
    class MTSDK
    {
        // c-tor
        public MTSDK() { }

        // vars
        protected static bool m_bDidInit = false;

        // init
        public bool Init()
        {
            try
            {
                if (!m_bDidInit)
                {
                    Int32 result = MathTypeSDK.Instance.MTAPIConnectMgn(MTApiStartValues.mtinitLAUNCH_AS_NEEDED, 30);
                    if (result == MathTypeReturnValue.mtOK)
                    {
                        m_bDidInit = true;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                m_bDidInit = false;
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

        public override string ToString() { return "Clipboard Text"; }
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
        public EquationInputFile(string text, string strInTrans)
            : base(strInTrans)
        {
            this.strFileName = text;
            iType = MTXFormEqn.mtxfmLOCAL;
        }
    }

    class EquationInputFileText : EquationInputFile
    {
        public EquationInputFileText(string text, string strInTrans)
            : base(text, strInTrans)
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

            if (m_ei.Get())
            {
                if (m_ei.GetMTEF())
                {
                    if (ConvertToOutput())
                    {
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
                if (MathTypeSDK.Instance.MTXFormResetMgn() == MathTypeReturnValue.mtOK && SetTranslator())
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
}

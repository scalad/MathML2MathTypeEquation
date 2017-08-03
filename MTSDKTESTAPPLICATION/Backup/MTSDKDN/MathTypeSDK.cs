// MathTypeSDK.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTSDKDN/MTSDKDN/MathTypeSDK.cs 7     4/07/10 11:00a Jimm $

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace MTSDKDN
{
    #region MathType SDK Types
    [StructLayout(LayoutKind.Sequential)]
    public struct MTAPI_PICT
    {
        public long _mm;
        public long _xExt;
        public long _yExt;
        public long _hMF;

        public MTAPI_PICT(long mm, long xExt, long yExt, long hMF)
        {
            _mm = mm;
            _xExt = xExt;
            _yExt = yExt;
            _hMF = hMF;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public long _left;
        public long _top;
        public long _right;
        public long _bottom;

        public RECT(long left, long top, long right, long bottom)
        {
            _left = left;
            _top = top;
            _right = right;
            _bottom = bottom;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MTAPI_DIMS
    {
        public int _baseline;
        public RECT _bounds;

        public MTAPI_DIMS(int baseline, ref RECT bounds)
        {
            _baseline = baseline;

            _bounds._left = bounds._left;
            _bounds._top = bounds._top;
            _bounds._right = bounds._right;
            _bounds._bottom = bounds._bottom;
        }
    }

    #endregion MathType SDK Types

    #region MathType Constants
	// Error codes returned by most API's
    public sealed class MathTypeReturnValue
    {
        public const short mtOK = 0;					// no error
        public const short mtNOT_FOUND = -1;			// trouble finding the MT erver, usually indicates a bad session ID
        public const short mtCANT_RUN = -2;				// could not start the MT server
        public const short mtBAD_VERSION = -3;			// server / DLL version mismatch
        public const short mtIN_USE = -4;				// server is busy/in-use
        public const short mtNOT_RUNNING = -5;			// server aborted
        public const short mtRUN_TIMEOUT = -6;			// connection to the server failed due to time-out 
        public const short mtNOT_EQUATION = -7;			// an API call that expects an equation could not find one
        public const short mtFILE_NOT_FOUND = -8;		// a preference, translator or other file could not be found
        public const short mtMEMORY = -9;				// a buffer too small to hold the result of an API call was passed in
		public const short mtBAD_FILE = -10;			// file found was not a translator
        public const short mtDATA_NOT_FOUND = -11;		// unable to read preferences from MTEF on the clipboard
        public const short mtTOO_MANY_SESSIONS = -12;	// too many open connections to the SDK
        public const short mtSUBSTITUTION_ERROR = -13;	// problem with substition error during a call to MTXFormEqn
        public const short mtTRANSLATOR_ERROR = -14;	// there was an error in compiling or in execution of a translator 
        public const short mtPREFERENCE_ERROR = -15;	// could not set preferences
        public const short mtBAD_PATH = -16;			// a bad path was encountered when trying to write to a file
        public const short mtFILE_ACCESS = -17;			// a file could not be written to
        public const short mtFILE_WRITE_ERROR = -18;	// a file could not be written to
		public const short mtBAD_DATA = -19;			// (deprecated)
        public const short mtERROR = -9999;				// other error
    }

    public sealed class MTXFormSetTranslator
    {
        // options values for MTXFormSetTranslator
		public const short mtxfmTRANSL_INC_NONE = 0;		// include neither the translator name nor equation data in output
		public const short mtxfmTRANSL_INC_NAME = 1;		// include the translator's name in translator output
		public const short mtxfmTRANSL_INC_DATA = 2;		// include MathType equation data in translator output
        public const short mtxfmTRANSL_INC_MTDEFAULT = 4;	// use defaults for translator name and equation data

        // append 'Z' to text equations placed on clipboard
        // kludge to fix Word's trailing CRLF truncation bug
        public const short mtxfmTRANSL_INC_CLIPBOARD_EXTRA = 8;
    }

    public sealed class MTTranslatorInfo
    {
        // options values for MTGetTranslatorsInfo
        public const short mttrnCOUNT = 1;		// get the total number of translators
		public const short mttrnMAX_NAME = 2;	// maximum size of any translator name
		public const short mttrnMAX_DESC = 3;	// maximum size of any translator description string
		public const short mttrnMAX_FILE = 4;	// maximum size of any translator file name
		public const short mttrnOPTIONS = 5;	// get translation options
    }

    public sealed class MTXTranslatorPreference
    {
        // options values for MTXFormSetTranslator
        public const short mtxfmPREF_EXISTING = 1;	// use existing preferences
		public const short mtxfmPREF_MTDEFAULT = 2;	// use MathType's default preferences
		public const short mtxfmPREF_USER = 3;		// use specified preferences
        public const short mtxfmPREF_LAST = 3;		// (internal use)
    }

    public sealed class MTXFormStatus
    {
        // return values from MTXFormGetStatus
		public const short mtxfmSTAT_ACTUAL_LEN = -1;	// number of bytes of data actually returned in dstData (if MTXformEqn succeeded) or
														// the number of bytes required (if MTXformEqn retuned mtMEMORY - not enough memory 
														// was specified for dstData), otherwise 0L.
		public const short mtxfmSTAT_TRANSL = -2;		// status for translation:
														//   mtOK - successful translation
														//   mtFILE_NOT_FOUND - could not find translator
														//   mtBAD_FILE - file found was not a translator
		public const short mtxfmSTAT_PREF = -3;			// status for set preferences:
														//   mtOK - sucess setting preferences 
														//   mtSUBSTITUTION_ERROR - bad preference data
    }

    public sealed class MTXFormEqn
    {
		// see comment for MTXFormEqnMgn below, for a description of the following constants

        // data sources/destinations for MTXFormEqn
		public const short mtxfmPREVIOUS = -1;
        public const short mtxfmCLIPBOARD = -2;
        public const short mtxfmLOCAL = -3;
        public const short mtxfmFILE = -4;

        // data formats for MTXFormEqn
        public const short mtxfmMTEF = 4;
        public const short mtxfmHMTEF = 5;
        public const short mtxfmPICT = 6;
        public const short mtxfmTEXT = 7;
        public const short mtxfmHTEXT = 8;
        public const short mtxfmGIF = 9;
        public const short mtxfmEPS_NONE = 10;
        public const short mtxfmEPS_WMF = 11;
        public const short mtxfmEPS_TIFF = 12;
    }

    public sealed class MTPreferences
    {
        // option values for MTSetMTPrefs
		public const short mtprfMODE_NEXT_EQN = 1;	// apply to next new equation
		public const short mtprfMODE_MTDEFAULT = 2;	// set MathType's defaults for new equations
		public const short mtprfMODE_INLINE = 4;	// makes next eqn inline
    }

    public sealed class MTXFormSubValues
    {
        // option values for MTXFormAddVarSub
        public const short mtxfmSUBST_ALL = 0;	// substitute all variables
        public const short mtxfmSUBST_ONE = 1;	// substitute one variable

        // find/replace types for MTXFormAddVarSub substitutions
        public const short mtxfmVAR_SUB_BAD = -1;
        public const short mtxfmVAR_SUB_PLAIN_TEXT = 0;
        public const short mtxfmVAR_SUB_MTEF_TEXT = 1;
        public const short mtxfmVAR_SUB_MTEF_BINARY = 2;
        public const short mtxfmVAR_SUB_DELETE = 3;
        public const short mtxfmVAR_SUB_MAX = 4;

        // replace styles for MTXFormAddVarSub substitutions where type = mtxfmVAR_SUB_PLAIN_TEXT
        public const short mtxfmSTYLE_FIRST = 1;
        public const short mtxfmSTYLE_TEXT = 1;
        public const short mtxfmSTYLE_FUNCTION = 2;
        public const short mtxfmSTYLE_VARIABLE = 3;
        public const short mtxfmSTYLE_LCGREEK = 4;
        public const short mtxfmSTYLE_UCGREEK = 5;
        public const short mtxfmSTYLE_SYMBOL = 6;
        public const short mtxfmSTYLE_VECTOR = 7;
        public const short mtxfmSTYLE_NUMBER = 8;
        public const short mtxfmSTYLE_LAST = 8;
    }

    public sealed class MTApiStartValues
    {
		public const short mtinitLAUNCH_AS_NEEDED = 0;	// launch MathType when first needed
		public const short mtinitLAUNCH_NOW = 1;		// launch MathType server immediately
    }

    public sealed class MTClipboardEqnTypes
    {
        public const short mtOLE_EQUATION = 1;		// equation OLE 1.0 object on clipboard
        public const short mtWMF_EQUATION = 2;		// Windows metafile equation graphic (not OLE object) on clipboard
        public const short mtMAC_PICT_EQUATION = 4; // Macintosh PICT equation graphic (not OLE object) on clipboard
        public const short mtOLE2_EQUATION = 8;		// equation OLE 2.0 object on clipboard
    }

    public sealed class MTURLTypes
    {
        // option values for MTGetURL
		public const short mturlMATHTYPE_HOME = 1;		// for MathType Homepage (e.g. http:www.dessci.com/en/products/mathtype)
		public const short mturlMATHTYPE_SUPPORT = 2;	// for Online Support (e.g. http:www.dessci.com/en/support/mathtype)
		public const short mturlMATHTYPE_FEEDBACK = 3;	// for Feedback Email To (e.g. mailto:support@dessci.com)
		public const short mturlMATHTYPE_ORDER = 4;		// for Order MathType (e.g. http://www.dessci.com/en/store/mathtype/options.asp?prodId=MT)
		public const short mturlMATHTYPE_FUTURE = 5;	// for Future MathType (to be determined)
        public const short mturlMATHTYPE_REGISTER = 6;	// for registering MathType
    }

    public sealed class MTDimensionValues
    {
        // options value for MTGetLastDimension
		// use the following to get a dimension from the last equation copied to the clipboard or written to a file
		public const short mtdimFIRST = 1;			// (internal use)
        public const short mtdimWIDTH = 1;			//
        public const short mtdimHEIGHT = 2;			//
        public const short mtdimBASELINE = 3;		//
        public const short mtdimHORIZ_POS_TYPE = 4;	//
        public const short mtdimHORIZ_POS = 5;		//
        public const short mtdimLast = 5;			// (internal use)
    }

	// clipboard formats
	public sealed class ClipboardFormats
	{
		public const string cfNative = "Native";
		public const string cfOwnerLink = "OwnerLink";
		public const string cfRTF = "Rich Text Format";
		public const string cfMTEF = "MathType EF";
		public const string cfMacPICT = "Mac PICT";
		public const string cfEmbedSrc = "Embed Source";
		public const string cfObjDesc = "Object Descriptor";
		public const string cfEmbeddedObj = "Embedded Object";
		public const string cfMTMacro = "MathType Macro";
		public const string cfHTML = "HTML Format";
		public const string cfMTMacroPict = "MathType Macro PICT";
		public const string cfMMLPres = "MathML Presentation";
		public const string cfMML = "MathML";
		public const string cfMMLXML = "application/mathml+xml";
		public const string cfTeX = "TeX Input Language";
	}

    #endregion MathType Constants

    public class MathTypeSDK
    {
		#region Singleton class
		private static readonly MathTypeSDK instance = new MathTypeSDK();
		private MathTypeSDK() { }
		public static MathTypeSDK Instance
		{
			get { return instance; }
		}
		#endregion

        #region MathType SDK Function Declarations

        [DllImport("MT6.dll", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTAPIConnect(short mtStart, short timeout);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTAPIVersion(ushort apiversion);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTAPIDisconnect();

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTEquationOnClipboard();

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTClearClipboard();

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetLastDimension(short dimIndex);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTOpenFileDialog(short fileType, [MarshalAs(UnmanagedType.LPStr)] string title, string dir, [MarshalAs(UnmanagedType.LPStr)] StringBuilder file, short fileLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetPrefsFromClipboard([MarshalAs(UnmanagedType.LPStr)] StringBuilder prefs, short prefsLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetPrefsFromFile([MarshalAs(UnmanagedType.LPStr)] string prefFile,
            [MarshalAs(UnmanagedType.LPStr)] StringBuilder prefs, short prefsLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTConvertPrefsToUIForm([MarshalAs(UnmanagedType.LPStr)] string inPrefs,
            [MarshalAs(UnmanagedType.LPStr)] StringBuilder outPrefs, short outPrefsLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetPrefsMTDefault([MarshalAs(UnmanagedType.LPStr)] StringBuilder prefs, short prefsLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTSetMTPrefs(short mode, [MarshalAs(UnmanagedType.LPStr)] string prefs, short timeout);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetTranslatorsInfo(short infoIndex);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTEnumTranslators(short index, [MarshalAs(UnmanagedType.LPStr)] StringBuilder transName,
            short transNameLen, [MarshalAs(UnmanagedType.LPStr)] StringBuilder transDesc, short transDescLen,
            [MarshalAs(UnmanagedType.LPStr)] StringBuilder transFile, short transFileLen);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTXFormReset();

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTXFormAddVarSub(short options, short findType, [MarshalAs(UnmanagedType.LPStr)] string find,
            int findLen, short replaceType, [MarshalAs(UnmanagedType.LPStr)] string replace, int replaceLen, short replaceStyle);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTXFormSetTranslator(ushort options, [MarshalAs(UnmanagedType.LPStr)] string transName);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTXFormSetPrefs(short prefType, [MarshalAs(UnmanagedType.LPStr)] string prefStr);

		[DllImport("MT6.dll", EntryPoint = "MTXFormEqn", ExactSpelling = false)]
        private static extern int MTXFormEqn(
			short src, 
			short srcFmt, 
			byte[] srcData, 
			int srcDataLen, 
			short dst,
			short dstFmt, 
			System.Text.StringBuilder dstData, 
			int dstDataLen, 
			[MarshalAs(UnmanagedType.LPStr)] string dstPath, 
			ref MTAPI_DIMS dims);

		[DllImport("MT6.dll", EntryPoint = "MTXFormEqn", ExactSpelling = false)]
		private static extern int MTXFormEqn(
			short src,
			short srcFmt,
			byte[] srcData,
			int srcDataLen,
			short dst,
			short dstFmt,
			IntPtr dstData,
			int dstDataLen,
			[MarshalAs(UnmanagedType.LPStr)] string dstPath,
			ref MTAPI_DIMS dims);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTXFormGetStatus(short index);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTPreviewDialog(IntPtr parent, [MarshalAs(UnmanagedType.LPStr)] string title,
            [MarshalAs(UnmanagedType.LPStr)] string prefs, [MarshalAs(UnmanagedType.LPStr)] string closeBtnText,
            [MarshalAs(UnmanagedType.LPStr)] string helpBtnText, int helpID, [MarshalAs(UnmanagedType.LPStr)] string helpFile);

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTShowAboutBox();

        [DllImport("MT6.DLL", CharSet = CharSet.Auto, PreserveSig = true, ExactSpelling = true, SetLastError = true)]
        private static extern int MTGetURL(int whichURL, bool bGoToURL, [MarshalAs(UnmanagedType.LPStr)] StringBuilder strURL, int sizeURL);

        #endregion MathType SDK Functions

		#region Kernel32 function calls
		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		public static extern IntPtr GlobalLock(HandleRef handle);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		public static extern bool GlobalUnlock(HandleRef handle);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		public static extern int GlobalSize(HandleRef handle);

		#endregion Kernel32 function calls

        #region MathType SDK Function Wrappers
        /// <summary>
        ///
        /// </summary>
        /// <param name="mtStart">mtinitLAUNCH_NOW => launch MathType server immediately
        ///	                      mtinitLAUNCH_AS_NEEDED => launch MathType when first needed</param>
        /// <param name="timeout"># of seconds to wait before timing out when attempting to
        ///                       launch MathType. If timeOut = -1 then will never timeout.
        ///                         This value is eventually passed to RPCConnectToServer
        ///                         where it is not currently used</param>
        /// <returns></returns>
        public int MTAPIConnectMgn(short mtStart, short timeout)
        {
            int retCode = MathTypeReturnValue.mtOK;

            retCode = MTAPIConnect(mtStart, timeout);
            return retCode;
        }

        /// <summary>
        /// Which version of the API is set.
        /// </summary>
        /// <param name="apiversion">Set the version of the API</param>
        /// <returns>0 if API set unknown or hi-byte (of lo-word) = major version,
        ///	                                 lo-byte (of lo-word) = minor version</returns>
        public int MTAPIVersionMgn(ushort apiversion)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTAPIVersion(apiversion);
            return retCode;
        }

        /// <summary>
        /// Disconnect the current instance of the running API server.
        /// </summary>
        /// <returns></returns>
        public int MTAPIDisconnectMgn()
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTAPIDisconnect();
            return retCode;
        }

        /// <summary>
        /// Check for the type of equation on the clipboard, if any
        /// </summary>
        /// <returns>If equation on the clipboard, returns type of eqn data:
        ///	            mtOLE_EQUATION, mtWMF_EQUATION, mtMAC_PICT_EQUATION,
        ///             Otherwise status value
        ///	             mtNOT_EQUATION - no eqn on clipboard
        ///              mtMEMORY - insufficient memory for prog ID
        ///              mtERROR - any other error</returns>
        public int MTEquationOnClipboardMgn()
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTEquationOnClipboard();
            return retCode;
        }

        /// <summary>
        /// clear the clipboard contents
        /// </summary>
        /// <returns>mtOK - everything was successful</returns>
        public int MTClearClipboardMgn()
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTClearClipboard();
            return retCode;
        }

        /// <summary>
        /// Get a dimension from the last equation copied to the clipboard or written
        ///		to a file.
        /// </summary>
        /// <param name="dimIndex">desired dimension (mtdimXXXX)</param>
        /// <returns>If successful (>0), value of desired dimension in 32nds of a point
        ///             Otherwise, error status
        ///		        mtNOT_EQUATION - no equation to take dimension from
        ///		        mtERROR - bad value for whichDim</returns>
        public int MTGetLastDimensionMgn(short dimIndex)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetLastDimension(dimIndex);
            return retCode;
        }

        /// <summary>
        /// Put up an open file dialog	(Win32 only)
        ///   Calls GetForegroundWindow for parent, upon which it gets centered
        /// </summary>
        /// <param name="fileType"> 1 for MT preference files</param>
        /// <param name="title">dialog window title</param>
        /// <param name="dir">default directory (may be empty or NULL)</param>
        /// <param name="file">result: new filename</param>
        /// <param name="fileLen">maximum number of characters in filename</param>
        /// <returns>Returns 1 for OK, 0 for Cancel</returns>
        public int MTOpenFileDialogMgn(short fileType, string title, string dir, StringBuilder file, short fileLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTOpenFileDialog(fileType, title, dir, file, fileLen);
            return retCode;
        }

        /// <summary>
        /// Get equation preferences from the MathType equation that is currently
        ///			on the clipboard
        /// </summary>
        /// <param name="prefs">[out] Preference string (if sizeStr > 0)</param>
        /// <param name="prefsLen">[in]  Size of prefStr (inc. null) or 0</param>
        /// <returns>If sizeStr == 0 then this is the size required for prefStr,
        ///	                 Otherwise it's a status
        ///	                    mtOK	        Success
        ///	                    mtMEMORY		Not enough memory for to store preferences
        ///	                    mtNOT_EQUATION	Not equation on clipboard
        ///	                    mtBAD_VERSION	No preference data found in equation
        ///	                    mtERROR		    Other error</returns>
        public int MTGetPrefsFromClipboardMgn(StringBuilder prefs, short prefsLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetPrefsFromClipboard(prefs, prefsLen);
            return retCode;
        }

        /// <summary>
        /// Get equation preferences from the specified preferences file
        /// </summary>
        /// <param name="prefFile">[in]  Pathname for the preference file</param>
        /// <param name="prefs">[out] Preference string (if sizeStr > 0)</param>
        /// <param name="prefsLen">[in]  Size of prefStr or 0 </param>
        /// <returns>If sizeStr == 0 then this is the size required for prefStr,
        ///	                Otherwise it's a status
        ///	                mtOK	            Success
        ///	                mtMEMORY	        Not enough memory for to store preferences
        ///	                mtFile_NOT_FOUND	File does not exist or bad pathname
        ///	                mtERROR				Other error</returns>
        public int MTGetPrefsFromFileMgn(string prefFile, StringBuilder prefs, short prefsLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetPrefsFromFile(prefFile, prefs, prefsLen);
            return retCode;
        }

        /// <summary>
        /// Convert internal preferences string to a form to be presented to the user
        /// </summary>
        /// <param name="inPrefs">[in]  internal preferences string</param>
        /// <param name="outPrefs">[out] Preference string (if sizeStr > 0)</param>
        /// <param name="outPrefsLen">[in]  Size of outPrefStr (inc. null) or 0 to get length</param>
        /// <returns>If outPrefsLen == 0 then this is the size required for outPrefStr, else it's a status
        ///	                mtOK		Success
        ///	                mtMEMORY	Not enough memory for to store preferences
        ///	                mtERROR		Other error</returns>
        public int MTConvertPrefsToUIFormMgn(string inPrefs, StringBuilder outPrefs, short outPrefsLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTConvertPrefsToUIForm(inPrefs, outPrefs, outPrefsLen);
            return retCode;
        }

        /// <summary>
        /// Get MathType's current default equation preferences
        /// </summary>
        /// <param name="prefs">[out] Preference string (if sizeStr > 0)</param>
        /// <param name="prefsLen">[in]  Size of prefStr or 0</param>
        /// <returns>If sizeStr == 0 then this is the size required for prefStr,
        ///	                Otherwise it's a status
        ///	                mtOK		Success
        ///	                mtMEMORY	Not enough memory for to store preferences
        ///	                mtERROR		Other error</returns>
        public int MTGetPrefsMTDefaultMgn(StringBuilder prefs, short prefsLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetPrefsMTDefault(prefs, prefsLen);
            return retCode;
        }

        /// <summary>
        /// Set MathType's default peferences for new equations
        /// </summary>
        /// <param name="mode">[in] Specifies the way the preferences will be applied
        ///                         mtprfMODE_NEXT_EQN => Apply to next new equation (see timeOut)
        ///	                        mtprfMODE_MTDEFAULT => Set MathType's defaults for new equations
        ///	                        mtprfMODE_INLINE => makes next eqn inline</param>
        /// <param name="prefs">[in] Null terminated preference string</param>
        /// <param name="timeout">[in] Number of seconds to wait for new equation (used only
        ///			                    when mode = 1), Note: -1 means wait forever</param>
        /// <returns>mtOK			Success,
        ///          mtBAD_DATA	    Bad pref string,
        ///          mtERROR		Any other error</returns>
        public int MTSetMTPrefsMgn(short mode, string prefs, short timeout)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTSetMTPrefs(mode, prefs, timeout);
            return retCode;
        }

        /// <summary>
        /// Get information about the current set of translators
        /// </summary>
        /// <param name="infoIndex">[In] A flag indicating what info to return:
        ///	                        1 => Total number of translators
        ///	                        2 => Maximum size of any translator name
        ///	                        3 => Maximum size of any translator description string
        ///	                        4 => Maximum size of any translator file name </param>
        /// <returns>If >= 0 then this value is the information specified by infoID
        ///	            Otherwise its a status
        ///	            mtERROR		Bad value for infoID</returns>
        public int MTGetTranslatorsInfoMgn(short infoIndex)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetTranslatorsInfo(infoIndex);
            return retCode;
        }

        /// <summary>
        /// Enumerate the available equation (TeX, etc.) translators
        /// </summary>
        /// <param name="index">[in] Index of the translator to enumerate
        ///			            (Must be initialized to 1 by the caller)</param>
        /// <param name="transName">[out] Translator name</param>
        /// <param name="transNameLen">[in]  Size of tShort. (May be set to zero)</param>
        /// <param name="transDesc">[out] Translator descriptor string</param>
        /// <param name="transDescLen">[in]  Size of transDesc. (May be set to zero)</param>
        /// <param name="transFile">[out] Translator file name</param>
        /// <param name="transFileLen">[in]  Size of transFile. (May be set to zero)</param>
        /// <returns>If >0 then this value is the index of next translator to enumerate,
        ///		        (i.e. the caller should pass this value in for indx - def below)
        ///	            Otherwise, a status
        ///		            mtOK		Success (no more translators in the list)
        ///		            mtMEMORY Not enough room in transName, transDesc, or transFile
        ///		            mtERROR  Any other failure</returns>
        public int MTEnumTranslatorsMgn(short index, StringBuilder transName, short transNameLen, StringBuilder transDesc,
            short transDescLen, StringBuilder transFile, short transFileLen)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTEnumTranslators(index, transName, transNameLen, transDesc, transDescLen, transFile, transFileLen);
            return retCode;
        }

        /// <summary>
        /// Resets to default options for MTXformEqn (i.e. no substitutions, no
        ///		translation, and use existing preferences)
        /// </summary>
        /// <returns>Only returns mtOK</returns>
        public int MTXFormResetMgn()
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormReset();
            return retCode;
        }

        /// <summary>
        /// Specify a variable substitution to be performed with next MTXformEqn
        ///		(may be called 0 or more times).
        /// </summary>
        /// <param name="options">mtxfmSUBST_ALL or mtxfmSUBST_ONE</param>
        /// <param name="findType">type of data in find arg (must be mtxfmVAR_SUB_PLAIN_TEXT for now)</param>
        /// <param name="find">equation text to be found and replaced (null-terminated text string for now)</param>
        /// <param name="findLen">length of find arg data (ignored for now)</param>
        /// <param name="replaceType">type of data in replace arg (mtxfmVAR_SUB_XXX; mtxfmVAR_SUB_DELETE to delete find)</param>
        /// <param name="replace">equation text to replace find arg</param>
        /// <param name="replaceLen">iff replaceType = mtxfmVAR_SUB_MTEF_BINARY, length of replace arg data</param>
        /// <param name="replaceStyle">if replaceType = mtxfmVAR_SUB_PLAIN_TEXT, style (fnXXXX)</param>
        /// <returns>mtOK - success
        ///          mtERROR - some other error</returns>
        public int MTXFormAddVarSubMgn(short options, short findType, string find, int findLen, short replaceType, string replace, int replaceLen, short replaceStyle)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormAddVarSub(options, findType, find, findLen, replaceType, replace, replaceLen, replaceStyle);
            return retCode;
        }

        /// <summary>
        /// Specify translation to be performed with the next MTXformEqn.
        /// </summary>
        /// <param name="options">[in] One or more (OR'd together) of:
        ///                         mtxfmTRANS_INC_NAME include the translator's name in
        ///		                    translator output
        ///                         mtxfmTRANS_INC_EQN include MathType equation data in
        ///		                    translator output</param>
        /// <param name="transName">[in] File name of translator to be used,
        ///			                    NULL for no translation</param>
        /// <returns>mtOK - success
        ///          mtFILE_NOT_FOUND - could not find translator
        ///          mtTRANSLATOR_ERROR - errors compiling translator
        ///          mtERROR - some other error</returns>
        public int MTXFormSetTranslatorMgn(ushort options, string transName)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormSetTranslator(options, transName);
            return retCode;
        }

        /// <summary>
        /// Specify a new set of preferences to be used with the next MTXformEqn.
        /// </summary>
        /// <param name="prefType">[in] One of the following,
        ///                         mtxfmPREF_EXISTING - use existing preferences
        ///                         mtxfmPREF_MTDEFAULT - use MathType's default preferences
        ///                         mtxfmPREF_USER - use specified preferences</param>
        /// <param name="prefStr">[in] Preferences to apply (mtxfmPREF_USER)</param>
        /// <returns>mtOK if the preference is set, else mtERROR</returns>
        public int MTXFormSetPrefsMgn(short prefType, string prefStr)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormSetPrefs(prefType, prefStr);
            return retCode;
        }

        /// <summary>
        /// Transform an equation (uses options specified via MTXAddVarSubst,
        ///		MTXSetTrans, and MTXSetPrefs)
        /// Note: Variations involving mtxform_SRC_PICT, mtxform_DST_PICT, or
        ///		mtxform_DST_HMTEF are not callable via Word Basic.
        /// </summary>
        /// <param name="src">[in] Equation data source, either
        ///                     mtxfmPREVIOUS => data from previous result
        ///                     mtxfmCLIPBOARD => data on clipboard
        ///                     mtxfmLOCAL => data passed (i.e. in srcData)</param>
        /// <param name="srcFmt">[in] Equation source data format (mtxfmXXX, see next)
        ///	                    Note: srcFmt, srcData, and srcDataLen are used only
        ///		                if src is mtfxmLOCAL</param>
        /// <param name="srcData">[in] Depends on data source (src)
        ///                         mtxfmMTEF => ptr to MTEF-binary (BYTE *)
        ///                         mtxfmPICT => ptr to pict (MTAPI_PICT *)
        ///                         mtxfmTEXT => ptr to text (CHAR *), either MTEF-text or plain text</param>
        /// <param name="srcDataLen">[in] # of bytes in srcData</param>
        /// <param name="dst">[in] Equation data destination, either
        ///                     mtxfmCLIPBOARD => transformed data placed on clipboard
        ///                     mtxfmLOCAL => transformed data in dstData
        ///                     mtxfmFILE => transformed data in the file specified by dstPath</param>
        /// <param name="dstFmt">[in] Equation data format (mtxfmXXX, see next)
        ///	                        Note: dstFmt, dstData, and dstDataLen are used only
        ///		                    if dst is mtfxmLOCAL (data placed on the clipboard
        ///		                    is either an OLE object or translator text)</param>
        /// <param name="dstData">[out] Depends on data destination (dstFmt)
        ///                         mtxfmMTEF => ptr to MTEF-binary (BYTE *)
        ///                         mtxfmHMTEF => ptr to handle to MTEF-binary (HANDLE *)
        ///                         mtxfmPICT => ptr to pict data (MTAPI_PICT *)
        ///                         mtxfmTEXT => ptr to translated text or, if no translator, MTEF-text (CHAR *)
        ///                         mtxfmHTEXT => ptr to handle to translated text or, if no translator, MTEF-text (HANDLE *)
        ///                         Note: If translator specified dst must be either
        ///		                        mtxfmTEXT or mtxfmHTEXT for the translation to be performed[out] Depends on data destination (dstFmt)
        ///                         mtxfmMTEF => ptr to MTEF-binary (BYTE *)
        ///                         mtxfmHMTEF => ptr to handle to MTEF-binary (HANDLE *)
        ///                         mtxfmPICT => ptr to pict data (MTAPI_PICT *)
        ///                         mtxfmTEXT => ptr to translated text or, if no translator, MTEF-text (CHAR *)
        ///                         mtxfmHTEXT => ptr to handle to translated text or, if no translator, MTEF-text (HANDLE *)
        ///                         Note: If translator specified dst must be either
        ///		                        mtxfmTEXT or mtxfmHTEXT for the translation to be performed</param>
        /// <param name="dstDataLen">[in] # of bytes in dstData (used for mtxfmLOCAL only)</param>
        /// <param name="dstPath">[in] destination pathname (used if dst == mtxfmFILE only, may be NULL if not used)</param>
        /// <param name="dims">[out] pict dimensions, may be NULL (valid only for
        ///		                    dst = mtxfmPICT)</param>
        /// <returns>mtOK - success
        ///          mtNOT_EQUATION - source data does not contain MTEF
        ///          mtSUBSTITUTION_ERROR - could not perform one or more subs
        ///          mtTRANSLATOR_ERROR - errors occured during translation
        ///									(translation not done)
        ///          mtPREFERENCE_ERROR - could not set perferences
        ///          mtMEMORY - not enough space in dstData
        ///          mtERROR - some other error </returns>
        public int MTXFormEqnMgn(
			short src, 
			short srcFmt, 
			byte[] srcData, 
			int srcDataLen, 
			short dst, 
			short dstFmt,
			System.Text.StringBuilder dstData, 
			int dstDataLen, 
			string dstPath, 
			ref MTAPI_DIMS dims)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormEqn(src, srcFmt, srcData, srcDataLen, dst, dstFmt, dstData, dstDataLen, dstPath, ref dims);
            return retCode;
        }

		public int MTXFormEqnMgn(
			short src,
			short srcFmt,
			byte[] srcData,
			int srcDataLen,
			short dst,
			short dstFmt,
			IntPtr dstData,
			int dstDataLen,
			string dstPath,
			ref MTAPI_DIMS dims)
		{
			int retCode = MathTypeReturnValue.mtOK;
			retCode = MTXFormEqn(src, srcFmt, srcData, srcDataLen, dst, dstFmt, dstData, dstDataLen, dstPath, ref dims);
			return retCode;
		}

        /// <summary>
        /// Check error/status after XformEqn
        /// </summary>
        /// <param name="index">[in] which status to get; described above</param>
        /// <returns>Depends on the value of 'which', as follows:
        ///         which = mtxfmSTAT_PREF, status for set preferences
        ///             mtOK - sucess setting preferences
        ///             mtBAD_DATA - bad preference data
        ///         which = mtxfmSTAT_TRANSL, status for translation
        ///             mtOK - successful translation
        ///             mtFILE_NOT_FOUND - could not find translator
        ///             mtBAD_FILE - file found was not a translator
        ///         which = mtxfmSTAT_ACTUAL_LEN, number of bytes of data
        ///             actually returned in dstData (if MTXformEqn succeeded) or
        ///             the number of bytes required (if MTXformEqn retuned
        ///             mtMEMORY - not enough memory was specified for dstData),
        ///	            otherwise 0L.
        ///         which >= 1, status of the i-th (i = which)
        ///             variable substitution, either # of times the
        ///             substitution was performed, or, if < 0, an error status
        ///
        /// NOTE: returns mtERROR for bad values of which</returns>
        public int MTXFormGetStatusMgn(Int16 index)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTXFormGetStatus(index);
            return retCode;
        }

        /// <summary>
        /// Puts up a preview dialog for displaying preferences
        /// </summary>
        /// <param name="parent">parent window</param>
        /// <param name="title">dialog title</param>
        /// <param name="prefs">text to preview</param>
        /// <param name="closeBtnText">text for Close button (can be NULL for English)</param>
        /// <param name="helpBtnText">text for Help button  (can be NULL for English)</param>
        /// <param name="helpID">Help topic ID</param>
        /// <param name="helpFile">help file</param>
        /// <returns>returns 0 if successful, non-zero if error</returns>
        public int MTPreviewDialogMgn(IntPtr parent, string title, string prefs, string closeBtnText, string helpBtnText, int helpID, string helpFile)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTPreviewDialog(parent, title, prefs, closeBtnText, helpBtnText, helpID, helpFile);
            return retCode;
        }

        /// <summary>
        /// Shows the about box for the current version of MathType
        /// </summary>
        /// <returns>Always returns mtOK</returns>
        public int MTShowAboutBoxMgn()
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTShowAboutBox();
            return retCode;
        }

        /// <summary>
        /// Displays the requested URL.
        /// </summary>
        /// <param name="whichURL">[in] One of --
        ///                              mturlMATHTYPE_HOME
        ///	                              for MathType Homepage (e.g. http:www.dessci.com/en/products/mathtype)
        ///                             mturlMATHTYPE_SUPPORT
        ///                                for Online Support (e.g. http:www.dessci.com/en/support/mathtype)
        ///                              mturlMATHTYPE_FEEDBACK
        ///                                for Feedback Email To (e.g. mailto:support@dessci.com)
        ///                              mturlMATHTYPE_ORDER
        ///                                for Order MathType (e.g. http://www.dessci.com/en/store/mathtype/options.asp?prodId=MT)
        ///                              mturlMATHTYPE_FUTURE
        ///                                for Future MathType (to be determined)</param>
        /// <param name="bGoToURL">[in] True if browser should be launched</param>
        /// <param name="strURL">[out] URL String (if sizeURL > 0)</param>
        /// <param name="sizeURL">[in] Size of strURL or 0</param>
        /// <returns>If >0 and sizeURL == 0 then this is the size required for strURL,
        ///	            Otherwise it's a status
        ///	                mtOK			Success
        ///                 mtMEMORY	Not enough memory to store the URL (in strURL)
        ///                 mtERROR		Could not find the URL</returns>
        public int MTGetURLMgn(int whichURL, bool bGoToURL, StringBuilder strURL, int sizeURL)
        {
            int retCode = MathTypeReturnValue.mtOK;
            retCode = MTGetURL(whichURL, bGoToURL, strURL, sizeURL);
            return retCode;
        }
        #endregion MathType Regular SDK Function calls

        #region IDataObject

        #region class OLECLOSE
        public sealed class OleClose
        {
            public const int OLECLOSE_SAVEIFDIRTY = 0;
            public const int OLECLOSE_NOSAVE = 1;
            public const int OLECLOSE_PROMPTSAVE = 2;
        }
        #endregion

        #region class tagSIZEL
        [StructLayout(LayoutKind.Sequential)]
        public sealed class tagSIZEL
        {
            public int cx;
            public int cy;
            public tagSIZEL() { }
            public tagSIZEL(int cx, int cy) { this.cx = cx; this.cy = cy; }
            public tagSIZEL(tagSIZEL o) { this.cx = o.cx; this.cy = o.cy; }
        }
        #endregion

        #region class COMRECT
        [ComVisible(true), StructLayout(LayoutKind.Sequential)]
        public class COMRECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;

            public COMRECT() { }

            public COMRECT(int left, int top, int right, int bottom)
            {
                this.left = left;
                this.top = top;
                this.right = right;
                this.bottom = bottom;
            }

            public static COMRECT FromXYWH(int x, int y, int width, int height)
            {
                return new COMRECT(x, y, x + width, y + height);
            }
        }
        #endregion

        #region class tagOLEVERB
        [StructLayout(LayoutKind.Sequential)]
        public sealed class tagOLEVERB
        {
            public int lVerb;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszVerbName;

            [MarshalAs(UnmanagedType.U4)]
            public int fuFlags;

            [MarshalAs(UnmanagedType.U4)]
            public int grfAttribs;

            public tagOLEVERB() { }
        }
        #endregion

        #region class FORMATETC
        public sealed class FORMATETC
        {
            [MarshalAs(UnmanagedType.I4)]
            public int cfFormat;

            public IntPtr ptd;

            [MarshalAs(UnmanagedType.I4)]
            public int dwAspect;

            [MarshalAs(UnmanagedType.I4)]
            public int lindex;

            [MarshalAs(UnmanagedType.I4)]
            public int tymed;
        }
        #endregion

        #region class STGMEDIUM
        [ComVisible(false), StructLayout(LayoutKind.Sequential)]
        public class STGMEDIUM
        {

            [MarshalAs(UnmanagedType.I4)]
            public int tymed;

            public IntPtr unionmember;

            public IntPtr pUnkForRelease;

        }
        #endregion

		#region class Ole32Methods
		public class Ole32Methods
		{
			[DllImport("ole32.Dll")]
			static public extern uint CoCreateInstance(ref Guid clsid,
			   [MarshalAs(UnmanagedType.IUnknown)] object inner,
			   uint context,
			   ref Guid uuid,
			   [MarshalAs(UnmanagedType.IUnknown)] out object rReturnedComObject);
		}
		#endregion

		#region interface IEnumOLEVERB
		[ComImport, Guid("00000104-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumOLEVERB
        {
            [PreserveSig]
            int Next([MarshalAs(UnmanagedType.U4)] int celt, [Out] tagOLEVERB rgelt, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pceltFetched);

            [PreserveSig]
            int Skip([In, MarshalAs(UnmanagedType.U4)] int celt);

            void Reset();

            void Clone(out IEnumOLEVERB ppenum);
        }
        #endregion

        #region interface IOleClientSite
        [ComImport, Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleClientSite
        {
            void SaveObject();
            void GetMoniker(uint dwAssign, uint dwWhichMoniker, ref object ppmk);
            void GetContainer(ref object ppContainer);
            void ShowObject();
            void OnShowWindow(bool fShow);
            void RequestNewObjectLayout();
        }
        #endregion

        #region interface IEnumFORMATETC
        [ComVisible(true), ComImport(), Guid("00000103-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumFORMATETC
        {

            [return: MarshalAs(UnmanagedType.I4)]
            [PreserveSig]
            int Next([In, MarshalAs(UnmanagedType.U4)] int celt, [Out] FORMATETC rgelt,
                [In, Out, MarshalAs(UnmanagedType.LPArray)] int[] pceltFetched);

            [return: MarshalAs(UnmanagedType.I4)]
            [PreserveSig]
            int Skip([In, MarshalAs(UnmanagedType.U4)]int celt);

            [return: MarshalAs(UnmanagedType.I4)]
            [PreserveSig]
            int Reset();

            [return: MarshalAs(UnmanagedType.I4)]
            [PreserveSig]
            int Clone([Out, MarshalAs(UnmanagedType.LPArray)] IEnumFORMATETC[] ppenum);
        }
        #endregion

        #region interface IDataObject
        [ComVisible(true), ComImport(), Guid("0000010E-0000-0000-C000-000000000046"),
            InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDataObject
        {
            int GetData(FORMATETC pFormatetc, [Out] STGMEDIUM pMedium);

            int GetDataHere(FORMATETC pFormatetc, [In, Out] STGMEDIUM pMedium);

            int QueryGetData(FORMATETC pFormatetc);

            int GetCanonicalFormatEtc(FORMATETC pformatectIn, [Out] FORMATETC pformatetcOut);

            int SetData(FORMATETC pFormatectIn, STGMEDIUM pmedium, int fRelease);

            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumFORMATETC EnumFormatEtc([In, MarshalAs(UnmanagedType.U4)] int dwDirection);

            int DAdvise(FORMATETC pFormatetc, [In, MarshalAs(UnmanagedType.U4)] int advf,
                [In, MarshalAs(UnmanagedType.Interface)] object pAdvSink,
                [Out, MarshalAs(UnmanagedType.LPArray)] int[] pdwConnection);

            int DUnadvise([In, MarshalAs(UnmanagedType.U4)] int dwConnection);

            int EnumDAdvise([Out, MarshalAs(UnmanagedType.LPArray)] object[] ppenumAdvise);
        }
        #endregion

        #region interface IAdviseSink
        [ComVisible(true), Guid("0000010f-0000-0000-C000-000000000046"),
            InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IAdviseSink
        {
            void OnDataChange([In] object pFormatetc, [In] object pStgmed);
            void OnViewChange([In, MarshalAs(UnmanagedType.U4)] int dwAspect,[In, MarshalAs(UnmanagedType.I4)] int lindex);
			void OnRename([In, MarshalAs(UnmanagedType.Interface)] UCOMIMoniker pmk);
            void OnSave();
            void OnClose();
        }
        #endregion

        #region interface IOleObject
        [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleObject
        {
            [PreserveSig]
            int SetClientSite([In, MarshalAs(UnmanagedType.Interface)] IOleClientSite pClientSite);

            IOleClientSite GetClientSite();

            [PreserveSig]
            int SetHostNames([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);

            [PreserveSig]
            int Close(int dwSaveOption);

            [PreserveSig]
            int SetMoniker([In, MarshalAs(UnmanagedType.U4)] int dwWhichMoniker, [In, MarshalAs(UnmanagedType.Interface)] object pmk);

            [PreserveSig]
            int GetMoniker([In, MarshalAs(UnmanagedType.U4)] int dwAssign, [In, MarshalAs(UnmanagedType.U4)] int dwWhichMoniker, [MarshalAs(UnmanagedType.Interface)] out object moniker);

            [PreserveSig]
            int InitFromData([In, MarshalAs(UnmanagedType.Interface)] IDataObject pDataObject, int fCreation, [In, MarshalAs(UnmanagedType.U4)] int dwReserved);

            [PreserveSig]
            int GetClipboardData([In, MarshalAs(UnmanagedType.U4)] int dwReserved, out IDataObject data);

            [PreserveSig]
            int DoVerb(int iVerb, [In] IntPtr lpmsg, [In, MarshalAs(UnmanagedType.Interface)] IOleClientSite pActiveSite, int lindex, IntPtr hwndParent, [In] COMRECT lprcPosRect);

            [PreserveSig]
            int EnumVerbs(out IEnumOLEVERB e);

            [PreserveSig]
            int OleUpdate();

            [PreserveSig]
            int IsUpToDate();

            [PreserveSig]
            int GetUserClassID([In, Out] ref Guid pClsid);

            [PreserveSig]
            int GetUserType([In, MarshalAs(UnmanagedType.U4)] int dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string userType);

            [PreserveSig]
            int SetExtent([In, MarshalAs(UnmanagedType.U4)] int dwDrawAspect, [In] tagSIZEL pSizel);

            [PreserveSig]
            int GetExtent([In, MarshalAs(UnmanagedType.U4)] int dwDrawAspect, [Out] tagSIZEL pSizel);

            [PreserveSig]
            int Advise(IAdviseSink pAdvSink, out int cookie);

            [PreserveSig]
            int Unadvise([In, MarshalAs(UnmanagedType.U4)] int dwConnection);

            //[PreserveSig]
            //int EnumAdvise(out IEnumSTATDATA e);

            //[PreserveSig]
            //int GetMiscStatus([In, MarshalAs(UnmanagedType.U4)] int dwAspect, out int misc);

            //[PreserveSig]
            //int SetColorScheme([In] Win32.tagLOGPALETTE pLogpal);
        }
        #endregion

		#region getIDataObject
		public static System.Runtime.InteropServices.ComTypes.IDataObject getIDataObject()
		{
			const uint CLSCTX_INPROC_SERVER = 4;

			//
			//  CLSID of the COM object
			//  FIXME: need the actual GUID to use here.
			//
			Guid clsid = new Guid("0002CE03-0000-0000-C000-000000000046");

			//
			//  GUID of the required interface
			//  TODO: Verify that this is the correct GUID
			//
			Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

			object instance = null;
			uint hResult = Ole32Methods.CoCreateInstance(ref clsid, null, CLSCTX_INPROC_SERVER, ref IID_IUnknown, out instance);

			if (hResult != 0)
				return null;

			return instance as System.Runtime.InteropServices.ComTypes.IDataObject;
		}
		#endregion

		#endregion
	}

}

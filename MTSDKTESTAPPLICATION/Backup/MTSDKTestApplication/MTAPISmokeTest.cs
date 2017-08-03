// MTAPISmokeTest.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTSDKTESTAPPLICATION/MTSDKTestApplication/MTAPISmokeTest.cs 4     3/30/10 8:39a Jimm $

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

#region MTSDK using statements
using MTSDKDN;
using mtReturnValues = MTSDKDN.MathTypeReturnValue;
using mtapiStart = MTSDKDN.MTApiStartValues;
using mtCBOLE = MTSDKDN.MTClipboardEqnTypes;
using mtDimensions = MTSDKDN.MTDimensionValues;
using mtPreferences = MTSDKDN.MTPreferences;
using mtTranslatorInfo = MTSDKDN.MTTranslatorInfo;
using mtxfrmSubValues = MTSDKDN.MTXFormSubValues;
using mtxFrmSetTrans = MTSDKDN.MTXFormSetTranslator;
using mtxfrmPreferences = MTSDKDN.MTXTranslatorPreference;
using mtxFormEqn = MTSDKDN.MTXFormEqn;
using mtxFormStatus = MTSDKDN.MTXFormStatus;
using mtURLTypes = MTSDKDN.MTURLTypes;

#endregion MTSDK using statements

namespace MTSDKTestApplication
{
    /// <summary>
    /// The purpose of this Form application is to provide a very simple
    /// means of conducting a smoke test of the MT SDK.
    ///
    /// The buttons show which of the API calls are the focus of that particular
    /// button push.  It was an attempt to demostrate what the defaults might be,
    /// feel free to change the calls to the individual API calls.  In making these
    /// changes please reference the MT SDK API documentation.
    /// </summary>
    public partial class MTAPISmokeTest : Form
    {
        private AppHelper _appHelper;
        private MathTypeSDK _MTSDK;
        private bool _isMTAPIConnected;
        private string _prefsFile;
        public MTAPISmokeTest()
        {
            _appHelper = new AppHelper();
            //_MTSDK = new MathTypeSDK();
			_MTSDK = MathTypeSDK.Instance;

            _isMTAPIConnected = false;

            InitializeComponent();
        }

        #region Internal Helpers

        /// <summary>
        /// A helper method to attach to the MT SDK.
        /// </summary>
        /// <param name="silentStart">True to display error message boxes, else false.</param>
        private void InternalStartAPI(bool silentStart)
        {
            int retVal = 0;
            retVal = _MTSDK.MTAPIConnectMgn(MTApiStartValues.mtinitLAUNCH_NOW, 100);

            string error;

            if (silentStart == false)
            {
                if (retVal != mtReturnValues.mtOK)
                {
                    error = "Could not start API";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "API started successfully";
                    _appHelper.ShowOKDialog(error);
                }
            }

            if (retVal == mtReturnValues.mtOK)
                _isMTAPIConnected = true;
        }

        /// <summary>
        /// Internal helper to get a preferences file.  This method is used
        /// by many different mouse click handlers in this form.
        /// </summary>
        /// <param name="prefs">Out parameter containing the full path to the selected
        /// preferences file.</param>
        /// <param name="silent">True to display return value message boxes.</param>
        private void GetPrefsFile(out string prefs, bool silent)
        {
            long retVal = 0;
            string error;

            string dialogName = "Test Dialog";
            StringBuilder outFileName = new StringBuilder("", 256);
            retVal = _MTSDK.MTOpenFileDialogMgn(1, dialogName, null, outFileName, 100);

            if (silent == false)
            {
                if (retVal == 1)
                {
                    error = "File dialog successfully opened, name: " + outFileName;
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "File dialog did not successfully open";
                    _appHelper.ShowErrorDialog(error);
                }
            }

            char[] outstring = new char[outFileName.Length];
            outFileName.CopyTo(0, outstring, 0, outFileName.Length);

            prefs = new string(outstring);
        }

        #endregion Internal Helpers

        /// <summary>
        /// Make the connection to the MT SDK API.  This will start the SDK server.
        /// </summary>
        /// <param name="sender">Paramater not used</param>
        /// <param name="e">Parameter not used</param>
        private void StartAPI_Click(object sender, EventArgs e)
        {
            InternalStartAPI(false);
        }

        /// <summary>
        /// Disconnect from the api, this will only disconnect if
        /// there is a previous connection.
        /// </summary>
        /// <param name="sender">Paramter not used</param>
        /// <param name="e">Parameter not used</param>
        private void EndPI_Click(object sender, EventArgs e)
        {
            int retVal = 0;

            retVal = _MTSDK.MTAPIDisconnectMgn();
            string error;

            if (retVal != mtReturnValues.mtOK)
            {
                error = "Could not disconnect from API";
                _appHelper.ShowErrorDialog(error);
            }
            else
            {
                _isMTAPIConnected = false;
                error = "API disconnected successfully";
                _appHelper.ShowOKDialog(error);
            }

        }

        /// <summary>
        /// Get the current version of the API.  Only version 5 of the API
        /// is supported.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void APIVersion_Click(object sender, EventArgs e)
        {
            int version = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                version = _MTSDK.MTAPIVersionMgn(5);
                if (version <= 0)
                {
                    error = "There is not a valid version";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {

                    error = "The current API verision is: " + version.ToString();
                    _appHelper.ShowOKDialog(error);
                }
            }
        }

        /// <summary>
        /// Determines if there is an equation currently on the clipboard.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void EqnOnClipboard_Click(object sender, EventArgs e)
        {
            int retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTEquationOnClipboardMgn();
                if (retVal == mtCBOLE.mtMAC_PICT_EQUATION ||
                    retVal == mtCBOLE.mtOLE_EQUATION ||
                    retVal == mtCBOLE.mtWMF_EQUATION )
                {
                    error = "There is an equation present on the clipboard";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {

                    error = "There is no equation present on the clipboard";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Clears all content from the clipboard.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void ClearCB_Click(object sender, EventArgs e)
        {
            int retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTClearClipboardMgn();

                if (retVal == mtReturnValues.mtOK)
                {
                    error = "The clipboard was cleared";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "The clipboard was not cleared";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Gets the requested dimension from the last equation copied to the
        /// clipboard or written to a file.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void GetLastDim_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTGetLastDimensionMgn(mtDimensions.mtdimWIDTH);

                if (retVal > 0)
                {
                    error = "Here is the last width dimension: " + retVal.ToString();
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "There was an error in retreiving the last dimension";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Puts up an open file dialog.  Calls GetForegroundParent (Win32)
        /// for a parent, upon which the resulting dialog is centered.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void MTOpenFile_Click(object sender, EventArgs e)
        {
            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                GetPrefsFile(out _prefsFile, false);
            }
        }

        /// <summary>
        /// Gets equation preferences from the MathType equation that is
        /// currently on the clipboard.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void GetCBPrefs_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                StringBuilder outPrefs = new StringBuilder("", 1024);
                retVal = _MTSDK.MTGetPrefsFromClipboardMgn(outPrefs, 1024);
                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Prefs about the clipboard were returned: " + outPrefs;
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtMEMORY)
                {
                    error = "Error: Not enough memory";
                    _appHelper.ShowErrorDialog(error);
                }
                else if(retVal == mtReturnValues.mtNOT_EQUATION)
                {
                    error = "Error: No equation on the clipboard";
                    _appHelper.ShowErrorDialog(error);
                }
                else if(retVal == mtReturnValues.mtBAD_VERSION)
                {
                    error = "Error: Bad version";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "Error: Some other error";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Get equation preferences from the specified preferences file.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void GetPrefsFromFile_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {

                if (_prefsFile == null)
                    GetPrefsFile(out _prefsFile, true);

                StringBuilder outPrefs = new StringBuilder("", 1024);
                retVal = _MTSDK.MTGetPrefsFromFileMgn(_prefsFile, outPrefs, 1024);
                if (retVal == mtReturnValues.mtOK)
                {
                    error = "These are the preferences returned: " + outPrefs;
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtMEMORY)
                {
                    error = "Error: Not enough memory to store preferences";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtFILE_NOT_FOUND)
                {
                    error = "Error: File does not exist or bad pathname";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "Error: Some other error";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Convert internal preferences string to a form to be presented
        /// to the user.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void PerfsToUIForm_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                if (_prefsFile == null)
                    GetPrefsFile(out _prefsFile, true);

                StringBuilder prefsForInput = new StringBuilder("", 1024);

                retVal = _MTSDK.MTGetPrefsFromFileMgn( _prefsFile, prefsForInput, 1024 );
                if( retVal == mtReturnValues.mtOK )
                {
                    long length = 0;
                    length  = _MTSDK.MTConvertPrefsToUIFormMgn(prefsForInput.ToString(), null, 0);

                    if (length >= 0)
                    {
                        StringBuilder outPrefs = new StringBuilder("", (int)length);
                        retVal = _MTSDK.MTConvertPrefsToUIFormMgn(prefsForInput.ToString(), outPrefs, (Int16)length);

                        if (retVal == mtReturnValues.mtOK)
                        {
                            error = "Status of Preferences: " + outPrefs;
                            _appHelper.ShowOKDialog(error);
                        }
                        else if (retVal == mtReturnValues.mtMEMORY)
                        {
                            error = "Error: Not memory to store preferences";
                            _appHelper.ShowErrorDialog(error);
                        }
                        else
                        {
                            error = "Error: Some other error type";
                            _appHelper.ShowErrorDialog(error);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get MathType's current default equation preferences.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void GetDefaultPrefs_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                long length = 0;
                length = _MTSDK.MTGetPrefsMTDefaultMgn(null, 0);
                if (length > 0)
                {
                    StringBuilder outPrefs = new StringBuilder("", (int)length);

                    retVal = _MTSDK.MTGetPrefsMTDefaultMgn(outPrefs, (Int16)length);

                    if (retVal == mtReturnValues.mtOK)
                    {
                        error = "These are MathType's default parameters: " + outPrefs;
                        _appHelper.ShowOKDialog(error);
                    }
                    else if (retVal == mtReturnValues.mtMEMORY)
                    {
                        error = "Error: Not enough memory to store preferences";
                        _appHelper.ShowOKDialog(error);
                    }
                    else
                    {
                        error = "Error: Some other error";
                        _appHelper.ShowOKDialog(error);
                    }
                }
            }
        }

        /// <summary>
        /// Set MathType's default preferences for new equations.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void SetMTPrefs_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                if (_prefsFile == null)
                    GetPrefsFile(out _prefsFile, true);

                StringBuilder prefsForInput = new StringBuilder("", 1024);

                retVal = _MTSDK.MTGetPrefsFromFileMgn(_prefsFile, prefsForInput, 1024);
                if (retVal == mtReturnValues.mtOK)
                {
                    retVal = _MTSDK.MTSetMTPrefsMgn( mtPreferences.mtprfMODE_MTDEFAULT, prefsForInput.ToString(), -1 );

                    if (retVal == mtReturnValues.mtOK)
                    {
                        error = "Setting the preferences was successful";
                        _appHelper.ShowOKDialog(error);
                    }
                    else if (retVal == mtReturnValues.mtBAD_DATA)
                    {
                        error = "Error: Bad preferences string";
                        _appHelper.ShowErrorDialog(error);
                    }
                    else
                    {
                        error = "Error: Some other error";
                        _appHelper.ShowErrorDialog(error);
                    }
                }
            }

        }

        /// <summary>
        /// Get information about the current set of translators.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void GetTranslatorsInfo_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTGetTranslatorsInfoMgn(mtTranslatorInfo.mttrnCOUNT);
                if (retVal >= 0)
                {
                    error = "Number of registered translators: " + retVal.ToString();
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: Some other error";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Enumerate the available equation translators.  Currently this
        /// implementation will only iterate for the first translator.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void EnumTranslators_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                StringBuilder transName = new StringBuilder("", 1024);
                StringBuilder transDesc = new StringBuilder("", 1024);
                StringBuilder transFile = new StringBuilder("", 1024);

                retVal = _MTSDK.MTEnumTranslatorsMgn(1, transName, 1024, transDesc, 1024, transFile, 1024);
                if (retVal == mtReturnValues.mtMEMORY)
                {
                    error = "Error not enough memory";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtERROR)
                {
                    error = "Error: Some other error";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "A single enumeration was successful";
                    _appHelper.ShowOKDialog(error);
                }

            }
        }

        /// <summary>
        /// Resets to default options for MTXFormEqn (i.e. no substitutions,
        /// no translation, and use existing preferences).
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void FormReset_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTXFormResetMgn();
                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Form successfully reset";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: An error occurred resetting the form";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Adds a variable substitution to be performed with the
        /// next MTXFormEqn (may be called 0 or more times).
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void FormAddVarSub_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                string sub = "x";
                string replace = "y";
                retVal = _MTSDK.MTXFormAddVarSubMgn(mtxfrmSubValues.mtxfmSUBST_ALL,
                                                     mtxfrmSubValues.mtxfmVAR_SUB_PLAIN_TEXT,
                                                     sub,
                                                     sub.Length,
                                                     mtxfrmSubValues.mtxfmVAR_SUB_PLAIN_TEXT,
                                                     replace,
                                                     replace.Length,
                                                     mtxfrmSubValues.mtxfmSTYLE_TEXT);
                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Value was successfully substituted";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: Value was not substituted";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Specify translation to be performed with the next
        /// MTXFormEqn.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void FormSetTrans_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                string translator = "Texvc.tdl";

                retVal = _MTSDK.MTXFormSetTranslatorMgn( (ushort)mtxFrmSetTrans.mtxfmTRANSL_INC_NAME, translator );

                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Changing translator value successful";
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtFILE_NOT_FOUND)
                {
                    error = "Error: Could not find translator";
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtTRANSLATOR_ERROR)
                {
                    error = "Error: Errors compiling translator";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "Error: Some other error";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Specify a new set of preferences to be used with the
        /// next MTXFormEqn.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void FormSetPrefs_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                string pref = "";
                retVal = _MTSDK.MTXFormSetPrefsMgn(mtxfrmPreferences.mtxfmPREF_EXISTING, pref);

                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Set preference was successful";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: Some error occurred";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Transforms an equation (uses options specified via MTXFormAddVarSub,
        /// MTXFormSetTranslator, and MTXFormSetPrefs).
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void FormEqn_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                string          srcString = "MathType!MTEF!2!1!+-feaaguart1ev2aaatCvAUfeBSjuyZL2yd9gzLbvyNv2CaerbuLwBLnhiov2DGi1BTfMBaeXatLxBI9gBaerbd9wDYLwzYbItLDharqqtubsr4rNCHbGeaGqiVu0Je9sqqrpepC0xbbL8F4rqqrFfpeea0xe9Lq-Jc9vqaqpepm0xbba9pwe9Q8fs0-yqaqpepae9pg0FirpepeKkFr0xfr-xfr-xb9adbaqaaeGaciGaaiaabeqaamaabaabaaGcbaGaaG4naiabgUcaRiaadIhaaaa!3892!";
                string			srcPath = @"";
				const int		iDstLen = 10000;
				StringBuilder	strDest = new StringBuilder(iDstLen);
                MTAPI_DIMS      dims = new MTAPI_DIMS();

/*				IntPtr ptr =  Marshal.AllocHGlobal(1024);
                int length = srcString.Length;
                length = 2 * (length + 1);
                IntPtr srcData = Marshal.AllocHGlobal(length);
                Marshal.Copy(srcString.ToCharArray(), 0, srcData, srcString.Length);*/
				
				byte[] srcData = System.Text.Encoding.ASCII.GetBytes(srcString);

                retVal = _MTSDK.MTXFormEqnMgn(mtxFormEqn.mtxfmLOCAL,
                                           mtxFormEqn.mtxfmMTEF,
                                           srcData,
                                           srcString.Length,
                                           mtxFormEqn.mtxfmLOCAL,
                                           mtxFormEqn.mtxfmTEXT,
                                           strDest,
                                           iDstLen,
                                           srcPath,
                                           ref dims);


                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Equation transformation was successful";
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtNOT_EQUATION)
                {
                    error = "Error: Source data does not contain MTEF";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtSUBSTITUTION_ERROR)
                {
                    error = "Error: Could not perform one or more rules";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtTRANSLATOR_ERROR)
                {
                    error = "Error: Errors occured during translation (translation not done)";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtPREFERENCE_ERROR)
                {
                    error = "Error: Could not set preferences";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtMEMORY)
                {
                    error = "Error: Not enough space in dstData";
                    _appHelper.ShowErrorDialog(error);
                }
                else
                {
                    error = "Error: Some other error occurred";
                    _appHelper.ShowErrorDialog(error);
                }

                // Free memory allocated from hglobal allocation
                // Marshal.FreeHGlobal(srcData);
            }
        }

        /// <summary>
        /// Check error/status after MTXFormEqn.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void MTXFormGetStatus_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTXFormGetStatusMgn(mtxFormStatus.mtxfmSTAT_TRANSL);

                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Successful translation";
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtFILE_NOT_FOUND)
                {
                    error = "Error: Could not find translator";
                    _appHelper.ShowOKDialog(error);
                }
                else if (retVal == mtReturnValues.mtBAD_FILE)
                {
                    error = "Error: File found was not a translator";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: Bad index value";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Puts up a preview dialog for displaying preferences.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void MTPreviewDialog_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                string  title = "API Test Preview Dialog";
                string  prefs = "";
                string  closeBtnText = "Close";
                string  helpButtonText = "Help";
                string  helpFile = "MT6enu.chm";

                retVal = _MTSDK.MTPreviewDialogMgn(this.Handle,
                                                    title,
                                                    prefs,
                                                    closeBtnText,
                                                    helpButtonText,
                                                    0,
                                                    helpFile); // Need to get a help file
                if (retVal == mtReturnValues.mtOK)
                {
                    error = "Successfully displayed Preview Dialog";
                    _appHelper.ShowOKDialog(error);
                }
                else
                {
                    error = "Error: Could not display Preview Dialog";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Shows the about box for the current version of MathType.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void MTShowAboutBox_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                retVal = _MTSDK.MTShowAboutBoxMgn();

                if (retVal != mtReturnValues.mtOK)
                {
                    error = "Error: An error showing the about box";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }

        /// <summary>
        /// Displays the requested URL.
        /// </summary>
        /// <param name="sender">Parameter not used</param>
        /// <param name="e">Parameter not used</param>
        private void MTGetURL_Click(object sender, EventArgs e)
        {
            long retVal = 0;
            string error;

            if (_isMTAPIConnected == false)
            {
                InternalStartAPI(true);
            }

            if (_isMTAPIConnected == true)
            {
                int length = 0;
                length = _MTSDK.MTGetURLMgn(mtURLTypes.mturlMATHTYPE_HOME, false, null, 0);

                if (length > 0)
                {
                    length += 1;
                    StringBuilder strURL = new StringBuilder("", length);
                    retVal = _MTSDK.MTGetURLMgn(mtURLTypes.mturlMATHTYPE_HOME, true, strURL, length);
                }
                else
                {
                    retVal = mtReturnValues.mtERROR;
                }

                if (retVal == mtReturnValues.mtERROR)
                {
                    error = "Error: An error occurred getting the requested URL";
                    _appHelper.ShowErrorDialog(error);
                }
                else if (retVal == mtReturnValues.mtMEMORY)
                {
                    error = "Error: Not enough memory to store the URL";
                    _appHelper.ShowErrorDialog(error);
                }
            }
        }
    }
}

// Connect.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTGETEQUATIONADDIN/MTGetEquationAddin/Connect.cs 8     1/19/10 2:01p Jimm $

namespace MTGetEquationAddin
{
	using System;
	using Extensibility;
    using System.IO;
	using System.Runtime.InteropServices;
    using System.Reflection;
    using System.Text;
    using System.Windows.Forms;
    using Microsoft.Office.Core;
    using Microsoft.Win32;

    // This is a technique to shorten naming of specific interfaces, enumerations, structures, and exceptions.
    //
    using Word = Microsoft.Office.Interop.Word;
    using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
    using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
    using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
    using ConnectIEnumETC = System.Runtime.InteropServices.ComTypes.IEnumFORMATETC;
    using COMException = System.Runtime.InteropServices.COMException;
    using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;

	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the MTGetEquationAddinSetup project,
	// right click the project in the Solution Explorer, then choose install.
	#endregion

	/// <summary>
	///   The object for implementing an Add-in.
    ///
    ///   The skeleton for the code was developed with the help of this support document from Microsoft
    ///     http://support.microsoft.com/Default.aspx?kbid=302901
    ///
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("280690B1-CC77-46EF-AE05-687F95691C58"), ProgId("MTGetEquationAddin.Connect")]
	public class Connect : Object, Extensibility.IDTExtensibility2
	{
        // Private member objects to support the toolbar addin.
        private CommandBarButton    _myGetMathMLDataButton;
        private CommandBarButton    _mySetMathMLDataButton;
        private CommandBarButton    _myGetTeXInputDataButton;
        private CommandBarButton    _mySetTeXInputDataButton;
        private object              _applicationObject;
        private object              _addInInstance;

        /// <summary>
        /// Define Kernel32 calls using C# types, this allows for C# to utilize PInvoke to call the methods
        /// enclosed in the OLE32 region below.
        /// For more information on PInvoke please see the following link: http://msdn.microsoft.com/en-us/library/aa288468.aspx
        /// </summary>
        #region Kernel32 function calls
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GlobalLock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern bool GlobalUnlock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int GlobalSize(HandleRef handle);

        #endregion Kernel32 function calls

        /// <summary>
        /// Define Ole32 calls using C# types, this allows for C# to utilize PInvoke to call the methods
        /// enclosed in the OLE32 region below.
        /// For more information on PInvoke please see the following link: http://msdn.microsoft.com/en-us/library/aa288468.aspx
        /// </summary>
        #region OLE32 function calls
        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int OleGetAutoConvert(ref Guid oCurrentCLSID, out Guid pConvertedClsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool IsEqualGUID( ref Guid rclsid1, ref Guid rclsid );

        #endregion OLE32 function calls
        /// <summary>
		///		Implements the constructor for the Add-in object.
		///		Place your initialization code within this method.
		/// </summary>
		public Connect()
		{
		}

		/// <summary>
		///      Implements the OnConnection method of the IDTExtensibility2 interface.
		///      Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param term='application'>
		///      Root object of the host application.
		/// </param>
		/// <param term='connectMode'>
		///      Describes how the Add-in is being loaded.
		/// </param>
		/// <param term='addInInst'>
		///      Object representing this Add-in.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
            // Assigns parameters to member variables
			_applicationObject = application;
			_addInInstance = addInInst;

            if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup) {
                OnStartupComplete(ref custom);
            }
		}

		/// <summary>
		///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
		///     Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param term='disconnectMode'>
		///      Describes how the Add-in is being unloaded.
		/// </param>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
            if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown) {
                OnBeginShutdown(ref custom);
            }
            _applicationObject = null;
		}

		/// <summary>
		///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		///      Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		/// <summary>
		///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		///      Receives notification that the host application has completed loading.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref System.Array custom)
		{
            CommandBars commandBars;
            CommandBar  standardBar;

            try {
                commandBars = (CommandBars)_applicationObject.GetType().InvokeMember("CommandBars", BindingFlags.GetProperty, null, _applicationObject, null);

            }
            catch (Exception) {
                // Outlook has the CommandBars collection on the Explorer object.
                object activeExplorer = _applicationObject.GetType().InvokeMember("ActiveExplorer", BindingFlags.GetProperty, null, _applicationObject, null);
                commandBars = (CommandBars)activeExplorer.GetType().InvokeMember("CommandBars", BindingFlags.GetProperty, null, activeExplorer, null);
            }

            // Set up a custom button on the "Standard" commandbar.
            try {
                standardBar = commandBars["Standard"];
            }
            catch (Exception) {
                // Access names its main toolbar Database.
                standardBar = commandBars["Database"];
            }

            // Add the MathML Get and Set Buttons
            AddMathMLInputButtons(ref commandBars, ref standardBar);

            // Add the TeX Input Language Get and Set Buttons
            AddTexInputButtons(ref commandBars, ref standardBar);

            object name = _applicationObject.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, _applicationObject, null);

            // Display a simple message to show which application you started in.
            standardBar = null;
            commandBars = null;
		}

        /// <summary>
        ///
        /// </summary>
        /// <param name="commandBars"></param>
        /// <param name="standardBar"></param>
        private void AddMathMLInputButtons(ref CommandBars commandBars, ref CommandBar standardBar)
        {
            // In case the button was not deleted, use the existing one.
            try
            {
                _myGetMathMLDataButton = (CommandBarButton)standardBar.Controls["Get MathML"];
            }
            catch (Exception)
            {
                object getDataMissing = System.Reflection.Missing.Value;
                _myGetMathMLDataButton = (CommandBarButton)standardBar.Controls.Add(1, getDataMissing, getDataMissing, getDataMissing, getDataMissing);
                _myGetMathMLDataButton.Caption = "Get MathML";
                _myGetMathMLDataButton.Style = MsoButtonStyle.msoButtonCaption;
            }


            // The following items are optional, but recommended.
            //  The Tag property lets you quickly find the control
            //  and helps MSO keep track of it when more than
            //  one application window is visible. The property is required
            //  by some Office applications and should be provided.
            _myGetMathMLDataButton.Tag = "Get MathML";

            // The OnAction property is optional but recommended.
            //  It should be set to the ProgID of the add-in, so that if
            //  the add-in is not loaded when a user presses the button,
            //  MSO loads the add-in automatically and then raises
            //  the Click event for the add-in to handle.
            _myGetMathMLDataButton.OnAction = "!<MTGetEquationAddin.Connect>";

            _myGetMathMLDataButton.Visible = true;
            _myGetMathMLDataButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.GetMathMLEquations_Click);

            // Add the Set Data Button to the toolbar.
            try
            {
                _mySetMathMLDataButton = (CommandBarButton)standardBar.Controls["Set MathML"];
            }
            catch (Exception)
            {
                object setDataMissing = System.Reflection.Missing.Value;
                _mySetMathMLDataButton = (CommandBarButton)standardBar.Controls.Add(1, setDataMissing, setDataMissing, setDataMissing, setDataMissing);
                _mySetMathMLDataButton.Caption = "Set MathML";
                _mySetMathMLDataButton.Style = MsoButtonStyle.msoButtonCaption;
            }


            // The following items are optional, but recommended.
            //  The Tag property lets you quickly find the control
            //  and helps MSO keep track of it when more than
            //  one application window is visible. The property is required
            //  by some Office applications and should be provided.
            _mySetMathMLDataButton.Tag = "Set MathML";

            // The OnAction property is optional but recommended.
            //  It should be set to the ProgID of the add-in, so that if
            //  the add-in is not loaded when a user presses the button,
            //  MSO loads the add-in automatically and then raises
            //  the Click event for the add-in to handle.
            _mySetMathMLDataButton.OnAction = "!<MTGetEquationAddin.Connect>";

            _mySetMathMLDataButton.Visible = true;
            _mySetMathMLDataButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(SetMathMLToMathType_Click);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="commandBars"></param>
        /// <param name="standardBar"></param>
        private void AddTexInputButtons( ref CommandBars commandBars, ref CommandBar  standardBar )
        {
            // In case the button was not deleted, use the existing one.
            try
            {
                _myGetTeXInputDataButton = (CommandBarButton)standardBar.Controls["Get TeX Input"];
            }
            catch (Exception)
            {
                object getDataMissing = System.Reflection.Missing.Value;
                _myGetTeXInputDataButton = (CommandBarButton)standardBar.Controls.Add(1, getDataMissing, getDataMissing, getDataMissing, getDataMissing);
                _myGetTeXInputDataButton.Caption = "Get TeX Input";
                _myGetTeXInputDataButton.Style = MsoButtonStyle.msoButtonCaption;
            }


            // The following items are optional, but recommended.
            //  The Tag property lets you quickly find the control
            //  and helps MSO keep track of it when more than
            //  one application window is visible. The property is required
            //  by some Office applications and should be provided.
            _myGetTeXInputDataButton.Tag = "Get TeX Input";

            // The OnAction property is optional but recommended.
            //  It should be set to the ProgID of the add-in, so that if
            //  the add-in is not loaded when a user presses the button,
            //  MSO loads the add-in automatically and then raises
            //  the Click event for the add-in to handle.
            _myGetTeXInputDataButton.OnAction = "!<MTGetEquationAddin.Connect>";

            _myGetTeXInputDataButton.Visible = true;
            _myGetTeXInputDataButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.GetTeXInputLanguageEquations_Click);

            // Add the Set Data Button to the toolbar.
            try
            {
                _mySetTeXInputDataButton = (CommandBarButton)standardBar.Controls["Set TeX Input"];
            }
            catch (Exception)
            {
                object setDataMissing = System.Reflection.Missing.Value;
                _mySetTeXInputDataButton = (CommandBarButton)standardBar.Controls.Add(1, setDataMissing, setDataMissing, setDataMissing, setDataMissing);
                _mySetTeXInputDataButton.Caption = "Set TeX Input";
                _mySetTeXInputDataButton.Style = MsoButtonStyle.msoButtonCaption;
            }


            // The following items are optional, but recommended.
            //  The Tag property lets you quickly find the control
            //  and helps MSO keep track of it when more than
            //  one application window is visible. The property is required
            //  by some Office applications and should be provided.
            _mySetTeXInputDataButton.Tag = "Set TeX Input";

            // The OnAction property is optional but recommended.
            //  It should be set to the ProgID of the add-in, so that if
            //  the add-in is not loaded when a user presses the button,
            //  MSO loads the add-in automatically and then raises
            //  the Click event for the add-in to handle.
            _mySetTeXInputDataButton.OnAction = "!<MTGetEquationAddin.Connect>";

            _mySetTeXInputDataButton.Visible = true;
            _mySetTeXInputDataButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(SetTeXInputLanguageToMathType_Click);
        }

		/// <summary>
		///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		///      Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref System.Array custom)
		{
            // Start the clean up of the button GetMathML
            object getDataMissing = System.Reflection.Missing.Value;
            _myGetMathMLDataButton.Delete(getDataMissing);
            _myGetMathMLDataButton = null;

            // Start the clean up of the button SetMathML
            object setDataMissing = System.Reflection.Missing.Value;
            _mySetMathMLDataButton.Delete(setDataMissing);
            _mySetMathMLDataButton = null;
		}

        /// <summary>
        /// This is the button event handler for the Get MathML button click.
        ///
        /// This method is registered as an event handler in the OnStartupComplete method, and
        /// associated with the Get MathML button.
        /// </summary>
        /// <param name="cmdBarButton"> The button object that receives the click event. Not used, part of the command handler signature.</param>
        /// <param name="cancel"> Not used, part of the command handler signature.</param>
        private void GetMathMLEquations_Click(CommandBarButton cmdBarButton, ref bool cancel)
        {
            // Get the document that is current in focus for the button click event.  This is the document that will be checked
            //  for a shape collection.
            object doc = _applicationObject.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, _applicationObject, null);

            if (doc != null) {
                // Retrieve the shapes collection from the current word document.
                Word.InlineShapes shapes = (Word.InlineShapes)doc.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, doc, null);
                int numShapesIterated = 0;

                // Iterate over all of the shapes in the collection.
                if (shapes != null &&
                    shapes.Count > 0) {
                    numShapesIterated = IterateShapes(ref shapes, false, true);
                }
            }
        }

        /// <summary>
        /// This is the button event handler for the Get TeX Input button click.
        ///
        /// This method is registered as an event handler in the OnStartupComplete method, and
        /// associated with the Get TeX Input button.
        /// </summary>
        /// <param name="cmdBarButton"> The button object that receives the click event. Not used, part of the command handler signature.</param>
        /// <param name="cancel"> Not used, part of the command handler signature.</param>
        private void GetTeXInputLanguageEquations_Click(CommandBarButton cmdBarButton, ref bool cancel)
        {
            // Get the document that is current in focus for the button click event.  This is the document that will be checked
            //  for a shape collection.
            object doc = _applicationObject.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, _applicationObject, null);

            if (doc != null)
            {
                // Retrieve the shapes collection from the current word document.
                Word.InlineShapes shapes = (Word.InlineShapes)doc.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, doc, null);
                int numShapesIterated = 0;

                // Iterate over all of the shapes in the collection.
                if (shapes != null &&
                    shapes.Count > 0)
                {
                    numShapesIterated = IterateShapes(ref shapes, false, false);
                }
            }
        }

        /// <summary>
        /// This is the button event handler for the Set MathML button click.
        ///
        /// This method is registered as an event handle in the OnStartupComplete method, and
        /// associated with the Set MathML button.
        /// </summary>
        /// <param name="cmdBarButton"> The button that receives the click event.</param>
        /// <param name="cancel">Not used, part of the command handler signature.</param>
        private void SetMathMLToMathType_Click(CommandBarButton cmdBarButton, ref bool cancel)
        {
            // Get the document that is current in focus for the button click event.  This is the document that will be checked
            //  for a shape collection.
            object doc = _applicationObject.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, _applicationObject, null);

            if (doc != null) {
                // Retrieve the shapes collection from the current word document.
                Word.InlineShapes shapes = (Word.InlineShapes)doc.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, doc, null);

                int numShapesIterated = 0;

                // Iterate over all of the shapes in the collection.
                if (shapes != null &&
                    shapes.Count > 0) {
                    numShapesIterated = IterateShapes(ref shapes, true, true);
                }
            }
        }

        /// <summary>
        /// This is the button event handler for the Set TeX Input button click.
        ///
        /// This method is registered as an event handle in the OnStartupComplete method, and
        /// associated with the Set TeX Input button.
        /// </summary>
        /// <param name="cmdBarButton"> The button that receives the click event.</param>
        /// <param name="cancel">Not used, part of the command handler signature.</param>
        private void SetTeXInputLanguageToMathType_Click(CommandBarButton cmdBarButton, ref bool cancel)
        {
            // Get the document that is current in focus for the button click event.  This is the document that will be checked
            //  for a shape collection.
            object doc = _applicationObject.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, _applicationObject, null);

            if (doc != null)
            {
                // Retrieve the shapes collection from the current word document.
                Word.InlineShapes shapes = (Word.InlineShapes)doc.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, doc, null);

                int numShapesIterated = 0;

                // Iterate over all of the shapes in the collection.
                if (shapes != null &&
                    shapes.Count > 0)
                {
                    numShapesIterated = IterateShapes(ref shapes, true, false);
                }
            }
        }

        /// <summary>
        /// Small DataFormats helper class.  This class is only used with
        ///  CanUtilizeMathML.
        /// </summary>
        public class formatFinder
        {
            // A simple default constrcutor to explicitly set the member verified to false.
            public formatFinder() {
                verified = false;
            }
            public DataFormats.Format format;
            public bool verified;
        };

        /// <summary>
        /// Checks to see if the IDataObject supports Tex Input Language.
        ///
        /// This function is used by both Equation_GetData and Equation_SetData.  In the case of Equation_SetData
        /// it does a query because most applications will support the format on the SetData side if it supports it
        /// on the GetData side.
        /// </summary>
        /// <param name="dataObject">The dataObject of type IDataObject that will be queried for TeX Input Language.</param>
        /// <param name="dataFormatRequested">This is the primary requested data format for getting TeX Input Language data.</param>
        /// <param name="dataFormat">The clipboard format of TeX Input Language type supported.  If TeX INput Language is not supported
        /// this value will be null.</param>
        /// <returns>Does the dataObject support TeX Input Language.  True if the dataObject supports: TeX Input Language. False otherwise.</returns>
        private bool CanUtilizeTexInputLanguage(ref IDataObject dataObject, ref string dataFormatRequested, out DataFormats.Format dataFormat)
        {
            bool canProvideTeXInputLanguage = false;

            // Initialize the output.  The initialization is required by the compiler else it gives
            //  a warning that a parameter of type out is not initialized.
            dataFormat = null;

            // Create an instance of FORMATETC for use with IDataObject.
            ConnectFORMATETC formatEtc = new ConnectFORMATETC();

            // DataFormats.Format will query and hold data regarding custom clipboard
            //  formats.
            DataFormats.Format dataFormatTeXInputLanguage;

            // Find within the clipboard system the registered custom clipboard formats for MathML.
            //  This data format would have been registered with MathType. If they have not be register with
            //  MathType, they will be with these calls.
            dataFormatTeXInputLanguage = DataFormats.GetFormat("TeX Input Language");

            // Initialize a FORMATETC structure to prepare for requesting data.
            formatEtc.cfFormat = (Int16)dataFormatTeXInputLanguage.Id;
            formatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
            formatEtc.lindex = -1;
            formatEtc.ptd = (IntPtr)0;
            formatEtc.tymed = TYMED.TYMED_HGLOBAL;

            // Return value for the data QueryGetData check.
            int queryReturn = -1;

            if (dataObject != null)
            {
                // Query for the TeX Input Language type data format
                queryReturn = dataObject.QueryGetData(ref formatEtc);
                if (queryReturn == 0)
                {
                    dataFormat = dataFormatTeXInputLanguage;
                    // Since we have a positive QueryGetData then there is at least one
                    //  TeX Input Language format supported.
                    canProvideTeXInputLanguage = true;
                }
            }

            return canProvideTeXInputLanguage;
        }

        /// <summary>
        /// Checks to see if the IDataObject supports MathML in any one of its three variants.  The three different
        /// versions are: MathML, MathML Presentation, application/mathml+xml.
        ///
        /// This function is used by both Equation_GetMathML and Equation_SetMathML.  In the case of Equation_SetMathML
        /// it does a query because most applications will support the format on the SetData side if it supports it
        /// on the GetData side.
        /// </summary>
        /// <param name="dataObject">The dataObject of type IDataObject that will be queried for supporting MathML types supported.</param>
        /// <param name="dataFormatRequested">This is the primary requested data format for getting MathML data.</param>
        /// <param name="dataFormat">The clipboard format of the MathML type supported.  If no MathML type is supported
        /// this value will be null.</param>
        /// <returns>Does the dataObject support MathML.  True if the dataObjects supports: MathML, MathML Presentation, or
        /// application/mathml+xml. False otherwise.</returns>
        private bool CanUtilizeMathML(ref IDataObject dataObject, ref string dataFormatRequested, out DataFormats.Format dataFormat)
        {
            bool canProvideMathML = false;

            // Initialize the output.  The initialization is required by the compiler else it gives
            //  a warning that a parameter of type out is not initialized.
            dataFormat = null;

            // Create an instance of FORMATETC for use with IDataObject.
            ConnectFORMATETC formatEtc = new ConnectFORMATETC();

            // DataFormats.Format will query and hold data regarding custom clipboard
            //  formats.
            DataFormats.Format dataFormatMathML;
            DataFormats.Format dataFormatMathMLPres;
            DataFormats.Format dataFormatAppMathMLXML;

            // Find within the clipboard system the registered custom clipboard formats for MathML.
            //  This data format would have been registered with MathType. If they have not be register with
            //  MathType, they will be with these calls.
            dataFormatMathML = DataFormats.GetFormat("MathML");
            dataFormatMathMLPres = DataFormats.GetFormat("MathML Presentation");
            dataFormatAppMathMLXML = DataFormats.GetFormat("application/mathml+xml");


            // create all three formats that are going to be checked.
            formatFinder[]  finder = new formatFinder[3];

            int x = 0;
            int countFormats = 3;

            // Allocate and populate the data structures to be used for querying the
            //  data object for the MathML formats.
            while (x < countFormats) {
                finder[x] = new formatFinder();
                x++;
            }

            finder[0].format = dataFormatMathML;
            finder[1].format = dataFormatMathMLPres;
            finder[2].format = dataFormatAppMathMLXML;


            // Initialize a FORMATETC structure to prepare for requesting data.
            formatEtc.cfFormat = (Int16)dataFormatMathML.Id;
            formatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
            formatEtc.lindex = -1;
            formatEtc.ptd = (IntPtr)0;
            formatEtc.tymed = TYMED.TYMED_HGLOBAL;

            x = 0;

            // Check all of the MathML formats, MathML, MathML Presetation, application/mathml+xml.
            while (x < countFormats) {
                // FORMATETC.cfFormat as defined in the API is a clipFormat, this is define by the CLR as
                //  an Int16.
                formatEtc.cfFormat = (Int16)finder[x].format.Id;

                // Return value for the data QueryGetData check.
                int queryReturn = -1;

                if (dataObject != null) {
                    // Query for the MathML type data format
                    queryReturn = dataObject.QueryGetData(ref formatEtc);
                    if (queryReturn == 0) {
                        // Check to see if the requested format was just queried.
                        //  If so leave the searching and return back to the caller.
                        if (finder[x].format.Name == dataFormatRequested) {
                            dataFormat = finder[x].format;
                            return true;
                        }
                        else {
                            // Since we have a positive QueryGetData then there is at least one
                            //  MathML format supported.
                            canProvideMathML = true;
                            finder[x].verified = true;
                        }
                    }
                }
                x++;
            }

            // Since the search did not yield the requested format.  Iterate and get the first verfied MathML
            //  format.
            if (canProvideMathML) {
                x = 0;
                while (x < countFormats) {
                    // Just get the first format that was verified supported by the server
                    if (finder[x].verified == true) {
                        // Set the data format for use by the IDataObject in the caller.
                        //  Remember dataFormat is an out parameter.
                        dataFormat = finder[x].format;
                        break;
                    }
                    x++;
                }
            }

            // This return value was set in the initial verification for the data formats.
            //  It was set to true because there was at least one MathML data format.
            return canProvideMathML;
        }


        /// <summary>
        /// Iterate over all of the shapes (InlineShape) in the collection,
        /// this collection is retrieved from the current active document.
        /// </summary>
        /// <param name="shapes"> The shape collection that will be iterated.
        /// This collection is from the currently active document.</param>
        /// <param name="isSetData"> Used as a switch for which handler is calling for the Get
        /// or Set data object.</param>
        /// <param name="isMathML">Is the type of operation requested a MathML operation</param>
        /// <returns>The number of shapes that have been either used with GetMathML or SetMathML.</returns>
        private int IterateShapes(ref Word.InlineShapes shapes, bool isSetData, bool isMathML)
        {
            int numShapesViewed = 0;

            // Validate there are a number of shape objects contained within the shapes collection.  If there is
            if (shapes != null &&
                shapes.Count > 0) {

                int                 count = 1;
                int                 numShapes = 0;
                Word.InlineShape    shape;

                numShapes = shapes.Count;

                // Count is the LCV and is used as the shapes accessor, it is 1 based.
                //  This is based on the Word object model.
                while (count <= numShapes) {
                    shape = shapes[count];

                    if (shape != null &&
                        shape.OLEFormat != null) {

                        bool     retVal = false;
                        string   progID;
                        Guid     autoConvert;

                        // Get the program ID which is a string from the current shape object and
                        //  find the ClassID (GUID) which is associated with this program ID.
                        progID = shape.OLEFormat.ProgID;

                        // Check the program ID to make sure that it is an equation shape we are editing.
                        if (progID.CompareTo("Equation.DSMT4") == 0) {
                            retVal = FindAutoConvert(ref progID, out autoConvert);

                            if (retVal == true) {
                                bool doesServerExists = false;

                                // If we are successful with the conversion of the CLSID we now need to query
                                //  the registry to see if there is an OLE server associated with the GUID (autoConvert)
                                //  If there is, verify that the server exists on the disk.
                                doesServerExists = DoesServerExist(ref autoConvert);

                                // If the server exists the next step is to validate if the server supports MathML.
                                if (doesServerExists) {
                                    string verbName = "RunForConversion";

                                    // The server supports MathML lets get the index for the requested verb.
                                    int indexForVerb = GetVerbIndex(verbName, ref autoConvert);

                                    if (isMathML)
                                    {
                                        bool supportsMathML = false;
                                        string dataFormat = "MathML";

                                        supportsMathML = DoesServerSupportFormat(ref autoConvert, ref dataFormat);

                                        // If the server supports
                                        if (supportsMathML)
                                        {
                                            if (indexForVerb != 999)
                                            {
                                                // Depending on the which event handler is calling iterate shapes
                                                //  will determine which Equation_X method will be called.  Please
                                                //  see Equation_SetData and Equation_GetData for more details.
                                                if (isSetData)
                                                {
                                                    Equation_SetData(ref shape, ref dataFormat, indexForVerb, true);
                                                }
                                                else
                                                {
                                                    Equation_GetData(ref shape, ref dataFormat, indexForVerb, true);
                                                }

                                                // Increment the number of shapes that have been affected by either Equation_GetData or
                                                //  Equation_SetData.
                                                numShapesViewed++;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        string dataFormat = "TeX Input Language";
                                        bool supportsTexInputLanguage = false;

                                        supportsTexInputLanguage = DoesServerSupportFormat(ref autoConvert, ref dataFormat);

                                        if (supportsTexInputLanguage)
                                        {
                                            if (indexForVerb != 999)
                                            {
                                                // Depending on the which event handler is calling iterate shapes
                                                //  will determine which Equation_X method will be called.  Please
                                                //  see Equation_SetData and Equation_GetData for more details.
                                                if (isSetData)
                                                {
                                                    Equation_SetData(ref shape, ref dataFormat, indexForVerb, false);
                                                }
                                                else
                                                {
                                                    Equation_GetData(ref shape, ref dataFormat, indexForVerb, false);
                                                }

                                                // Increment the number of shapes that have been affected by either Equation_GetData or
                                                //  Equation_SetData.
                                                numShapesViewed++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Increment the LCV (count)
                    count++;
                }
            }

            // Return the number of shapes visited from the shapes collection.
            return numShapesViewed;
        }

        #region Server Support Methods
        /// <summary>
        /// Verifies the existance for the server, if the server does exist it will return true.
        /// Else it will return false if there is not a valid path to the server.
        /// </summary>
        /// <param name="guidToCheck">The GUID that represents the server to check if it exists.</param>
        /// <returns>True if the server does exist on disk, else false</returns>
        private bool DoesServerExist(ref Guid guidToCheck)
        {
            bool        doesServerExists = false;

            // Find the registry key cooresponding with the guidToCheck for the LocalServer32 entry.
            string      regLocation = @"Software\Classes\CLSID\" + @"{" + guidToCheck.ToString() + @"}" + @"\" + @"LocalServer32";
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey(regLocation);

            if (regKey != null) {
                // Get the value(s) associated with the LocalServer32 registry entry.
                string[] valueNames = regKey.GetValueNames();
                // There should only be one entry and that is the absolute path to the application.
                string pathToExe = (string)regKey.GetValue(valueNames[0]);

                if (pathToExe.Length > 0) {
                    // If we have a string representing the absolute path to the directory then,
                    //  verify that the application exists.
                    if( File.Exists(pathToExe) )
                        doesServerExists = true;
                }
            }

            return doesServerExists;
        }

        /// <summary>
        /// For the server that is found via a GUID determine the verb to
        /// start the server from the requested string.
        /// </summary>
        /// <param name="verbToFind">The verb to be found.</param>
        /// <param name="guidToCheck">GUID of the currently found server for the verb.</param>
        /// <returns>The value for the verb searched.  If the verb cannot be found
        /// or there is an error then we will return an error state for the verb which is denoted as
        /// the value 999.</returns>
        private int GetVerbIndex(string verbToFind, ref Guid guidToCheck)
        {
            int         indexForVerb = 999;

            // Find the registry key cooresponding with the guidToCheck for the verb entry.
            string      regLocation = @"Software\Classes\CLSID\" + "{" + guidToCheck.ToString() + "}" + @"\Verb";
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey(regLocation);

            if (regKey != null) {

                // Are there any subkeys to check for the existance of a verb.  If so
                //  iterate over the subkeys to find the verb of interest.  The verb of
                //  interest is the parameter verbToFind.
                if (regKey.SubKeyCount > 0) {
                    int outerCount = 0;
                    int subKeyCount = 0;

                    string[] valueNames = regKey.GetSubKeyNames();

                    // Iterate over all of the regkey(s) from the Verb regkey associated with the guidToCheck.
                    //  There maybe more than one sub kley contained within the regkey.
                    while (outerCount < regKey.SubKeyCount) {
                        RegistryKey subKey;
                        if (regKey.SubKeyCount > 0) {
                            subKey = regKey.OpenSubKey(valueNames[outerCount]);
                            if (subKey != null) {
                                int innerValueCount = 0;
                                string[] verbs = subKey.GetValueNames();
                                subKeyCount = subKey.ValueCount;
                                string  verb;

                                // For each subkey see if there is a verb that matches the one
                                //  that is being searched for.  If so, convert the subkey name which is a
                                //  string to that of an integer.  This integer (verb index) will be returned and used
                                //  for creating and starting the server.  The index value is used by the server so
                                //  the server will know how to create itself appropriately.
                                while (innerValueCount < subKeyCount) {
                                    verb = (string)subKey.GetValue(verbs[innerValueCount]);
                                    if (verb.Contains(verbToFind) == true) {
                                        string numVerb;
                                        numVerb = valueNames[outerCount].ToString();
                                        indexForVerb = int.Parse( numVerb );
                                        break;
                                    }
                                    innerValueCount++;
                                }
                            }
                        }

                        // If the indexForVerb is not 999, which indicates we have found a matching verb.
                        //  If it is found lets break out of the current loop and return.
                        if (indexForVerb != 999)
                            break;

                        outerCount++;
                    }
                }
            }

            return indexForVerb;
        }

        /// <summary>
        /// As applications get upgraded or replaced, there will be documents that are associated with
        ///  older versions of the application via GUID.  This method will find the GUID that represents
        ///  the application that is backward compatibile with the current IDataObject.
        /// </summary>
        /// <param name="guidToCheck"> This is the initial Guid that starts the search</param>
        /// <param name="autoConvert"> This parameter contains the converted Guid, it could be the
        /// same as the parameter guidToCheck.</param>
        private void RecurseAutoConvert(ref Guid guidToCheck, out Guid autoConvert)
        {
            int comRetVal = 0;

            // Get the next GUID in the auto conversion chain for the GUIDs.
            //  If there is a valid GUID that can be used for conversion check
            //  for equality.
            //
            // OleGetAutoConvert is a call to the OLE32 DLL.
            comRetVal = OleGetAutoConvert(ref guidToCheck, out autoConvert);
            if (comRetVal == 0) {

                bool isGuidTheSame = false;
                try {
                    // This is a call to OLE32 to check if the two Guids are equal.
                    isGuidTheSame = IsEqualGUID(ref guidToCheck, ref autoConvert);
                }
                catch (COMException comException) {

                    // Catch and display the COM exception message if thrown by IsEqualGUID.
                    MessageBox.Show(comException.Message);
                }
                catch (Exception e) {
                    // Catch and display the exception message if anything other than a COM exception is thrown.
                    MessageBox.Show(e.Message);
                }

                // If the GUIDs are the same recurse until the conversion chain has
                //  been exhausted.
                if (isGuidTheSame == false) {
                    guidToCheck = autoConvert;
                    RecurseAutoConvert(ref guidToCheck, out autoConvert);
                }
            }
            else {
                // The end of the conversion chain has been reached.  The in parameter is
                //  now assigned to the out parameter.
                autoConvert = guidToCheck;
            }
        }

        /// <summary>
        /// From the Program ID find the actual GUID that represents
        ///  the server that is to be created for use by the OLE objects.
        /// </summary>
        /// <param name="progID"> The program ID that will get converted to the GUID which
        ///  represents the application that will consume the IDataObject.</param>
        /// <param name="autoConvert"> The GUID which represents the application that will
        ///   consume the IDataObjects.</param>
        /// <returns>True if the guid is found, else false.</returns>
        private bool FindAutoConvert(ref string progID, out Guid autoConvert)
        {
            bool    retVal = false;
            Guid    outputGuid;
            int     comRetVal = 0;

            // A call to OLE32.  Converts the program ID to a CLSIS(Guid).
            comRetVal = CLSIDFromProgID(progID, out outputGuid);

            // If a successful conversion occurs, then recurse down the
            //  GUID auto conversion chain to find the GUID that is actually
            //  used as the server for the OLE objects.
            if (comRetVal == 0) {
                RecurseAutoConvert(ref outputGuid, out autoConvert);
                retVal = true;
            }
            else {
                autoConvert = outputGuid;
            }

            return retVal;
        }
        #endregion Server Support Methods

        #region Equation methods

        /// <summary>
        /// This method will check the found server to see if this application
        ///  is registered to support the requested format.  This check is done using DataFormats
        ///  registry subkey of the main application registry key.
        /// </summary>
        /// <param name="guidToCheck">The application GUID to check to verify if it supports
        ///   the requested data format.</param>
        /// <param name="dataFormat">The data format to verify, if it is supported by the server.</param>
        /// <returns></returns>
        private bool DoesServerSupportFormat(ref Guid guidToCheck, ref string dataFormat)
        {
            // Find the registry subkey for the GUID to verify which data formats the server supports.
            string regLocation = @"Software\Classes\CLSID\" + "{" + guidToCheck.ToString() + "}" + @"\DataFormats\GetSet";
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey(regLocation);

            if (regKey != null) {
                string[] valueNames = regKey.GetSubKeyNames();
                int x = 0;

                // Iterate over the keys to determine if requested format is supported.
                while (x < regKey.SubKeyCount) {
                    RegistryKey subKey;
                    if (regKey.SubKeyCount > 0) {
                        subKey = regKey.OpenSubKey(valueNames[x]);
                        if (subKey != null) {
                            string[] dataFormats = subKey.GetValueNames();
                            int y = 0;
                            while (y < subKey.ValueCount) {
                                string format = (string)subKey.GetValue(dataFormats[y]);

                                // Compare to verify if requested dataFormat is the same as the subkey.
                                if (format.Contains(dataFormat)) {
                                    return true;
                                }
                                y++;
                            }
                        }
                    }
                    x++;
                }
            }

            return false;
        }

        /// <summary>
        /// Called from iterate shapes.  It will get the requested data from the MathType server
        ///  and will display the resulting data (depending on the requested format) in a message box.
        /// </summary>
        /// <param name="shape">The shape which is used as an IDataObject.</param>
        /// <param name="dataFormatRequested">This is the requested data format for getting data from IDataObject::GetData.</param>
        /// <param name="indexForVerb">The index to start (Do Verb) the application for the give shape.</param>
        private void Equation_GetData(ref Word.InlineShape shape, ref string dataFormatRequested, int indexForVerb, bool isMathML)
        {
            if (shape != null) {
                IDataObject  oleDataObject;
                object       dataObject = null;
                object       objVerb;

                objVerb = indexForVerb;

                try {
                    // Start MathType, and get the dataobject that is connected to the server.
                    shape.OLEFormat.DoVerb(ref objVerb);

                    // The DoVerb operation must be successful in order to get the object
                    dataObject = shape.OLEFormat.Object;
                }
                catch (Exception e) {
                    // Display the exception message to the user.
                    MessageBox.Show(e.Message);
                }

                MTSDKDN.MathTypeSDK.IOleObject oleObject = null;

                // This is a C# version of a QueryInterface.  The "as" operator for C# in this case will
                //  act as a QueryInterface to get the appropriate interface.  If the appropriate interface
                //  is not present, it will return null in the assignment.
                if (dataObject != null) {
                    oleDataObject = dataObject as IDataObject;
                    oleObject = dataObject as MTSDKDN.MathTypeSDK.IOleObject;
                }
                else {
                    // There was an issue with the addin trying to start with the verb we
                    //  knew.  A backup is to call the with the primary verb and start the
                    //  application normally.  This will start the application in a non-hidden
                    //  mode.  It will appear as MathType has been started from the start menu.
                    objVerb = Word.WdOLEVerb.wdOLEVerbPrimary;
                    shape.OLEFormat.DoVerb(ref objVerb);

                    dataObject = shape.OLEFormat.Object;
                    oleDataObject = dataObject as IDataObject;
                    oleObject = dataObject as MTSDKDN.MathTypeSDK.IOleObject;
                }

                // Create instances of FORMATETC and STGMEDIUM for use with IDataObject
                ConnectFORMATETC    formatEtc = new ConnectFORMATETC();
                ConnectSTGMEDIUM    stgMedium = new ConnectSTGMEDIUM();
                DataFormats.Format  dataFormat;

                bool isFormatSupported = false;

                // Verify the appropriate requested data format.
                if (isMathML == true)
                {
                    isFormatSupported = CanUtilizeMathML(ref oleDataObject, ref dataFormatRequested, out dataFormat);
                }
                else
                {
                    isFormatSupported = CanUtilizeTexInputLanguage(ref oleDataObject, ref dataFormatRequested, out dataFormat);
                }

                if (isFormatSupported == true)
                {
                    if (oleDataObject != null)
                    {
                        // Initialize a FORMATETC structure to get the requested data
                        formatEtc.cfFormat = (Int16)dataFormat.Id;
                        formatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
                        formatEtc.lindex = -1;
                        formatEtc.ptd = (IntPtr)0;
                        formatEtc.tymed = TYMED.TYMED_HGLOBAL;

                        // By setting this member to NULL, states there is no data to process.  If this was
                        //  non-null an error would occur with GetData.
                        stgMedium.tymed = TYMED.TYMED_NULL;

                        try
                        {
                            // Get the data requested from the IDataObject.  formatEtc is the data type requested,
                            //  stgMedium is the output parameter that will contain the data.
                            oleDataObject.GetData(ref formatEtc, out stgMedium);
                        }
                        catch (System.Runtime.InteropServices.COMException e)
                        {
                            MessageBox.Show(e.ToString());
                            throw;
                        }

                        // If there is data in the STGMEDIUMS tymed member, display it to the caller in a message box.
                        if (stgMedium.tymed == TYMED.TYMED_HGLOBAL &&
                            stgMedium.unionmember != null)
                        {

                            // This function does nothing more than displaying the equation in a MessageBox.
                            WriteOutEquationFromStgMedium(ref stgMedium);

                            // The data object reference must be closed.  Having too many IDataObjects open will
                            //  cause the system to run short of resources.
                            //
                            // The oleObject was a C# query interface earlier in this function.
                            if (oleObject != null)
                            {
                                oleObject.Close(MTSDKDN.MathTypeSDK.OleClose.OLECLOSE_NOSAVE);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Called from iterate shapes.  It will set the MathML/Tex Input Language data to the equation in the MathType server.
        ///  Once the equation has the data set into it the server will force a redrawing of
        ///  the equation within word to show the changes.
        /// </summary>
        /// <param name="shape">The shape which is used as an IDataObject.</param>
        /// <param name="dataFormatRequested">This is the primary requested data format for getting MathML/Tex Input Language data.</param>
        /// <param name="indexForVerb">The index to start the applcation for the give shape.</param>
        private void Equation_SetData(ref Word.InlineShape shape, ref string dataFormatRequested, int indexForVerb, bool isMathML)
        {
            if (shape != null) {
                IDataObject oleDataObject;
                object dataObject = null;
                object objVerb;

                objVerb = indexForVerb;

                try {
                    // Start MathType, and get the dataobject that is connected to the server.
                    shape.OLEFormat.DoVerb(ref objVerb);

                    // The DoVerb operation must be successful in order to get the object
                    dataObject = shape.OLEFormat.Object;
                }
                catch (Exception e) {
                    // we have an issue with trying to get the verb,
                    //  There will be a attempt at another way to start the application.
                    MessageBox.Show(e.Message);
                }

                MTSDKDN.MathTypeSDK.IOleObject oleObject = null;

                // This is a C# version of a QueryInterface.  The "as" operator for C# in this case will
                //  act as a QueryInterface to get the appropriate interface.  If the appropriate interface
                //  is not present, it will return null in the assignment.
                if (dataObject != null) {
                    oleDataObject = dataObject as IDataObject;
                    oleObject = dataObject as MTSDKDN.MathTypeSDK.IOleObject;
                }
                else {

                    // There was an issue with the addin trying to start with the verb we
                    //  knew.  A backup is to call the with the primary verb and start the
                    //  application normally.
                    objVerb = Word.WdOLEVerb.wdOLEVerbPrimary;
                    shape.OLEFormat.DoVerb(ref objVerb);

                    dataObject = shape.OLEFormat.Object;
                    oleDataObject = dataObject as IDataObject;
                    oleObject = dataObject as MTSDKDN.MathTypeSDK.IOleObject;
                }

                // Create instances of FORMATETC and STGMEDIUM for use with IDataObject
                ConnectFORMATETC    formatEtc = new ConnectFORMATETC();
                ConnectSTGMEDIUM    stgMedium = new ConnectSTGMEDIUM();
                DataFormats.Format  dataFormat;

                bool isFormatSupported = false;

                // Verify the appropriate requested data format.
                if (isMathML == true)
                {
                    isFormatSupported = CanUtilizeMathML(ref oleDataObject, ref dataFormatRequested, out dataFormat);
                }
                else
                {
                    isFormatSupported = CanUtilizeTexInputLanguage(ref oleDataObject, ref dataFormatRequested, out dataFormat);
                }

                if (isFormatSupported == true)
                {
                    if (oleDataObject != null)
                    {
                        //Initialize a FORMATETC structure to set the data
                        formatEtc.cfFormat = (Int16)dataFormat.Id;
                        formatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
                        formatEtc.lindex = -1;
                        formatEtc.ptd = (IntPtr)0;
                        formatEtc.tymed = TYMED.TYMED_HGLOBAL;

                        string wideMathEqn;

                        // Get a simple equation of the appropriate type.
                        GetSimpleEqn(isMathML, out wideMathEqn);

                        // Convert the hard coded string to a HGlobal which will be assigned to a STGMEDIUM structure
                        //  for setting into the oleDataObject.
                        try
                        {
                            // Create marshalable data for the stgMedium object.
                            stgMedium.unionmember = Marshal.StringToHGlobalAuto(wideMathEqn);
                            stgMedium.tymed = TYMED.TYMED_HGLOBAL;
                            stgMedium.pUnkForRelease = 0;
                        }
                        catch (Exception exp)
                        {
                            System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + exp.ToString());
                        }

                        // Set the hard coded Simple Equation data into this data object.
                        oleDataObject.SetData(ref formatEtc, ref stgMedium, false);

                        // The data object reference must be closed.  Having too many IDataObjects open will
                        //  cause the system to run short of resources.
                        //
                        // The oleObject was a C# query interface earlier in this function.
                        if (oleObject != null)
                        {
                            oleObject.Close(MTSDKDN.MathTypeSDK.OleClose.OLECLOSE_NOSAVE);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This will take the data from the STGMEDIUM and display the MathML that is
        ///  in the TYMED_HGLOBAL.  It will display this data in a message box.
        /// </summary>
        /// <param name="stgMedium">The structure contains the HGLOBAL data that has been
        ///  retrieved from a call to GetData.</param>
        private void WriteOutEquationFromStgMedium(ref ConnectSTGMEDIUM stgMedium)
        {
            IntPtr ptr;
            byte[] rawArray = null;

            // Verify that our data contained within the STGMEDIUM is non-null
            if (stgMedium.unionmember != null) {

                // Get the pointer to the data that is contained
                //  within the STGMEDIUM
                ptr = stgMedium.unionmember;

                // The pointer now becomes a Handle reference.
                HandleRef handleRef = new HandleRef(null, ptr);

                try {
                    // Lock in the handle to get the pointer to the data
                    IntPtr ptrToHandle = GlobalLock(handleRef);

                    // Get the size of the memory block
                    int length = GlobalSize(handleRef);

                    // New an array of bytes and Marshal the data across.
                    rawArray = new byte[length];
                    Marshal.Copy(ptrToHandle, rawArray, 0, length);

                    // Display the text by creating a string from the rawArray
                    string stringToShow = Encoding.ASCII.GetString(rawArray);

                    // Display it with a MessageBox
                    System.Windows.Forms.MessageBox.Show(stringToShow);
                }
                catch (Exception e) {
                    System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + e.ToString());
                }
                finally {
                    // This always gets called.  It does not require that an exception be thrown.
                    GlobalUnlock(handleRef);
                }
            }
        }

        private void GetSimpleEqn(bool isMathML, out string simpleEqn)
        {
            simpleEqn = "";

            if (isMathML)
            {
                simpleEqn = "<math><mi>x</mi></math>";
            }
            else
            {
                simpleEqn = "<math>\\sqrt{a^2 + b^2} </math>";
            }
        }
        #endregion Equation methods

    }
}

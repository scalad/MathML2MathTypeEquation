// AppHelper.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTSDKTESTAPPLICATION/MTSDKTestApplication/AppHelper.cs 3     1/19/10 2:02p Jimm $

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace MTSDKTestApplication
{
    /// <summary>
    /// A very simple helper class.  Used to wrap calls to the
    /// message box.  Message box is from System.Windows.Forms
    /// </summary>
    class AppHelper
    {
        public AppHelper()
        {
        }

        public void ShowOKDialog(string stringToShow)
        {
            string mtsdkString = "MT SDK Test Application";
            MessageBox.Show(stringToShow, mtsdkString);
        }

        public void ShowErrorDialog(string stringError)
        {
            string mtsdkString = "MT SDK Test Application";
            MessageBox.Show(stringError, mtsdkString);
        }
    }
}

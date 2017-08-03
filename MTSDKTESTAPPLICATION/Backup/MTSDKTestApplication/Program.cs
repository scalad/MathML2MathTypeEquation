// Program.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTSDKTESTAPPLICATION/MTSDKTestApplication/Program.cs 3     1/19/10 2:02p Jimm $

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MTSDKTestApplication
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MTAPISmokeTest());
        }
    }
}

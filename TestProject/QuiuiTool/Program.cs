using Hamilton.TestFramework.QuiuiTool;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TestProject
{
    /// <summary>
    /// Initializes the QuiuiTool application.
    /// </summary>
    /// <remarks>
    /// The Microsoft Public License (Ms-PL) Copyright (c) 2018 Hamilton Marques.
    /// Meet Hamilton at http://bit.do/HamiltonLinkedin
    /// Videos about QuiuiTool at http://bit.do/QuiuiTool
    /// </remarks>
    static class Program
    {
        /// <summary>
        /// Calls a test method.
        /// </summary>
        /// <param name="testName">Name of the test method.</param>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        /// <returns>null when the test passes; the exception thrown when the test fails, is inconclusive or is unknown.</returns>
        public delegate Exception CallTestMethod(string testName, string docxName, List<string> logs);

        /// <summary>
        /// Quits a WebDriver instance, closing every associated window.
        /// </summary>
        /// <param name="driverKey">Key of the ChromeDriver instance.</param>
        public delegate void QuitDriver(string driverKey);

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            CallTestMethod call = QuiuiBase.CallTestMethod;
            QuitDriver quit = QuiuiBase.QuitDriver;

            QuiuiApp myQuiui = new QuiuiApp(QuiuiBase.solutionFolder, QuiuiBase.hash, QuiuiBase.testClasses, call, quit);

            SingleInstanceApplication.Run(myQuiui, OnStartupNextInstance);
        }

        /// <summary>
        /// Provides the method that runs a single-instance application.
        /// </summary>
        public class SingleInstanceApplication : WindowsFormsApplicationBase
        {
            /// <summary>
            /// Initializes a new instance of the SingleInstanceApplication class.
            /// </summary>
            private SingleInstanceApplication()
            {
                base.IsSingleInstance = true;
            }

            /// <summary>
            /// Runs a single-instance application.
            /// </summary>
            /// <param name="form">Application to be run.</param>
            /// <param name="handler">Event to bring the application to the foreground.</param>
            public static void Run(Form form, StartupNextInstanceEventHandler handler)
            {
                SingleInstanceApplication app = new SingleInstanceApplication();
                app.MainForm = form;
                app.StartupNextInstance += handler;
                app.Run(Environment.GetCommandLineArgs());
            }
        }

        /// <summary>
        /// Brings the application to the foreground.
        /// Occurs when attempting to start a single-instance application and the application is already active.
        /// </summary>
        /// <param name="sender">Default Visual Basic parameter.</param>
        /// <param name="e">Default Visual Basic parameter.</param>
        private static void OnStartupNextInstance(object sender, StartupNextInstanceEventArgs e)
        {
            e.BringToForeground = true;
        }
    }
}
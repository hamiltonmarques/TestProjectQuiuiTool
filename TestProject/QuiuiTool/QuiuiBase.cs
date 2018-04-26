using Hamilton.TestFramework.QuiuiTool;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestProject
{
    /// <summary>
    /// Base class for the QuiuiTool application.
    /// Provides methods to:
    /// call test methods;
    /// initialize WebDriver instances;
    /// take screenshot of a web page;
    /// quit WebDriver instances.
    /// </summary>
    /// <remarks>
    /// The Microsoft Public License (Ms-PL) Copyright (c) 2018 Hamilton Marques.
    /// Meet Hamilton at http://bit.do/HamiltonLinkedin
    /// Videos about QuiuiTool at http://bit.do/QuiuiTool
    /// </remarks>
    public class QuiuiBase
    {
        #region + Essential steps ツ +
        /// <summary>
        /// URL of the website to be loaded.
        /// Change the URL value (essential step).
        /// </summary>
        private const string URL = "https://google.com/ncr";

        /// <summary>
        /// Path of the solution in your pc or in a team repository, etc.
        /// Change the solutionFolder value and rebuild solution (essential step).
        /// </summary>
        public static string solutionFolder = @"D:\Projects\Meowth\TestProject";

        /// <summary>
        /// Used by encrypt and decrypt methods.
        /// Change the hash value (essential step).
        /// </summary>
        public static string hash = "Al*$pio&7";

        /// <summary>
        /// Used to keep the test classes.
        /// Open the Test Explorer (Ctrl+E, T), group by class and then add the test classes (essential step).
        /// </summary>
        public static List<Type> testClasses = new List<Type>
        {
            typeof(TestClassA),
            typeof(TestClassB)
        };

        // Create one object of each test class (essential step).

        /// <summary>
        /// Contains its test methods.
        /// </summary>
        private static TestClassA testClassA = new TestClassA();

        /// <summary>
        /// Contains its test methods.
        /// </summary>
        private static TestClassB testClassB = new TestClassB();

        /// <summary>
        /// Calls a test method.
        /// A test is unknown when the call of the method has not yet been added.
        /// </summary>
        /// <param name="testName">Name of the test method.</param>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        /// <returns>null when the test passes; the exception thrown when the test fails, is inconclusive or is unknown.</returns>
        public static Exception CallTestMethod(string testName, string docxName, List<string> logs)
        {
            try
            {
                switch (testName)
                {
                    // Call all test methods (essential step).
                    case "FranceTest":
                        testClassA.FranceTest(docxName, logs);
                        break;
                    case "PanamaTest":
                        testClassA.PanamaTest(docxName, logs);
                        break;
                    case "PortugalTest":
                        testClassA.PortugalTest(docxName, logs);
                        break;
                    case "SpainTest":
                        testClassA.SpainTest(docxName, logs);
                        break;
                    case "SingaporeTest":
                        testClassA.SingaporeTest(docxName, logs);
                        break;
                    case "SenegalTest":
                        testClassA.SenegalTest(docxName, logs);
                        break;
                    case "Japan5minTest":
                        testClassA.Japan5minTest(docxName, logs);
                        break;
                    case "FrenchFriesTest":
                        testClassB.FrenchFriesTest(docxName, logs);
                        break;
                    case "FruitSaladTest":
                        testClassB.FruitSaladTest(docxName, logs);
                        break;
                    case "MashedPotatoesTest":
                        testClassB.MashedPotatoesTest(docxName, logs);
                        break;
                    case "PuddingTest":
                        testClassB.PuddingTest(docxName, logs);
                        break;
                    case "WatermelonTest":
                        testClassB.WatermelonTest(docxName, logs);
                        break;
                    case "ApplePieTest":
                        testClassB.ApplePieTest(docxName, logs);
                        break;
                    default:
                        return new Exception(QuiuiApp.TestUnknown);
                }
            }
            catch (Exception ex)
            {
                return ex;
            }

            return null;
        }
        #endregion

        #region Others Fields
        /// <summary>
        /// Used to add arguments in the list of arguments to be appended to the Chrome.exe command line.
        /// </summary>
        private static ChromeOptions options = new ChromeOptions();

        /// <summary>
        /// Used to keep the WebDriver instances while the run of the tests.
        /// </summary>
        private static Dictionary<string, IWebDriver> drivers = new Dictionary<string, IWebDriver>();
        #endregion

        #region QuiuiBase
        /// <summary>
        /// Adds the arguments in the list of arguments to be appended to the Chrome.exe command line.
        /// </summary>
        static QuiuiBase()
        {
            options.AddArgument("--start-maximized");
            options.AddArgument("--disable-infobars");
        }
        #endregion

        #region GetDriver
        /// <summary>
        /// Initializes a ChromeDriver instance and loads the website of the URL string.
        /// </summary>
        /// <param name="driverKey">Key of the ChromeDriver instance.</param>
        /// <returns>A ChromeDriver instance.</returns>
        public static IWebDriver GetDriver(string driverKey)
        {
            QuiuiApp.Token.ThrowIfCancellationRequested();

            ChromeDriverService ChromeService = ChromeDriverService.CreateDefaultService();
            ChromeService.HideCommandPromptWindow = true;

            QuiuiApp.Token.ThrowIfCancellationRequested();

            IWebDriver driver = new ChromeDriver(ChromeService, options);

            driver.Manage().Timeouts().PageLoad = TimeSpan.FromMilliseconds(25000);
            driver.Navigate().GoToUrl(URL);

            drivers.Add(driverKey, driver);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            return driver;
        }
        #endregion

        #region QuitDriver
        /// <summary>
        /// Quits a WebDriver instance contained in the drivers dictionary, closing every associated window.
        /// </summary>
        /// <param name="driverKey">Key of the ChromeDriver instance.</param>
        public static void QuitDriver(string driverKey)
        {
            if (drivers.ContainsKey(driverKey))
            {
                drivers[driverKey].Quit();
                drivers.Remove(driverKey);
            }
        }
        #endregion

        #region TakeScreenshot
        /// <summary>
        /// Takes a screenshot of a web page and adds the image in a Microsoft Word file.
        /// </summary>
        /// <param name="driver">IWebDriver intance.</param>
        /// <param name="docxName">Name of the Microsoft Word file.</param>
        public static void TakeScreenshot(IWebDriver driver, string docxName)
        {
            if (!Directory.Exists(QuiuiApp.ScreenshotsFolder))
            {
                Directory.CreateDirectory(QuiuiApp.ScreenshotsFolder);
            }

            string pngPath = QuiuiApp.ScreenshotsFolder + @"\" + docxName + ".png";

            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(pngPath, ScreenshotImageFormat.Png);

            Word.AddImage(pngPath, docxName);
        }
        #endregion
    }
}
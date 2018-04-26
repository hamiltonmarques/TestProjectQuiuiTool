using Hamilton.TestFramework.QuiuiTool;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Threading;
using Xceed.Words.NET;

namespace TestProject
{
    /// <summary>
    /// Generic test class for the QuiuiTool application.
    /// Notice: a test must contain the [TestMethod] attribute to load its name inside QuiuiTool application.
    /// Notice: call the method QuiuiApp.Token.ThrowIfCancellationRequested() inside the tests more than once, otherwise the cancellation of the run will take much longer.
    /// </summary>
    /// <remarks>
    /// The Microsoft Public License (Ms-PL) Copyright (c) 2018 Hamilton Marques.
    /// Meet Hamilton at http://bit.do/HamiltonLinkedin
    /// Videos about QuiuiTool at http://bit.do/QuiuiTool
    /// </remarks>
    [TestClass]
    public class TestClassA
    {
        /// <summary>
        /// Loads the Google web page, searches by France, takes a screenshot and adds it in the evidences file.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void FranceTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("France");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            logs.Add("[FranceTest]");
            logs.Add("Open google page.");
            logs.Add("Fill field search.");
            logs.Add("Enter pressed.");
            logs.Add("Validating links.");

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            QuiuiBase.TakeScreenshot(driver, docxName);

            Assert.IsTrue(true, "Some message.");

            QuiuiApp.Token.ThrowIfCancellationRequested();

            Word.AddText("France, in Western Europe, encompasses medieval cities, alpine villages and Mediterranean beaches.", Alignment.left, docxName);

            logs.Add("Test finished.");
        }

        /// <summary>
        /// Loads the Google web page, searches by Panama, takes a screenshot and adds it in the evidences file.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void PanamaTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Panama");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            QuiuiBase.TakeScreenshot(driver, docxName);

            Assert.IsTrue(true, "Some message.");
        }

        /// <summary>
        /// Loads the Google web page, searches by Portugal, takes a screenshot and adds it in the evidences file.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void PortugalTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Portugal");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            logs.Add("[PortugalTest]");
            logs.Add("Finding web element.");
            logs.Add("Element found.");
            logs.Add("Scroll up page.");
            logs.Add("Taking print.");
            logs.Add("Validating data base select.");
            logs.Add("Updating field search.");

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            QuiuiBase.TakeScreenshot(driver, docxName);

            Assert.AreEqual("Lisbon", "Porto", "Incorrect largest city name.");

            logs.Add("Test finished.");
        }

        /// <summary>
        /// Loads the Google web page and searches by Spain.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void SpainTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Spain");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            logs.Add("[SpainTest]");
            logs.Add("Loading service requests.");
            logs.Add("Fill field search with invalid characters.");
            logs.Add("Click on search button.");
            logs.Add("Select value 30 on Size select field.");
            logs.Add("Click on Send button.");
            logs.Add("Validating successfull addition message.");
            logs.Add("It is a very long log. I am writing something only to look scrolls on text field and to validate how is displayed. Let me see if it works with a very long log.");
            logs.Add("Copy value from comments field.");
            logs.Add("Checking backwards compatibility.");
            logs.Add("Select only Spain facts.");

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            Assert.IsTrue(false, "Page did not load castles architecture.");

            logs.Add("Test finished.");
        }

        /// <summary>
        /// Loads the Google web page, searches by Singapore, takes a screenshot and adds it in the evidences file.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void SingaporeTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Singapore");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            logs.Add("[SingaporeTest]");
            logs.Add("Initializing inconclusive test.");
            logs.Add("Field Adress filled successfully.");
            logs.Add("Searching by partial name.");
            logs.Add("Validating select items.");
            logs.Add("Loading page results.");
            logs.Add("Select menu Orders.");

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            QuiuiBase.TakeScreenshot(driver, docxName);

            Word.AddText("Singapore, an island city-state off southern Malaysia, is a global financial center with a tropical climate and multicultural population.", Alignment.center, docxName);

            Assert.Inconclusive("Singapore flag not found. It must be displayed.");

            logs.Add("Test finished.");
        }

        /// <summary>
        /// Loads the Google web page, searches by Senegal, takes a screenshot and adds it in the evidences file.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void SenegalTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Senegal");
            TxtSearch.Submit();

            QuiuiApp.Token.ThrowIfCancellationRequested();

            logs.Add("[SenegalTest]");
            logs.Add("Waiting page loading.");
            logs.Add("Settings are ready. Allow export file.");

            Thread.Sleep(3000);

            QuiuiApp.Token.ThrowIfCancellationRequested();

            QuiuiBase.TakeScreenshot(driver, docxName);

            Assert.Inconclusive("Account without access authorization.");

            logs.Add("Test finished.");
        }

        /// <summary>
        /// Loads the Google web page, searches by Japan and waits 5 minutes.
        /// </summary>
        /// <param name="docxName">Name of the evidences file.</param>
        /// <param name="logs">List to add the logs of the test.</param>
        [TestMethod]
        public void Japan5minTest(string docxName, List<string> logs)
        {
            IWebDriver driver = QuiuiBase.GetDriver(docxName);

            IWebElement TxtSearch = driver.FindElement(By.Name("q"));
            TxtSearch.SendKeys("Japan");
            TxtSearch.Submit();

            for (int i = 0; i < 60; i++)
            {
                QuiuiApp.Token.ThrowIfCancellationRequested();
                Thread.Sleep(5000);
            }
        }
    }
}
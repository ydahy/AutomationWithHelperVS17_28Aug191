using System;
using System.IO;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using Shouldly;
using Config;
using AutomationWithHelperVS17_28Aug191.Properties;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AutomationWithHelperVS17_28Aug191.Reporting;
using System.Text;

namespace AutomationWithHelperVS17_28Aug191
{
    [SetUpFixture]
    public class OneTimeSetupAndTearDown : Common
    {
        /// <summary>
        /// Assembly Initialize Runs One Time at the Start of Run/Debug Test Methods
        /// Deletes the Request Log Files and Issues if Found in the Project's Folder
        /// </summary>
        [OneTimeSetUp]
        public void AssemblyInitialize()
        {
            DeleteFilesAndScreenshots();

            GetEnvironmentVariablesForSettings();

            try
            {
                //To create report directory and add HTML report into it

                Directory.CreateDirectory(reportDir);
                var htmlReporter = new ExtentHtmlReporter(reportDir + "\\Automation_Report.html");
                htmlReporter.Config.EnableTimeline = true;
                htmlReporter.Config.DocumentTitle = typeof(Common).Namespace + " Document";
                htmlReporter.Config.ReportName = typeof(Common).Namespace + " Report";
                ReportingManager.extentInstance.AttachReporter(htmlReporter);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static void GetEnvironmentVariablesForSettings()
        {
            Settings.Default.Frontend_URL = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("Frontend_URL")) ? Environment.GetEnvironmentVariable("Frontend_URL") : Settings.Default.Frontend_URL;

            Settings.Default.Backend_URL = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("Backend_URL")) ? Environment.GetEnvironmentVariable("Backend_URL") : Settings.Default.Backend_URL;

            Settings.Default.SENDGRID_API_KEY = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("SENDGRID_API_KEY")) ? Environment.GetEnvironmentVariable("SENDGRID_API_KEY") : Settings.Default.SENDGRID_API_KEY;

            Settings.Default.FromMail = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("FromMail")) ? Environment.GetEnvironmentVariable("FromMail") : Settings.Default.FromMail;

            Settings.Default.timeToWaitInMinutes = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("timeToWaitInMinutes")) ? double.Parse(Environment.GetEnvironmentVariable("timeToWaitInMinutes")) : Settings.Default.timeToWaitInMinutes;

            Settings.Default.RunHeadless = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("RunHeadless")) ? bool.Parse(Environment.GetEnvironmentVariable("RunHeadless")) : Settings.Default.RunHeadless;
        }

        /// <summary>
        /// Assembly Cleanup Runs One Time at the End of Run/Debug Test Methods
        /// Sends Mail with Test Result based on Log Files Created while Running Code
        /// </summary>
        [OneTimeTearDown]
        public void AssemblyCleanup()
        {
            try
            {
                ReportingManager.extentInstance.Flush();
            }
            catch (Exception e)
            {
                throw e;
            }

            StringBuilder plainString = new StringBuilder("Dear,");
            StringBuilder htmlString = new StringBuilder("Dear,");

            plainString.Append(Environment.NewLine);
            plainString.Append(Environment.NewLine);

            htmlString.Append("<br></br>");
            htmlString.Append("<br></br>");

            plainString.Append("Kindly note that " + typeof(Common).Namespace + " Test Run is Done.");
            plainString.Append(Environment.NewLine);

            htmlString.Append("Kindly note that " + typeof(Common).Namespace + " Test Run is Done.");
            htmlString.Append("<br></br>");

            plainString.Append("Test Result is: " + (TestContext.CurrentContext.Result.Outcome == ResultState.Error || TestContext.CurrentContext.Result.Outcome == ResultState.Failure ? "Some Tests Failed" : "All Tests Passed") + ".");
            plainString.Append(Environment.NewLine);

            htmlString.Append("Test Result is: <b>" + (TestContext.CurrentContext.Result.Outcome == ResultState.Error || TestContext.CurrentContext.Result.Outcome == ResultState.Failure ? "Some Tests Failed" : "All Tests Passed") + "</b>.");
            htmlString.Append("<br></br>");

            plainString.Append("Please check attached files \"index.html\" and \"dashboard.html\" for all logs and attached Screenshots if any were taken.");

            htmlString.Append("Please check attached files \"index.html\" and \"dashboard.html\" for all logs and attached Screenshots if any were taken.");

            plainString.Append(Environment.NewLine);
            plainString.Append(Environment.NewLine);

            htmlString.Append("<br></br>");
            htmlString.Append("<br></br>");

            plainString.Append("Regards");

            htmlString.Append("Regards");

            SendMail(Settings.FromMail, "[Automated Mail] Automation Run Result", plainString.ToString(), htmlString.ToString(), Settings.Default.toMails, Settings.Default.SENDGRID_API_KEY).Wait();
        }
    }
}
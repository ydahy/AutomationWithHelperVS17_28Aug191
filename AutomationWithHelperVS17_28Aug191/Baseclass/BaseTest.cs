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

namespace AutomationWithHelperVS17_28Aug191
{
    public class BaseTest : Common
    {
        public TestContext TestContext { get; set; }

        /// <summary>
        /// Initialize Runs at the Start of Run/Debug of Each Test Method
        /// Opens New Driver and Initializes its Wait
        /// </summary>
        [SetUp]
        public void Setup()
        {
            try
            {
                if (GetCurrentMethod(2).Equals("InvokeMethod"))
                    _test = ReportingManager.extentInstance.CreateTest(TestContext.CurrentContext.Test.Name);
            }
            catch (Exception e)
            {
                throw e;
            }

            Driver = IsDriverOpen() ? Driver : Initialize(Browser.chrome, out Wait, timeToWaitInMinutes: Settings.timeToWaitInMinutes, runHeadless: Settings.RunHeadless);
        }

        /// <summary>
        /// Cleanup Runs at the End of Run/Debug of Each Test Method
        /// Closes Open Driver
        /// </summary>
        [TearDown]
        public void CleanUp()
        {
            try
            {
                var status = TestContext.CurrentContext.Result.Outcome.Status;
                var stacktrace = TestContext.CurrentContext.Result.StackTrace;
                var errorMessage = TestContext.CurrentContext.Result.Message;
                Status logstatus;
                switch (status)
                {
                    case TestStatus.Failed:
                        logstatus = Status.Fail;
                        string fileName = TakeScreenshot(TestContext.CurrentContext.Test.Name + ".png");
                        CloseDriver();
                        _test.Log(logstatus, "Test ended with " + logstatus + " - " + errorMessage);
                        _test.Log(logstatus, "Snapshot below: " + fileName);
                        break;
                    case TestStatus.Skipped:
                        CloseDriver();
                        logstatus = Status.Skip;
                        _test.Log(logstatus, "Test ended with " + logstatus);
                        break;
                    default:
                        CloseDriver();
                        logstatus = Status.Pass;
                        _test.Log(logstatus, "Test ended with " + logstatus);
                        break;
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            if (!string.IsNullOrEmpty(errors))
                LogIssue("", errors, false, true);
        }
    }
}
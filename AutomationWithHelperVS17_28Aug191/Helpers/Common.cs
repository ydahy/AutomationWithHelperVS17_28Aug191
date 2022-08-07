using System;
using System.Globalization;
using System.IO;
using AutomationWithHelperVS17_28Aug191.Properties;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AutomationWithHelperVS17_28Aug191.Reporting;

namespace AutomationWithHelperVS17_28Aug191
{
    public class Common : TestHelper
    {
        internal Settings Settings = Properties.Settings.Default;

        public enum LocalizationKeys
        {
            RequestSubmittedSuccessfullyMessage
        }

        public enum LogFiles
        {
            RequestsLogFile,
            IssuesLogFile,
            CommandsLogFile
        }

        public string LocalizedValueOf(Enum value)
        {
            return Resources.ResourceManager.GetString(value.ToString(), new CultureInfo(Settings.Language.ToString()));
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using AventStack.ExtentReports;

namespace AutomationWithHelperVS17_28Aug191.Reporting
{
    /// <summary>
    /// Creates a single instance of Extent Report 
    /// </summary>
    public class ReportingManager
    {
        /// <summary>
        /// Create new instance of Extent report
        /// </summary>
        private static readonly ExtentReports _instance = new ExtentReports();

        static ReportingManager() { }
        private ReportingManager() { }

        /// <summary>
        /// Property to return the instance of the report.
        /// </summary>
        /// <value>
        /// The instance.
        /// </value>
        public static ExtentReports extentInstance
        {
            get
            {
                return _instance;
            }
        }
    }
}

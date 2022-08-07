using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Xml;
using AutoIt;
using Microsoft.Win32;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using ExpectedConditions = SeleniumExtras.WaitHelpers.ExpectedConditions;
using AventStack.ExtentReports;
using SendGrid;
using SendGrid.Helpers.Mail;
using System.Threading.Tasks;

namespace AutomationWithHelperVS17_28Aug191
{
    /// <summary>
    /// Automation Test Helper including many functions to ease writing Selenium WebDriver Test Automation Code
    /// </summary>
    public class TestHelper : ITestHelper
    {
        /// <summary>
        /// Global Driver used throughout the TestHelper Functions
        /// </summary>
        protected IWebDriver Driver;
        /// <summary>
        /// Global Wait used throughout the TestHelper Functions
        /// </summary>
        protected WebDriverWait Wait;
        private int driverInstanceNumber = 0;
        /// <summary>
        /// Console Errors found in Driver
        /// </summary>
        protected string errors;
        /// <summary>
        /// Path to Report Created by ExtentReports Framework
        /// </summary>
        protected string reportDir = SetDir("Test_Execution_Reports");
        /// <summary>
        /// Extent Test Instance Used in Reporting with ExtentReports Framework
        /// </summary>
        protected static ExtentTest _test;

        private static Random _random = new Random();
        private static object _syncLock = new object();


        #region Getters and Setters

        /// <summary>
        /// Set Driver Instance to gived driver
        /// </summary>
        /// <param name="driver"></param>
        public virtual void SetDriver(IWebDriver driver)
        {
            Driver = driver;
        }

        /// <summary>
        /// Set Wait Instance to given wait
        /// </summary>
        /// <param name="wait"></param>
        public virtual void SetWait(WebDriverWait wait)
        {
            Wait = wait;
        }

        /// <summary>
        /// Get Driver Instance
        /// </summary>
        /// <returns>Driver Instance</returns>
        public virtual IWebDriver GetDriver()
        {
            return Driver;
        }

        /// <summary>
        /// Get Wait Instance
        /// </summary>
        /// <returns>Wait Instance</returns>
        public virtual WebDriverWait GetWait()
        {
            return Wait;
        }

        #endregion

        #region Wait Functions

        /// <summary>
        /// Wait For: Element to meet the given condition
        /// <br>Usage Example: WaitFor(By.Id(""), Condition.Visible);
        /// <br>               WaitFor(null, Condition.PageReadyState);</br></br>
        /// </summary>
        /// <param name="locator">By locator of the element</param>
        /// <param name="condition">Condition to be satisfied</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitFor(By locator, Condition condition, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(FrontendLogin)) && !GetCurrentMethod(level: 2).Contains(nameof(BackendLogin)))
                LogCommands("", GetCurrentMethod() + ((!condition.Equals(Condition.PageReadyState) && !condition.Equals(Condition.AlertIsPresent)) ? ((string.IsNullOrEmpty(locatorDescription) ? (" Element with Locator: " + locator.ToString()) : " " + locatorDescription) + " with") : "") + " Condition " + condition.ToString(), false, true);

            try
            {
                switch (condition)
                {
                    case Condition.Visible:
                        Wait.Until(ExpectedConditions.ElementIsVisible(locator));
                        break;
                    case Condition.Exist:
                        Wait.Until(ExpectedConditions.ElementExists(locator));
                        break;
                    case Condition.Invisible:
                        Wait.Until(ExpectedConditions.InvisibilityOfElementLocated(locator));
                        break;
                    case Condition.Clickable:
                        Wait.Until(ExpectedConditions.ElementToBeClickable(locator));
                        break;
                    case Condition.Enabled:
                        Wait.Until(d =>
                        {
                            try
                            {
                                return FindElement(locator).Enabled;
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                        });
                        break;
                    case Condition.FrameAvailabilityAndSwitchToIt:
                        Wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(locator));
                        break;
                    case Condition.NotExist:
                        Wait.Until(d => !IsElementPresent(locator));
                        break;
                    case Condition.SelectListLoaded:
                        Wait.Until(d => new SelectElement(FindElement(locator)).Options.Count > 1);
                        break;
                    case Condition.PageReadyState:
                        WaitForPageReadyState();
                        break;
                    case Condition.PresenceOfAllElementsLocatedBy:
                        Wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(locator));
                        break;
                    case Condition.VisibilityOfAllElementsLocatedBy:
                        Wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(locator));
                        break;
                    case Condition.ElementToBeSelected:
                        Wait.Until(ExpectedConditions.ElementToBeSelected(locator));
                        break;
                    case Condition.AlertIsPresent:
                        Wait.Until(ExpectedConditions.AlertIsPresent());
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("condition", condition, null);
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait For Element to meet the given condition
        /// <br>Usage Example: WaitFor(FindElement(By.Id("")), Condition.Visible);
        /// <br>               WaitFor(null, Condition.PageReadyState);</br></br>
        /// </summary>
        /// <param name="element">Element to Wait For</param>
        /// <param name="condition">Condition to be satisfied</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void WaitFor(IWebElement element, Condition condition, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + ((!condition.Equals(Condition.PageReadyState) && !condition.Equals(Condition.AlertIsPresent)) ? ((string.IsNullOrEmpty(elementDescription) ? (" Element: " + element.ToString()) : " " + elementDescription) + " with") : "") + " Condition " + condition.ToString(), false, true);

            try
            {
                switch (condition)
                {
                    case Condition.Visible:
                        Wait.Until(d =>
                        {
                            try
                            {
                                return element.Displayed;
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                        });
                        break;
                    case Condition.Invisible:
                        Wait.Until(d =>
                        {
                            try
                            {
                                return !element.Displayed;
                            }
                            catch (Exception)
                            {
                                return true;
                            }
                        });
                        break;
                    case Condition.Clickable:
                        Wait.Until(ExpectedConditions.ElementToBeClickable(element));
                        break;
                    case Condition.Enabled:
                        Wait.Until(d =>
                        {
                            try
                            {
                                return element.Enabled;
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                        });
                        break;
                    case Condition.FrameAvailabilityAndSwitchToIt:
                        Wait.Until(d =>
                        {
                            try
                            {
                                if (!element.Displayed) return false;
                                string id = element.GetAttribute("id");
                                string name = element.GetAttribute("name");
                                string frameId = string.IsNullOrEmpty(id) ? name : id;
                                SwitchToNestedFrame(frameId);
                                return true;
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                        });
                        break;
                    case Condition.NotExist:
                        Wait.Until(d =>
                        {
                            try
                            {
                                var txt = element.Text;
                                return false;
                            }
                            catch (Exception)
                            {
                                return true;
                            }
                        });
                        break;
                    case Condition.SelectListLoaded:
                        Wait.Until(d => new SelectElement(element).Options.Count > 1);
                        break;
                    case Condition.PageReadyState:
                        WaitForPageReadyState();
                        break;
                    case Condition.ElementToBeSelected:
                        Wait.Until(ExpectedConditions.ElementToBeSelected(element));
                        break;
                    case Condition.AlertIsPresent:
                        Wait.Until(ExpectedConditions.AlertIsPresent());
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("condition", condition, null);
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait for Page Ready State: Waits till the page stops loading using javascript executor to get the value of "readyState" property in the page
        /// </summary>
        public virtual void WaitForPageReadyState()
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitFor)) && !GetCurrentMethod(level: 2).Contains(nameof(UploadAttachments)) && !GetCurrentMethod(level: 2).Contains(nameof(FrontendLogin)) && !GetCurrentMethod(level: 2).Contains(nameof(BackendLogin)))
                LogCommands("", GetCurrentMethod(), false, true);

            SwitchToDefault();
            string documentReadyState;
            do
            {
                try
                {
                    documentReadyState = (string)ExecuteScript("return document.readyState");
                }
                catch (Exception)
                {
                    break;
                }
                Thread.Sleep(500);
            } while (!documentReadyState.Equals("complete") && !documentReadyState.Equals("interactive"));

            var iframes = FindElements(By.TagName("iframe"));

            for (var iframesCount = 0; iframesCount < iframes.Count; iframesCount++)
            {
                iframes = FindElements(By.TagName("iframe"));
                string frameId;

                try
                {
                    frameId = iframes[iframesCount].GetAttribute("id");
                }
                catch (ArgumentOutOfRangeException)
                {
                    break;
                }
                catch (StaleElementReferenceException)
                {
                    continue;
                }

                string frameReadyState;
                do
                {
                    try
                    {
                        frameReadyState =
                            (string)
                            ExecuteScript("return document.getElementById('" + frameId + "').contentDocument.readyState");
                    }
                    catch (Exception)
                    {
                        break;
                    }
                    Thread.Sleep(500);
                } while (!frameReadyState.Equals("complete") && !frameReadyState.Equals("interactive"));
            }
        }

        /// <summary>
        /// Wait till Element is Displayed: Uses explicit wait till the element is displayed
        /// </summary>
        /// <param name="element">Element that should be displayed</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitTillElementIsDisplayed(IWebElement element, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitFor)))
                LogCommands("", GetCurrentMethod() + " for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element: " + element.ToString()) : locatorDescription), false, true);

            try
            {
                Wait.Until(d =>
                {
                    try
                    {
                        return element.Displayed;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait for provided text to be visible in the page
        /// </summary>
        /// <param name="text">Text that should be visible</param>
        public virtual void WaitForTextToBeVisible(string text)
        {
            LogCommands("", "Check if " + GetCurrentMethod() + $" where text is: {text}", false, true);

            Wait.Until(d => IsTextVisible(text));
        }

        /// <summary>
        /// Wait for provided text to not be visible in the page
        /// </summary>
        /// <param name="text">Text that should not be visible</param>
        public virtual void WaitForTextToBeInvisible(string text)
        {
            LogCommands("", "Check if " + GetCurrentMethod() + $" where text is: {text}", false, true);

            Wait.Until(d => IsTextInvisible(text));
        }

        /// <summary>
        /// Wait for provided text to be visible in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the text</param>
        /// <param name="text">Text that should be visible</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitForTextToBeVisible(By locator, string text, string locatorDescription = "")
        {
            LogCommands("", "Check if " + GetCurrentMethod() + " for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" where text is: {text}", false, true);

            Wait.Until(d => IsTextVisible(locator, text));
        }

        /// <summary>
        /// Wait for provided text to not be visible in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should not contain the text</param>
        /// <param name="text">Text that should not be visible</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitForTextToBeInvisible(By locator, string text, string locatorDescription = "")
        {
            LogCommands("", "Check if " + GetCurrentMethod() + " for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" where text is: {text}", false, true);

            Wait.Until(d => IsTextInvisible(locator, text));
        }

        /// <summary>
        /// Wait till the given attribute appears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the attribute</param>
        /// <param name="attr">Attribute that should be found in the element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitTillAttributeAppears(By locator, Attribute attr, string locatorDescription = "")
        {
            LogCommands("", $"WaitTillAttribute \"{attr.ToString()}\" Appears for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Wait.Until(d =>
                {
                    try
                    {
                        FindElement(locator).GetAttribute(attr.ToString().Replace('_', '-'));
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait till the given attribute appears in the element
        /// </summary>
        /// <param name="element">Element that should contain the attribute</param>
        /// <param name="attr">Attribute that should be found in the element</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void WaitTillAttributeAppears(IWebElement element, Attribute attr, string elementDescription = "")
        {
            LogCommands("", $"WaitTillAttribute \"{attr.ToString()}\" Appears for " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                Wait.Until(d =>
                {
                    try
                    {
                        element.GetAttribute(attr.ToString().Replace('_', '-'));
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait till the value of the given attribute appears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the value of the attribute</param>
        /// <param name="attr">Attribute that should contain the value</param>
        /// <param name="val">Value that should appear in the Attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitTillValueOfAttributeAppears(By locator, Attribute attr, string val, string locatorDescription = "")
        {
            LogCommands("", $"WaitTillValue \"{val}\" OfAttribute \"{attr.ToString()}\" Appears for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Wait.Until(d =>
                {
                    try
                    {
                        var el = FindElement(locator).GetAttribute(attr.ToString().Replace('_', '-'));
                        return el.Contains(val);
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Wait till the value of the given attribute disappears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should not contain the value of the attribute</param>
        /// <param name="attr">Attribute that should not contain the value</param>
        /// <param name="val">Value that should disappear from the Attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void WaitTillValueOfAttributeDisappears(By locator, Attribute attr, string val, string locatorDescription = "")
        {
            LogCommands("", $"WaitTillValue \"{val}\" OfAttribute \"{attr.ToString()}\" Disppears for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Wait.Until(d =>
                {
                    try
                    {
                        var el = FindElement(locator).GetAttribute(attr.ToString().Replace('_', '-'));
                        return !el.Contains(val);
                    }
                    catch (Exception)
                    {
                        return true;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Thread Sleep for given milliseconds
        /// </summary>
        /// <param name="millisecondsTimeout">Milliseconds to sleep</param>
        public virtual void SleepFor(int millisecondsTimeout)
        {
            LogCommands("", GetCurrentMethod() + millisecondsTimeout, false, true);

            Thread.Sleep(millisecondsTimeout);
        }

        /// <summary>
        /// Get dom content loaded time in seconds at any given time
        /// </summary>
        /// <returns>dom content loaded time in seconds</returns>
        public virtual string GetDomContentLoadedTime()
        {
            WaitForPageReadyState();

            var domContentLoadedTimeInMilliSeconds = (long)ExecuteScript("try{window.performance = window.performance || window.mozPerformance || window.msPerformance || window.webkitPerformance || {};" +
                "return(parseInt(window.performance.timing.domContentLoadedEventEnd)-parseInt(window.performance.timing.navigationStart));}catch(e){alert(e);}");

            return (domContentLoadedTimeInMilliSeconds / 1000.0).ToString();
        }

        #endregion

        #region Finding Elements Functions

        /// <summary>
        /// Finds the first element using the given locator
        /// Usage Example: FindElement(By.Id(""));
        /// </summary>
        /// <param name="locator">By locator of the element</param>
        /// <returns>Element found</returns>
        public virtual IWebElement FindElement(By locator)
        {
            return Driver.FindElement(locator);
        }

        /// <summary>
        /// Finds all elements using the given locator
        /// Usage Example: FindElements(By.Id(""));
        /// </summary>
        /// <param name="locator">By locator of the elements</param>
        /// <returns>List of elements found</returns>
        public virtual IList<IWebElement> FindElements(By locator)
        {
            return Driver.FindElements(locator);
        }

        /// <summary>
        /// Finds the first element that is a link that is a child of the given element
        /// Usage Example: FindLinkInElement(FindElement(By.Id("")));
        /// </summary>
        /// <param name="element">Parent element that contains the link element as a child</param>
        /// <returns>Element of the link</returns>
        public virtual IWebElement FindLinkInElement(IWebElement element)
        {
            try
            {
                return element.FindElement(By.TagName("a"));
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Parent Element: gets the parent of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Child Element that the function should find its Parent</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Parent Element</returns>
        public virtual IWebElement GetParentElement(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                var parentElement = FindElement(locator).FindElement(By.XPath(".."));

                return parentElement;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Parent Element: gets the parent of the element sent as parameter
        /// </summary>
        /// <param name="element">Child Element that the function should find its Parent</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Parent Element</returns>
        public virtual IWebElement GetParentElement(IWebElement element, string elementDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(FindCertainParentOf)) && !GetCurrentMethod(level: 2).Contains(nameof(FindCertainParentWithAttribute)))
                LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                var parentElement = element.FindElement(By.XPath(".."));

                return parentElement;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Find Certain Parent Of Element: Loops through the Parents of the Given Element to Find element with the given tag name.
        /// Example: Can be used in workspace in requests table for example where the locator is the request id and we want to get the row containing the request (hence the default value of the tagName parameter is tr) to get the status related to this request from the same row and assert on it
        /// </summary>
        /// <param name="locator">Locator of the Element whose Parent is to be Found</param>
        /// <param name="tagName">Tag Name of Parent to be Found. Default "tr" (table row in html)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>Parent Containing the Given Element or Null in case no Parent with the provided tag name was Found after Checking All Parents Till the HTML Element</returns>
        public virtual IWebElement FindCertainParentOf(By locator, string tagName = "tr", string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" whose Tag Name is {tagName}", false, true);

            try
            {
                var found = false;
                var counter = 0;

                var elements = FindElements(locator);
                var element = FindElement(locator);

                foreach (var displayedElement in elements)
                {
                    if (!displayedElement.Displayed) continue;
                    element = displayedElement;
                    break;
                }

                while (!found && !element.TagName.Equals("html"))
                {
                    element = GetParentElement(element);
                    found = element.TagName.Equals(tagName);
                    counter++;
                }

                if (element.TagName.Equals("html") && !found)
                {
                    return null;
                }

                return element;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Find Certain Parent Of Element: Loops through the Parents of the Given Element to Find element with the given tag name
        /// </summary>
        /// <param name="element">Element whose Parent is to be Found</param>
        /// <param name="tagName">Tag Name of Parent to be Found. Default "tr"</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>Parent Containing the Given Element or Null in case no Parent with the provided tag name was Found after Checking 10 Parents of the Element</returns>
        public virtual IWebElement FindCertainParentOf(IWebElement element, string tagName = "tr", string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription) + $" whose Tag Name is {tagName}", false, true);

            try
            {
                var found = false;
                var counter = 0;

                while (!found && counter < 10)
                {
                    element = GetParentElement(element);
                    found = element.TagName.Equals(tagName);
                    counter++;
                }

                if (counter == 10 && found == false)
                {
                    return null;
                }

                return element;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Find a parent of an element that contains value of a specific attribute
        /// </summary>
        /// <param name="locator">Locator of the element whose parent will be located</param>
        /// <param name="attr">Attribute found in the parent that should be located</param>
        /// <param name="val">Value of the attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The located parent element with the attribute and value provided</returns>
        public virtual IWebElement FindCertainParentWithAttribute(By locator, string attr, string val, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" whose Attribute is {attr} and its Value is {val}", false, true);

            try
            {
                var element = FindElement(locator);

                while (!element.GetAttribute(attr).Contains(val))
                {
                    element = GetParentElement(element);
                }

                return element;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Find a parent of an element that contains value of a specific attribute
        /// </summary>
        /// <param name="element">Element whose parent will be located</param>
        /// <param name="attr">Attribute found in the parent that should be located</param>
        /// <param name="val">Value of the attribute</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The located parent element with the attribute and value provided</returns>
        public virtual IWebElement FindCertainParentWithAttribute(IWebElement element, string attr, string val, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription) + $" whose Attribute is {attr} and its Value is {val}", false, true);

            try
            {
                while (!element.GetAttribute(attr).Contains(val))
                {
                    element = GetParentElement(element);
                }

                return element;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Child Elements: gets the direct children of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Parent Element that the function should find its Children</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Child Elements</returns>
        public virtual IList<IWebElement> GetChildElements(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                IList<IWebElement> childElements = FindElement(locator).FindElements(By.XPath("*"));

                return childElements;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Child Elements: gets the children of the element sent as parameter
        /// </summary>
        /// <param name="element">Parent Element that the function should find its Children</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Child Elements</returns>
        public virtual IList<IWebElement> GetChildElements(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                IList<IWebElement> parentElement = element.FindElements(By.XPath("*"));

                return parentElement;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #endregion

        #region Interacting with Elements Functions

        /// <summary>
        /// Clears text in a given text field/area
        /// </summary>
        /// <param name="locator">By locator of the element whose text should be cleared</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void Clear(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Driver.FindElement(locator).Clear();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Clears text in a given text field/area
        /// </summary>
        /// <param name="element">Element whose text should be cleared</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void Clear(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                element.Clear();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Send Keys to a given text field/area
        /// </summary>
        /// <param name="locator">By locator of the element that the text should be sent to</param>
        /// <param name="text">Text that should be entered in the text field</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void SendKeys(By locator, string text, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{text}\" to " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Wait.Until(d => FindElement(locator).Displayed);
                Wait.Until(d => FindElement(locator).Enabled);
                FindElement(locator).SendKeys(text);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Send Keys to a given text field/area
        /// </summary>
        /// <param name="element">Element that the text should be sent to</param>
        /// <param name="text">Text that should be entered in the text field</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void SendKeys(IWebElement element, string text, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{text}\" to " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                Wait.Until(d => element.Displayed);
                Wait.Until(d => element.Enabled);
                element.SendKeys(text);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Text of: gets the text of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Element whose text will be returned</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Text of the Element</returns>
        public virtual string GetTextOf(By locator, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(SetDate)))
                LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                return FindElement(locator).Text;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get Text of: gets the text of the element sent as parameter
        /// </summary>
        /// <param name="element">Element whose text will be returned</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Text of the Element</returns>
        public virtual string GetTextOf(IWebElement element, string elementDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(SetDate)))
                LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                return element.Text;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Scroll to Element: scrolls to an element in the page (horizontal or vertical)
        /// <br>Automatically detects if there is a Custom Scroll Bar (mCustomScrollbar) in the page and uses the correct scroll function in this case</br>
        /// </summary>
        /// <param name="locator">Locator of Element that should be scrolled to</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void ScrollTo(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                var counter = 0;
                bool elementInView;

                do
                {
                    if (IsElementPresent(By.ClassName("mCustomScrollbar")))
                    {
                        ExecuteScript("$('.mCustomScrollbar').mCustomScrollbar(\"scrollTo\",arguments[0]);", FindElement(locator));
                    }
                    else
                    {
                        ExecuteScript("arguments[0].scrollIntoView();", FindElement(locator));
                    }

                    counter++;
                    elementInView = (bool)ExecuteScript(
                        "var elem = arguments[0],                 " +
                        "  box = elem.getBoundingClientRect(),    " +
                        "  cx = box.left + box.width / 2,         " +
                        "  cy = box.top + box.height / 2,         " +
                        "  e = document.elementFromPoint(cx, cy); " +
                        "for (; e; e = e.parentElement) {         " +
                        "  if (e === elem)                        " +
                        "    return true;                         " +
                        "}                                        " +
                        "return false;                            "
                        , FindElement(locator));
                } while (!elementInView && counter < 5);

                Thread.Sleep(1000);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Scroll to Element: scrolls to an element in the page (horizontal or vertical)
        /// <br>Automatically detects if there is a Custom Scroll Bar (mCustomScrollbar) in the page and uses the correct scroll function in this case</br>
        /// </summary>
        /// <param name="element">Element that should be scrolled to</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void ScrollTo(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                var counter = 0;
                bool elementInView;

                do
                {
                    if (IsElementPresent(By.ClassName("mCustomScrollbar")))
                    {
                        ExecuteScript("$('.mCustomScrollbar').mCustomScrollbar(\"scrollTo\",arguments[0]);", element);
                    }
                    else
                    {
                        ExecuteScript("arguments[0].scrollIntoView();", element);
                    }

                    counter++;
                    elementInView = (bool)ExecuteScript(
                        "var elem = arguments[0],                 " +
                        "  box = elem.getBoundingClientRect(),    " +
                        "  cx = box.left + box.width / 2,         " +
                        "  cy = box.top + box.height / 2,         " +
                        "  e = document.elementFromPoint(cx, cy); " +
                        "for (; e; e = e.parentElement) {         " +
                        "  if (e === elem)                        " +
                        "    return true;                         " +
                        "}                                        " +
                        "return false;                            "
                        , element);
                } while (!elementInView && counter < 5);

                Thread.Sleep(1000);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Click on Element: Tries to Click on an Element using its By Locator and waits if there is InvalidOperationException or StaleElementReferenceException
        /// Clicks on Element using javascript if usingJavascript is true and using .Click() Function if usingJavascript is false
        /// </summary>
        /// <param name="locator">By Locator of Element that needs to be Clicked</param>
        /// <param name="usingJavascript">Click on the button using Javascript or not</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void ClickOn(By locator, bool usingJavascript, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(SetDate)) && !GetCurrentMethod(level: 2).Contains(nameof(FrontendLogin)) && !GetCurrentMethod(level: 2).Contains(nameof(BackendLogin)))
                LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + (usingJavascript ? " using Javascript" : ""), false, true);

            try
            {
                var type = Driver.GetType();

                Wait.Until(d =>
                {
                    try
                    {
                        var element = FindElement(locator);
                        if (element != null && element.Enabled)
                        {
                            if (usingJavascript || type.Name.Equals("InternetExplorerDriver"))
                            {
                                ExecuteScript("arguments[0].click();", element);
                                return true;
                            }
                            if (element.Displayed)
                            {
                                element.Click();
                                return true;
                            }
                        }
                        return false;
                    }
                    catch (InvalidOperationException)
                    {
                        return false;
                    }
                    catch (StaleElementReferenceException)
                    {
                        return false;
                    }
                    catch (WebDriverException)
                    {
                        return true;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Click on Element: Tries to click on an element using the browser's Javascript or not depending on the boolean value
        /// and waits if there is InvalidOperationException or StaleElementReferenceException
        /// </summary>
        /// <param name="element">Element that needs to be Clicked</param>
        /// <param name="usingJavascript">Click on the button using Javascript or not</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void ClickOn(IWebElement element, bool usingJavascript, string elementDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(SetDate)))
                LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription) + (usingJavascript ? " using Javascript" : ""), false, true);

            try
            {
                var type = Driver.GetType();

                Wait.Until(d =>
                {
                    try
                    {
                        if (element != null && element.Enabled)
                        {
                            if (usingJavascript || type.Name.Equals("InternetExplorerDriver"))
                            {
                                ExecuteScript("arguments[0].click();", element);
                                return true;
                            }
                            if (element.Displayed)
                            {
                                element.Click();
                                return true;
                            }
                        }
                        return false;
                    }
                    catch (InvalidOperationException)
                    {
                        return false;
                    }
                    catch (StaleElementReferenceException)
                    {
                        return false;
                    }
                    catch (WebDriverException)
                    {
                        return true;
                    }
                });
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #region Select List Elements Interactions

        /// <summary>
        /// Select option from drop down list by value
        /// Usage Example: SelectByValue(By.Id("selectElementId"), "optionValue");
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="value">Value of the option's Value Attribute that we are searching for (e.g. <option value="value"></option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void SelectByValue(By locator, string value, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{value}\" from " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                new SelectElement(FindElement(locator)).SelectByValue(value);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Select option from drop down list by value
        /// Usage Example: SelectByValue(FindElement(By.Id("selectElementId")), "optionValue");
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="value">Value of the option's Value Attribute that we are searching for (e.g. <option value="value"></option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void SelectByValue(IWebElement element, string value, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{value}\" from " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                new SelectElement(element).SelectByValue(value);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Select option from drop down list by its text
        /// Usage Example: SelectByText(By.Id("selectElementId"), "optionText");
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="text">Value of the option's Text Attribute that we are searching for (e.g. <option>text</option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void SelectByText(By locator, string text, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{text}\" from " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                new SelectElement(FindElement(locator)).SelectByText(text);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Select option from drop down list by its text
        /// Usage Example: SelectByText(FindElement(By.Id("selectElementId")), "optionText");
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="text">Value of the option's Text Attribute that we are searching for (e.g. <option>text</option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void SelectByText(IWebElement element, string text, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{text}\" from " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                new SelectElement(element).SelectByText(text);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Select option from drop down list by its index
        /// Usage Example: SelectByIndex(By.Id("selectElementId"), 1);
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="index">Number of the option in the list that we are searching for (e.g. <option>option1</option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void SelectByIndex(By locator, int index, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{index}\" from " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                new SelectElement(FindElement(locator)).SelectByIndex(index);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Select option from drop down list by its index
        /// Usage Example: SelectByIndex(FindElement(By.Id("selectElementId")), 1);
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="index">Number of the option in the list that we are searching for (e.g. <option>option1</option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void SelectByIndex(IWebElement element, int index, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" \"{index}\" from " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                new SelectElement(element).SelectByIndex(index);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Gets the options of the select list
        /// Usage Example: GetSelectOptions(By.Id("selectElementId"));
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual IList<IWebElement> GetSelectOptions(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" of " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                return new SelectElement(FindElement(locator)).Options;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Gets the options of the select list
        /// Usage Example: GetSelectOptions(FindElement(By.Id("selectElementId")));
        /// </summary>
        /// <param name="element">The select element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual IList<IWebElement> GetSelectOptions(IWebElement element, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + $" of " + (string.IsNullOrEmpty(locatorDescription) ? ("Element: " + element.ToString()) : locatorDescription), false, true);

            try
            {
                return new SelectElement(element).Options;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #endregion

        #region Mouse Interactions with Elements

        /// <summary>
        /// Double Click on an element using its locator
        /// </summary>
        /// <param name="locator">Locator of the element to click on twice</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void DoubleClickOn(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Actions().MoveToElement(FindElement(locator)).DoubleClick().Build().Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Double Click on an element
        /// </summary>
        /// <param name="element">Element to click on twice</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void DoubleClickOn(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                Actions().MoveToElement(element).DoubleClick().Build().Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Hover on an element using its locator
        /// Note: Mouse will not move in the page but the action will be executed
        /// </summary>
        /// <param name="locator">Locator of element to hover on</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void HoverOn(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Actions().MoveToElement(FindElement(locator)).Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Hover on an element
        /// Note: Mouse will not move in the page but the action will be executed
        /// </summary>
        /// <param name="element">Element to hover on</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void HoverOn(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                Actions().MoveToElement(element).Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Right click on element using its locator
        /// </summary>
        /// <param name="locator">Locator of the element to context click on</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        public virtual void RightClickOn(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                Actions().MoveToElement(FindElement(locator)).ContextClick().Build().Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Right click on element
        /// </summary>
        /// <param name="element">Element to context click on</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        public virtual void RightClickOn(IWebElement element, string elementDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(elementDescription) ? ("Element: " + element.ToString()) : elementDescription), false, true);

            try
            {
                Actions().MoveToElement(element).ContextClick().Build().Perform();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Actions: Returns Actions Object for the provided _driver
        /// Can be used for Clicking, Double Clicking, Focus On Element by Moving to It, Send Keys to an Element and other Functions
        /// How to Use?
        ///     example:
        ///         Actions().MoveToElement(element).Click().Perform();
        /// </summary>
        /// <returns>Actions Object</returns>
        public virtual Actions Actions()
        {
            return new Actions(Driver);
        }

        #endregion

        #region Checks Done on Elements

        /// <summary>
        /// Is Element Visible: Checks whether an Element is Visible or Not
        /// </summary>
        /// <param name="locator">By Locator of Element that need to be checked</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case element is visible and False in case element is not visible</returns>
        public virtual bool IsElementVisible(By locator, string locatorDescription = "")
        {
            LogCommands("", "Check if " + GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                return FindElement(locator).Displayed;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Is Element Present: Checks whether an Element is Present or Not
        /// </summary>
        /// <param name="locator">By Locator of Element that needs to be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case element is present and False in case element is not present</returns>
        public virtual bool IsElementPresent(By locator, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(OpenRequestFromRequestsList)) && !GetCurrentMethod(level: 2).Contains(nameof(ScrollTo)) && !GetCurrentMethod(level: 2).Contains(nameof(WaitFor)))
                LogCommands("", "Check if " + GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription), false, true);

            try
            {
                FindElement(locator);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Is Text Visible: Checks whether some Text is Visible in the page or Not
        /// </summary>
        /// <param name="text">text that needs to be found</param>
        /// <returns>True in case text is visible and False in case text is not visible</returns>
        public virtual bool IsTextVisible(string text)
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForTextToBeVisible)))
                LogCommands("", "Check if " + GetCurrentMethod() + $" where text is: {text}", false, true);

            try
            {
                var element = FindElement(By.XPath($"//*[contains(text(), '{text}')]"));
                return element.Displayed;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Is Text Invisible: Checks whether the provided text is visible in the page or not
        /// </summary>
        /// <param name="text">Text that should be invisible</param>
        /// <returns>True in case text is invisible and False in case it is visible</returns>
        public virtual bool IsTextInvisible(string text)
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForTextToBeInvisible)))
                LogCommands("", "Check if " + GetCurrentMethod() + $" where text is: {text}", false, true);

            try
            {
                var element = FindElement(By.XPath($"//*[contains(text(), '{text}')]"));
                return !element.Displayed;
            }
            catch (Exception)
            {
                return true;
            }
        }

        /// <summary>
        /// Is Text Visible: Checks that the provided element's text contains the provided text
        /// </summary>
        /// <param name="locator">Locator that the text should be found in</param>
        /// <param name="text">Text that needs to be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case text is found in the element and False in case text is not found</returns>
        public virtual bool IsTextVisible(By locator, string text, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForTextToBeVisible)))
                LogCommands("", "Check if " + GetCurrentMethod() + " for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" where text is: {text}", false, true);

            try
            {
                var element = FindElement(locator);
                return element.Text.Contains(text);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Is Text Invisible: Checks that the provided element's text doesn't contain the provided text
        /// </summary>
        /// <param name="locator">Locator that the text should not be found in</param>
        /// <param name="text">Text that should not be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case text is not found in the element and False in case text is found</returns>
        public virtual bool IsTextInvisible(By locator, string text, string locatorDescription = "")
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForTextToBeInvisible)))
                LogCommands("", "Check if " + GetCurrentMethod() + " for " + (string.IsNullOrEmpty(locatorDescription) ? ("Element with Locator: " + locator.ToString()) : locatorDescription) + $" where text is: {text}", false, true);

            try
            {
                var element = FindElement(locator);
                return !element.Text.Contains(text);
            }
            catch (Exception)
            {
                return true;
            }
        }

        #endregion

        #endregion

        #region Send Scripts to Browser Functions

        /// <summary>
        /// Scripts: Returns IJavaScriptExecutor Object for the provided _driver
        /// Can be used for Clicking, Scrolling or Executing any Javascript Code in the Browser
        /// How to Use?
        ///     example:
        ///         ExecuteScript("arguments[0].click();", element);
        /// </summary>
        /// <returns>IJavaScriptExecutor Object</returns>
        public virtual IJavaScriptExecutor Scripts()
        {
            return (IJavaScriptExecutor)Driver;
        }

        /// <summary>
        /// Executes Javascript in the current page
        /// Usage Example: ExecuteScript("arguments[0].scrollIntoView();", FindElement(By.Id("elementToScrollToId")));
        /// </summary>
        /// <param name="script">Script to be executed in the page</param>
        /// <param name="args">Parameters used in the script</param>
        /// <returns></returns>
        public virtual object ExecuteScript(string script, params object[] args)
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForPageReadyState)) && !GetCurrentMethod(level: 2).Contains(nameof(ClickOn)) && !GetCurrentMethod(level: 2).Contains(nameof(ScrollTo)) && !GetCurrentMethod(level: 2).Contains(nameof(OpenNewTab)))
                LogCommands("", GetCurrentMethod() + $" \"{script}\"", false, true);

            return Scripts().ExecuteScript(script, args);
        }

        #endregion

        #region Functions Implementing Specific Scenarios

        /// <summary>
        /// Click on Multiple Elements : Works on Navigating to my services, my tasks, my pending tasks or any threads of links
        /// Usage Example: ClickOnMultipleLinks(new By[] {By.Id("link1Id"), By.Id("link2Id")});
        /// </summary>
        /// <param name="links">links[]: Array of Links Text that will be clicked by "By.LinkText"</param>
        public virtual void ClickOnMultipleElements(By[] links)
        {
            try
            {
                foreach (var linkLocator in links)
                {
                    WaitFor(linkLocator, Condition.Visible);

                    ScrollTo(FindElement(linkLocator));

                    ClickOn(linkLocator, false);

                    WaitForPageReadyState();
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Clicks on sign in button/link if available (start with username locator if no sign in button available),
        /// enters the username in the first loginLocator,
        /// enters the password in the second loginLocator and clicks on the third loginLocator
        /// Usage Example:
        ///     FrontendLogin("username", "password", new By[] { By.Id("usernameId"), By.Id("passwordId"), By.Id("loginButtonId")}); >> Basic 3 locators that should be available
        /// or  FrontendLogin("username", "password", new By[] { By.Id("signInButtonIdToGoToLoginForm"), By.Id("usernameId"), By.Id("passwordId"), By.Id("loginButtonId")});
        /// </summary>
        /// <param name="username">Username used in Login</param>
        /// <param name="password">Password of the User</param>
        /// <param name="loginLocators">List of By Locators ordered as follows: 1. sign in button By locator if available, 2. username field By locator, 3. password field By locator, 4. login button By locator</param>
        public virtual void FrontendLogin(string username, string password, By[] loginLocators)
        {
            LogCommands("", $"Frontend Login with username \"{username}\" and password \"{password}\"", false, true);

            try
            {
                Driver.Manage().Window.Maximize();

                var index = 0;

                WaitFor(loginLocators[index], Condition.Visible);

                if (loginLocators.Length > 3)
                {
                    while (index < loginLocators.Length - 3)
                    {
                        ClickOn(loginLocators[index], false);

                        WaitForPageReadyState();

                        index++;
                    }
                }

                SendKeys(loginLocators[index], username, "Username Field");
                index++;
                SendKeys(loginLocators[index], password, "Password Field");
                index++;

                ClickOn(loginLocators[index], false, "Login Button");

                WaitForPageReadyState();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Backend Login: Login to any Backend Portal that requires Windows Authentication
        /// Works with IE and Chrome
        /// </summary>
        /// <param name="backendUrl">Url to go to</param>
        /// <param name="domain">Domain of the User to be sent to Windows Authentication Popup</param>
        /// <param name="username">Username to be sent to Windows Authentication Popup</param>
        /// <param name="password">Password to be sent to Windows Authentication Popup</param>
        public virtual void BackendLogin(string backendUrl, string domain, string username, string password)
        {
            try
            {
                SwitchToWindow(Driver.CurrentWindowHandle);

                try
                {
                    GoToUrl(backendUrl.Trim());
                }
                catch (Exception)
                {
                    // ignored
                }

                var usernameWithDomain = string.IsNullOrEmpty(domain) ? username : (domain + "\\" + username);
                Thread.Sleep(5000);
                SendBackendLoginCredentials(usernameWithDomain, password);

                Thread.Sleep(5000);

                MaximizeWindow();

                WaitForPageReadyState();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        private static void SendBackendLoginCredentials(string usernameWithDomain, string password)
        {
            lock (_syncLock)
            {
                LogCommands("", GetCurrentMethod() + $" username \"{usernameWithDomain}\" and password \"{password}\"", false, true);

                AutoItX.Send(usernameWithDomain + "{TAB}");
                AutoItX.Send(password + "{TAB}{Enter}");
            }
        }

        /// <summary>
        /// Wait Till Page Loads: Waits till the loader element disappears.
        /// Default Number of Times to Wait is 2 because sometimes the Loader appears more than once before the page is fully loaded
        /// </summary>
        /// <param name="loaderLocator">By Locator for the Loader Element</param>
        /// /// <param name="numberOfWaits">Number of Times to Wait for the Loader to Disappear</param>
        public virtual void WaitTillPageLoad(By loaderLocator, int numberOfWaits = 2)
        {
            try
            {
                int counter = 0;
                while (counter < numberOfWaits)
                {
                    WaitFor(loaderLocator, Condition.Invisible, "Loader");
                    Thread.Sleep(2000);
                    counter++;
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Finds the Request ID in the given container locator
        /// Usage Example: 
        ///     <example id="requestIdLocatorId">Request Submitted Successfully. Request ID is Req-18-0011.</example>
        ///     GetRequestId(By.Id("requestIdLocatorId"), "is", "."); // should return Req-18-0011
        /// 
        ///     <example id="requestIdLocatorId">Req-18-0011 was submitted successfully</example>
        ///     GetRequestId(By.Id("requestIdLocatorId"), "", "was"); // should return Req-18-0011
        /// 
        ///     <example id="requestIdLocatorId">Request Submitted Successfully. Request ID is Req-18-0011</example>
        ///     GetRequestId(By.Id("requestIdLocatorId"), "is", ""); // should return Req-18-0011
        /// </summary>
        /// <param name="requestIdLocator">By locator of the element whose text contains the Request ID</param>
        /// <param name="separatorBefore">String found before the Request ID that should be retrieved. Leave Empty ("") if the Request ID is the first thing in the text</param>
        /// <param name="separatorAfter">String found after the Request ID that should be retrieved. Leave Empty ("") if the Request ID is the last thing in the text</param>
        /// <returns>Text found between the 2 provided separators with Spaces Trimmed</returns>
        public virtual string GetRequestId(By requestIdLocator, string separatorBefore, string separatorAfter)
        {
            try
            {
                WaitFor(requestIdLocator, Condition.Visible);
                var requestIdContainerText = GetTextOf(requestIdLocator);
                var extractedRequestId = GetRequestId(requestIdContainerText, separatorBefore, separatorAfter);
                return extractedRequestId;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Finds the Request ID in the given container locator
        /// Usage Example: 
        ///     GetRequestId("Request Submitted Successfully. Request ID is Req-18-0011.", "is", "."); // should return Req-18-0011
        /// 
        ///     GetRequestId("Req-18-0011 was submitted successfully", "", "was"); // should return Req-18-0011
        /// 
        ///     GetRequestId("Request Submitted Successfully. Request ID is Req-18-0011", "is", ""); // should return Req-18-0011
        /// </summary>
        /// <param name="requestIdContainerText">Text that contains the Request ID</param>
        /// <param name="separatorBefore">String found before the Request ID that should be retrieved. Leave Empty ("") if the Request ID is the first thing in the text</param>
        /// <param name="separatorAfter">String found after the Request ID that should be retrieved. Leave Empty ("") if the Request ID is the last thing in the text</param>
        /// <returns>Text found between the 2 provided separators with Spaces Trimmed</returns>
        public virtual string GetRequestId(string requestIdContainerText, string separatorBefore, string separatorAfter)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                string extractedRequestId;
                if (string.IsNullOrEmpty(separatorBefore) && string.IsNullOrEmpty(separatorAfter))
                {
                    extractedRequestId = requestIdContainerText;
                }
                else if (string.IsNullOrEmpty(separatorBefore))
                {
                    extractedRequestId =
                        requestIdContainerText.Split(new[] { separatorAfter }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                }
                else if (string.IsNullOrEmpty(separatorAfter))
                {
                    extractedRequestId =
                        requestIdContainerText.Split(new[] { separatorBefore }, StringSplitOptions.RemoveEmptyEntries)[1].Trim
                            ();
                }
                else
                {
                    extractedRequestId =
                        requestIdContainerText.Split(new[] { separatorBefore }, StringSplitOptions.RemoveEmptyEntries)[1]
                            .Split(
                                new[] { separatorAfter }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                }

                LogInfo("", "Request Number Found is: " + extractedRequestId, false, true);

                return extractedRequestId;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Check Status: Finds the status that needs to be checked using the By locator and compares it to the given expected status
        /// </summary>
        /// <param name="appStatusLocator">By Locator of the Status Element in the Page</param>
        /// <param name="expectedStatus">String of the Status that is Expected to Appear</param>
        public virtual bool CheckStatus(By appStatusLocator, string expectedStatus)
        {
            try
            {
                WaitFor(appStatusLocator, Condition.Visible);
                var visibleStatus = GetTextOf(appStatusLocator);
                var statusIsCorrect = visibleStatus.Trim().ToLower().Contains(expectedStatus.Trim().ToLower());
                LogInfo("",
                    (statusIsCorrect ? "Correct Request Status: " : "Wrong Request Status: Expected Status is \"" + expectedStatus + "\" and ") + "Visible Status is \"" + visibleStatus + "\"", false, true);
                if (!statusIsCorrect)
                    LogIssue("", "Wrong Request Status: Expected Status is \"" + expectedStatus + "\" but " + "Visible Status is \"" + visibleStatus + "\"", false, true);
                return statusIsCorrect;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Check Status: Finds the status that needs to be checked using the given element and compares it to the given expected status
        /// </summary>
        /// <param name="appStatusElement">Status Element in the Page</param>
        /// <param name="expectedStatus">String of the Status that is Expected to Appear</param>
        public virtual bool CheckStatus(IWebElement appStatusElement, string expectedStatus)
        {
            try
            {
                Wait.Until(d => appStatusElement.Displayed);
                var visibleStatus = GetTextOf(appStatusElement);
                var statusIsCorrect = visibleStatus.Trim().ToLower().Contains(expectedStatus.Trim().ToLower());
                LogInfo("", (statusIsCorrect ? "Correct Request Status: " : "Wrong Request Status: Expected Status is \"" + expectedStatus + "\" and ") + "Visible Status is \"" + visibleStatus + "\"", false, true);
                return statusIsCorrect;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Upload Attachments: Finds all the Attachments using the given common part of the attribute value and uploads attachments to them
        /// <br>Note 1: May need to be overridden to add waits and checks that the file is uploaded</br>
        /// <br>Note 2: Function assumes that the uploader doesn't need to click on a button after selecting the file for it to be uploaded</br>
        /// </summary>
        /// <param name="commonFileLocator">Locator of the Common/Repeated Part of the Files</param>
        /// <param name="fileType">File Type to be Uploaded</param>
        public virtual void UploadAttachments(By commonFileLocator, FileType fileType)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                var pdfAttachmentPath = SetDir("pdf.pdf");
                var imageAttachmentPath = SetDir("jpg.jpg");
                var documentAttachmentPath = SetDir("docx.docx");
                var textAttachmentPath = SetDir("txt.txt");

                string documentPath;

                switch (fileType)
                {
                    case FileType.Pdf:
                        documentPath = pdfAttachmentPath;
                        break;
                    case FileType.Docx:
                        documentPath = documentAttachmentPath;
                        break;
                    case FileType.Jpg:
                        documentPath = imageAttachmentPath;
                        break;
                    case FileType.Txt:
                        documentPath = textAttachmentPath;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("fileType", fileType, null);
                }

                var fileUploadersList = FindElements(commonFileLocator);

                foreach (var fileUpload in fileUploadersList)
                {
                    fileUpload.SendKeys(documentPath);

                    WaitForPageReadyState();
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Open Request from My Tasks Or My Requests With Search control: Works on opening the request or tasks
        /// from the grid of tasks using the search control where the Request ID is a Link
        /// </summary>
        /// <param name="requestId">The specific request ID that you want to open</param>
        /// <param name="statusIndexinGrid">Index of the cell in the grid row that contains the status that will be checked</param>
        /// <param name="requestTextBoxLocator">By Locator of the request ID textbox that is found in the Search Control</param>
        /// <param name="searchButtonLocator">By Locator of the search button that is found in the Search Control</param>
        /// <param name="expectedStatus">The expected status of the request in the Tasks Grid</param>
        public virtual void OpenRequestFromRequestsList(string requestId, int statusIndexinGrid, By requestTextBoxLocator, By searchButtonLocator, string expectedStatus)
        {
            try
            {
                do
                {
                    SearchRequest(requestId, requestTextBoxLocator, searchButtonLocator);
                    Thread.Sleep(1000);
                } while (!IsElementPresent(By.LinkText(requestId), requestId + " Link"));

                WaitFor(By.LinkText(requestId), Condition.Visible, requestId + " Link");

                CheckStatusInGrid(requestId, statusIndexinGrid, expectedStatus);

                ClickOn(By.LinkText(requestId), false, requestId + " Link");
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Take Action: Takes action from a selection list using action value
        /// And enters comment in given comment text box locator if available else send it as null
        /// </summary>
        /// <param name="actionValue">Value of the Action that should be taken</param>
        /// <param name="actionControlLocator">Locator of the action selection list control</param>
        /// <param name="commentTextboxLocator">Locator of the comment text box</param>
        /// <param name="saveButtonLocator">Locator of the save button to be clicked after selecting action</param>
        public virtual void TakeAction(string actionValue, By actionControlLocator, By commentTextboxLocator, By saveButtonLocator)
        {
            try
            {
                WaitFor(actionControlLocator, Condition.Visible, "Action Control");
                try
                {
                    SendKeys(commentTextboxLocator, "Action: " + actionValue, "Comment Text Box");
                }
                catch (Exception)
                {
                    // ignored
                }

                SelectByValue(actionControlLocator, actionValue, "Action Control");

                ClickOn(saveButtonLocator, false, "Save/Submit Button");
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Search Request: Enters the Request ID to the text box of request ID in the search control and Clicks on the Search button
        /// </summary>
        /// <param name="requestId">The specific request ID that you want to search for</param>
        /// <param name="requestTextBoxLocator">By Locator of the request ID textbox that is found in the Search Control</param>
        /// <param name="searchButtonLocator">By Locator of the search button that is found in the Search Control</param>
        public virtual void SearchRequest(string requestId, By requestTextBoxLocator, By searchButtonLocator)
        {
            try
            {
                WaitFor(requestTextBoxLocator, Condition.Visible, "Request ID Text Box");
                Clear(requestTextBoxLocator, "Request ID Text Box");
                SendKeys(requestTextBoxLocator, requestId, "Request ID Text Box");
                WaitFor(searchButtonLocator, Condition.Visible, "Search/Filter Button");
                SendKeys(searchButtonLocator, Keys.Enter, "Search/Filter Button");
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Check Status in Grid: Check given Status of a given Request ID in a Grid where there Request ID is a Link
        /// Concern: to be used after searching for request in Grid
        /// </summary>
        /// <param name="requestId">Request ID whose Status should be checked</param>
        /// <param name="statusIndexinGrid">Index of the cell in the grid row that contains the status that will be checked</param>
        /// <param name="expectedRequestStatus">Expected Request Status that should be compared to the found Status</param>
        public virtual bool CheckStatusInGrid(string requestId, int statusIndexinGrid, string expectedRequestStatus)
        {
            try
            {
                WaitFor(By.PartialLinkText(requestId), Condition.Visible, requestId + " Link");
                ScrollTo(By.PartialLinkText(requestId), requestId + " Link");
                var reqRow = FindCertainParentOf(By.PartialLinkText(requestId), locatorDescription: requestId + " Link");
                var tdElements = reqRow.FindElements(By.TagName("td"));

                var tableCellofStatus = tdElements.ElementAt(statusIndexinGrid);
                var status = GetTextOf(tableCellofStatus, "Status Cell in Grid");
                var statusIsCorrect = status.ToLower().Trim().Equals(expectedRequestStatus.ToLower().Trim());
                LogInfo(requestId, (statusIsCorrect ? "Correct Status in Grid: " : "Wrong Status in Grid: Expected Status is \"" + expectedRequestStatus + "\" and ") + "Visible Status is \"" + status + "\"", false, false);

                return statusIsCorrect;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Set Date: Takes the Date that should be selected and the elements of the Calendar Popup and selects the date
        /// </summary>
        /// <param name="date">Date that should be selected</param>
        /// <param name="calenderButton">button element to be clicked to open the calendar popup</param>
        /// <param name="calendarPopup">Calendar Popup Element</param>
        /// <param name="previousMonth">Previous Month button Element</param>
        /// <param name="nextMonth">Next Month button Element</param>
        /// <param name="dateFieldDescription">Description of the Date Element for Logging</param>
        public virtual void SetDate(DateTime date, By calenderButton, By calendarPopup, By previousMonth, By nextMonth, string dateFieldDescription = "")
        {
            try
            {
                ClickOn(calenderButton, false);

                var noOfClicks = ((date.Year - DateTime.Now.Year) * 12) + (date.Month - DateTime.Now.Month);

                for (var i = 0; i < Math.Abs(noOfClicks); i++)
                {
                    var locator = noOfClicks > 0 ? nextMonth : previousMonth;
                    ClickOn(locator, false);
                }

                Thread.Sleep(1000);

                var calendarTable = FindElement(calendarPopup);
                var calendarTbody = calendarTable.FindElement(By.TagName("tbody"));
                var rows = calendarTbody.FindElements(By.TagName("tr"));
                var clicked = false;

                for (var rowsCount = 0; rowsCount < rows.Count; rowsCount++)
                {
                    var cells = rows[rowsCount].FindElements(By.TagName("td"));

                    for (var cellsCount = 0; cellsCount < cells.Count; cellsCount++)
                    {
                        var dayText = date.Day.ToString();
                        if (GetTextOf(cells[cellsCount]).Equals(dayText))
                        {
                            ClickOn(cells[cellsCount], false);
                            clicked = true;
                            break;
                        }

                        cells = rows[rowsCount].FindElements(By.TagName("td"));
                    }

                    if (clicked)
                    {
                        break;
                    }

                    calendarTable = FindElement(calendarPopup);
                    calendarTbody = calendarTable.FindElement(By.TagName("tbody"));
                    rows = calendarTbody.FindElements(By.TagName("tr"));
                }

                LogCommands("", GetCurrentMethod() + " of " + (string.IsNullOrEmpty(dateFieldDescription) ? ("Element with Locator: " + calenderButton.ToString()) : dateFieldDescription) + " to " + date.ToString("dd/MM/yyyy"), false, true);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #endregion

        #region CssSelector and XPath Helper Functions

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the end of the attribute
        /// Usage Example: FindElement(Suffix(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Suffix(Attribute attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.SuffixOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the start of the attribute
        /// Usage Example: FindElement(Prefix(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Prefix(Attribute attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.PrefixOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be any part of the attribute
        /// Usage Example: FindElement(Part(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Part(Attribute attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.PartOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the full value of the attribute
        /// Usage Example: FindElement(Full(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Full(Attribute attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.FullName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the end of the attribute
        /// Usage Example: FindElement(Suffix("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Suffix(string attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.SuffixOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the start of the attribute
        /// Usage Example: FindElement(Prefix("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Prefix(string attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.PrefixOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be any part of the attribute
        /// Usage Example: FindElement(Part("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Part(string attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.PartOfName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the full value of the attribute
        /// Usage Example: FindElement(Full("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By Full(string attr, string val)
        {
            try
            {
                return CssSelectorMaker(attr, AttributeValueType.FullName, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find a link element whose title attribute contains certain text using the XPath Function
        /// Usage Example: FindElement(LinkTitleContains("linkTitle"));
        /// </summary>
        /// <param name="val">Value of the TItle Attribute that we are searching for (e.g. title="...AttributeValue...")</param>
        /// <returns>The By XPath Locator</returns>
        public virtual By LinkTitleContains(string val)
        {
            try
            {
                return XPathMaker(HtmlTag.a, Attribute.title, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Builds the String used to find a span element whose text attribute contains certain text using the XPath Function
        /// Usage Example: FindElement(SpanTextContains("text"));
        /// </summary>
        /// <param name="val">Value of the Title Attribute that we are searching for (e.g. <span>...AttributeValue...</span>)</param>
        /// <returns>The By XPath Locator</returns>
        public virtual By SpanTextContains(string val)
        {
            try
            {
                return XPathMaker(HtmlTag.span, Attribute.text, val);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// XPath Maker: Builds the String used to find an Element using the XPath Function and returns By xpath locator
        /// </summary>
        /// <param name="htmlTag">HTML Tags  used to find the Attribute within, coming from the enum HTMLTag (e.g. <input> </input>)</param>
        /// <param name="attribute">Attribute to be found, coming from the enum Attribute (e.g. id = "")</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By XPath Locator</returns>
        public virtual By XPathMaker(HtmlTag htmlTag, Attribute attribute, string attributeValue)
        {
            try
            {
                // Example: "//HTMLTag[contains(@Attribute, 'AttributeValue')]"
                var elementBuilder = "//";
                elementBuilder += htmlTag.ToString().Replace('_', '-').ToLower().Trim();
                elementBuilder += "[contains(" + (!attribute.Equals(Attribute.text) ? "@" : "");
                elementBuilder += attribute.ToString().Replace('_', '-').ToLower().Trim() +
                                  (attribute.Equals(Attribute.text) ? "()" : "");
                elementBuilder += ", '";
                elementBuilder += attributeValue;
                elementBuilder += "')]";

                return By.XPath(elementBuilder);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// XPath Maker: Builds the String used to find an Element using the XPath Function and returns By xpath locator
        /// This is an Overload Function to be used in case of composite Tags or Attributes
        /// </summary>
        /// <param name="htmlTag">HTML Tags  used to find the Attribute within, entered as String (e.g. <input> </input>)</param>
        /// <param name="attribute">Attribute to be found, entered as String (e.g. data-bind = "")</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By XPath Locator</returns>
        public virtual By XPathMaker(string htmlTag, string attribute, string attributeValue)
        {
            try
            {
                // Example: "//HTMLTag[contains(@Attribute, 'AttributeValue')]"
                // Example: "//input[contains(@id, 'AttributeValue')]"
                // Example: "//input[contains(text(), 'AttributeValue')]"
                var elementBuilder = "//";
                elementBuilder += htmlTag.ToLower().Trim();
                elementBuilder += "[contains(" + (!attribute.ToLower().Contains("text") ? "@" : "");
                elementBuilder += attribute.ToLower().Trim();
                if (attribute.ToLower().Contains("text") && !attribute.ToLower().Contains("()"))
                    elementBuilder += "()";
                elementBuilder += ", '";
                elementBuilder += attributeValue;
                elementBuilder += "')]";

                return By.XPath(elementBuilder);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// CssSelector Maker: Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// </summary>
        /// <param name="attribute">Attribute to be found, coming from the enum Attribute (e.g. id = "")</param>
        /// <param name="attributeValueType">Type of the attribute value to be found, coming from the enum AttributeValueType (e.g. could be full, prefix. suffix or part of the value)</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By CssSelectorMaker(Attribute attribute, AttributeValueType attributeValueType, string attributeValue)
        {
            try
            {
                string attributeNameTypeString;

                switch (attributeValueType)
                {
                    case AttributeValueType.FullName:
                        attributeNameTypeString = "";
                        break;
                    case AttributeValueType.PrefixOfName:
                        attributeNameTypeString = "^";
                        break;
                    case AttributeValueType.SuffixOfName:
                        attributeNameTypeString = "$";
                        break;
                    case AttributeValueType.PartOfName:
                        attributeNameTypeString = "*";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("attributeValueType", attributeValueType, null);
                }
                return
                    By.CssSelector("[" + attribute.ToString().Replace('_', '-').ToLower().Trim() + attributeNameTypeString + "='" + attributeValue + "']");
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// CssSelector Maker: Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// This is an Overload Function to be used in case of composite Attributes
        /// </summary>
        /// <param name="attribute">Attribute to be found, entered as String (e.g. data-bind = "")</param>
        /// <param name="attributeValueType">Type of the attribute value to be found, coming from the enum AttributeValueType (e.g. could be full, prefix. suffix or part of the value)</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        public virtual By CssSelectorMaker(string attribute, AttributeValueType attributeValueType, string attributeValue)
        {
            try
            {
                string attributeValueTypeString;

                switch (attributeValueType)
                {
                    case AttributeValueType.FullName:
                        attributeValueTypeString = "";
                        break;
                    case AttributeValueType.PrefixOfName:
                        attributeValueTypeString = "^";
                        break;
                    case AttributeValueType.SuffixOfName:
                        attributeValueTypeString = "$";
                        break;
                    case AttributeValueType.PartOfName:
                        attributeValueTypeString = "*";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("attributeValueType", attributeValueType, null);
                }
                return
                    By.CssSelector("[" + attribute.ToLower().Trim() + attributeValueTypeString + "='" + attributeValue + "']");
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #endregion

        #region Driver, URL, Alerts and Tabs/Windows Handling Functions

        /// <summary>
        /// Initialize: Open New Browser of Given Type and Return the Driver and the Explicit Wait
        /// <br>Chrome Driver is Opened with Multiple Chrome Options</br>
        /// <br>Firefox Driver is Opened with the Default Profile</br>
        /// <br>Note: If the RunHeadless Setting is set to "True", the browser will not open and everything will be running in background (works with Chrome and Firefox only) but it will not run correctly if there is need to use AutoIt (if there is windows authentication or uploading through windows explorer)</br>
        /// <br>Example: Use Profile can be required in case the pages take too much time loading and caching and history are better kept to minimize the load time or can also be required in case there is login with otp to keep the user logged in even when the browser is closed instead of always using a clean session and having to enter the otp each time</br>
        /// <br>    Note for using the profile: Logout will have to be handled as the login may be cached even after closing and reopening a new session with the same profile</br>
        /// </summary>
        /// <param name="browser">Browser Type (Chrome, Firefox, IE, Edge)</param>
        /// <param name="wait">Predefined WebDriverWait for the Created Web Driver returned as an out parameter</param>
        /// <param name="pageLoadStrategy">Page Load Strategy Set to Default by Default. Should be Updated to None in Case Browser Page will Contain Windows Authentication Handling</param>
        /// <param name="timeToWaitInMinutes">Time in Minutes to Define the Wait (Explicit and Implicit) with</param>
        /// <param name="useProfile">Boolean to Confirm whether to Use an existing Profile in Browser Initialization or Open a Clean Session, Default Value is False</param>
        /// <param name="runHeadless">Boolean to Open the Browser in Interactive Mode or in Background, Default Value is False</param>
        /// <returns>Web Driver that will be Used</returns>
        public virtual IWebDriver Initialize(Browser browser, out WebDriverWait wait, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Default, double timeToWaitInMinutes = 2, bool useProfile = false, bool runHeadless = false)
        {
            LogCommands("", GetCurrentMethod() + " New Driver and Wait", false, true);

            try
            {
                var driversDirectory = TestContext.CurrentContext.TestDirectory;

                if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("RunHeadless")))
                    runHeadless = bool.Parse(Environment.GetEnvironmentVariable("RunHeadless"));

                IWebDriver driver;

                switch (browser)
                {
                    case Browser.ie:
                        var ieOptions = new InternetExplorerOptions
                        {
                            PageLoadStrategy = pageLoadStrategy
                        };
                        ieOptions.SetLoggingPreference(LogType.Browser, LogLevel.Severe);
                        driver = new InternetExplorerDriver(driversDirectory, ieOptions, TimeSpan.FromMinutes(timeToWaitInMinutes));
                        break;
                    case Browser.chrome:
                        if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ChromeWebDriver")))
                            driversDirectory = Environment.GetEnvironmentVariable("ChromeWebDriver");
                        var chromeOptions = new ChromeOptions();
                        chromeOptions.AddArguments("test-type");
                        chromeOptions.AddArguments("start-maximized");
                        chromeOptions.AddArguments("--js-flags=--expose-gc");
                        chromeOptions.AddArguments("--enable-precise-memory-info");
                        chromeOptions.AddArguments("--disable-popup-blocking");
                        chromeOptions.AddArguments("--disable-default-apps");
                        chromeOptions.AddArguments("test-type=browser");
                        chromeOptions.AddArguments("disable-infobars");
                        chromeOptions.AddArguments("--disable-notifications");
                        chromeOptions.AddArguments("--disable-device-discovery-notifications");
                        chromeOptions.AddExcludedArgument("enable-automation");
                        chromeOptions.AddAdditionalCapability("useAutomationExtension", false);
                        chromeOptions.SetLoggingPreference(LogType.Browser, LogLevel.Severe);
                        if (useProfile)
                        {
                            var localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                            localAppDataPath += "\\Google\\Chrome\\User Data";
                            chromeOptions.AddArguments("--user-data-dir=" + localAppDataPath + "\\Default");
                        }
                        if (runHeadless)
                            chromeOptions.AddArguments("--headless");
                        chromeOptions.PageLoadStrategy = pageLoadStrategy;
                        driver = new ChromeDriver(driversDirectory, chromeOptions, TimeSpan.FromMinutes(timeToWaitInMinutes));
                        break;
                    case Browser.firefox:
                        var ffOptions = new FirefoxOptions();
                        ffOptions.SetPreference("dom.webnotifications.enabled", false);
                        Environment.SetEnvironmentVariable("webdriver.gecko.driver", driversDirectory);
                        var profManager = new FirefoxProfileManager();
                        var profile = profManager.GetProfile(profManager.ExistingProfiles[0]);
                        ffOptions.Profile = profile;
                        ffOptions.PageLoadStrategy = pageLoadStrategy;
                        ffOptions.SetLoggingPreference(LogType.Browser, LogLevel.Severe);
                        if (runHeadless)
                            ffOptions.AddArguments("-headless");
                        Thread.Sleep(1000);
                        driver = new FirefoxDriver(driversDirectory, ffOptions, TimeSpan.FromMinutes(timeToWaitInMinutes));
                        break;
                    case Browser.edge:
                        var edgeOptions = new EdgeOptions { PageLoadStrategy = pageLoadStrategy };
                        Environment.SetEnvironmentVariable("webdriver.edge.driver", driversDirectory);
                        edgeOptions.SetLoggingPreference(LogType.Browser, LogLevel.Severe);
                        driver = new EdgeDriver(driversDirectory, edgeOptions, TimeSpan.FromMinutes(timeToWaitInMinutes));
                        break;
                    default:
                        driver = new ChromeDriver(driversDirectory);
                        break;
                }

                wait = new WebDriverWait(driver, TimeSpan.FromMinutes(timeToWaitInMinutes));
                driverInstanceNumber++;

                if (!browser.Equals(Browser.chrome))
                    driver.Manage().Window.Maximize();

                return driver;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Opens a new empty tab and switch to it
        /// </summary>
        /// <param name="numberOfTabs">Number of tabs or windows open in the current driver after opening the new tab</param>
        public void OpenNewTab(int numberOfTabs)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            ExecuteScript("window.open();");
            Wait.Until(d => d.WindowHandles.Count == numberOfTabs);
            SwitchToWindowWithUrlContaining("about:blank");
        }

        /// <summary>
        /// Wait for tab or window to open by giving the expected number of tabs/windows that should be found in the driver
        /// </summary>
        /// <param name="numberOfTabs">Number of tabs or windows open in the current driver after opening or closing the tab or window</param>
        public virtual void WaitForTabToOpenOrClose(int numberOfTabs)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            Wait.Until(d => d.WindowHandles.Count == numberOfTabs);
        }

        /// <summary>
        /// Instructs the driver to send future commands to the given window
        /// </summary>
        /// <param name="windowName">Window name to be selected</param>
        public virtual void SwitchToWindow(string windowName)
        {
            LogCommands("", GetCurrentMethod() + $" {windowName}", false, true);

            try
            {
                Driver.SwitchTo().Window(windowName);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Switch to window containing part of URL
        /// </summary>
        /// <param name="urlPart">Part of the URL that should be in the page</param>
        public virtual void SwitchToWindowWithUrlContaining(string urlPart)
        {
            LogCommands("", GetCurrentMethod() + $" {urlPart}", false, true);

            try
            {
                foreach (var window in Driver.WindowHandles)
                {
                    Driver.SwitchTo().Window(window);
                    if (Driver.Url.Contains(urlPart))
                    {
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Switch to window Not containing part of URL
        /// </summary>
        /// <param name="urlPart">Part of the URL that shouldn't be in the page</param>
        public virtual void SwitchToWindowWithUrlNotContaining(string urlPart)
        {
            LogCommands("", GetCurrentMethod() + $" {urlPart}", false, true);

            try
            {
                foreach (var window in Driver.WindowHandles)
                {
                    Driver.SwitchTo().Window(window);
                    if (!Driver.Url.Contains(urlPart))
                    {
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Navigate to given URL
        /// </summary>
        /// <param name="url">URL to go to</param>
        public virtual void GoToUrl(string url)
        {
            LogCommands("", GetCurrentMethod() + $" {url}", false, true);

            try
            {
                Driver.Url = url;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// Get Url of the Current Tab/Window
        /// </summary>
        /// <returns></returns>
        public virtual string GetUrl()
        {
            LogCommands("", GetCurrentMethod(), false, true);

            return Driver.Url;
        }

        /// <summary>
        /// Maximize the current browser window
        /// </summary>
        public virtual void MaximizeWindow()
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                Driver.Manage().Window.Maximize();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Refreshes the current page
        /// </summary>
        public virtual void RefreshPage()
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                Driver.Navigate().Refresh();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Close and Dispose the driver if opened
        /// </summary>
        public virtual void CloseDriver()
        {
            try
            {
                GetBrowserErrors();

                bool driverWasOpen;

                if (driverWasOpen = IsDriverOpen() && Driver.WindowHandles.Count == 1)
                    Driver.Close();
                else
                    Driver.Quit();

                Driver.Dispose();

                if (driverWasOpen)
                    LogCommands("", GetCurrentMethod(), false, true);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Checks if the driver is open or not
        /// </summary>
        /// <returns></returns>
        public virtual bool IsDriverOpen()
        {
            if (Driver == null) return false;

            try
            {
                var str = Driver.CurrentWindowHandle;
                if (!GetCurrentMethod(level: 2).Contains(nameof(GetBrowserErrors)) && !GetCurrentMethod(level: 2).Contains(nameof(CloseDriver)))
                    LogCommands("", GetCurrentMethod(), false, true);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Closes the current tab and xes back to the first tab
        /// </summary>
        public void CloseTab()
        {
            LogCommands("", GetCurrentMethod(), false, true);

            Driver.Close();
            SwitchToWindow(Driver.WindowHandles[0]);
        }

        /// <summary>
        /// Detect Browser Errors Found in the Console
        /// </summary>
        public virtual void GetBrowserErrors()
        {
            try
            {
                if (IsDriverOpen())
                {
                    var browserErrors = Driver.Manage().Logs.GetLog(LogType.Browser);

                    if (browserErrors.Any())
                    {
                        if (!string.IsNullOrEmpty(errors))
                            errors += Environment.NewLine;
                        errors = $"Browser error(s) in Browser #{driverInstanceNumber}:" + Environment.NewLine + browserErrors.Aggregate("", (s, entry) => s + entry.Message + Environment.NewLine);
                    }
                }
            }
            catch (NullReferenceException)
            {

            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Accept alert if present
        /// </summary>
        public virtual void AcceptAlert()
        {
            try
            {
                Driver.SwitchTo().Alert().Accept();
                LogCommands("", GetCurrentMethod(), false, true);
            }
            catch (NoAlertPresentException)
            {
                //ignored
            }
        }

        /// <summary>
        /// Dismiss alert if present
        /// </summary>
        public virtual void DismissAlert()
        {
            try
            {
                Driver.SwitchTo().Alert().Dismiss();
                LogCommands("", GetCurrentMethod(), false, true);
            }
            catch (NoAlertPresentException)
            {
                //ignored
            }
        }

        #endregion

        #region Frames Handling Functions

        /// <summary>
        /// Get Iframe Containing Certain Element: Finds and switches to the iframe that contains the element provided by its Locator
        /// <br>Note: The frames have to be on the same level under the root (default content), the function won't search in nested frames</br>
        /// </summary>
        /// <param name="locator">Locator of the element that will be searched for in the iframes</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>Returns the id of the iframe if found</returns>
        public virtual string SwitchToFrameOfElement(By locator, string locatorDescription = "")
        {
            LogCommands("", GetCurrentMethod() + " " + (string.IsNullOrEmpty(locatorDescription) ? ("with Locator: " + locator.ToString()) : locatorDescription), false, true);

            WaitForPageReadyState();

            SwitchToDefault();
            try
            {
                FindElement(locator);
                return null;
            }
            catch (Exception)
            {
                var iframeElementsList = FindElements(By.TagName("iframe"));
                foreach (var iframeElement in iframeElementsList)
                {
                    if (!iframeElement.Displayed && !iframeElement.GetAttribute("style").Contains("visible")) continue;
                    var iframeId = iframeElement.GetAttribute("id");
                    if (string.IsNullOrEmpty(iframeId)) continue;
                    SwitchToFrame(iframeId);
                    try
                    {
                        FindElement(locator);
                        return iframeId;
                    }
                    catch (NoSuchElementException)
                    {
                        SwitchToDefault();
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Instructs the driver to send future commands to the main document (outside of all frames)
        /// </summary>
        public virtual void SwitchToDefault()
        {
            if (!GetCurrentMethod(level: 2).Contains(nameof(WaitForPageReadyState)) && !GetCurrentMethod(level: 2).Contains(nameof(SwitchToFrame)) && !GetCurrentMethod(level: 2).Contains(nameof(SwitchToFrameOfElement)))
                LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                Driver.SwitchTo().DefaultContent();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Instructs the driver to send future commands to the parent frame of the current frame
        /// </summary>
        public virtual void SwitchToParentFrame()
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                Driver.SwitchTo().ParentFrame();
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Instructs the driver to send future commands to the given frame
        /// the frame should be directly under the main document, not nested
        /// </summary>
        /// <param name="frameId">Frame ID or name to be selected</param>
        public virtual void SwitchToFrame(string frameId)
        {
            try
            {
                SwitchToDefault();
                WaitFor(By.Id(frameId), Condition.FrameAvailabilityAndSwitchToIt, "Frame: " + frameId);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }

            LogCommands("", GetCurrentMethod() + $" {frameId}", false, true);
        }

        /// <summary>
        /// Instructs the driver to send future commands to the given frame which may be nested inside another frame
        /// </summary>
        /// <param name="frameId">Frame ID or name to be selected</param>
        public virtual void SwitchToNestedFrame(string frameId)
        {
            WaitFor(By.Id(frameId), Condition.FrameAvailabilityAndSwitchToIt, "Frame: " + frameId);

            LogCommands("", GetCurrentMethod() + $" {frameId}", false, true);
        }

        #endregion

        #region Dealing with Numbers Functions

        /// <summary>
        /// Random Number: Generate random number and return it
        /// </summary>
        /// <param name="startNumber">Start of the Interval of Numbers to find from</param>
        /// <param name="endNumber">End of the Interval of Numbers to find from (not included in the random number to be generated)</param>
        /// <returns>The Chosen Random Number</returns>
        public static int RandomNumber(int startNumber, int endNumber)
        {
            lock (_syncLock)
            {
                return _random.Next(startNumber, endNumber);
            }
        }

        /// <summary>
        /// Generate random double number and return it
        /// </summary>
        /// <param name="minValue">Start of the Interval of Numbers to find from</param>
        /// <param name="maxValue">End of the Interval of Numbers to find from</param>
        /// <returns>The Chosen Random Number</returns>
        public static double RandomDoubleNumber(double minValue, double maxValue)
        {
            lock (_syncLock)
            {
                return minValue + (_random.NextDouble() * (maxValue - minValue));
            }
        }

        /// <summary>
        /// Get Numbers In Text: Returns the Given Number as Text (e.g. number=123, Output is: One Two Three)
        /// </summary>
        /// <param name="number">Given Number that will be Converted to Text</param>
        /// <returns>Text of the Given Number</returns>
        public static string GetNumbersInText(int number)
        {
            lock (_syncLock)
            {
                var numberInText = "";
                var numberString = number.ToString();
                var numberStringList = numberString.ToList();

                for (var numbersCount = 0; numbersCount < numberStringList.Count; numbersCount++)
                {
                    switch (numberStringList[numbersCount])
                    {
                        case '0':
                            numberInText += "Zero";
                            break;
                        case '1':
                            numberInText += "One";
                            break;
                        case '2':
                            numberInText += "Two";
                            break;
                        case '3':
                            numberInText += "Three";
                            break;
                        case '4':
                            numberInText += "Four";
                            break;
                        case '5':
                            numberInText += "Five";
                            break;
                        case '6':
                            numberInText += "Six";
                            break;
                        case '7':
                            numberInText += "Seven";
                            break;
                        case '8':
                            numberInText += "Eight";
                            break;
                        case '9':
                            numberInText += "Nine";
                            break;
                    }
                    if (numbersCount != (numberStringList.Count - 1))
                    {
                        numberInText += " ";
                    }
                }

                return numberInText;
            }
        }

        /// <summary>
        /// Get Numbers In Text: Returns the Given Number as Text (e.g. number=123, Output is: One Two Three)
        /// </summary>
        /// <param name="number">Given Number that will be Converted to Text</param>
        /// <returns>Text of the Given Number</returns>
        public static string GetNumbersInTextArabic(int number)
        {
            lock (_syncLock)
            {
                var numberInText = "";
                var numberString = number.ToString();
                var numberStringList = numberString.ToList();

                for (var numbersCount = 0; numbersCount < numberStringList.Count; numbersCount++)
                {
                    switch (numberStringList[numbersCount])
                    {
                        case '0':
                            numberInText += "صفر";
                            break;
                        case '1':
                            numberInText += "واحد";
                            break;
                        case '2':
                            numberInText += "اثنان";
                            break;
                        case '3':
                            numberInText += "ثلاثة";
                            break;
                        case '4':
                            numberInText += "أربعة";
                            break;
                        case '5':
                            numberInText += "خمسة";
                            break;
                        case '6':
                            numberInText += "ستة";
                            break;
                        case '7':
                            numberInText += "سبعة";
                            break;
                        case '8':
                            numberInText += "ثمانية";
                            break;
                        case '9':
                            numberInText += "تسعة";
                            break;
                    }
                    if (numbersCount != (numberStringList.Count - 1))
                    {
                        numberInText += " ";
                    }
                }

                return numberInText;
            }
        }

        /// <summary>
        /// Extract Number from Given Text
        /// </summary>
        /// <param name="text">Text containing the number</param>
        /// <returns>Number found in text</returns>
        public virtual string ExtractNumberFromText(string text)
        {
            return new string(text.Where(char.IsDigit).ToArray());
        }

        #endregion

        #region Reporting/Logging and Send Mail Functions

        /// <summary>
        /// set Directory: Get the full path of certain Folder inside the Project
        /// </summary>
        /// <param name="folderName">Container Folder inside the Project (e.g. Drivers or Attachments)</param>
        /// <returns></returns>
        public static string SetDir(string folderName)
        {
            lock (_syncLock)
            {
                return TestContext.CurrentContext.TestDirectory + (string.IsNullOrEmpty(folderName) ? "" : "\\" + folderName);
            }
        }

        /// <summary>
        /// Create Or Update File To Save Request IDs And Text: Creates or Updates "RequestsLogFile.txt" File based on the Given Parameters
        /// </summary>
        /// <param name="requestId">Request Number to be Logged if Available</param>
        /// <param name="textToBeSaved">Text that is Required to be Logged</param>
        /// <param name="newLineOnly">Boolean to indicate if only a New Line Separator will be entered in the Log File (Line Consists of Multiple = Signs)</param>
        /// <param name="textOnly">Boolean to indicate if only the entered text will be Logged without the Request Number</param>
        public static void LogInfo(string requestId, string textToBeSaved, bool newLineOnly, bool textOnly)
        {
            Log("RequestsLogFile", requestId, textToBeSaved, newLineOnly, textOnly);
        }

        /// <summary>
        /// Create Or Update Issues And Errors File To Save Text: Creates or Updates "IssuesLogFile.txt" File based on the Given Parameters
        /// </summary>
        /// <param name="requestId">Request Number to be Logged if Available</param>
        /// <param name="textToBeSaved">Text that is Required to be Logged</param>
        /// <param name="newLineOnly">Boolean to indicate if only a New Line Separator will be entered in the Log File (Line Consists of Multiple = Signs)</param>
        /// <param name="textOnly">Boolean to indicate if only the entered text will be Logged without the Request Number</param>
        public static void LogIssue(string requestId, string textToBeSaved, bool newLineOnly, bool textOnly)
        {
            Log("IssuesLogFile", requestId, textToBeSaved, newLineOnly, textOnly);
        }

        /// <summary>
        /// Create Or Update File To Save Request IDs And Text: Creates or Updates "CommandsLogFile.txt" File based on the Given Parameters
        /// </summary>
        /// <param name="requestId">Request Number to be Logged if Available</param>
        /// <param name="textToBeSaved">Text that is Required to be Logged</param>
        /// <param name="newLineOnly">Boolean to indicate if only a New Line Separator will be entered in the Log File (Line Consists of Multiple = Signs)</param>
        /// <param name="textOnly">Boolean to indicate if only the entered text will be Logged without the Request Number</param>
        private static void LogCommands(string requestId, string textToBeSaved, bool newLineOnly, bool textOnly)
        {
            Log("CommandsLogFile", requestId, textToBeSaved, newLineOnly, textOnly);
        }

        /// <summary>
        /// Create Or Update File To Save Request IDs And Text: Creates or Updates "RequestsLogFile.txt" File based on the Given Parameters
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="requestId">Request Number to be Logged if Available</param>
        /// <param name="textToBeSaved">Text that is Required to be Logged</param>
        /// <param name="newLineOnly">Boolean to indicate if only a New Line Separator will be entered in the Log File (Line Consists of Multiple = Signs)</param>
        /// <param name="textOnly">Boolean to indicate if only the entered text will be Logged without the Request Number</param>
        public static void Log(string fileName, string requestId, string textToBeSaved, bool newLineOnly, bool textOnly)
        {
            lock (_syncLock)
            {
                var path = SetDir(fileName.Contains(".txt") ? fileName : (fileName + ".txt"));
                TextWriter tw;

                if (!File.Exists(path))
                {
                    File.Create(path).Dispose();
                    tw = new StreamWriter(path);
                }
                else
                {
                    tw = new StreamWriter(path, true);
                }

                StringBuilder stringBuilder = new StringBuilder();

                if (!newLineOnly && !textOnly)
                {
                    stringBuilder.Append("Request Number: " + requestId + Environment.NewLine);
                    stringBuilder.Append("" + textToBeSaved + Environment.NewLine);
                }
                if (textOnly)
                {
                    stringBuilder.Append("" + textToBeSaved + Environment.NewLine);
                }
                if (newLineOnly)
                {
                    stringBuilder.Append("===========================================" + Environment.NewLine);
                }

                stringBuilder.AppendLine();

                if (_test != null && !newLineOnly)
                    _test.Log(Status.Info, stringBuilder.ToString());

                if (!newLineOnly)
                    stringBuilder.Insert(0, "Date and Time: " + DateTime.Now.ToString(CultureInfo.InvariantCulture) + Environment.NewLine);

                tw.WriteLine(stringBuilder.ToString());
                TestContext.WriteLine(stringBuilder.ToString());

                tw.Close();
            }
        }

        /// <summary>
        /// Take Screenshot: Takes a Screenshot of the Browser's Screen
        /// And Saved the Image in Folder "Screenshots" in the Project's Folder
        /// </summary>
        /// <param name="screenshotName">Screenshot Name, will be filled randomly if not provided</param>
        /// <returns>Returns Created Image Name or "No File Created" if there was an Exception while Taking the Screenshot</returns>
        public virtual string TakeScreenshot(string screenshotName = "")
        {
            try
            {
                Directory.CreateDirectory(SetDir("\\Screenshots"));
            }
            catch (IOException)
            {
            }

            var previousMethod = GetCurrentMethod(level: 2);
            string fileName;

            try
            {
                var screenshot = ((ITakesScreenshot)Driver).GetScreenshot();
                fileName = string.IsNullOrEmpty(screenshotName) ? "Screenshot_" + previousMethod + "_" + RandomNumber(100000, 999999) + ".png" : (screenshotName + (screenshotName.Contains(".png") ? "" : ".png"));
                screenshot.SaveAsFile(SetDir("Screenshots\\" + fileName), ScreenshotImageFormat.Png);
                _test.AddScreenCaptureFromPath(SetDir("Screenshots\\" + fileName));
                LogCommands("", GetCurrentMethod(), false, true);
            }
            catch (Exception)
            {
                fileName = "No File Created";
            }

            return fileName;
        }

        /// <summary>
        /// Get Current Method: Returns the Name of the Method in Given Level, By Default it Returns the Name of the Method Calling the GetCurrentMethod Function
        /// </summary>
        /// <param name="level">Index of the Method whose Name will be Returned in the Call Stack</param>
        /// <param name="insertMethodNameInLog">Boolean to indicate if the Name of the Method will be entered in the Log File</param>
        /// <returns></returns>
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static string GetCurrentMethod(int level = 1, bool insertMethodNameInLog = false)
        {
            lock (_syncLock)
            {
                var st = new StackTrace();
                var sf = st.GetFrame(level);

                var testMethodName = sf.GetMethod().Name;

                if (insertMethodNameInLog)
                {
                    TestContext.WriteLine("Running Test Method Name: " + testMethodName);
                    LogInfo("", "Running Test Method Name: " + testMethodName, false,
                        true);
                }

                return sf.GetMethod().Name;
            }
        }

        /// <summary>
        /// Send Mail: Sends Mail through SendGrid Api using account "linkdev.automationtest@gmail.com" while attaching log files and screenshots if any were taken
        /// </summary>
        /// <param name="fromMail">Email Address that will Send the Mail</param>
        /// <param name="subject">Subject of the Mail</param>
        /// <param name="plainBody">Body of the Mail in Plain Text Format</param>
        /// <param name="htmlBody">Body of the Mail in HTML Format</param>
        /// <param name="toMails">List of Emails that the mail will be sent to</param>
        /// <param name="apiKey">SendGrid Api key for the account that is used to send the email</param>
        /// <param name="filesToBeUploadedList">List of files that need to be uploaded in the email, should be in the output directory folder</param>
        public virtual async Task SendMail(string fromMail, string subject, string plainBody, string htmlBody, StringCollection toMails, string apiKey, params string[] filesToBeUploadedList)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());

                var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
                var productName = "";
                if (reg != null)
                {
                    productName = (string)reg.GetValue("ProductName");
                }

                LogInfo("",
                    "System Configuration Settings:\r\n    Machine Windows Version: " + productName +
                    "\r\n    Machine Name: " + Environment.MachineName, false, true);

                var client = new SendGridClient(apiKey);
                var from = new EmailAddress(fromMail);
                List<EmailAddress> emailAddresses = new List<EmailAddress>();
                var to = new EmailAddress("linkdev.automationtest@gmail.com");
                emailAddresses.Add(to);
                if (toMails != null && toMails.Count > 0)
                {
                    foreach (var toMail in toMails)
                    {
                        emailAddresses.Add(new EmailAddress(toMail));
                    }
                }
                var msg = MailHelper.CreateSingleEmailToMultipleRecipients(from, emailAddresses, subject, plainBody, htmlBody);

                if (File.Exists(SetDir("RequestsLogFile.txt")))
                    msg.AddAttachment("RequestsLogFile.txt", Convert.ToBase64String(File.ReadAllBytes(SetDir("RequestsLogFile.txt"))), "text/plain", "attachment");

                if (File.Exists(SetDir("IssuesLogFile.txt")))
                    msg.AddAttachment("IssuesLogFile.txt", Convert.ToBase64String(File.ReadAllBytes(SetDir("IssuesLogFile.txt"))), "text/plain", "attachment");

                if (File.Exists(SetDir("CommandsLogFile.txt")))
                    msg.AddAttachment("CommandsLogFile.txt", Convert.ToBase64String(File.ReadAllBytes(SetDir("CommandsLogFile.txt"))), "text/plain", "attachment");

                for (int filesCounter = 0; filesCounter < filesToBeUploadedList.Length; filesCounter++)
                {
                    var fileName = filesToBeUploadedList[filesCounter];

                    if (fileName.Contains("RequestsLogFile") || fileName.Contains("IssuesLogFile"))
                        continue;

                    fileName = fileName.Contains(".txt") ? fileName : (fileName + ".txt");

                    if (File.Exists(SetDir(fileName)))
                        msg.AddAttachment(fileName, Convert.ToBase64String(File.ReadAllBytes(SetDir(fileName))), "text/plain", "attachment");
                }

                if (Directory.Exists(SetDir("Screenshots")))
                {
                    var files = Directory.GetFiles(SetDir("Screenshots"), "*", SearchOption.AllDirectories);
                    foreach (var file in files)
                    {
                        msg.AddAttachment(file.Split('\\').Last(), Convert.ToBase64String(File.ReadAllBytes(file)), "image/png", "attachment");
                    }
                }

                if (Directory.Exists(reportDir))
                {
                    var files = Directory.GetFiles(reportDir);
                    foreach (var file in files)
                    {
                        msg.AddAttachment(file.Split('\\').Last(), Convert.ToBase64String(File.ReadAllBytes(file)), "	text/html", "attachment");
                    }
                }

                var response = await client.SendEmailAsync(msg);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Send Mail: Sends Mail through Linkdev SMTP Server "10.2.2.10" Port 25
        /// While Attaching Log Files and Screenshots if any were Taken
        /// Logs Message in "EmailIssuesLogFile.txt" in case Sending Mail Fails
        /// </summary>
        /// <param name="subject">Subject of the Mail</param>
        /// <param name="body">Body of the Mail</param>
        /// <param name="toMails">List of Emails that the mail will be sent to</param>
        /// <param name="filesToBeUploadedList">List of files that need to be uploaded in the email, should be in the output directory folder</param>
        public virtual void SendMail(string subject, string body, StringCollection toMails, params string[] filesToBeUploadedList)
        {
            LogCommands("", GetCurrentMethod(), false, true);

            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());

                var ips = (from ip in host.AddressList where ip.AddressFamily == AddressFamily.InterNetwork select ip.ToString()).ToList();

                var useInternalSMTP = ips.Find(i => i.StartsWith("10.2.")) != null;

                var fromEmail = useInternalSMTP ? "nataly.wadie@linkdev.com" : "Linkdev.automationtest@gmail.com";//"nataly.wadie@linkdev.com";
                var mailMessage = new MailMessage(fromEmail, "Linkdev.automationtest@gmail.com", subject, body);

                if (toMails != null && toMails.Count > 0)
                {
                    foreach (var toMail in toMails)
                    {
                        mailMessage.To.Add(toMail.ToLower().Trim());
                    }
                }

                var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
                var productName = "";
                if (reg != null)
                {
                    productName = (string)reg.GetValue("ProductName");
                }

                LogInfo("",
                    "System Configuration Settings:\r\n    Machine Windows Version: " + productName +
                    "\r\n    Machine Name: " + Environment.MachineName, false, true);

                if (File.Exists(SetDir("RequestsLogFile.txt")))
                    mailMessage.Attachments.Add(new System.Net.Mail.Attachment(SetDir("RequestsLogFile.txt")));

                if (File.Exists(SetDir("IssuesLogFile.txt")))
                    mailMessage.Attachments.Add(new System.Net.Mail.Attachment(SetDir("IssuesLogFile.txt")));

                if (File.Exists(SetDir("CommandsLogFile.txt")))
                    mailMessage.Attachments.Add(new System.Net.Mail.Attachment(SetDir("CommandsLogFile.txt")));

                for (int filesCounter = 0; filesCounter < filesToBeUploadedList.Length; filesCounter++)
                {
                    var fileName = filesToBeUploadedList[filesCounter];

                    if (fileName.Contains("RequestsLogFile") || fileName.Contains("IssuesLogFile"))
                        continue;

                    fileName = fileName.Contains(".txt") ? fileName : (fileName + ".txt");

                    if (File.Exists(SetDir(fileName)))
                        mailMessage.Attachments.Add(new System.Net.Mail.Attachment(SetDir(fileName)));
                }

                if (Directory.Exists(SetDir("Screenshots")))
                {
                    var files = Directory.GetFiles(SetDir("Screenshots"), "*", SearchOption.AllDirectories);
                    foreach (var file in files)
                    {
                        mailMessage.Attachments.Add(new System.Net.Mail.Attachment(file));
                    }
                }

                if (Directory.Exists(reportDir))
                {
                    var files = Directory.GetFiles(reportDir);
                    foreach (var file in files)
                    {
                        mailMessage.Attachments.Add(new System.Net.Mail.Attachment(file));
                    }
                }

                var counter = 0;
                var mailSent = false;

                do
                {
                    var smtpClient = useInternalSMTP ? new SmtpClient("10.2.2.10", 25) : new SmtpClient("smtp.gmail.com", 587);//new SmtpClient("smtp-pulse.com", 465);
                    smtpClient.EnableSsl = !useInternalSMTP;//true;
                    smtpClient.UseDefaultCredentials = useInternalSMTP;//false;
                    if (!useInternalSMTP)
                        smtpClient.Credentials = new NetworkCredential("linkdev.automationtest@gmail.com", DecodePassword("TGlua0AxMjM0"));//"q4T7H3GqP2t5nCE"
                    try
                    {
                        smtpClient.Send(mailMessage);
                        mailSent = true;
                    }
                    catch (Exception ex)
                    {
                        counter++;
                        TestContext.WriteLine(ex.Message);

                        var path = SetDir("EmailIssuesLogFile.txt");
                        TextWriter tw;

                        if (!File.Exists(path))
                        {
                            File.Create(path).Dispose();
                            tw = new StreamWriter(path);
                        }
                        else
                        {
                            tw = new StreamWriter(path, true);
                        }

                        tw.WriteLine("Date and Time: " + DateTime.Now.ToString(CultureInfo.InvariantCulture));
                        tw.WriteLine("Sending Mail Failed. Mail Content is:");
                        tw.WriteLine(mailMessage.Body);
                        tw.WriteLine("===========================================");
                        tw.Close();
                    }
                } while (!mailSent && counter < 5);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Delete Files and Screenshots: Deletes RequestsLogFile.txt,
        /// IssuesLogFile.txt and Screeshots Folder and All its Contents if they Exist
        /// </summary>
        public static void DeleteFilesAndScreenshots(params string[] filesToBeDeletedList)
        {
            lock (_syncLock)
            {
                LogCommands("", GetCurrentMethod(), false, true);

                if (File.Exists(SetDir("RequestsLogFile.txt")))
                    File.Delete(SetDir("RequestsLogFile.txt"));

                if (File.Exists(SetDir("IssuesLogFile.txt")))
                    File.Delete(SetDir("IssuesLogFile.txt"));

                if (File.Exists(SetDir("CommandsLogFile.txt")))
                    File.Delete(SetDir("CommandsLogFile.txt"));

                if (Directory.Exists(SetDir("Screenshots")))
                    Directory.Delete(SetDir("Screenshots"), true);

                if (Directory.Exists(SetDir("Test_Execution_Reports")))
                    Directory.Delete(SetDir("Test_Execution_Reports"), true);

                for (int filesCounter = 0; filesCounter < filesToBeDeletedList.Length; filesCounter++)
                {
                    var fileName = filesToBeDeletedList[filesCounter];

                    if (fileName.Contains("RequestsLogFile") || fileName.Contains("IssuesLogFile"))
                        continue;

                    if (File.Exists(SetDir(fileName)))
                        File.Delete(SetDir(fileName));
                }
            }
        }

        #endregion

        #region Rarely Used Functions

        /// <summary>
        /// Encodes a string in Base64
        /// </summary>
        /// <param name="password">String to be encoded</param>
        /// <returns>Encoded string</returns>
        public virtual string EncodePassword(string password)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(password));
        }

        /// <summary>
        /// Decodes a string from Base64
        /// </summary>
        /// <param name="encodedPassword">Encoded string that should be decoded</param>
        /// <returns>Decoded string</returns>
        public virtual string DecodePassword(string encodedPassword)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(encodedPassword));
        }

        /// <summary>
        /// Generates an Egyptian National ID
        /// </summary>
        /// <param name="olderThan21">Default value is true to generate a national id of older than 21 years person</param>
        /// <param name="isMale">Default value is true to generate a national id for a isMale person</param>
        /// <returns>The generated national id</returns>
        public virtual string GenerateEgyptianNationalId(bool olderThan21 = true, bool isMale = true)
        {
            try
            {
                var hasToBeBefore2000 = false;
                var minYear = DateTime.Today.Year.ToString().Substring(2);

                if (olderThan21)
                {
                    var yearOver21 = DateTime.Today.AddYears(-21).Year;
                    hasToBeBefore2000 = yearOver21 < 2000;
                    minYear = yearOver21.ToString().Substring(2);
                }

                var nationalId = "";

                var firstNumber = hasToBeBefore2000 ? 2 : RandomNumber(2, 4);
                int secondNumber;
                int thirdNumber;

                switch (firstNumber)
                {
                    case 2:
                        secondNumber = RandomNumber(8, 10);
                        thirdNumber = RandomNumber(0, secondNumber.ToString().Equals(minYear.Substring(0, 1)) ? int.Parse(minYear.Substring(1)) : 10);
                        break;
                    case 3:
                        secondNumber = RandomNumber(0, 2);
                        thirdNumber = RandomNumber(0, 8);
                        break;
                    default:
                        secondNumber = RandomNumber(8, 10);
                        thirdNumber = RandomNumber(0, 10);
                        break;
                }
                nationalId += firstNumber + secondNumber.ToString() + thirdNumber;

                var fourthNumber = RandomNumber(0, 2);
                int fifthNumber;

                switch (fourthNumber)
                {
                    case 0:
                        fifthNumber = RandomNumber(1, 10);
                        break;
                    case 1:
                        fifthNumber = RandomNumber(0, 3);
                        break;
                    default:
                        fifthNumber = 0;
                        break;
                }

                var month = fourthNumber + "" + fifthNumber;

                nationalId += month;

                int sixthNumber;
                int seventhNumber;

                switch (month)
                {
                    case "02":
                        sixthNumber = RandomNumber(0, 3);
                        switch (sixthNumber)
                        {
                            case 2:
                                seventhNumber = RandomNumber(0, 9);
                                break;
                            default:
                                seventhNumber = RandomNumber(0, 10);
                                break;
                        }
                        break;
                    case "04":
                    case "06":
                    case "09":
                    case "11":
                        sixthNumber = RandomNumber(0, 4);
                        switch (sixthNumber)
                        {
                            case 3:
                                seventhNumber = 0;
                                break;
                            default:
                                seventhNumber = RandomNumber(0, 10);
                                break;
                        }
                        break;
                    default:
                        sixthNumber = RandomNumber(0, 4);
                        switch (sixthNumber)
                        {
                            case 3:
                                seventhNumber = RandomNumber(0, 2);
                                break;
                            default:
                                seventhNumber = RandomNumber(0, 10);
                                break;
                        }
                        break;
                }

                var day = sixthNumber + "" + seventhNumber;

                nationalId += day;

                var cityIdsList = new[]
                {
                "01", "02", "03", "04", "11", "12", "13", "14", "15", "16", "17", "18", "19", "21", "22", "23", "24",
                "25", "26", "27", "28", "29", "31", "32", "33", "34", "35", "88"
            };

                nationalId += cityIdsList[RandomNumber(0, cityIdsList.Length)];

                nationalId += RandomNumber(0, 10);
                nationalId += RandomNumber(0, 10);
                nationalId += RandomNumber(0, 10);

                var maleNumbers = new[] { "1", "3", "5", "7", "9" };
                var femaleNumbers = new[] { "2", "4", "6", "8" };

                nationalId += isMale
                    ? maleNumbers[RandomNumber(0, maleNumbers.Length)]
                    : femaleNumbers[RandomNumber(0, femaleNumbers.Length)];
                nationalId += RandomNumber(0, 10);

                return nationalId;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Form Generator Function: Generates Code for all the fields found in the Form and writes the output in a class ready to be added to the project
        /// </summary>
        /// <param name="elementsContainerDivLocator">html element that carries all the fields that need to be filled</param>
        /// /// <param name="doTheAction">If set to true, the element's action will be executed, not just written in the output file. e.g. if it is an input then the field will be filled with test data</param>
        /// /// <param name="fileName">Name of the output class, default is "FormFiller"</param>
        /// <param name="IsAngularJS">If set to true, the select elements will be selected first in case any were hidden</param>
        public void FormGenerator(By elementsContainerDivLocator, bool doTheAction = false, string fileName = "FormFiller", bool IsAngularJS = false)
        {
            fileName = fileName.Replace(" ", "");

            if (File.Exists(SetDir(fileName + ".cs")))
                File.Delete(SetDir(fileName + ".cs"));

            using (StreamWriter file = new StreamWriter(SetDir(fileName + ".cs")))
            {
                file.WriteLine("using System;");
                file.WriteLine("using System.Collections.Generic;");
                file.WriteLine("using System.Collections.ObjectModel;");
                file.WriteLine("using System.Collections.Specialized;");
                file.WriteLine("using System.Diagnostics;");
                file.WriteLine("using System.Globalization;");
                file.WriteLine("using System.IO;");
                file.WriteLine("using System.Linq;");
                file.WriteLine("using System.Text;");
                file.WriteLine("using System.Threading;");
                file.WriteLine("using NUnit.Framework;");
                file.WriteLine("using OpenQA.Selenium;");
                file.WriteLine("using OpenQA.Selenium.Chrome;");
                file.WriteLine("using OpenQA.Selenium.Interactions;");
                file.WriteLine("using OpenQA.Selenium.Support.UI;");
                file.WriteLine("using ExpectedConditions = SeleniumExtras.WaitHelpers.ExpectedConditions;");
                file.WriteLine("");

                file.WriteLine("namespace " + typeof(Common).Namespace);
                file.WriteLine("{");
                file.WriteLine($"\tpublic class {fileName} : Common");
                file.WriteLine("\t{");
                file.WriteLine("\t\tpublic void Fill()");
                file.WriteLine("\t\t{");
                file.WriteLine("");

                WaitForPageReadyState();

                WaitFor(elementsContainerDivLocator, Condition.Visible);

                IList<IWebElement> FieldsContainerList = FindElements(elementsContainerDivLocator);
                IWebElement FieldsContainer = FieldsContainerList[0];

                if (IsAngularJS)
                {
                    IList<IWebElement> SelectsListForAngular = FieldsContainer.FindElements(By.TagName("select"));
                    for (int AngularSelectsCount = 0; AngularSelectsCount < SelectsListForAngular.Count; AngularSelectsCount++)
                    {
                        if (SelectsListForAngular[AngularSelectsCount].Displayed && string.IsNullOrEmpty(SelectsListForAngular[AngularSelectsCount].GetAttribute("disabled")))
                        {
                            WaitFor(SelectsListForAngular[AngularSelectsCount], Condition.SelectListLoaded);
                            new SelectElement(SelectsListForAngular[AngularSelectsCount]).SelectByIndex(1);
                            Thread.Sleep(3000);
                        }
                    }
                }

                FieldsContainer = FindElement(elementsContainerDivLocator);
                IList<IWebElement> containerChildren = FieldsContainer.FindElements(By.XPath("descendant-or-self::*"));
                for (int childrenCounter = 0; childrenCounter < containerChildren.Count; childrenCounter++)
                {
                    try
                    {
                        IWebElement child = containerChildren[childrenCounter];

                        if (child.TagName.ToLower().Equals("input"))
                            AddAndFillInput(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("select"))
                            AddAndFillSelect(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("textarea"))
                            AddAndFillTextArea(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("a"))
                            AddAndFillLink(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("button"))
                            AddAndFillButton(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("ul"))
                            AddAndFillUnOrderedList(file, child, doTheAction);

                        if (child.TagName.ToLower().Equals("li"))
                            AddAndFillListItem(file, child, doTheAction);

                        containerChildren = FieldsContainer.FindElements(By.XPath("descendant-or-self::*"));
                    }
                    catch (Exception e)
                    {
                        LogIssue("", $"Exception occurred in function \"{GetCurrentMethod()}\" with message \"{e.Message}\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                        break;
                    }
                }

                file.WriteLine("");
                file.WriteLine("\t\t}");
                file.WriteLine("\t}");
                file.WriteLine("}");
            }
        }

        private void AddAndFillListItem(StreamWriter file, IWebElement li, bool doTheAction = false)
        {
            string FieldType = li.GetAttribute("type");
            if (string.IsNullOrEmpty(FieldType))
            {
                FieldType = li.GetAttribute("role");
            }

            if (string.IsNullOrEmpty(li.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = li.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = li.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = li.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = li.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = li.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + FieldType);
                file.WriteLine("\t\t\tClickOn(XPathMaker(\"li\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                if (doTheAction)
                    ClickOn(li, false, attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillUnOrderedList(StreamWriter file, IWebElement ul, bool doTheAction = false)
        {
            string FieldType = ul.GetAttribute("type");
            if (string.IsNullOrEmpty(FieldType))
            {
                FieldType = ul.GetAttribute("role");
            }

            if (string.IsNullOrEmpty(ul.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = ul.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = ul.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = ul.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = ul.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = ul.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + FieldType);
                file.WriteLine("\t\t\tClickOn(XPathMaker(\"ul\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                if (doTheAction)
                    ClickOn(ul, false, attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillButton(StreamWriter file, IWebElement button, bool doTheAction = false)
        {
            if (string.IsNullOrEmpty(button.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = button.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = button.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = button.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = button.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = button.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + button.Text);
                file.WriteLine("\t\t\tClickOn(XPathMaker(\"button\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                if (doTheAction)
                    ClickOn(button, false, attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillLink(StreamWriter file, IWebElement link, bool doTheAction = false)
        {
            if (string.IsNullOrEmpty(link.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = link.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = link.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = link.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = link.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = link.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + link.Text);
                file.WriteLine("\t\t\tClickOn(XPathMaker(\"a\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                if (doTheAction)
                    ClickOn(link, false, attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillTextArea(StreamWriter file, IWebElement textArea, bool doTheAction = false)
        {
            if (string.IsNullOrEmpty(textArea.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = textArea.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = textArea.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = textArea.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = textArea.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = textArea.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + textArea.GetAttribute("placeholder"));
                file.WriteLine("\t\t\tSendKeys(XPathMaker(\"textarea\", \"" + attribute + "\", \"" + attributeValue + "\"), \"test textarea\");");
                if (doTheAction)
                    SendKeys(textArea, "test textarea", attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillSelect(StreamWriter file, IWebElement select, bool doTheAction = false)
        {
            if (string.IsNullOrEmpty(select.GetAttribute("disabled")))
            {
                string attribute = "id";
                string attributeValue = select.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = select.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = select.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = select.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = select.GetAttribute("class");
                            }
                        }
                    }
                }

                file.WriteLine("\t\t\t//" + attributeValue);
                file.WriteLine("\t\t\tSelectByIndex(XPathMaker(\"select\", \"" + attribute + "\", \"" + attributeValue + "\"), 1);");
                if (doTheAction)
                    SelectByIndex(select, 1, attributeValue);
                file.WriteLine();
            }
        }

        private void AddAndFillInput(StreamWriter file, IWebElement input, bool doTheAction = false)
        {
            if (string.IsNullOrEmpty(input.GetAttribute("disabled")))
            {
                string FieldType = input.GetAttribute("type");
                if (string.IsNullOrEmpty(FieldType))
                {
                    FieldType = input.GetAttribute("role");
                }

                string attribute = "id";
                string attributeValue = input.GetAttribute("id");
                if (string.IsNullOrEmpty(attributeValue))
                {
                    attribute = "name";
                    attributeValue = input.GetAttribute("name");
                    if (string.IsNullOrEmpty(attributeValue))
                    {
                        attribute = "ng-model";
                        attributeValue = input.GetAttribute("ng-model");
                        if (string.IsNullOrEmpty(attributeValue))
                        {
                            attribute = "ng-click";
                            attributeValue = input.GetAttribute("ng-click");
                            if (string.IsNullOrEmpty(attributeValue))
                            {
                                attribute = "class";
                                attributeValue = input.GetAttribute("class");
                            }
                        }
                    }
                }

                if (FieldType.Equals("text"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tSendKeys(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), \"test text\");");
                    if (doTheAction)
                        SendKeys(input, "test text", attributeValue);
                }

                if (FieldType.Equals("submit"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tClickOn(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                    if (doTheAction)
                        ClickOn(input, false, attributeValue);
                }

                if (FieldType.Equals("radio"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tClickOn(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                    if (doTheAction)
                        ClickOn(input, false, attributeValue);
                }

                if (FieldType.Equals("password"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tSendKeys(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), \"P@ssw0rd\");");
                    if (doTheAction)
                        SendKeys(input, "P@ssw0rd", attributeValue);
                }

                if (FieldType.Equals("file"))
                {
                    if (string.IsNullOrEmpty(input.GetAttribute("title")))
                    {
                        IWebElement FileParent = input.FindElement(By.XPath(".."));
                        file.WriteLine("\t\t\t//" + FileParent.Text);
                    }
                    else
                    {
                        file.WriteLine("\t\t\t//" + input.GetAttribute("title"));
                    }
                    file.WriteLine("\t\t\tUploadAttachment(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), FileType.Pdf);");
                    if (doTheAction)
                        UploadAttachments(XPathMaker("input", attribute, attributeValue), FileType.Pdf);
                }

                if (FieldType.Equals("checkbox"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tClickOn(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), false);");
                    if (doTheAction)
                        ClickOn(input, false, attributeValue);
                }

                if (FieldType.Equals("email"))
                {
                    file.WriteLine("\t\t\t//" + attributeValue);
                    file.WriteLine("\t\t\tSendKeys(XPathMaker(\"input\", \"" + attribute + "\", \"" + attributeValue + "\"), \"testmail@mail.com\");");
                    if (doTheAction)
                        SendKeys(input, "testmail@mail.com", attributeValue);
                }

                file.WriteLine();
            }
        }

        #endregion

        #region UI Testing Trial Functions (Not Used)

        /// <summary>
        /// Export UI Elements to Excel Sheet named "UIElements.xlsx"
        /// First Row contains: WindowWidth(X) | WindowHeight(Y)
        /// Second Row contains: URL | Locator | Location X | Location Y | Scroll Direction | Is Homepage | Is Repeated | Screenshot
        /// Locator is defined as follows: "attributeName:AttributeValue"
        /// </summary>
        /// <param name="baseUrl">Base Url of the Portal being Tested</param>
        public virtual void ExportUIElementsAndPositionsToExcelSheet(string baseUrl)
        {
            try
            {
                var fileName = "UIElements.xlsx";

                if (File.Exists(SetDir(fileName)))
                    File.Delete(SetDir(fileName));

                var excelSheetInputs = new List<ExcelSheetInput>();

                var windowHeight = (long)ExecuteScript("return window.innerHeight");
                var windowWidth = (long)ExecuteScript("return window.innerWidth");

                excelSheetInputs.Add(new ExcelSheetInput(fileName, "WindowWidth(X):" + windowWidth, 1, "A"));
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "WindowHeight(Y):" + windowHeight, 1, "B"));

                excelSheetInputs.Add(new ExcelSheetInput(fileName, "URL", 2, "A"));
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Locator", 2, "B"));
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Location X", 2, "C"));
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Location Y", 2, "D"));
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Scroll Direction", 2, "E")); // Enter H for Horizontal, V for Vertical and N for No Scroll (C for Check)
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Is Homepage", 2, "F")); // Y for Yes, N for No
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Is Repeated", 2, "G")); // Y for Yes, N for No
                excelSheetInputs.Add(new ExcelSheetInput(fileName, "Screenshot", 2, "H")); // Screenshot Name, to be found in Screenshots Folder in Debug Folder

                XmlDocument doc = new XmlDocument();
                doc.Load(SetDir("portalUrls.xml"));

                uint rowNumber = 3;

                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    string url = node.ChildNodes[0].InnerText;

                    url = url.Contains("http") ? "" : (baseUrl + ((!baseUrl.EndsWith("/") && !url.StartsWith("/")) ? "/" : "") + url);

                    GoToUrl(url);

                    var lastPartOfUrl = url.Split('/').Last();
                    var isHomepage = lastPartOfUrl.Equals("en") || lastPartOfUrl.Equals("ar");

                    excelSheetInputs.Add(new ExcelSheetInput(fileName, url, rowNumber, "A"));
                    rowNumber++;

                    var elements = (IList<IWebElement>)ExecuteScript("return $('body *:visible:not(meta):not(script):not(style):not(head):not(link):not(div)')");

                    rowNumber = GetElementsLocators(ref excelSheetInputs, rowNumber, fileName, elements, false, windowHeight, windowWidth);
                }

                new ExcelHelper().UpdateCells(SetDir(fileName), excelSheetInputs);
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Validate UI Elements Found in the UIElements.xlsx Excel Sheet, Follows the Specific Formatting of the Sheet
        /// </summary>
        public virtual void ValidateUIElementsFromExcelSheet()
        {
            try
            {
                var excelHeaders = new Dictionary<string, string>();
                var windowWidth = "";
                var windowHeight = "";
                var excelRows = new ExcelHelper().GetExcelSheetData(out excelHeaders, out windowWidth, out windowHeight);

                for (int rowsCounter = 0; rowsCounter < excelRows.Count; rowsCounter++)
                {
                    var url = excelRows[rowsCounter][0];

                    var containerClass = "";

                    if (!string.IsNullOrEmpty(url))
                    {
                        if (url.Contains("Click:"))
                        {
                            var clickLocSplit = url.Replace("Click:", "").Split(':');
                            ClickOn(Part(clickLocSplit[0], clickLocSplit[1]), false);

                            containerClass = "menu";

                            WaitForPageReadyState();

                            LogInfo("", "Clicked on Element with Locator: " + url.Replace("Click:", ""), false, true);
                        }
                        else
                        {
                            GoToUrl(url);

                            LogInfo("", "Navigated to URL: " + url, false, true);
                        }
                    }

                    var locatorStr = excelRows[rowsCounter][1];

                    if (!string.IsNullOrEmpty(locatorStr) && locatorStr.Contains(":"))
                    {
                        var locSplit = locatorStr?.Split(':');

                        By loc;

                        if (locatorStr.StartsWith("text:"))
                        {
                            loc = XPathMaker(locSplit[1], "text", locSplit[2]);
                        }
                        else if (locatorStr.StartsWith("CssSelector:"))
                        {
                            loc = By.CssSelector(locSplit[1]);
                        }
                        else if (locatorStr.StartsWith("src:../"))
                        {
                            loc = Suffix(locSplit[0], locSplit[1].Replace("../", ""));
                        }
                        else
                        {
                            loc = Full(locSplit[0], locSplit[1]);
                        }

                        var scroll = excelRows[rowsCounter][4];

                        if (!scroll.Equals("N"))
                        {
                            var scrollHorizontalTimes = 0;
                            var scrollVerticalTimes = 0;

                            if (scroll.Contains(":"))
                            {
                                var scrollSplit = scroll.Split(':');
                                scrollHorizontalTimes = int.Parse(scrollSplit[0].Replace("H", ""));
                                scrollVerticalTimes = int.Parse(scrollSplit[1].Replace("V", ""));
                            }
                            else if (scroll.Contains("H"))
                            {
                                scrollHorizontalTimes = int.Parse(scroll.Replace("H", ""));
                            }
                            else
                            {
                                scrollVerticalTimes = int.Parse(scroll.Replace("V", ""));
                            }

                            long windowWidthValue = long.Parse(windowWidth);
                            long windowHeightValue = long.Parse(windowHeight);

                            for (int horizontalScrollCounter = 0; horizontalScrollCounter < scrollHorizontalTimes; horizontalScrollCounter++)
                            {
                                ExecuteScript("window.scrollBy(" + windowWidthValue + ", 0)");
                            }

                            for (int verticalScrollCounter = 0; verticalScrollCounter < scrollVerticalTimes; verticalScrollCounter++)
                            {
                                ExecuteScript("window.scrollBy(0, " + windowHeightValue + ")");
                            }
                        }

                        var elementLocationX = excelRows[rowsCounter][2];
                        var elementLocationY = excelRows[rowsCounter][3];
                        var isHomepage = excelRows[rowsCounter][5].Equals("Y");
                        var isRepeated = excelRows[rowsCounter][6].Equals("Y");

                        if (isRepeated)
                        {
                            var elementFoundAtLocationCorrectly = false;

                            var elements = string.IsNullOrEmpty(containerClass) ? FindElements(loc) : FindElement(By.ClassName(containerClass)).FindElements(loc);
                            elements = elements.Where(e => e.Location.X.ToString().Equals(elementLocationX) && e.Location.Y.ToString().Equals(elementLocationY)).ToList();

                            if (elements.Count > 0)
                            {
                                if (elements.Count == 1)
                                {
                                    elementFoundAtLocationCorrectly = true;
                                }
                            }

                            if (!elementFoundAtLocationCorrectly)
                            {
                                LogInfo("", "Wrong Location: Element with Locator \"" + locatorStr + "\" was not Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\"", false, true);
                                LogIssue("", "Element with Locator \"" + locatorStr + "\" was not Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\"\r\nPlease Check Screenshot \"" + TakeScreenshot() + "\"", false, true);
                            }
                            else
                                LogInfo("", "Correct Location: Element with Locator \"" + locatorStr + "\" was Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\"", false, true);
                        }
                        else
                        {
                            try
                            {
                                var element = string.IsNullOrEmpty(containerClass) ? FindElement(loc) : FindElement(By.ClassName(containerClass)).FindElement(loc);

                                if (element.Location.X.ToString().Equals(elementLocationX) && element.Location.Y.ToString().Equals(elementLocationY))
                                {
                                    LogInfo("", "Correct Location: Element with Locator \"" + locatorStr + "\" was Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\"", false, true);
                                }
                                else
                                {
                                    LogInfo("", "Wrong Location: Element with Locator \"" + locatorStr + "\" was Expected to be Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\", but was Found at Location \"X: " + element.Location.X.ToString() + ", Y: " + element.Location.Y.ToString() + "\"", false, true);
                                    LogIssue("", "Element with Locator \"" + locatorStr + "\" was Expected to be Found at Location \"X: " + elementLocationX + ", Y: " + elementLocationY + "\", but was Found at Location \"X: " + element.Location.X.ToString() + ", Y: " + element.Location.Y.ToString() + "\"\r\nPlease Check Screenshot \"" + TakeScreenshot() + "\"", false, true);
                                }
                            }
                            catch (NoSuchElementException)
                            {
                                LogIssue("", "Element with Locator \"" + locatorStr + "\" was not Found in the Page\r\nPlease Check Screenshot \"" + TakeScreenshot() + "\"", false, true);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Fill excelSheetInputs with Elements that will be added to the UIElements.xlsx
        /// </summary>
        /// <param name="excelSheetInputs">List of Inputs that will be filled in the Excel Sheet UIElements.xlsx</param>
        /// <param name="rowNumber">Row Number of Excel Sheet</param>
        /// <param name="fileName">File Name of Excel Sheet</param>
        /// <param name="elements">List of Elements Found in the Current Page</param>
        /// <param name="isHomepage">Boolean Identifying if the Current Page is Homepage as it may have Unique Handling</param>
        /// <param name="windowHeight">Inner Window Height to Check if Vertical Scroll is Required</param>
        /// <param name="windowWidth">Inner Window Width to Check if Horizontal Scroll is Required</param>
        /// <returns>Row Number</returns>
        private uint GetElementsLocators(ref List<ExcelSheetInput> excelSheetInputs, uint rowNumber, string fileName, IList<IWebElement> elements, bool isHomepage, long windowHeight, long windowWidth)
        {
            try
            {
                var dontIncreaseRowNumber = false;

                for (int elementsCounter = 0; elementsCounter < elements.Count; elementsCounter++)
                {
                    if (elements[elementsCounter].Displayed && elements[elementsCounter].FindElements(By.XPath("*")).Count == 0)
                    {
                        var fillCells = false;

                        if (elementsCounter == elements.Count - 1)
                            dontIncreaseRowNumber = true;
                        var attributesList = (ReadOnlyCollection<object>)ExecuteScript("return arguments[0].attributes", elements[elementsCounter]);

                        var attributes = GetElementAttributes(attributesList);

                        var elementX = elements[elementsCounter].Location.X;
                        var elementY = elements[elementsCounter].Location.Y;

                        var scrollVertical = false;
                        var scrollVerticalTimes = 0;
                        var scrollHorizontal = false;
                        var scrollHorizontalTimes = 0;

                        while (elementY > windowHeight)
                        {
                            ExecuteScript("window.scrollBy(0, " + windowHeight + ")");

                            elementX = elements[elementsCounter].Location.X;
                            elementY = elements[elementsCounter].Location.Y;

                            scrollVertical = true;
                            scrollVerticalTimes++;
                        }

                        while (elementX > windowWidth)
                        {
                            ExecuteScript("window.scrollBy(" + windowWidth + ", 0)");

                            elementX = elements[elementsCounter].Location.X;
                            elementY = elements[elementsCounter].Location.Y;

                            scrollHorizontal = true;
                            scrollHorizontalTimes++;
                        }

                        var elementValue = elements[elementsCounter].GetAttribute("value");
                        var elementText = GetTextOf(elements[elementsCounter]);

                        var uniqueAttributeFound = false;

                        for (int attributesCounter = 0; attributesCounter < attributes.Count; attributesCounter++)
                        {
                            var nodeName = attributes[attributesCounter].Split(':')[0];
                            var nodeValue = attributes[attributesCounter].Split(':')[1];

                            if (!string.IsNullOrEmpty(nodeName) && !string.IsNullOrEmpty(nodeValue))
                                if (FindElements(Full(nodeName, nodeValue)).Count == 1)
                                {
                                    uniqueAttributeFound = true;
                                    excelSheetInputs.Add(new ExcelSheetInput(fileName, attributes[attributesCounter], rowNumber, "B"));
                                    break;
                                }
                        }

                        if (!uniqueAttributeFound)
                        {
                            var tagName = elements[elementsCounter].TagName;

                            if (!string.IsNullOrEmpty(elementValue) && FindElements(Full("value", elementValue)).Count == 1)
                            {
                                excelSheetInputs.Add(new ExcelSheetInput(fileName, "value:" + elementValue, rowNumber, "B"));
                                uniqueAttributeFound = true;
                            }
                            else if (!string.IsNullOrEmpty(elementText) && FindElements(XPathMaker(tagName, "text", elementText)).Count == 1)
                            {
                                excelSheetInputs.Add(new ExcelSheetInput(fileName, "text:" + tagName + ":" + elementText, rowNumber, "B"));
                                uniqueAttributeFound = true;
                            }
                        }

                        if (uniqueAttributeFound)
                        {
                            fillCells = true;
                        }
                        else
                        {
                            var uniqueParentFound = false;
                            for (int attributesCounter = 0; attributesCounter < attributes.Count && !uniqueParentFound; attributesCounter++)
                            {
                                var nodeName = attributes[attributesCounter].Split(':')[0];
                                var nodeValue = attributes[attributesCounter].Split(':')[1];

                                if (!string.IsNullOrEmpty(nodeName) && !string.IsNullOrEmpty(nodeValue))
                                {
                                    var parent = GetParentElement(elements[elementsCounter]);
                                    var counter = 0;

                                    while (!uniqueParentFound && counter < 10)
                                    {
                                        var parentAttributesList = (ReadOnlyCollection<object>)ExecuteScript("return arguments[0].attributes", parent);
                                        var parentAttributes = GetElementAttributes(parentAttributesList);

                                        for (int parentAttributesCounter = 0; parentAttributesCounter < parentAttributes.Count; parentAttributesCounter++)
                                        {
                                            var parentNodeName = parentAttributes[parentAttributesCounter].Split(':')[0];
                                            var parentNodeValue = parentAttributes[parentAttributesCounter].Split(':')[1];

                                            if (!string.IsNullOrEmpty(parentNodeName) && !string.IsNullOrEmpty(parentNodeValue))
                                                if (FindElements(Full(parentNodeName, parentNodeValue)).Count == 1)
                                                {
                                                    uniqueParentFound = true;
                                                    excelSheetInputs.Add(new ExcelSheetInput(fileName, "CssSelector:[" + parentNodeName + "='" + parentNodeValue + "'] [" + nodeName + "='" + nodeValue + "']", rowNumber, "B"));
                                                    break;
                                                }
                                        }

                                        parent = GetParentElement(elements[elementsCounter]);
                                        counter++;
                                    }
                                }
                            }

                            if (uniqueParentFound)
                            {
                                fillCells = true;
                            }
                            else
                            {
                                var elementAdded = false;
                                for (int attributesCounter = 0; attributesCounter < attributes.Count; attributesCounter++)
                                {
                                    var nodeName = attributes[attributesCounter].Split(':')[0];
                                    var nodeValue = attributes[attributesCounter].Split(':')[1];

                                    if (!string.IsNullOrEmpty(nodeName) && !string.IsNullOrEmpty(nodeValue))
                                    {
                                        elementAdded = true;
                                        excelSheetInputs.Add(new ExcelSheetInput(fileName, attributes[attributesCounter], rowNumber, "B"));
                                        break;
                                    }
                                }

                                if (elementAdded)
                                {
                                    fillCells = true;
                                }
                            }
                        }

                        if (fillCells)
                        {
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, "", rowNumber, "A"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, elementX.ToString(), rowNumber, "C"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, elementY.ToString(), rowNumber, "D"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, ((!scrollVertical && !scrollHorizontal) ? "N" : (scrollHorizontal ? (scrollHorizontalTimes + "H") : "") + ((scrollVertical && scrollHorizontal) ? ":" : "") + (scrollVertical ? (scrollVerticalTimes + "V") : "")), rowNumber, "E"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, isHomepage ? "Y" : "N", rowNumber, "F"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, "Y", rowNumber, "G"));
                            excelSheetInputs.Add(new ExcelSheetInput(fileName, TakeScreenshot(), rowNumber, "H"));

                            rowNumber++;
                        }
                    }

                    if (!dontIncreaseRowNumber && elementsCounter.Equals(elements.Count - 1))
                        rowNumber++;
                }

                return rowNumber;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        /// <summary>
        /// Get List of Attributes in Following Format: {"attributeName:attributeValue", "attributeName:attributeValue"}
        /// </summary>
        /// <param name="attributesList">List of Attributes of an Element</param>
        /// <returns>List of Attributes with Specific Format</returns>
        private List<string> GetElementAttributes(ReadOnlyCollection<object> attributesList)
        {
            try
            {
                var attributes = new List<string>();

                for (int attributesCounter = 0; attributesCounter < attributesList.Count; attributesCounter++)
                {
                    var attribute = (Dictionary<string, object>)attributesList[attributesCounter];
                    var nodeName = attribute["nodeName"].ToString();
                    var nodeValue = attribute["nodeValue"].ToString().Replace("'", "\\'");

                    attributes.Add(nodeName + ":" + nodeValue);
                }

                return attributes;
            }
            catch (Exception e)
            {
                LogIssue("", "Exception occurred in function \"" + GetCurrentMethod() + "\". Please check attached screenshot: " + TakeScreenshot(), false, true);
                throw e;
            }
        }

        #endregion

        #region Enums

        /// <summary>
        /// List of Conditions that are used in WaitFor Function
        /// </summary>
        public enum Condition
        {
            /// <summary>
            /// Wait till given element exists in html
            /// </summary>
            Exist,
            /// <summary>
            /// Wait till given element doesn't exist in html
            /// </summary>
            NotExist,
            /// <summary>
            /// Wait till given element is visible
            /// </summary>
            Visible,
            /// <summary>
            /// Wait till given element is invisible
            /// </summary>
            Invisible,
            /// <summary>
            /// Wait till given element is enabled
            /// </summary>
            Enabled,
            /// <summary>
            /// Wait till given element is clickable
            /// </summary>
            Clickable,
            /// <summary>
            /// Wait till frame is available and switch to it
            /// </summary>
            FrameAvailabilityAndSwitchToIt,
            /// <summary>
            /// Wait till select list of given element contains more than 1 option
            /// </summary>
            SelectListLoaded,
            /// <summary>
            /// Wait till page ready state is interactive or complete
            /// </summary>
            PageReadyState,
            /// <summary>
            /// Wait till all elements with given locator exist in html
            /// </summary>
            PresenceOfAllElementsLocatedBy,
            /// <summary>
            /// Wait till all elements with given locator are visible
            /// </summary>
            VisibilityOfAllElementsLocatedBy,
            /// <summary>
            /// Wait till given element is selected
            /// </summary>
            ElementToBeSelected,
            /// <summary>
            /// Wait till an alert is present
            /// </summary>
            AlertIsPresent
        }

        /// <summary>
        /// List of Possible Tag Names in HTML. Used in XPathMaker Function
        /// </summary>
        public enum HtmlTag
        {
            input,
            textarea,
            textField,
            a,
            link,
            button,
            div,
            span,
            label,
            ul,
            li,
            select,
            dropdown,
            option,
            table,
            thead,
            tbody,
            tr,
            td,
            img,
            strong,
            h2,
            h3
        }

        /// <summary>
        /// List of Possible Attributes in HTML. Used in XPathMaker and CssSelectorMaker Functions and their derivatives
        /// </summary>
        public enum Attribute
        {
            id,
            Class,
            href,
            type,
            name,
            value,
            For,
            src,
            text,
            title,
            style,
            tabindex,
            role,
            tooltip,
            ng_click,
            ng_dblclick,
            ng_controller,
            ng_hide,
            ng_show,
            ng_include,
            ng_form,
            ng_init,
            ng_class,
            ng_if,
            ng_repeat,
            ng_model,
            ng_switch,
            aria_controls,
            aria_selected,
            aria_owns,
            aria_labelledby,
            formcontrolname,
            formarrayname,
            wj_part
        }

        /// <summary>
        /// List of Type Attribute Values in HTML
        /// </summary>
        public enum Type
        {
            submit,
            button,
            checkbox,
            radio,
            option,
            text
        }

        /// <summary>
        /// List of Attribute Value Types. Used in CssSelectorMaker Function and its derivatives
        /// </summary>
        public enum AttributeValueType
        {
            FullName,
            PrefixOfName,
            SuffixOfName,
            PartOfName
        }

        /// <summary>
        /// List of Browser Types. Used in Initialize Function
        /// </summary>
        public enum Browser
        {
            ie,
            chrome,
            firefox,
            edge
        }

        /// <summary>
        /// List of File Types that can be Uploaded. Used in UploadAttachments Function
        /// </summary>
        public enum FileType
        {
            Pdf,
            Docx,
            Jpg,
            Txt
        }

        #endregion

    }
}
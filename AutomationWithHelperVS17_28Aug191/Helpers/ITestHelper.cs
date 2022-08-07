using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace AutomationWithHelperVS17_28Aug191
{
    /// <summary>
    /// Automation Test Helper including many functions to ease writing Selenium WebDriver Test Automation Code
    /// </summary>
    public interface ITestHelper
    {
        /// <summary>
        /// Accept alert if present
        /// </summary>
        void AcceptAlert();
        /// <summary>
        /// Actions: Returns Actions Object for the provided _driver
        /// Can be used for Clicking, Double Clicking, Focus On Element by Moving to It, Send Keys to an Element and other Functions
        /// How to Use?
        ///     example:
        ///         Actions().MoveToElement(element).Click().Perform();
        /// </summary>
        /// <returns>Actions Object</returns>
        Actions Actions();
        /// <summary>
        /// Backend Login: Login to any Backend Portal that requires Windows Authentication
        /// Works with IE and Chrome
        /// </summary>
        /// <param name="backendUrl">Url to go to</param>
        /// <param name="domain">Domain of the User to be sent to Windows Authentication Popup</param>
        /// <param name="username">Username to be sent to Windows Authentication Popup</param>
        /// <param name="password">Password to be sent to Windows Authentication Popup</param>
        void BackendLogin(string backendUrl, string domain, string username, string password);
        /// <summary>
        /// Check Status: Finds the status that needs to be checked using the By locator and compares it to the given expected status
        /// </summary>
        /// <param name="appStatusLocator">By Locator of the Status Element in the Page</param>
        /// <param name="expectedStatus">String of the Status that is Expected to Appear</param>
        bool CheckStatus(By appStatusLocator, string expectedStatus);
        /// <summary>
        /// Check Status: Finds the status that needs to be checked using the given element and compares it to the given expected status
        /// </summary>
        /// <param name="appStatusElement">Status Element in the Page</param>
        /// <param name="expectedStatus">String of the Status that is Expected to Appear</param>
        bool CheckStatus(IWebElement appStatusElement, string expectedStatus);
        /// <summary>
        /// Check Status in Grid: Check given Status of a given Request ID in a Grid where there Request ID is a Link
        /// Concern: to be used after searching for request in Grid
        /// </summary>
        /// <param name="requestId">Request ID whose Status should be checked</param>
        /// <param name="statusIndexinGrid">Index of the cell in the grid row that contains the status that will be checked</param>
        /// <param name="expectedRequestStatus">Expected Request Status that should be compared to the found Status</param>
        bool CheckStatusInGrid(string requestId, int statusIndexinGrid, string expectedRequestStatus);
        /// <summary>
        /// Clears text in a given text field/area
        /// </summary>
        /// <param name="locator">By locator of the element whose text should be cleared</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void Clear(By locator, string locatorDescription = "");
        /// <summary>
        /// Clears text in a given text field/area
        /// </summary>
        /// <param name="element">Element whose text should be cleared</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void Clear(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Click on Element: Tries to Click on an Element using its By Locator and waits if there is InvalidOperationException or StaleElementReferenceException
        /// Clicks on Element using javascript if usingJavascript is true and using .Click() Function if usingJavascript is false
        /// </summary>
        /// <param name="locator">By Locator of Element that needs to be Clicked</param>
        /// <param name="usingJavascript">Click on the button using Javascript or not</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void ClickOn(By locator, bool usingJavascript, string locatorDescription = "");
        /// <summary>
        /// Click on Element: Tries to click on an element using the browser's Javascript or not depending on the boolean value
        /// and waits if there is InvalidOperationException or StaleElementReferenceException
        /// </summary>
        /// <param name="element">Element that needs to be Clicked</param>
        /// <param name="usingJavascript">Click on the button using Javascript or not</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void ClickOn(IWebElement element, bool usingJavascript, string elementDescription = "");
        /// <summary>
        /// Click on Multiple Elements : Works on Navigating to my services, my tasks, my pending tasks or any threads of links
        /// Usage Example: ClickOnMultipleLinks(new By[] {By.Id("link1Id"), By.Id("link2Id")});
        /// </summary>
        /// <param name="links">links[]: Array of Links Text that will be clicked by "By.LinkText"</param>
        void ClickOnMultipleElements(By[] links);
        /// <summary>
        /// Close and Dispose the driver if opened
        /// </summary>
        void CloseDriver();
        /// <summary>
        /// Closes the current tab and switches back to the first tab
        /// </summary>
        void CloseTab();
        /// <summary>
        /// CssSelector Maker: Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// </summary>
        /// <param name="attribute">Attribute to be found, coming from the enum Attribute (e.g. id = "")</param>
        /// <param name="attributeValueType">Type of the attribute value to be found, coming from the enum AttributeValueType (e.g. could be full, prefix. suffix or part of the value)</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By CssSelectorMaker(TestHelper.Attribute attribute, TestHelper.AttributeValueType attributeValueType, string attributeValue);
        /// <summary>
        /// CssSelector Maker: Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// This is an Overload Function to be used in case of composite Attributes
        /// </summary>
        /// <param name="attribute">Attribute to be found, entered as String (e.g. data-bind = "")</param>
        /// <param name="attributeValueType">Type of the attribute value to be found, coming from the enum AttributeValueType (e.g. could be full, prefix. suffix or part of the value)</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By CssSelectorMaker(string attribute, TestHelper.AttributeValueType attributeValueType, string attributeValue);
        /// <summary>
        /// Decodes a string from Base64
        /// </summary>
        /// <param name="encodedPassword">Encoded string that should be decoded</param>
        /// <returns>Decoded string</returns>
        string DecodePassword(string encodedPassword);
        /// <summary>
        /// Dismiss alert if present
        /// </summary>
        void DismissAlert();
        /// <summary>
        /// Double Click on an element using its locator
        /// </summary>
        /// <param name="locator">Locator of the element to click on twice</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void DoubleClickOn(By locator, string locatorDescription = "");
        /// <summary>
        /// Double Click on an element
        /// </summary>
        /// <param name="element">Element to click on twice</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void DoubleClickOn(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Encodes a string in Base64
        /// </summary>
        /// <param name="password">String to be encoded</param>
        /// <returns>Encoded string</returns>
        string EncodePassword(string password);
        /// <summary>
        /// Executes Javascript in the current page
        /// Usage Example: ExecuteScript("arguments[0].scrollIntoView();", FindElement(By.Id("elementToScrollToId")));
        /// </summary>
        /// <param name="script">Script to be executed in the page</param>
        /// <param name="args">Parameters used in the script</param>
        /// <returns></returns>
        object ExecuteScript(string script, params object[] args);
        /// <summary>
        /// Export UI Elements to Excel Sheet named "UIElements.xlsx"
        /// First Row contains: WindowWidth(X) | WindowHeight(Y)
        /// Second Row contains: URL | Locator | Location X | Location Y | Scroll Direction | Is Homepage | Is Repeated | Screenshot
        /// Locator is defined as follows: "attributeName:AttributeValue"
        /// </summary>
        /// <param name="baseUrl">Base Url of the Portal being Tested</param>
        void ExportUIElementsAndPositionsToExcelSheet(string baseUrl);
        /// <summary>
        /// Extract Number from Given Text
        /// </summary>
        /// <param name="text">Text containing the number</param>
        /// <returns>Number found in text</returns>
        string ExtractNumberFromText(string text);
        /// <summary>
        /// Find Certain Parent Of Element: Loops through the Parents of the Given Element to Find element with the given tag name.
        /// Example: Can be used in workspace in requests table for example where the locator is the request id and we want to get the row containing the request (hence the default value of the tagName parameter is tr) to get the status related to this request from the same row and assert on it
        /// </summary>
        /// <param name="locator">Locator of the Element whose Parent is to be Found</param>
        /// <param name="tagName">Tag Name of Parent to be Found. Default "tr" (table row in html)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>Parent Containing the Given Element or Null in case no Parent with the provided tag name was Found after Checking All Parents Till the HTML Element</returns>
        IWebElement FindCertainParentOf(By locator, string tagName = "tr", string locatorDescription = "");
        /// <summary>
        /// Find Certain Parent Of Element: Loops through the Parents of the Given Element to Find element with the given tag name
        /// </summary>
        /// <param name="element">Element whose Parent is to be Found</param>
        /// <param name="tagName">Tag Name of Parent to be Found. Default "tr"</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>Parent Containing the Given Element or Null in case no Parent with the provided tag name was Found after Checking 10 Parents of the Element</returns>
        IWebElement FindCertainParentOf(IWebElement element, string tagName = "tr", string elementDescription = "");
        /// <summary>
        /// Find a parent of an element that contains value of a specific attribute
        /// </summary>
        /// <param name="locator">Locator of the element whose parent will be located</param>
        /// <param name="attr">Attribute found in the parent that should be located</param>
        /// <param name="val">Value of the attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The located parent element with the attribute and value provided</returns>
        IWebElement FindCertainParentWithAttribute(By locator, string attr, string val, string locatorDescription = "");
        /// <summary>
        /// Find a parent of an element that contains value of a specific attribute
        /// </summary>
        /// <param name="element">Element whose parent will be located</param>
        /// <param name="attr">Attribute found in the parent that should be located</param>
        /// <param name="val">Value of the attribute</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The located parent element with the attribute and value provided</returns>
        IWebElement FindCertainParentWithAttribute(IWebElement element, string attr, string val, string elementDescription = "");
        /// <summary>
        /// Finds the first element using the given locator
        /// Usage Example: FindElement(By.Id(""));
        /// </summary>
        /// <param name="locator">By locator of the element</param>
        /// <returns>Element found</returns>
        IWebElement FindElement(By locator);
        /// <summary>
        /// Finds all elements using the given locator
        /// Usage Example: FindElements(By.Id(""));
        /// </summary>
        /// <param name="locator">By locator of the elements</param>
        /// <returns>List of elements found</returns>
        IList<IWebElement> FindElements(By locator);
        /// <summary>
        /// Finds the first element that is a link that is a child of the given element
        /// Usage Example: FindLinkInElement(FindElement(By.Id("")));
        /// </summary>
        /// <param name="element">Parent element that contains the link element as a child</param>
        /// <returns>Element of the link</returns>
        IWebElement FindLinkInElement(IWebElement element);
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
        void FrontendLogin(string username, string password, By[] loginLocators);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the full value of the attribute
        /// Usage Example: FindElement(Full(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Full(TestHelper.Attribute attr, string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the full value of the attribute
        /// Usage Example: FindElement(Full("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Full(string attr, string val);
        /// <summary>
        /// Generates an Egyptian National ID
        /// </summary>
        /// <param name="olderThan21">Default value is true to generate a national id of older than 21 years person</param>
        /// <param name="isMale">Default value is true to generate a national id for a isMale person</param>
        /// <returns>The generated national id</returns>
        string GenerateEgyptianNationalId(bool olderThan21 = true, bool isMale = true);
        /// <summary>
        /// Detect Browser Errors Found in the Console
        /// </summary>
        void GetBrowserErrors();
        /// <summary>
        /// Get Child Elements: gets the direct children of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Parent Element that the function should find its Children</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Child Elements</returns>
        IList<IWebElement> GetChildElements(By locator, string locatorDescription = "");
        /// <summary>
        /// Get Child Elements: gets the children of the element sent as parameter
        /// </summary>
        /// <param name="element">Parent Element that the function should find its Children</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Child Elements</returns>
        IList<IWebElement> GetChildElements(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Get dom content loaded time in seconds at any given time
        /// </summary>
        /// <returns>dom content loaded time in seconds</returns>
        string GetDomContentLoadedTime();
        /// <summary>
        /// Get Driver Instance
        /// </summary>
        /// <returns>Driver Instance</returns>
        IWebDriver GetDriver();
        /// <summary>
        /// Get Parent Element: gets the parent of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Child Element that the function should find its Parent</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Parent Element</returns>
        IWebElement GetParentElement(By locator, string locatorDescription = "");
        /// <summary>
        /// Get Parent Element: gets the parent of the element sent as parameter
        /// </summary>
        /// <param name="element">Child Element that the function should find its Parent</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Parent Element</returns>
        IWebElement GetParentElement(IWebElement element, string elementDescription = "");
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
        string GetRequestId(By requestIdLocator, string separatorBefore, string separatorAfter);
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
        string GetRequestId(string requestIdContainerText, string separatorBefore, string separatorAfter);
        /// <summary>
        /// Gets the options of the select list
        /// Usage Example: GetSelectOptions(By.Id("selectElementId"));
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        IList<IWebElement> GetSelectOptions(By locator, string locatorDescription = "");
        /// <summary>
        /// Gets the options of the select list
        /// Usage Example: GetSelectOptions(FindElement(By.Id("selectElementId")));
        /// </summary>
        /// <param name="element">The select element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        IList<IWebElement> GetSelectOptions(IWebElement element, string locatorDescription = "");
        /// <summary>
        /// Get Text of: gets the text of the element sent as parameter
        /// </summary>
        /// <param name="locator">Locator of the Element whose text will be returned</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>The Text of the Element</returns>
        string GetTextOf(By locator, string locatorDescription = "");
        /// <summary>
        /// Get Text of: gets the text of the element sent as parameter
        /// </summary>
        /// <param name="element">Element whose text will be returned</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        /// <returns>The Text of the Element</returns>
        string GetTextOf(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Get Url of the Current Tab/Window
        /// </summary>
        /// <returns></returns>
        string GetUrl();
        /// <summary>
        /// Get Wait Instance
        /// </summary>
        /// <returns>Wait Instance</returns>
        WebDriverWait GetWait();
        /// <summary>
        /// Navigate to given URL
        /// </summary>
        /// <param name="url">URL to go to</param>
        void GoToUrl(string url);
        /// <summary>
        /// Hover on an element using its locator
        /// Note: Mouse will not move in the page but the action will be executed
        /// </summary>
        /// <param name="locator">Locator of element to hover on</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void HoverOn(By locator, string locatorDescription = "");
        /// <summary>
        /// Hover on an element
        /// Note: Mouse will not move in the page but the action will be executed
        /// </summary>
        /// <param name="element">Element to hover on</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void HoverOn(IWebElement element, string elementDescription = "");
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
        IWebDriver Initialize(TestHelper.Browser browser, out WebDriverWait wait, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Default, double timeToWaitInMinutes = 2, bool useProfile = false, bool runHeadless = false);
        /// <summary>
        /// Checks if the driver is open or not
        /// </summary>
        /// <returns></returns>
        bool IsDriverOpen();
        /// <summary>
        /// Is Element Present: Checks whether an Element is Present or Not
        /// </summary>
        /// <param name="locator">By Locator of Element that needs to be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case element is present and False in case element is not present</returns>
        bool IsElementPresent(By locator, string locatorDescription = "");
        /// <summary>
        /// Is Element Visible: Checks whether an Element is Visible or Not
        /// </summary>
        /// <param name="locator">By Locator of Element that need to be checked</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case element is visible and False in case element is not visible</returns>
        bool IsElementVisible(By locator, string locatorDescription = "");
        /// <summary>
        /// Is Text Invisible: Checks that the provided element's text doesn't contain the provided text
        /// </summary>
        /// <param name="locator">Locator that the text should not be found in</param>
        /// <param name="text">Text that should not be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case text is not found in the element and False in case text is found</returns>
        bool IsTextInvisible(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Is Text Invisible: Checks whether the provided text is visible in the page or not
        /// </summary>
        /// <param name="text">Text that should be invisible</param>
        /// <returns>True in case text is invisible and False in case it is visible</returns>
        bool IsTextInvisible(string text);
        /// <summary>
        /// Is Text Visible: Checks that the provided element's text contains the provided text
        /// </summary>
        /// <param name="locator">Locator that the text should be found in</param>
        /// <param name="text">Text that needs to be found</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>True in case text is found in the element and False in case text is not found</returns>
        bool IsTextVisible(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Is Text Visible: Checks whether some Text is Visible in the page or Not
        /// </summary>
        /// <param name="text">text that needs to be found</param>
        /// <returns>True in case text is visible and False in case text is not visible</returns>
        bool IsTextVisible(string text);
        /// <summary>
        /// Builds the String used to find a link element whose title attribute contains certain text using the XPath Function
        /// Usage Example: FindElement(LinkTitleContains("linkTitle"));
        /// </summary>
        /// <param name="val">Value of the TItle Attribute that we are searching for (e.g. title="...AttributeValue...")</param>
        /// <returns>The By XPath Locator</returns>
        By LinkTitleContains(string val);
        /// <summary>
        /// Maximize the current browser window
        /// </summary>
        void MaximizeWindow();
        /// <summary>
        /// Opens a new empty tab and switch to it
        /// </summary>
        /// <param name="numberOfTabs">Number of tabs or windows open in the current driver after opening the new tab</param>
        void OpenNewTab(int numberOfTabs);
        /// <summary>
        /// Open Request from My Tasks Or My Requests With Search control: Works on opening the request or tasks
        /// from the grid of tasks using the search control where the Request ID is a Link
        /// </summary>
        /// <param name="requestId">The specific request ID that you want to open</param>
        /// <param name="statusIndexinGrid">Index of the cell in the grid row that contains the status that will be checked</param>
        /// <param name="requestTextBoxLocator">By Locator of the request ID textbox that is found in the Search Control</param>
        /// <param name="searchButtonLocator">By Locator of the search button that is found in the Search Control</param>
        /// <param name="expectedStatus">The expected status of the request in the Tasks Grid</param>
        void OpenRequestFromRequestsList(string requestId, int statusIndexinGrid, By requestTextBoxLocator, By searchButtonLocator, string expectedStatus);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be any part of the attribute
        /// Usage Example: FindElement(Part(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Part(TestHelper.Attribute attr, string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be any part of the attribute
        /// Usage Example: FindElement(Part("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Part(string attr, string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the start of the attribute
        /// Usage Example: FindElement(Prefix(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Prefix(TestHelper.Attribute attr, string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the start of the attribute
        /// Usage Example: FindElement(Prefix("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="AttributeValue...")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Prefix(string attr, string val);
        /// <summary>
        /// Refreshes the current page
        /// </summary>
        void RefreshPage();
        /// <summary>
        /// Right click on element using its locator
        /// </summary>
        /// <param name="locator">Locator of the element to context click on</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void RightClickOn(By locator, string locatorDescription = "");
        /// <summary>
        /// Right click on element
        /// </summary>
        /// <param name="element">Element to context click on</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void RightClickOn(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Scripts: Returns IJavaScriptExecutor Object for the provided _driver
        /// Can be used for Clicking, Scrolling or Executing any Javascript Code in the Browser
        /// How to Use?
        ///     example:
        ///         ExecuteScript("arguments[0].click();", element);
        /// </summary>
        /// <returns>IJavaScriptExecutor Object</returns>
        IJavaScriptExecutor Scripts();
        /// <summary>
        /// Scroll to Element: scrolls to an element in the page (horizontal or vertical)
        /// <br>Automatically detects if there is a Custom Scroll Bar (mCustomScrollbar) in the page and uses the correct scroll function in this case</br>
        /// </summary>
        /// <param name="locator">Locator of Element that should be scrolled to</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void ScrollTo(By locator, string locatorDescription = "");
        /// <summary>
        /// Scroll to Element: scrolls to an element in the page (horizontal or vertical)
        /// <br>Automatically detects if there is a Custom Scroll Bar (mCustomScrollbar) in the page and uses the correct scroll function in this case</br>
        /// </summary>
        /// <param name="element">Element that should be scrolled to</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void ScrollTo(IWebElement element, string elementDescription = "");
        /// <summary>
        /// Search Request: Enters the Request ID to the text box of request ID in the search control and Clicks on the Search button
        /// </summary>
        /// <param name="requestId">The specific request ID that you want to search for</param>
        /// <param name="requestTextBoxLocator">By Locator of the request ID textbox that is found in the Search Control</param>
        /// <param name="searchButtonLocator">By Locator of the search button that is found in the Search Control</param>
        void SearchRequest(string requestId, By requestTextBoxLocator, By searchButtonLocator);
        /// <summary>
        /// Select option from drop down list by its index
        /// Usage Example: SelectByIndex(By.Id("selectElementId"), 1);
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="index">Number of the option in the list that we are searching for (e.g. <option>option1</option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void SelectByIndex(By locator, int index, string locatorDescription = "");
        /// <summary>
        /// Select option from drop down list by its index
        /// Usage Example: SelectByIndex(FindElement(By.Id("selectElementId")), 1);
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="index">Number of the option in the list that we are searching for (e.g. <option>option1</option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void SelectByIndex(IWebElement element, int index, string elementDescription = "");
        /// <summary>
        /// Select option from drop down list by its text
        /// Usage Example: SelectByText(By.Id("selectElementId"), "optionText");
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="text">Value of the option's Text Attribute that we are searching for (e.g. <option>text</option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void SelectByText(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Select option from drop down list by its text
        /// Usage Example: SelectByText(FindElement(By.Id("selectElementId")), "optionText");
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="text">Value of the option's Text Attribute that we are searching for (e.g. <option>text</option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void SelectByText(IWebElement element, string text, string elementDescription = "");
        /// <summary>
        /// Select option from drop down list by value
        /// Usage Example: SelectByValue(By.Id("selectElementId"), "optionValue");
        /// </summary>
        /// <param name="locator">By locator of the select element</param>
        /// <param name="value">Value of the option's Value Attribute that we are searching for (e.g. <option value="value"></option>)</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void SelectByValue(By locator, string value, string locatorDescription = "");
        /// <summary>
        /// Select option from drop down list by value
        /// Usage Example: SelectByValue(FindElement(By.Id("selectElementId")), "optionValue");
        /// </summary>
        /// <param name="element">Element of the select element</param>
        /// <param name="value">Value of the option's Value Attribute that we are searching for (e.g. <option value="value"></option>)</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void SelectByValue(IWebElement element, string value, string elementDescription = "");
        /// <summary>
        /// Send Keys to a given text field/area
        /// </summary>
        /// <param name="locator">By locator of the element that the text should be sent to</param>
        /// <param name="text">Text that should be entered in the text field</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void SendKeys(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Send Keys to a given text field/area
        /// </summary>
        /// <param name="element">Element that the text should be sent to</param>
        /// <param name="text">Text that should be entered in the text field</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void SendKeys(IWebElement element, string text, string elementDescription = "");
        /// <summary>
        /// Send Mail: Sends Mail through Linkdev SMTP Server "10.2.2.10" Port 25
        /// While Attaching Log Files and Screenshots if any were Taken
        /// Logs Message in "EmailIssuesLogFile.txt" in case Sending Mail Fails
        /// </summary>
        /// <param name="subject">Subject of the Mail</param>
        /// <param name="body">Body of the Mail</param>
        /// <param name="toMails">List of Emails that the mail will be sent to</param>
        /// <param name="filesToBeUploadedList">List of files that need to be uploaded in the email, should be in the output directory folder</param>
        void SendMail(string subject, string body, StringCollection toMails, params string[] filesToBeUploadedList);
        /// <summary>
        /// Set Date: Takes the Date that should be selected and the elements of the Calendar Popup and selects the date
        /// </summary>
        /// <param name="date">Date that should be selected</param>
        /// <param name="calenderButton">button element to be clicked to open the calendar popup</param>
        /// <param name="calendarPopup">Calendar Popup Element</param>
        /// <param name="previousMonth">Previous Month button Element</param>
        /// <param name="nextMonth">Next Month button Element</param>
        /// <param name="dateFieldDescription">Description of the Date Element for Logging</param>
        void SetDate(DateTime date, By calenderButton, By calendarPopup, By previousMonth, By nextMonth, string dateFieldDescription = "");
        /// <summary>
        /// Set Driver Instance to gived driver
        /// </summary>
        /// <param name="driver"></param>
        void SetDriver(IWebDriver driver);
        /// <summary>
        /// Set Wait Instance to given wait
        /// </summary>
        /// <param name="wait"></param>
        void SetWait(WebDriverWait wait);
        /// <summary>
        /// Thread Sleep for given milliseconds
        /// </summary>
        /// <param name="millisecondsTimeout">Milliseconds to sleep</param>
        void SleepFor(int millisecondsTimeout);
        /// <summary>
        /// Builds the String used to find a span element whose text attribute contains certain text using the XPath Function
        /// Usage Example: FindElement(SpanTextContains("text"));
        /// </summary>
        /// <param name="val">Value of the Title Attribute that we are searching for (e.g. <span>...AttributeValue...</span>)</param>
        /// <returns>The By XPath Locator</returns>
        By SpanTextContains(string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the end of the attribute
        /// Usage Example: FindElement(Suffix(Attribute.href, "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, coming from the enum Attribute (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Suffix(TestHelper.Attribute attr, string val);
        /// <summary>
        /// Builds the String used to find an Element using the CssSelector Function and returns By CssSelector locator
        /// The given value should be the end of the attribute
        /// Usage Example: FindElement(Suffix("href", "/default.aspx"));
        /// </summary>
        /// <param name="attr">Attribute to be found, entered as string (e.g. value="")</param>
        /// <param name="val">Value of the Attribute that we are searching for (e.g. href="...AttributeValue")</param>
        /// <returns>The By CssSelector Locator</returns>
        By Suffix(string attr, string val);
        /// <summary>
        /// Instructs the driver to send future commands to the main document (outside of all frames)
        /// </summary>
        void SwitchToDefault();
        /// <summary>
        /// Instructs the driver to send future commands to the given frame
        /// the frame should be directly under the main document, not nested
        /// </summary>
        /// <param name="frameId">Frame ID or name to be selected</param>
        void SwitchToFrame(string frameId);
        /// <summary>
        /// Get Iframe Containing Certain Element: Finds and switches to the iframe that contains the element provided by its Locator
        /// <br>Note: The frames have to be on the same level under the root (default content), the function won't search in nested frames</br>
        /// </summary>
        /// <param name="locator">Locator of the element that will be searched for in the iframes</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        /// <returns>Returns the id of the iframe if found</returns>
        string SwitchToFrameOfElement(By locator, string locatorDescription = "");
        /// <summary>
        /// Instructs the driver to send future commands to the given frame which may be nested inside another frame
        /// </summary>
        /// <param name="frameId">Frame ID or name to be selected</param>
        void SwitchToNestedFrame(string frameId);
        /// <summary>
        /// Instructs the driver to send future commands to the parent frame of the current frame
        /// </summary>
        void SwitchToParentFrame();
        /// <summary>
        /// Instructs the driver to send future commands to the given window
        /// </summary>
        /// <param name="windowName">Window name to be selected</param>
        void SwitchToWindow(string windowName);
        /// <summary>
        /// Switch to window containing part of URL
        /// </summary>
        /// <param name="urlPart">Part of the URL that should be in the page</param>
        void SwitchToWindowWithUrlContaining(string urlPart);
        /// <summary>
        /// Switch to window Not containing part of URL
        /// </summary>
        /// <param name="urlPart">Part of the URL that shouldn't be in the page</param>
        void SwitchToWindowWithUrlNotContaining(string urlPart);
        /// <summary>
        /// Take Action: Takes action from a selection list using action value
        /// And enters comment in given comment text box locator if available else send it as null
        /// </summary>
        /// <param name="actionValue">Value of the Action that should be taken</param>
        /// <param name="actionControlLocator">Locator of the action selection list control</param>
        /// <param name="commentTextboxLocator">Locator of the comment text box</param>
        /// <param name="saveButtonLocator">Locator of the save button to be clicked after selecting action</param>
        void TakeAction(string actionValue, By actionControlLocator, By commentTextboxLocator, By saveButtonLocator);
        /// <summary>
        /// Take Screenshot: Takes a Screenshot of the Browser's Screen
        /// And Saved the Image in Folder "Screenshots" in the Project's Folder
        /// </summary>
        /// <param name="screenshotName">Screenshot Name, will be filled randomly if not provided</param>
        /// <returns>Returns Created Image Name or "No File Created" if there was an Exception while Taking the Screenshot</returns>
        string TakeScreenshot(string screenshotName = "");
        /// <summary>
        /// Upload Attachments: Finds all the Attachments using the given common part of the attribute value and uploads attachments to them
        /// <br>Note 1: May need to be overridden to add waits and checks that the file is uploaded</br>
        /// <br>Note 2: Function assumes that the uploader doesn't need to click on a button after selecting the file for it to be uploaded</br>
        /// </summary>
        /// <param name="commonFileLocator">Locator of the Common/Repeated Part of the Files</param>
        /// <param name="fileType">File Type to be Uploaded</param>
        void UploadAttachments(By commonFileLocator, TestHelper.FileType fileType);
        /// <summary>
        /// Validate UI Elements Found in the UIElements.xlsx Excel Sheet, Follows the Specific Formatting of the Sheet
        /// </summary>
        void ValidateUIElementsFromExcelSheet();
        /// <summary>
        /// Wait For: Element to meet the given condition
        /// <br>Usage Example: WaitFor(By.Id(""), Condition.Visible);
        /// <br>               WaitFor(null, Condition.PageReadyState);</br></br>
        /// </summary>
        /// <param name="locator">By locator of the element</param>
        /// <param name="condition">Condition to be satisfied</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitFor(By locator, TestHelper.Condition condition, string locatorDescription = "");
        /// <summary>
        /// Wait For Element to meet the given condition
        /// <br>Usage Example: WaitFor(FindElement(By.Id("")), Condition.Visible);
        /// <br>               WaitFor(null, Condition.PageReadyState);</br></br>
        /// </summary>
        /// <param name="element">Element to Wait For</param>
        /// <param name="condition">Condition to be satisfied</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void WaitFor(IWebElement element, TestHelper.Condition condition, string elementDescription = "");
        /// <summary>
        /// Wait for Page Ready State: Waits till the page stops loading using javascript executor to get the value of "readyState" property in the page
        /// </summary>
        void WaitForPageReadyState();
        /// <summary>
        /// Wait for tab or window to open by giving the expected number of tabs/windows that should be found in the driver
        /// </summary>
        /// <param name="numberOfTabs">Number of tabs or windows open in the current driver after opening or closing the tab or window</param>
        void WaitForTabToOpenOrClose(int numberOfTabs);
        /// <summary>
        /// Wait for provided text to not be visible in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should not contain the text</param>
        /// <param name="text">Text that should not be visible</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitForTextToBeInvisible(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Wait for provided text to not be visible in the page
        /// </summary>
        /// <param name="text">Text that should not be visible</param>
        void WaitForTextToBeInvisible(string text);
        /// <summary>
        /// Wait for provided text to be visible in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the text</param>
        /// <param name="text">Text that should be visible</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitForTextToBeVisible(By locator, string text, string locatorDescription = "");
        /// <summary>
        /// Wait for provided text to be visible in the page
        /// </summary>
        /// <param name="text">Text that should be visible</param>
        void WaitForTextToBeVisible(string text);
        /// <summary>
        /// Wait till the given attribute appears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the attribute</param>
        /// <param name="attr">Attribute that should be found in the element</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitTillAttributeAppears(By locator, TestHelper.Attribute attr, string locatorDescription = "");
        /// <summary>
        /// Wait till the given attribute appears in the element
        /// </summary>
        /// <param name="element">Element that should contain the attribute</param>
        /// <param name="attr">Attribute that should be found in the element</param>
        /// <param name="elementDescription">Description of the Element for Logging</param>
        void WaitTillAttributeAppears(IWebElement element, TestHelper.Attribute attr, string elementDescription = "");
        /// <summary>
        /// Wait till Element is Displayed: Uses explicit wait till the element is displayed
        /// </summary>
        /// <param name="element">Element that should be displayed</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitTillElementIsDisplayed(IWebElement element, string locatorDescription = "");
        /// <summary>
        /// Wait Till Page Loads: Waits till the loader element disappears.
        /// Default Number of Times to Wait is 2 because sometimes the Loader appears more than once before the page is fully loaded
        /// </summary>
        /// <param name="loaderLocator">By Locator for the Loader Element</param>
        /// /// <param name="numberOfWaits">Number of Times to Wait for the Loader to Disappear</param>
        void WaitTillPageLoad(By loaderLocator, int numberOfWaits = 2);
        /// <summary>
        /// Wait till the value of the given attribute appears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should contain the value of the attribute</param>
        /// <param name="attr">Attribute that should contain the value</param>
        /// <param name="val">Value that should appear in the Attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitTillValueOfAttributeAppears(By locator, TestHelper.Attribute attr, string val, string locatorDescription = "");
        /// <summary>
        /// Wait till the value of the given attribute disappears in the element
        /// </summary>
        /// <param name="locator">Locator of the element that should not contain the value of the attribute</param>
        /// <param name="attr">Attribute that should not contain the value</param>
        /// <param name="val">Value that should disappear from the Attribute</param>
        /// <param name="locatorDescription">Description of the Element for Logging</param>
        void WaitTillValueOfAttributeDisappears(By locator, TestHelper.Attribute attr, string val, string locatorDescription = "");
        /// <summary>
        /// XPath Maker: Builds the String used to find an Element using the XPath Function and returns By xpath locator
        /// </summary>
        /// <param name="htmlTag">HTML Tags  used to find the Attribute within, coming from the enum HTMLTag (e.g. <input> </input>)</param>
        /// <param name="attribute">Attribute to be found, coming from the enum Attribute (e.g. id = "")</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By XPath Locator</returns>
        By XPathMaker(TestHelper.HtmlTag htmlTag, TestHelper.Attribute attribute, string attributeValue);
        /// <summary>
        /// XPath Maker: Builds the String used to find an Element using the XPath Function and returns By xpath locator
        /// This is an Overload Function to be used in case of composite Tags or Attributes
        /// </summary>
        /// <param name="htmlTag">HTML Tags  used to find the Attribute within, entered as String (e.g. <input> </input>)</param>
        /// <param name="attribute">Attribute to be found, entered as String (e.g. data-bind = "")</param>
        /// <param name="attributeValue">Value of the Attribute that we are searching for (e.g. id = "AttributeValue")</param>
        /// <returns>The By XPath Locator</returns>
        By XPathMaker(string htmlTag, string attribute, string attributeValue);
    }
}
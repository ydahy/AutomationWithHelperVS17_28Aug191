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
using AutomationWithHelperVS17_28Aug191.Properties;
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

// ReSharper disable LocalizableElement

namespace AutomationWithHelperVS17_28Aug191
{
    /// <summary>
    /// CRM Test Helper including many functions to ease writing Selenium WebDriver Test CRM Automation Code
    /// </summary>
    public class CrmHelper : TestHelper, ICrmHelper
    {
        /// <summary>
        /// Enum predefined in the template including most Command Items 
        /// </summary>
        public enum CrmCommands
        {
            NewRecord,
            Save,
            SavePrimary,
            SaveAndClose,
            Pick,
            Delete,
            DeleteMenu,
            Edit,
            Activate,
            Deactivate,
            SaveAndRunRoutingRule,
            RunRoutingRule,
            Assign,
            AssignSelectedRecord,
            AddToQueue,
            Send,
            SendDirectEmail,
            Sharing,
            CaseFlow,
            RunReport,
            DocumentTemplate,
            ExportToExcel,
            Import,
            Charts,
            DesignView,
            SystemView,
            CreateChildcase,
            QueueItemDetails,
            moreCommands
        }

        /// <summary>
        /// CRM Dismiss Popup: Switch to the Frame of the Popup and Close it
        /// </summary>
        /// <param name="iframeId">IFrame ID of the Popup</param>
        public virtual void CrmDismissPopup(string iframeId)
        {
            WaitForPageReadyState();

            SwitchToFrame(iframeId);

            Wait.Until(d => IsElementVisible(By.Id("buttonClose")) || IsElementVisible(By.Id("butBegin")) ||
                            IsElementVisible(By.Id("ok_id")));

            if (IsElementVisible(By.Id("buttonClose")))
                ClickOn(By.Id("buttonClose"), true);
            else if (IsElementVisible(By.Id("butBegin")))
                ClickOn(By.Id("butBegin"), true);
            else if (IsElementVisible(By.Id("ok_id")))
                ClickOn(By.Id("ok_id"), true);
            else
                Assert.Fail("Can't Click on Either button");

            SwitchToDefault();
            WaitFor(By.Id(iframeId), Condition.Invisible);
        }

        /// <summary>
        /// Switch to InlineDialog Iframes and Click on Ok/Close
        /// Dismisses at most 2 Warnings
        /// </summary>
        public virtual void CrmDismissEmailWarnings()
        {
            SwitchToDefault();

            if (IsElementPresent(By.Id("InlineDialog1_Iframe")))
            {
                CrmDismissPopup("InlineDialog1_Iframe");
                CrmDismissPopup("InlineDialog_Iframe");
            }
            else
            {
                if (IsElementPresent(By.Id("InlineDialog_Iframe")))
                    CrmDismissPopup("InlineDialog_Iframe");
            }
        }

        /// <summary>
        /// CRM Pick Request: Goes to items available to work on and Picks the Request using the Request ID then switches to items i am working on
        /// </summary>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        public virtual void CrmPickRequest(string requestId)
        {
            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items available to work on"), Condition.Visible);

            ClickOn(FindElement(By.PartialLinkText("Items available to work on")), true);

            CrmSearchInFindField(requestId);

            Wait.Until(ExpectedConditions.ElementExists(By.ClassName("ms-crm-RowCheckBox")));
            ClickOn(By.ClassName("ms-crm-RowCheckBox"), true);

            Thread.Sleep(2000);
            SwitchToDefault();

            CrmClickOnCommandBarItem(CrmCommands.Pick);

            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items I am working on"), Condition.Visible);
            ClickOn(By.PartialLinkText("Items I am working on"), true);
        }

        /// <summary>
        /// CRM Pick Request: Goes to items available to work on and given queue name and Picks the Request using the Request ID
        /// then switches to items i am working on
        /// </summary>
        /// <param name="queueName">Queue Name to pick the request from in case it is assigned to a specific queue</param>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        public virtual void CrmPickRequest(string queueName, string requestId)
        {
            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items available to work on"), Condition.Visible);
            ClickOn(FindElement(By.PartialLinkText("Items available to work on")), true);

            if (queueName != "")
            {
                SelectByText(By.Id("crmQueueSelector"), queueName);
            }

            CrmSearchInFindField(requestId);

            WaitFor(By.ClassName("ms-crm-RowCheckBox"), Condition.Exist);
            ClickOn(By.ClassName("ms-crm-RowCheckBox"), true);

            Thread.Sleep(2000);
            SwitchToDefault();

            CrmClickOnCommandBarItem(CrmCommands.Pick);

            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items I am working on"), Condition.Visible);
            ClickOn(By.PartialLinkText("Items I am working on"), true);
        }

        /// <summary>
        /// CRM Pick Request: Goes to available items query name provided,
        /// Picks the Request using the Request ID found in the provided index of the table without Search
        /// Then switches to items i am working on
        /// Returns the task title in case it is different than the request ID
        /// </summary>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        /// <param name="taskTitleIndexInGrid">Index of the Task Title in the Grid</param>
        /// <returns>Returns the task title</returns>
        public virtual string CRMPickRequest_WithoutSearch(string requestId, int taskTitleIndexInGrid)
        {
            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);

            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items available to work on"), Condition.Visible);

            ClickOn(FindElement(By.PartialLinkText("Items available to work on")), true);

            var row = FindCertainParentOf(By.LinkText(requestId));
            var cell = row.FindElements(By.TagName("td"))[taskTitleIndexInGrid];
            var link = FindLinkInElement(cell);
            var taskTitle = GetTextOf(link);
            ClickOn(row.FindElements(By.TagName("td"))[0], true);

            Thread.Sleep(2000);
            SwitchToDefault();

            CrmClickOnCommandBarItem(CrmCommands.Pick);

            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText("Items I am working on"), Condition.Visible);
            ClickOn(FindElement(By.PartialLinkText("Items I am working on")), false);

            return taskTitle;
        }

        /// <summary>
        /// CRM Pick Request: Goes to available items query name provided and queue name if provided
        /// Picks the Request using the Request ID found in the provided index of the table without Search
        /// Then opens the provided working items query name
        /// Returns the task title in case it is different than the request ID
        /// </summary>
        /// <param name="availableQueryName">Query Name of Available Items to Work On</param>
        /// <param name="workingQueryName">Query Name of Items Working On</param>
        /// <param name="queueName">Queue Name to pick the request from in case it is assigned to a specific queue</param>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        /// <param name="taskTitleIndexInGrid">Index of the Task Title in the Grid</param>
        /// <returns>Returns the task title</returns>
        public virtual string CRMPickRequest_WithoutSearch(string availableQueryName, string workingQueryName, string queueName, string requestId, int taskTitleIndexInGrid)
        {
            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText(availableQueryName), Condition.Visible);
            ClickOn(FindElement(By.PartialLinkText(availableQueryName)), true);

            if (queueName != "")
            {
                SelectByText(By.Id("crmQueueSelector"), queueName);
            }

            var row = FindCertainParentOf(By.LinkText(requestId));
            var cell = row.FindElements(By.TagName("td"))[taskTitleIndexInGrid];
            var link = FindLinkInElement(cell);
            var taskTitle = GetTextOf(link);
            ClickOn(row.FindElements(By.TagName("td"))[0], true);

            Thread.Sleep(2000);
            SwitchToDefault();

            CrmClickOnCommandBarItem(CrmCommands.Pick);

            SwitchToFrameOfElement(By.Id("crmGrid_SavedNewQuerySelector"));

            WaitFor(By.Id("crmGrid_SavedNewQuerySelector"), Condition.Visible);
            ClickOn(FindElement(By.Id("crmGrid_SavedNewQuerySelector")), true);

            WaitFor(By.PartialLinkText(workingQueryName), Condition.Visible);
            ClickOn(FindElement(By.PartialLinkText(workingQueryName)), false);

            return taskTitle;
        }

        /// <summary>
        /// CRM Open Request: Searches for a request by its ID, Opens it using the link of the requestID provided
        /// and Waits for the loading image that appears inside the request to disappear
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        public virtual void CrmOpenRequest(string requestId)
        {
            CrmSearchInFindField(requestId.Trim());

            WaitFor(By.PartialLinkText(requestId.Trim()), Condition.Visible);
            ClickOn(FindElements(By.PartialLinkText(requestId.Trim()))[0], true);

            CrmWaitForCaseOrRequestToBeLoaded();
        }

        /// <summary>
        /// CRM Open Request: Searches for a request by its ID, Clicks on the Cell that contains the link to the request using the Clickable Cell Index in Grid Parameter
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        /// <param name="clickableCellIndexInGrid">Position of the Clickable Element in the Grid to Navigate to the Request</param>
        public virtual void CrmOpenRequest(string requestId, int clickableCellIndexInGrid)
        {
            CrmSearchInFindField(requestId.Trim());

            SwitchToFrameOfElement(By.Id("crmGrid_divDataArea"));
            var tableDiv = FindElement(By.Id("crmGrid_divDataArea"));
            var table = tableDiv.FindElement(By.TagName("table"));
            var tableTbody = table.FindElement(By.TagName("tbody"));
            var tableTRs = tableTbody.FindElements(By.TagName("tr"));

            if (IsElementPresent(LinkTitleContains("Sort by Created On")) && tableTRs.Count > 1)
            {
                ClickOn(LinkTitleContains("Sort by Created On"), true);
                Thread.Sleep(2500);
                ClickOn(LinkTitleContains("Sort by Created On"), true);
                Thread.Sleep(2500);
            }

            tableDiv = FindElement(By.Id("crmGrid_divDataArea"));
            table = tableDiv.FindElement(By.TagName("table"));
            tableTbody = table.FindElement(By.TagName("tbody"));
            tableTRs = tableTbody.FindElements(By.TagName("tr"));
            var tableFirstTrTDs = tableTRs[0].FindElements(By.TagName("td"));

            ClickOn(FindLinkInElement(tableFirstTrTDs[clickableCellIndexInGrid]), true);

            CrmWaitForCaseOrRequestToBeLoaded();
        }

        /// <summary>
        /// CRM Open Request Without Search: Clicks on the Request ID by link text without search
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        public virtual void CrmOpenRequest_WithoutSearch(string requestId)
        {
            SwitchToFrameOfElement(By.Id("crmGrid_divDataArea"));

            ClickOn(LinkTitleContains(requestId.Trim()), true);

            CrmWaitForCaseOrRequestToBeLoaded();
        }

        /// <summary>
        /// CRM Save Request through Footer Save: Switches to Footer Save's iframe
        /// Then clicks on the Save found in the footer and waits till the saving disappears from footer
        /// </summary>
        public virtual void CrmFooterSaveRequest()
        {
            SwitchToFrameOfElement(By.Id("savefooter_statuscontrol"));
            WaitFor(By.Id("savefooter_statuscontrol"), Condition.Visible);
            ClickOn(By.Id("savefooter_statuscontrol"), false);
            CrmWaitTillPageIsSaved();
        }

        /// <summary>
        /// CRM Save Request through Header Save: Switches to default content, clicks on the save found in the header
        /// And waits till the saving disappears from the footer
        /// </summary>
        public virtual void CrmSaveRequest()
        {
            CrmClickOnCommandBarItem(CrmCommands.Save);
        }

        /// <summary>
        /// CRM Save And Close Request: Switches to default content, clicks on the Save and Close found in the header
        /// And waits till the saving disappears from the footer
        /// </summary>
        public virtual void CrmSaveAndCloseRequest()
        {
            CrmClickOnCommandBarItem(CrmCommands.SaveAndClose);
        }

        /// <summary>
        /// CRM Wait Till Page is Saved: Waits till the saving word disappears from the footer
        /// </summary>
        public virtual void CrmWaitTillPageIsSaved()
        {
            SwitchToFrameOfElement(By.Id("footer_statuscontrol"));

            Wait.Until(d =>
            {
                try
                {
                    var appStatus = GetTextOf(By.Id("titlefooter_statuscontrol"));
                    return appStatus == "Read only" || appStatus == "" || appStatus == "unsaved changes";
                }
                catch (StaleElementReferenceException)
                {
                    var appStatus = GetTextOf(By.Id("titlefooter_statuscontrol"));
                    return appStatus == "Read only" || appStatus == "" || appStatus == "unsaved changes";
                }
            });

            WaitForPageReadyState();
        }

        /// <summary>
        /// CRM Upload Attachments: Uploads attachments to all folders in the request through the sharepoint
        /// </summary>
        /// <param name="docUploaderUsername">Username that has privilege to upload documents to sharepoint</param>
        /// <param name="docUploaderPassword">Password of User that has privilege to upload documents to sharepoint</param>
        public virtual void CrmUploadAttachments(string docUploaderUsername, string docUploaderPassword)
        {
            Thread.Sleep(5000);
            SwitchToDefault();

            ClickOn(FindElement(By.Id("TabNode_tab0Tab")), true);

            WaitFor(By.Id("Node_navDocument"), Condition.Visible);
            ClickOn(FindElement(By.Id("Node_navDocument")), true);

            Thread.Sleep(5000);
            AutoItX.Send(docUploaderUsername + "{TAB}");
            AutoItX.Send(docUploaderPassword + "{TAB}{Enter}");

            Thread.Sleep(5000);


            SwitchToFrameOfElement(By.Id("areaDocumentFrame"));
            SwitchToNestedFrame("areaDocumentFrame");
            SwitchToNestedFrame("gridIframe");

            //Open your document location in SharePoint
            WaitFor(By.Id("openSharepointButton"), Condition.Visible);
            ClickOn(FindElement(By.Id("openSharepointButton")), true);

            Wait.Until(d => Driver.WindowHandles.Count > 1);

            SwitchToWindow(Driver.WindowHandles[Driver.WindowHandles.Count - 1]);

            WaitFor(By.Id("onetidDoclibViewTbl0"), Condition.Visible);
            var attachmentsRows = FindElement(By.Id("onetidDoclibViewTbl0")).FindElements(By.TagName("tr"));

            for (var attachmentsCounter = 1; attachmentsCounter < attachmentsRows.Count; attachmentsCounter++)
            {
                WaitFor(By.Id("onetidDoclibViewTbl0"), Condition.Visible);
                attachmentsRows = FindElement(By.Id("onetidDoclibViewTbl0")).FindElements(By.TagName("tr"));
                var attachmentCells = attachmentsRows[attachmentsCounter].FindElements(By.TagName("td"));
                ClickOn(attachmentCells[1], false);

                WaitFor(By.Id("idHomePageNewDocument-WPQ2"), Condition.Visible);
                ClickOn(By.Id("idHomePageNewDocument-WPQ2"), false);

                WaitFor(Prefix(Attribute.id, "DlgFrame"), Condition.Visible);
                SwitchToFrame("DlgFrame");

                Thread.Sleep(3000);
                var documentAttachmentPath = SetDir("pdf.pdf");
                WaitFor(Suffix(Attribute.id, "InputFile"), Condition.Visible);
                ClickOn(Suffix(Attribute.id, "InputFile"), false);
                Thread.Sleep(5000);
                AutoItX.Send(documentAttachmentPath + "{Enter}");

                WaitFor(Suffix(Attribute.id, "btnOK"), Condition.Visible);
                ClickOn(Suffix(Attribute.id, "btnOK"), false);
                FindElement(By.CssSelector("[id$='btnOK']")).Click();

                Thread.Sleep(3000);
                SwitchToDefault();
                WaitFor(By.Id("onetidDoclibViewTbl0"), Condition.Visible);
                attachmentsRows = FindElement(By.Id("onetidDoclibViewTbl0")).FindElements(By.TagName("tr"));
                attachmentCells = attachmentsRows[attachmentsCounter].FindElements(By.TagName("td"));
                ClickOn(attachmentCells[3], false);
                WaitFor(By.Id("Ribbon.Document-title"), Condition.Visible);
                ClickOn(By.Id("Ribbon.Document-title"), false);

                WaitFor(By.Id("Ribbon.Documents.EditCheckout.CheckOut-Medium"), Condition.Visible);
                ClickOn(By.Id("Ribbon.Documents.EditCheckout.CheckOut-Medium"), false);

                WaitFor(By.XPath("//*[@id=\"DeltaPlaceHolderPageTitleInTitleArea\"]/span/span[1]/a"), Condition.Visible);
                ClickOn(By.XPath("//*[@id=\"DeltaPlaceHolderPageTitleInTitleArea\"]/span/span[1]/a"), false);
            }

            Driver.Close();

            Wait.Until(d => Driver.WindowHandles.Count == 1);
            SwitchToWindow(Driver.WindowHandles[Driver.WindowHandles.Count - 1]);
        }

        /// <summary>
        /// CRM Upload Attachments: Uploads attachments to the specified folder in the request through the CRM Frame
        /// </summary>
        /// <param name="folderName">Name of the Folder where the attachment will be uploaded</param>
        /// <param name="docUploaderUsername">Username that has privilege to upload documents to sharepoint</param>
        /// <param name="docUploaderPassword">Password of User that has privilege to upload documents to sharepoint</param>
        public virtual void CrmUploadAttachments(string folderName, string docUploaderUsername, string docUploaderPassword)
        {
            Thread.Sleep(5000);
            SwitchToDefault();

            ClickOn(By.Id("TabNode_tab0Tab"), true);

            WaitFor(By.Id("Node_navDocument"), Condition.Visible);
            ClickOn(By.Id("Node_navDocument"), true);

            Thread.Sleep(5000);
            AutoItX.Send(docUploaderUsername + "{TAB}");
            AutoItX.Send(docUploaderPassword + "{TAB}{Enter}");

            Thread.Sleep(5000);

            SwitchToFrameOfElement(By.Id("areaDocumentFrame"));
            SwitchToNestedFrame("areaDocumentFrame");
            SwitchToNestedFrame("gridIframe");

            var divDataArea = FindElement(By.Id("divDataArea"));
            var tableOfDivDataArea = divDataArea.FindElement(By.TagName("table"));
            var tbodyOfTable = tableOfDivDataArea.FindElement(By.TagName("tbody"));
            IList<IWebElement> trsOfTbody = tbodyOfTable.FindElements(By.TagName("tr"));
            var trOfFolder = trsOfTbody[0];
            foreach (var folderElement in trsOfTbody)
            {
                var trsFolderName = folderElement.GetAttribute("foldername");
                if (!folderName.Equals(trsFolderName)) continue;
                trOfFolder = folderElement;
                break;
            }
            Actions().DoubleClick(trOfFolder).Perform();
            Thread.Sleep(2000);
            var liOfUploadButton = FindElement(By.Id("uploadBtn"));
            var spanOfUploadButton = liOfUploadButton.FindElement(By.TagName("span"));
            var linkOfUploadButton = spanOfUploadButton.FindElement(By.TagName("a"));
            linkOfUploadButton.Click();
            Thread.Sleep(3000);
            SwitchToWindow(Driver.WindowHandles[1]);
            WaitTillPageLoad(By.Id("spPageFrame"));
            SwitchToFrame("spPageFrame");
            Thread.Sleep(5000);

            try
            {
                FindElement(XPathMaker(HtmlTag.input, Attribute.id, "InputFile")).SendKeys(SetDir("pdf.pdf"));
                ClickOn(FindElement(XPathMaker(HtmlTag.input, Attribute.id, "btnOK")), true);
            }
            catch (NoSuchElementException)
            {
                Driver.SwitchTo().Window(Driver.WindowHandles[0]);
                FindElement(XPathMaker(HtmlTag.input, Attribute.id, "InputFile")).SendKeys(SetDir("pdf.pdf"));
                ClickOn(FindElement(XPathMaker(HtmlTag.input, Attribute.id, "btnOK")),
                    true);
            }
            Wait.Until(d => Driver.WindowHandles.Count < 2);
            Thread.Sleep(3000);
            SwitchToWindow(Driver.WindowHandles[0]);
            SwitchToDefault();
        }

        /// <summary>
        /// CRM Send Keys: Clicks on the element using its ID and enters the text to its text box
        /// Usage Example: CRMSendKeys("id", "text");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the text box to appear</param>
        /// <param name="text">Text that will be entered in the text box</param>
        public virtual void CrmSendKeys(string id, string text)
        {
            var elementLocator = By.Id(id);
            WaitFor(elementLocator, Condition.Visible);
            ClickOn(elementLocator, false);

            elementLocator = By.Id(id + "_i");
            WaitFor(elementLocator, Condition.Visible);
            Thread.Sleep(2000);
            SendKeys(elementLocator, text);
            SendKeys(elementLocator, Keys.Tab);
        }

        /// <summary>
        /// CRM Drop Down Select By Text: Clicks on the element using its ID and selects the option by the text sent to the function
        /// Usage Example: CRMDropDownSelectByText("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select list to appear</param>
        /// <param name="selectedText">Text of the option that will to be selected</param>
        public virtual void CrmDropDownSelectByText(string id, string selectedText)
        {
            var elementLocator = By.Id(id);
            WaitFor(elementLocator, Condition.Visible);
            ClickOn(elementLocator, false);

            elementLocator = By.Id(id + "_i");
            WaitFor(elementLocator, Condition.Visible);
            SelectByText(elementLocator, selectedText);
            //_driver.FindElement(elementLocator).SendKeys(Keys.Tab);
        }

        /// <summary>
        /// CRM Select Text: Clicks on the element using its ID and selects the link by the text sent to the function
        /// Usage Example: CRMSelectText("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        /// <param name="selectedText">Text of the link that will to be selected</param>
        public virtual void CrmSelectText(string id, string selectedText)
        {
            var elementLocator = By.Id(id);
            ClickOn(elementLocator, false);
            WaitFor(By.PartialLinkText(selectedText), Condition.Visible);
            ClickOn(By.PartialLinkText(selectedText), false);
        }

        /// <summary>
        /// CRM Select Text From Lookup: Clicks on the element using its ID and selects the link by the text sent to the function
        /// Usage Example: CrmSelectTextFromLookup("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        /// <param name="selectedText">Text of the link that will to be selected</param>
        public virtual void CrmSelectTextFromLookup(string id, string selectedText)
        {
            var elementLocator = By.Id(id);
            ClickOn(elementLocator, true);
            elementLocator = By.Id(id + "_ledit");
            //Actions().MoveToElement(_driver.FindElement(elementLocator)).SendKeys(selectedText + "\r\n");
            SendKeys(FindElement(elementLocator), selectedText + "\r\n");
            selectedText = selectedText.Replace("*", "");
            WaitFor(By.PartialLinkText(selectedText), Condition.Visible);
            ClickOn(By.PartialLinkText(selectedText), true);
        }

        /// <summary>
        /// CRM Select Random Item From Lookup: Clicks on the element using its ID and selects the first link
        /// Usage Example: CrmSelectRandomItemFromLookup("id");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        public virtual void CrmSelectRandomItemFromLookup(string id)
        {
            var elementLocator = By.Id(id);
            ClickOn(elementLocator, true);
            elementLocator = By.Id(id + "_ledit");
            FindElement(elementLocator).SendKeys("\r\n");
            elementLocator = By.Id(id + "_IMenu");
            WaitFor(elementLocator, Condition.Visible);
            var menu = FindElement(elementLocator);
            var linksInMenu = menu.FindElements(By.TagName("a"));
            foreach (var linkElement in linksInMenu)
            {
                if (!linkElement.Displayed) continue;
                ClickOn(linkElement, true);
                break;
            }
        }

        /// <summary>
        /// CRM Set Date: Clicks on the element using its ID and enters the given date
        /// Usage Example: CrmSetDate("id", DateTime.Now);
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the date field to appear</param>
        /// <param name="date">Date that will be Selected</param>
        public virtual void CrmSetDate(string id, DateTime date)
        {
            var elementLocator = By.Id(id);
            ClickOn(elementLocator, true);
            elementLocator = By.Id(id + "_iDateInput");
            var dateFormated = string.Format("{0:M/d/yyyy}", date);
            SendKeys(elementLocator, dateFormated);
            SendKeys(elementLocator, Keys.Tab);
        }

        /// <summary>
        /// Wait for Stage to be Active: Takes the Stage ID and waits till the Stage is Active
        /// </summary>
        /// <param name="stageId">ID of the stage that will _wait for</param>
        public virtual void CrmWaitForStageToBeActive(string stageId)
        {
            SwitchToFrameOfElement(By.Id(stageId));

            var seconds = 0;
            var stageActive = false;

            do
            {
                try
                {
                    stageActive = FindElement(By.Id(stageId)).GetAttribute("class").Contains("activeStage");
                    seconds++;
                }
                catch (Exception)
                {
                    seconds++;
                }
            } while (!stageActive && seconds < 30);

            var counter = 0;

            while (!CrmCheckIfStageIsActive(stageId) && counter < 5)
            {
                RefreshPage();

                CrmWaitForCaseOrRequestToBeLoaded();

                SwitchToDefault();

                CrmDismissEmailWarnings();
                counter++;
            }
        }

        /// <summary>
        /// Check if Stage is Active: Takes the Stage ID and checks whether the Stage is Active or Not
        /// </summary>
        /// <param name="stageId">ID of the stage that will be checked</param>
        /// <returns>true or false indicating the activity state of the stage</returns>
        public virtual bool CrmCheckIfStageIsActive(string stageId)
        {
            SwitchToFrameOfElement(By.Id(stageId));
            WaitFor(By.Id(stageId), Condition.Exist);
            var stageActive = FindElement(By.Id(stageId)).GetAttribute("class").Contains("activeStage");
            return stageActive;
        }

        /// <summary>
        /// Next Stage: Clicks on the Next Stage button in CRM Business Workflow and waits till the page is saved
        /// </summary>
        public virtual void CrmNextStage()
        {
            SwitchToFrameOfElement(By.Id("stageAdvanceActionContainer"));

            Wait.Until(d => IsElementVisible(By.Id("stageAdvanceActionContainer")) ||
                            IsElementVisible(By.Id("stageNavigateActionContainer")));

            if (!FindElement(By.Id("stageAdvanceActionContainer")).GetAttribute("class").Contains("hidden"))
            {
                ClickOn(By.Id("stageAdvanceActionContainer"), true);
            }
            else
            {
                if (!FindElement(By.Id("stageNavigateActionContainer")).GetAttribute("class").Contains("hidden"))
                {
                    ClickOn(FindElement(By.Id("stageNavigateActionContainer")), true);
                }
            }

            CrmWaitTillPageIsSaved();
        }

        /// <summary>
        /// CRM search in find field: Search for given text after switching to find field's iframe
        /// </summary>
        /// <param name="searchText">Text that will be searched for</param>
        public virtual void CrmSearchInFindField(string searchText)
        {
            var searchField = Suffix(Attribute.id, "_findCriteria");
            SwitchToFrameOfElement(searchField);

            ClickOn(searchField, true);
            FindElement(searchField).Clear();
            FindElement(searchField).SendKeys(searchText);
            FindElement(searchField).SendKeys("\r\n");
            Thread.Sleep(1000);

            SwitchToDefault();
            ExecuteScript("arguments[0].style.display='none';", FindElement(By.Id("navBarOverlay")));

            SwitchToFrameOfElement(searchField);

            if (!IsElementVisible(Full(Attribute.Class, "ms-crm-Floating-Div"))) return;
            ExecuteScript("arguments[0].style.display='none';",
                FindElement(Full(Attribute.Class, "ms-crm-Floating-Div")));
            ClickOn(By.Id("crmGrid_SavedNewQuerySelector"), true);
        }

        /// <summary>
        /// Click on Multiple Links to reach a listing page
        /// Usage Example: CrmGoToList(new By[] {By.Id("navTabId"), By.Id("Project"), By.Id("Services")});
        /// </summary>
        /// <param name="navigationLocators">List of Locators of all the elements that will be clicked to reach the list</param>
        public virtual void CrmGoToList(By[] navigationLocators)
        {
            SwitchToDefault();

            foreach (var navigationItem in navigationLocators)
            {
                WaitFor(navigationItem, Condition.Exist);
                ScrollTo(navigationItem);
                ClickOn(navigationItem, true);
                WaitForPageReadyState();
            }

            var loadingLocator = By.Id("loading");
            SwitchToFrameOfElement(loadingLocator);
            WaitFor(loadingLocator, Condition.Invisible);

            var containerLoader = By.Id("containerLoadingProgress");
            SwitchToFrameOfElement(containerLoader);
            WaitFor(containerLoader, Condition.Invisible);
            Thread.Sleep(2000);
        }

        /// <summary>
        /// CRM Check Status: Switches to Frame of the Request Status Locator and Checks that the Visible Status is same as the provided Expected Request Status
        /// </summary>
        /// <param name="requestStatusLocator">By Locator of the Field Containing the Request Status</param>
        /// <param name="expectedRequestStatus">Expected Status of the Request</param>
        public virtual bool CrmCheckStatus(By requestStatusLocator, string expectedRequestStatus)
        {
            Thread.Sleep(4000);
            SwitchToFrameOfElement(requestStatusLocator);
            var visibleStatus = GetTextOf(requestStatusLocator);
            var statusIsCorrect = visibleStatus.ToLower().Trim().Contains(expectedRequestStatus.ToLower().Trim());

            LogInfo("",
                (statusIsCorrect ? "Correct Status" : "Wrong Status") + " : Visible Status is " + visibleStatus, false,
                true);

            return statusIsCorrect;
        }

        /// <summary>
        /// CRM Check Email and SMS Title: Goes to Activities tab
        /// and searches for the given Email and SMS subject
        /// and checks if the same number as expected count is appearing in the Activities List
        /// </summary>
        /// <param name="expectedActivityTitle">Expected Subject of the Email and/or SMS (can be partial by adding Asterisk "*")</param>
        /// <param name="expectedCountOfActivities">Number of Activities that should appear based on the search of the title</param>
        /// <returns>True if the count is correct</returns>
        public virtual bool CrmCheckActivities(string expectedActivityTitle, int expectedCountOfActivities)
        {
            CrmSearchForActivities(expectedActivityTitle);

            var activitiesList = FindElements(By.PartialLinkText(expectedActivityTitle.Replace("*", "")));

            var countIsCorrect = activitiesList.Count >= expectedCountOfActivities;
            LogInfo("", "Email and/or SMS with Title \"" + expectedActivityTitle.Replace("*", "") + "\" are " + (countIsCorrect ? "Found in Activities and their count is " + activitiesList.Count : "Missing"), false, true);

            return countIsCorrect;
        }

        /// <summary>
        /// CRM Search for Activities: Go to Activities List in the Opened Page
        /// And Search for the Given Title
        /// </summary>
        /// <param name="expectedActivityTitle">Activity Title that should be found</param>
        public virtual void CrmSearchForActivities(string expectedActivityTitle)
        {
            SwitchToDefault();

            ClickOn(FindElement(By.Id("TabNode_tab0Tab")).FindElement(By.TagName("a")), true);

            WaitFor(By.Id("Node_navActivities"), Condition.Visible);
            ClickOn(FindElement(By.Id("Node_navActivities")), true);

            SwitchToFrameOfElement(By.Id("areaActivitiesFrame"));
            SwitchToNestedFrame("areaActivitiesFrame");

            if (IsElementVisible(Suffix(Attribute.id, "_datefilter")))
            {
                SelectByValue(Suffix(Attribute.id, "_datefilter"), "All");

                Thread.Sleep(2000);
            }

            ClickOn(Suffix(Attribute.id, "_findCriteria"), false);
            SendKeys(Suffix(Attribute.id, "_findCriteria"), expectedActivityTitle + "\r\n");

            Thread.Sleep(6000);
        }

        /// <summary>
        /// CRM Expand Section: Finds the Section's Arrow using the Locator sent to the function and Expands it if it is collapsed
        /// </summary>
        /// <param name="sectionLocator">By Locator of the Section that should be Expanded</param>
        public virtual void CrmExpandSection(By sectionLocator)
        {
            SwitchToFrameOfElement(sectionLocator);
            var sectionImage = FindElement(sectionLocator).FindElement(By.TagName("img"));
            if (sectionImage.GetAttribute("alt").Contains("Expand this tab"))
            {
                ClickOn(sectionImage, true);
            }
        }

        /// <summary>
        /// CRM Wait For Case Or Request To Be Loaded: Waits until the Case or the Request is Loaded
        /// </summary>
        public virtual void CrmWaitForCaseOrRequestToBeLoaded()
        {
            WaitForPageReadyState();

            var loadingLocator = By.Id("loading");

            SwitchToFrameOfElement(loadingLocator);

            WaitFor(loadingLocator, Condition.Invisible);

            SwitchToDefault();

            WaitFor(By.Id("TabNode_tab0Tab-main"), Condition.Visible);

            var containerLoader = By.Id("containerLoadingProgress");

            SwitchToFrameOfElement(containerLoader);

            WaitFor(containerLoader, Condition.Invisible);

            SwitchToDefault();

            WaitForPageReadyState();
        }

        /// <summary>
        /// CRM Click On Command Bar Item: Clicks on the Given Command Bar Item (e.g. New Record, Pick, Save, Save and Close, etc..)
        /// </summary>
        /// <param name="command">Command Item that will be Clicked</param>
        public virtual void CrmClickOnCommandBarItem(CrmCommands command)
        {
            SwitchToDefault();
            var liLocator = CrmCommandBarLocator(command);
            WaitFor(liLocator, Condition.Visible);
            ClickOn(liLocator, false);

            switch (command)
            {
                case CrmCommands.NewRecord:
                    CrmWaitForCaseOrRequestToBeLoaded();
                    break;
                case CrmCommands.Pick:
                case CrmCommands.DeleteMenu:
                    CrmDismissPopup("InlineDialog_Iframe");
                    break;
                case CrmCommands.Save:
                case CrmCommands.SaveAndClose:
                case CrmCommands.SavePrimary:
                    CrmWaitTillPageIsSaved();
                    break;
            }

            WaitForPageReadyState();
        }

        /// <summary>
        /// Gets the Command Bar Item locator for the given command
        /// </summary>
        /// <param name="command">Command Item that will be Located</param>
        /// <returns>Returns Command Bar Item Locator</returns>
        public virtual By CrmCommandBarLocator(CrmCommands command)
        {
            return Suffix(Attribute.id, "." + command);
        }

        /// <summary>
        /// CRM Open Hidden and Locked Fields: Opens all the fields in the page in CRM
        /// </summary>
        public virtual void CrmOpenHiddenAndLockedFields()
        {
            ExecuteScript(
                "javascript:var form=$(\"iframe\").filter(function(){return $(this).css(\"visibility\")==\"visible\"})[0].contentWindow;try{form.Mscrm.InlineEditDataService.get_dataService().validateAndFireSaveEvents=function(){return new Mscrm.SaveResponse(5,\"\")}}catch(e){}var attrs=form.Xrm.Page.data.entity.attributes.get();for(var i in attrs){attrs[i].setRequiredLevel(\"none\")}var contrs=form.Xrm.Page.ui.controls.get();for(var i in contrs){try{contrs[i].setVisible(true);contrs[i].setDisabled(false);contrs[i].clearNotification()}catch(e){}}var tabs=form.Xrm.Page.ui.tabs.get();for(var i in tabs){tabs[i].setVisible(true);tabs[i].setDisplayState(\"expanded\");var sects=tabs[i].sections.get();for(var i in sects){sects[i].setVisible(true)}}");
        }
    }
}
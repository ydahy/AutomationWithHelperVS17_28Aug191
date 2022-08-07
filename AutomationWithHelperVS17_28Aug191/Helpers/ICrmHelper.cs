using System;
using OpenQA.Selenium;

namespace AutomationWithHelperVS17_28Aug191
{
    /// <summary>
    /// CRM Helper Interface
    /// </summary>
    public interface ICrmHelper
    {
        /// <summary>
        /// CRM Check Email and SMS Title: Goes to Activities tab
        /// and searches for the given Email and SMS subject
        /// and checks if the same number as expected count is appearing in the Activities List
        /// </summary>
        /// <param name="expectedActivityTitle">Expected Subject of the Email and/or SMS (can be partial by adding Asterisk "*")</param>
        /// <param name="expectedCountOfActivities">Number of Activities that should appear based on the search of the title</param>
        /// <returns>True if the count is correct</returns>
        bool CrmCheckActivities(string expectedActivityTitle, int expectedCountOfActivities);
        /// <summary>
        /// Check if Stage is Active: Takes the Stage ID and checks whether the Stage is Active or Not
        /// </summary>
        /// <param name="stageId">ID of the stage that will be checked</param>
        /// <returns>true or false indicating the activity state of the stage</returns>
        bool CrmCheckIfStageIsActive(string stageId);
        /// <summary>
        /// CRM Check Status: Switches to Frame of the Request Status Locator and Checks that the Visible Status is same as the provided Expected Request Status
        /// </summary>
        /// <param name="requestStatusLocator">By Locator of the Field Containing the Request Status</param>
        /// <param name="expectedRequestStatus">Expected Status of the Request</param>
        bool CrmCheckStatus(By requestStatusLocator, string expectedRequestStatus);
        /// <summary>
        /// CRM Click On Command Bar Item: Clicks on the Given Command Bar Item (e.g. New Record, Pick, Save, Save and Close, etc..)
        /// </summary>
        /// <param name="command">Command Item that will be Clicked</param>
        void CrmClickOnCommandBarItem(CrmHelper.CrmCommands command);
        /// <summary>
        /// Gets the Command Bar Item locator for the given command
        /// </summary>
        /// <param name="command">Command Item that will be Located</param>
        /// <returns>Returns Command Bar Item Locator</returns>
        By CrmCommandBarLocator(CrmHelper.CrmCommands command);
        /// <summary>
        /// Switch to InlineDialog Iframes and Click on Ok/Close
        /// Dismisses at most 2 Warnings
        /// </summary>
        void CrmDismissEmailWarnings();
        /// <summary>
        /// CRM Dismiss Popup: Switch to the Frame of the Popup and Close it
        /// </summary>
        /// <param name="iframeId">IFrame ID of the Popup</param>
        void CrmDismissPopup(string iframeId);
        /// <summary>
        /// CRM Drop Down Select By Text: Clicks on the element using its ID and selects the option by the text sent to the function
        /// Usage Example: CRMDropDownSelectByText("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select list to appear</param>
        /// <param name="selectedText">Text of the option that will to be selected</param>
        void CrmDropDownSelectByText(string id, string selectedText);
        /// <summary>
        /// CRM Expand Section: Finds the Section's Arrow using the Locator sent to the function and Expands it if it is collapsed
        /// </summary>
        /// <param name="sectionLocator">By Locator of the Section that should be Expanded</param>
        void CrmExpandSection(By sectionLocator);
        /// <summary>
        /// CRM Save Request through Footer Save: Switches to Footer Save's iframe
        /// Then clicks on the Save found in the footer and waits till the saving disappears from footer
        /// </summary>
        void CrmFooterSaveRequest();
        /// <summary>
        /// Click on Multiple Links to reach a listing page
        /// Usage Example: CrmGoToList(new By[] {By.Id("navTabId"), By.Id("Project"), By.Id("Services")});
        /// </summary>
        /// <param name="navigationLocators">List of Locators of all the elements that will be clicked to reach the list</param>
        void CrmGoToList(By[] navigationLocators);
        /// <summary>
        /// Next Stage: Clicks on the Next Stage button in CRM Business Workflow and waits till the page is saved
        /// </summary>
        void CrmNextStage();
        /// <summary>
        /// CRM Open Hidden and Locked Fields: Opens all the fields in the page in CRM
        /// </summary>
        void CrmOpenHiddenAndLockedFields();
        /// <summary>
        /// CRM Open Request: Searches for a request by its ID, Opens it using the link of the requestID provided
        /// and Waits for the loading image that appears inside the request to disappear
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        void CrmOpenRequest(string requestId);
        /// <summary>
        /// CRM Open Request: Searches for a request by its ID, Clicks on the Cell that contains the link to the request using the Clickable Cell Index in Grid Parameter
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        /// <param name="clickableCellIndexInGrid">Position of the Clickable Element in the Grid to Navigate to the Request</param>
        void CrmOpenRequest(string requestId, int clickableCellIndexInGrid);
        /// <summary>
        /// CRM Open Request Without Search: Clicks on the Request ID by link text without search
        /// </summary>
        /// <param name="requestId">ID of the Request that will be opened</param>
        void CrmOpenRequest_WithoutSearch(string requestId);
        /// <summary>
        /// CRM Pick Request: Goes to items available to work on and Picks the Request using the Request ID then switches to items i am working on
        /// </summary>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        void CrmPickRequest(string requestId);
        /// <summary>
        /// CRM Pick Request: Goes to items available to work on and given queue name and Picks the Request using the Request ID
        /// then switches to items i am working on
        /// </summary>
        /// <param name="queueName">Queue Name to pick the request from in case it is assigned to a specific queue</param>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        void CrmPickRequest(string queueName, string requestId);
        /// <summary>
        /// CRM Pick Request: Goes to available items query name provided,
        /// Picks the Request using the Request ID found in the provided index of the table without Search
        /// Then switches to items i am working on
        /// Returns the task title in case it is different than the request ID
        /// </summary>
        /// <param name="requestId">ID of the Request that will be picked and opened</param>
        /// <param name="taskTitleIndexInGrid">Index of the Task Title in the Grid</param>
        /// <returns>Returns the task title</returns>
        string CRMPickRequest_WithoutSearch(string requestId, int taskTitleIndexInGrid);
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
        string CRMPickRequest_WithoutSearch(string availableQueryName, string workingQueryName, string queueName, string requestId, int taskTitleIndexInGrid);
        /// <summary>
        /// CRM Save And Close Request: Switches to default content, clicks on the Save and Close found in the header
        /// And waits till the saving disappears from the footer
        /// </summary>
        void CrmSaveAndCloseRequest();
        /// <summary>
        /// CRM Save Request through Header Save: Switches to default content, clicks on the save found in the header
        /// And waits till the saving disappears from the footer
        /// </summary>
        void CrmSaveRequest();
        /// <summary>
        /// CRM Search for Activities: Go to Activities List in the Opened Page
        /// And Search for the Given Title
        /// </summary>
        /// <param name="expectedActivityTitle">Activity Title that should be found</param>
        void CrmSearchForActivities(string expectedActivityTitle);
        /// <summary>
        /// CRM search in find field: Search for given text after switching to find field's iframe
        /// </summary>
        /// <param name="searchText">Text that will be searched for</param>
        void CrmSearchInFindField(string searchText);
        /// <summary>
        /// CRM Select Random Item From Lookup: Clicks on the element using its ID and selects the first link
        /// Usage Example: CrmSelectRandomItemFromLookup("id");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        void CrmSelectRandomItemFromLookup(string id);
        /// <summary>
        /// CRM Select Text: Clicks on the element using its ID and selects the link by the text sent to the function
        /// Usage Example: CRMSelectText("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        /// <param name="selectedText">Text of the link that will to be selected</param>
        void CrmSelectText(string id, string selectedText);
        /// <summary>
        /// CRM Select Text From Lookup: Clicks on the element using its ID and selects the link by the text sent to the function
        /// Usage Example: CrmSelectTextFromLookup("id", "selectedText");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the select to appear</param>
        /// <param name="selectedText">Text of the link that will to be selected</param>
        void CrmSelectTextFromLookup(string id, string selectedText);
        /// <summary>
        /// CRM Send Keys: Clicks on the element using its ID and enters the text to its text box
        /// Usage Example: CRMSendKeys("id", "text");
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the text box to appear</param>
        /// <param name="text">Text that will be entered in the text box</param>
        void CrmSendKeys(string id, string text);
        /// <summary>
        /// CRM Set Date: Clicks on the element using its ID and enters the given date
        /// Usage Example: CrmSetDate("id", DateTime.Now);
        /// </summary>
        /// <param name="id">ID of the element that will be clicked for the date field to appear</param>
        /// <param name="date">Date that will be Selected</param>
        void CrmSetDate(string id, DateTime date);
        /// <summary>
        /// CRM Upload Attachments: Uploads attachments to all folders in the request through the sharepoint
        /// </summary>
        /// <param name="docUploaderUsername">Username that has privilege to upload documents to sharepoint</param>
        /// <param name="docUploaderPassword">Password of User that has privilege to upload documents to sharepoint</param>
        void CrmUploadAttachments(string docUploaderUsername, string docUploaderPassword);
        /// <summary>
        /// CRM Upload Attachments: Uploads attachments to the specified folder in the request through the CRM Frame
        /// </summary>
        /// <param name="folderName">Name of the Folder where the attachment will be uploaded</param>
        /// <param name="docUploaderUsername">Username that has privilege to upload documents to sharepoint</param>
        /// <param name="docUploaderPassword">Password of User that has privilege to upload documents to sharepoint</param>
        void CrmUploadAttachments(string folderName, string docUploaderUsername, string docUploaderPassword);
        /// <summary>
        /// CRM Wait For Case Or Request To Be Loaded: Waits until the Case or the Request is Loaded
        /// </summary>
        void CrmWaitForCaseOrRequestToBeLoaded();
        /// <summary>
        /// Wait for Stage to be Active: Takes the Stage ID and waits till the Stage is Active
        /// </summary>
        /// <param name="stageId">ID of the stage that will _wait for</param>
        void CrmWaitForStageToBeActive(string stageId);
        /// <summary>
        /// CRM Wait Till Page is Saved: Waits till the saving word disappears from the footer
        /// </summary>
        void CrmWaitTillPageIsSaved();
    }
}